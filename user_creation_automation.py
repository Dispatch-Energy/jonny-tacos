"""
User creation automation for QuickBase onboarding tickets.

This module handles:
- Parsing onboarding details from the ticket description
- Creating or resetting a Microsoft 365 user account
- Assigning a standard license
- Granting access to a SharePoint-backed group
- Adding the new user to a QuickBase app as a Participant
- Queueing a due-date email with credentials for later delivery
"""

import asyncio
import json
import logging
import os
import re
import secrets
import string
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, time, timezone
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import requests

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover - Python 3.11 has zoneinfo
    ZoneInfo = None

try:
    from azure.storage.blob import BlobServiceClient
except ImportError:  # pragma: no cover - optional in tests
    BlobServiceClient = None

from m365_provisioning import M365GraphClient
from quickbase_manager import QuickBaseManager


USER_CREATION_CATEGORY = "User Creation"
DEFAULT_QB_APP_ID = "btq7amgsn"
DEFAULT_QB_ROLE_NAME = "Participant"
DEFAULT_EMAIL_TIMEZONE = "America/New_York"
DEFAULT_USAGE_LOCATION = "US"
DEFAULT_OPENAI_ROLE = "reader"
DEFAULT_ANTHROPIC_ROLE = "user"


def is_user_creation_category(category: str) -> bool:
    """Return True when the ticket category should trigger onboarding automation."""
    return (category or "").strip().lower() == USER_CREATION_CATEGORY.lower()


def slugify_name(value: str) -> str:
    """Normalize a name component into an ASCII identifier."""
    normalized = unicodedata.normalize("NFKD", value or "")
    ascii_value = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z]", "", ascii_value.lower())


def build_username_local_part(first_name: str, last_name: str) -> str:
    """Build the M365 username as first initial + full last name."""
    first_slug = slugify_name(first_name)
    last_slug = slugify_name(last_name)
    if not first_slug or not last_slug:
        raise ValueError("First and last name are required to build the username")
    return f"{first_slug[0]}{last_slug}"


def normalize_display_name(value: str) -> str:
    """Normalize display names for safe equality checks."""
    normalized = unicodedata.normalize("NFKD", value or "")
    ascii_value = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z]", "", ascii_value.lower())


def extract_personal_email(description: str) -> Optional[str]:
    """Extract a personal email address from the ticket description."""
    text = (description or "").strip()
    if not text:
        return None

    compact_match = re.search(
        r"[A-Za-z][A-Za-z' -]*\s*,\s*[A-Za-z][A-Za-z' -]*\s*;\s*"
        r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
        text,
        re.IGNORECASE,
    )
    if compact_match:
        return compact_match.group(1).strip().lower()

    labeled_match = re.search(
        r"email\s*[:\-]\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
        text,
        re.IGNORECASE,
    )
    if labeled_match:
        return labeled_match.group(1).strip().lower()

    generic_match = re.search(
        r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
        text,
        re.IGNORECASE,
    )
    if generic_match:
        return generic_match.group(1).strip().lower()

    return None


def extract_user_creation_details(description: str) -> Tuple[str, str, Optional[str]]:
    """Extract first name, last name, and personal email from the description."""
    text = (description or "").strip()
    if not text:
        raise ValueError("Ticket description is empty")

    compact_match = re.search(
        r"([A-Za-z][A-Za-z' -]*)\s*,\s*([A-Za-z][A-Za-z' -]*)\s*;\s*"
        r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
        text,
        re.IGNORECASE,
    )
    if compact_match:
        first_name = compact_match.group(1).strip().split()[0]
        last_name = compact_match.group(2).strip().split()[-1]
        personal_email = compact_match.group(3).strip().lower()
        return first_name, last_name, personal_email

    first_name, last_name = extract_first_last_name(text)
    personal_email = extract_personal_email(text)
    return first_name, last_name, personal_email


def get_user_creation_source_text(ticket_data: Dict[str, Any]) -> str:
    """Pick the best available source field for structured user-creation details."""
    description = (ticket_data.get("description") or "").strip()
    subject = (ticket_data.get("subject") or "").strip()

    description_has_email = bool(extract_personal_email(description))
    subject_has_email = bool(extract_personal_email(subject))

    if description_has_email:
        return description
    if subject_has_email:
        return subject
    if description:
        return description
    return subject


def extract_first_last_name(description: str) -> Tuple[str, str]:
    """Extract first and last name from a QuickBase ticket description."""
    text = (description or "").strip()
    if not text:
        raise ValueError("Ticket description is empty")

    first_match = re.search(r"first\s*name\s*[:\-]\s*([A-Za-z][A-Za-z' -]*)", text, re.IGNORECASE)
    last_match = re.search(r"last\s*name\s*[:\-]\s*([A-Za-z][A-Za-z' -]*)", text, re.IGNORECASE)
    if first_match and last_match:
        return first_match.group(1).strip().split()[0], last_match.group(1).strip().split()[-1]

    full_name_patterns = [
        r'on\s*boarding\s*for\s*[:\-"]+\s*([A-Za-z][A-Za-z\'-]+(?:\s+[A-Za-z][A-Za-z\'-]+)+)',
        r'new\s+user\s*[:\-]\s*([A-Za-z][A-Za-z\'-]+(?:\s+[A-Za-z][A-Za-z\'-]+)+)',
        r'name\s*[:\-]\s*([A-Za-z][A-Za-z\'-]+(?:\s+[A-Za-z][A-Za-z\'-]+)+)',
        r'employee\s*[:\-]\s*([A-Za-z][A-Za-z\'-]+(?:\s+[A-Za-z][A-Za-z\'-]+)+)',
    ]
    for pattern in full_name_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if not match:
            continue
        tokens = [token for token in match.group(1).replace('"', "").split() if token]
        if len(tokens) >= 2:
            return tokens[0], tokens[-1]

    camel_match = re.search(r"\b([A-Z][a-z]+)([A-Z][a-z]+)\b", text)
    if camel_match:
        return camel_match.group(1), camel_match.group(2)

    raise ValueError(
        "Could not extract first and last name from the ticket description. "
        "Include either 'First Name'/'Last Name' or a full name like 'On boarding for: Jane Smith'."
    )


def parse_due_date(value: str) -> date:
    """Parse a QuickBase due-date string into a date."""
    raw = (value or "").strip()
    if not raw:
        raise ValueError("Ticket due date is required for scheduled onboarding email")

    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).date()
    except ValueError:
        pass

    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Unsupported due date format: {value}") from exc


def calculate_due_date_send_at(due_date_value: str, timezone_name: str = DEFAULT_EMAIL_TIMEZONE) -> datetime:
    """Calculate the UTC timestamp for 5:00 AM on the ticket due date."""
    due = parse_due_date(due_date_value)
    if ZoneInfo is None:
        raise ValueError("ZoneInfo is not available in this Python runtime")

    tz = ZoneInfo(timezone_name)
    local_dt = datetime.combine(due, time(hour=5, minute=0), tzinfo=tz)
    return local_dt.astimezone(timezone.utc)


def generate_temporary_password(length: int = 16) -> str:
    """Generate a strong temporary password for a new user."""
    if length < 12:
        raise ValueError("Password length must be at least 12 characters")

    alphabet = string.ascii_letters + string.digits
    symbols = "!@#$%^*-_+"

    required = [
        secrets.choice(string.ascii_lowercase),
        secrets.choice(string.ascii_uppercase),
        secrets.choice(string.digits),
        secrets.choice(symbols),
    ]
    remaining = [
        secrets.choice(alphabet + symbols)
        for _ in range(length - len(required))
    ]
    password_chars = required + remaining
    secrets.SystemRandom().shuffle(password_chars)
    return "".join(password_chars)


def build_onboarding_email_subject(display_name: str) -> str:
    """Build the scheduled email subject."""
    return f"New user account details for {display_name}"


def build_onboarding_email_body(
    display_name: str,
    username: str,
    temporary_password: str,
    ticket_number: str,
) -> str:
    """Build the scheduled email body."""
    return (
        f"New user provisioning is complete for {display_name}.\n\n"
        f"Username: {username}\n"
        f"Temporary Password: {temporary_password}\n\n"
        "The user will be prompted to change this password at first sign-in.\n\n"
        f"Ticket: {ticket_number}\n"
        "This message was sent automatically by the IT support bot."
    )


def build_initial_ticket_resolution(
    display_name: str,
    username: str,
    recipient_email: str,
    send_at_utc: datetime,
) -> str:
    """Resolution note written after provisioning and queueing the email."""
    return (
        "[USER_CREATION_AUTOMATION]\n"
        f"Provisioned Microsoft 365 account for {display_name}.\n"
        f"Username: {username}\n"
        f"Credential email queued for {recipient_email} at {send_at_utc.isoformat()}."
    )


def build_completion_ticket_resolution(
    display_name: str,
    username: str,
    recipient_email: str,
    sent_at_utc: datetime,
) -> str:
    """Resolution note written once the queued email has been sent."""
    return (
        "[USER_CREATION_COMPLETED]\n"
        f"Provisioning completed for {display_name}.\n"
        f"Username: {username}\n"
        f"Credential email sent to {recipient_email} at {sent_at_utc.isoformat()}."
    )


class ScheduledOnboardingEmailStore:
    """Persist queued onboarding emails in Azure Blob Storage."""

    def __init__(self):
        self.storage_connection_string = (
            os.environ.get("TEAMS_STORAGE_CONNECTION_STRING", "")
            or os.environ.get("AzureWebJobsStorage", "")
            or os.environ.get("AZURE_STORAGE_CONNECTION_STRING", "")
        )
        self.container_name = os.environ.get(
            "USER_CREATION_EMAIL_CONTAINER",
            "scheduled-onboarding-emails",
        )
        self.executor = ThreadPoolExecutor(max_workers=2)
        self._blob_service_client = None
        self._container_ready = False

    def is_configured(self) -> bool:
        return bool(BlobServiceClient and self.storage_connection_string)

    def _blob_name(self, ticket_number: str) -> str:
        safe_ticket = re.sub(r"[^a-z0-9._-]+", "_", (ticket_number or "").lower())
        return f"{safe_ticket}.json"

    async def _get_container_client(self):
        if not self.is_configured():
            return None

        if self._blob_service_client is None:
            try:
                self._blob_service_client = BlobServiceClient.from_connection_string(
                    self.storage_connection_string
                )
            except Exception as exc:
                logging.error(f"Failed to initialize blob client for onboarding emails: {exc}")
                return None

        container_client = self._blob_service_client.get_container_client(self.container_name)

        if not self._container_ready:
            loop = asyncio.get_event_loop()

            def ensure_container():
                try:
                    container_client.create_container()
                except Exception:
                    pass
                return True

            await loop.run_in_executor(self.executor, ensure_container)
            self._container_ready = True

        return container_client

    async def queue_job(self, job: Dict[str, Any]) -> bool:
        """Create or overwrite a queued onboarding email job."""
        container_client = await self._get_container_client()
        if not container_client:
            return False

        blob_name = self._blob_name(job.get("ticket_number", ""))
        loop = asyncio.get_event_loop()

        def upload():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                blob_client.upload_blob(json.dumps(job), overwrite=True)
                return True
            except Exception as exc:
                logging.error(f"Failed to queue onboarding email job {blob_name}: {exc}")
                return False

        return await loop.run_in_executor(self.executor, upload)

    async def get_job(self, ticket_number: str) -> Optional[Dict[str, Any]]:
        """Load a queued onboarding email job for a ticket."""
        container_client = await self._get_container_client()
        if not container_client:
            return None

        blob_name = self._blob_name(ticket_number)
        loop = asyncio.get_event_loop()

        def download():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                raw = blob_client.download_blob().readall()
                return json.loads(raw)
            except Exception:
                return None

        return await loop.run_in_executor(self.executor, download)

    async def has_job(self, ticket_number: str) -> bool:
        """Return True if a queued job already exists for the ticket."""
        return await self.get_job(ticket_number) is not None

    async def delete_job(self, ticket_number: str) -> bool:
        """Delete a queued onboarding email job."""
        container_client = await self._get_container_client()
        if not container_client:
            return False

        blob_name = self._blob_name(ticket_number)
        loop = asyncio.get_event_loop()

        def delete():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                blob_client.delete_blob()
                return True
            except Exception as exc:
                logging.error(f"Failed to delete onboarding email job {blob_name}: {exc}")
                return False

        return await loop.run_in_executor(self.executor, delete)

    async def list_due_jobs(self, now_utc: Optional[datetime] = None) -> List[Dict[str, Any]]:
        """Return queued jobs whose send time has arrived."""
        container_client = await self._get_container_client()
        if not container_client:
            return []

        now = now_utc or datetime.now(timezone.utc)
        loop = asyncio.get_event_loop()

        def load_due_jobs():
            due_jobs: List[Dict[str, Any]] = []
            try:
                for blob in container_client.list_blobs():
                    blob_client = container_client.get_blob_client(blob.name)
                    raw = blob_client.download_blob().readall()
                    job = json.loads(raw)
                    send_at_raw = job.get("send_at_utc")
                    if not send_at_raw:
                        continue
                    send_at = datetime.fromisoformat(send_at_raw.replace("Z", "+00:00"))
                    if send_at <= now:
                        due_jobs.append(job)
            except Exception as exc:
                logging.error(f"Failed to load queued onboarding emails: {exc}")
            return due_jobs

        return await loop.run_in_executor(self.executor, load_due_jobs)


class UserCreationApprovalStore:
    """Persist pending user-creation approvals in Azure Blob Storage."""

    def __init__(self):
        self.storage_connection_string = (
            os.environ.get("TEAMS_STORAGE_CONNECTION_STRING", "")
            or os.environ.get("AzureWebJobsStorage", "")
            or os.environ.get("AZURE_STORAGE_CONNECTION_STRING", "")
        )
        self.container_name = os.environ.get(
            "USER_CREATION_APPROVAL_CONTAINER",
            "user-creation-approvals",
        )
        self.executor = ThreadPoolExecutor(max_workers=2)
        self._blob_service_client = None
        self._container_ready = False

    def is_configured(self) -> bool:
        return bool(BlobServiceClient and self.storage_connection_string)

    def _blob_name(self, request_id: str) -> str:
        safe_request = re.sub(r"[^a-z0-9._-]+", "_", (request_id or "").lower())
        return f"{safe_request}.json"

    async def _get_container_client(self):
        if not self.is_configured():
            return None

        if self._blob_service_client is None:
            try:
                self._blob_service_client = BlobServiceClient.from_connection_string(
                    self.storage_connection_string
                )
            except Exception as exc:
                logging.error(f"Failed to initialize blob client for approval store: {exc}")
                return None

        container_client = self._blob_service_client.get_container_client(self.container_name)

        if not self._container_ready:
            loop = asyncio.get_event_loop()

            def ensure_container():
                try:
                    container_client.create_container()
                except Exception:
                    pass
                return True

            await loop.run_in_executor(self.executor, ensure_container)
            self._container_ready = True

        return container_client

    async def save_request(self, request: Dict[str, Any]) -> bool:
        container_client = await self._get_container_client()
        if not container_client:
            return False

        blob_name = self._blob_name(request.get("request_id", ""))
        loop = asyncio.get_event_loop()

        def upload():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                blob_client.upload_blob(json.dumps(request), overwrite=True)
                return True
            except Exception as exc:
                logging.error(f"Failed to save approval request {blob_name}: {exc}")
                return False

        return await loop.run_in_executor(self.executor, upload)

    async def get_request(self, request_id: str) -> Optional[Dict[str, Any]]:
        container_client = await self._get_container_client()
        if not container_client:
            return None

        blob_name = self._blob_name(request_id)
        loop = asyncio.get_event_loop()

        def download():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                raw = blob_client.download_blob().readall()
                return json.loads(raw)
            except Exception:
                return None

        return await loop.run_in_executor(self.executor, download)

    async def delete_request(self, request_id: str) -> bool:
        container_client = await self._get_container_client()
        if not container_client:
            return False

        blob_name = self._blob_name(request_id)
        loop = asyncio.get_event_loop()

        def delete():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                blob_client.delete_blob()
                return True
            except Exception as exc:
                logging.error(f"Failed to delete approval request {blob_name}: {exc}")
                return False

        return await loop.run_in_executor(self.executor, delete)


class UserCreationGraphClient(M365GraphClient):
    """Graph client extensions for onboarding users."""

    async def get_user_by_upn(self, user_principal_name: str) -> Optional[Dict[str, Any]]:
        encoded_upn = quote(user_principal_name)
        result = await self._graph_request(
            "get",
            f"/users/{encoded_upn}?$select=id,displayName,userPrincipalName",
        )
        status = result.get("_status_code", 0)
        if status == 200:
            return result
        if status == 404:
            return None
        if status:
            logging.error(f"Graph user lookup failed for {user_principal_name}: {status} {result}")
        return None

    async def create_or_reset_user(
        self,
        first_name: str,
        last_name: str,
        user_principal_name: str,
        password: str,
        usage_location: str,
    ) -> Dict[str, Any]:
        """Create a cloud-only M365 user or reset the password if it already exists for the same person."""
        display_name = f"{first_name} {last_name}".strip()
        mail_nickname = user_principal_name.split("@", 1)[0]

        existing_user = await self.get_user_by_upn(user_principal_name)
        if existing_user:
            existing_display = existing_user.get("displayName", "")
            if normalize_display_name(existing_display) != normalize_display_name(display_name):
                return {
                    "success": False,
                    "error": (
                        f"Username collision: {user_principal_name} already belongs to "
                        f"{existing_display or 'another user'}."
                    ),
                }

            reset_result = await self._graph_request(
                "patch",
                f"/users/{quote(existing_user['id'])}",
                {
                    "usageLocation": usage_location,
                    "passwordProfile": {
                        "forceChangePasswordNextSignIn": True,
                        "password": password,
                    },
                },
            )
            status = reset_result.get("_status_code", 0)
            if status == 204:
                return {
                    "success": True,
                    "created": False,
                    "reset_password": True,
                    "user_id": existing_user.get("id"),
                    "display_name": existing_display or display_name,
                    "user_principal_name": existing_user.get("userPrincipalName", user_principal_name),
                }

            return {
                "success": False,
                "error": reset_result.get("error", {}).get("message", f"HTTP {status}"),
            }

        result = await self._graph_request(
            "post",
            "/users",
            {
                "accountEnabled": True,
                "displayName": display_name,
                "givenName": first_name,
                "surname": last_name,
                "mailNickname": mail_nickname,
                "userPrincipalName": user_principal_name,
                "usageLocation": usage_location,
                "passwordProfile": {
                    "forceChangePasswordNextSignIn": True,
                    "password": password,
                },
            },
        )
        status = result.get("_status_code", 0)
        if status in (200, 201):
            return {
                "success": True,
                "created": True,
                "user_id": result.get("id"),
                "display_name": result.get("displayName", display_name),
                "user_principal_name": result.get("userPrincipalName", user_principal_name),
            }

        return {
            "success": False,
            "error": result.get("error", {}).get("message", f"HTTP {status}"),
        }

    async def resolve_license_sku_id(self, sku_hint: str) -> Tuple[Optional[str], Optional[str]]:
        """Resolve a configured license hint into a concrete SKU ID."""
        hint = (sku_hint or "").strip()
        if not hint:
            return None, "USER_CREATION_STANDARD_LICENSE_SKU is not configured"

        if re.fullmatch(r"[0-9a-fA-F-]{36}", hint):
            return hint, None

        result = await self._graph_request("get", "/subscribedSkus")
        status = result.get("_status_code", 0)
        if status != 200:
            return None, result.get("error", {}).get("message", f"HTTP {status}")

        for sku in result.get("value", []):
            if sku.get("skuPartNumber", "").upper() == hint.upper():
                return sku.get("skuId"), None

        return None, f"Could not find subscribed SKU matching {hint}"

    async def assign_licenses(self, user_id: str, sku_ids: List[str]) -> Dict[str, Any]:
        """Assign one or more licenses to the new user."""
        unique_sku_ids: List[str] = []
        seen = set()
        for sku_id in sku_ids:
            if sku_id and sku_id not in seen:
                unique_sku_ids.append(sku_id)
                seen.add(sku_id)

        if not unique_sku_ids:
            return {"success": False, "error": "At least one license SKU is required"}

        result = await self._graph_request(
            "post",
            f"/users/{quote(user_id)}/assignLicense",
            {
                "addLicenses": [
                    {
                        "skuId": sku_id,
                        "disabledPlans": [],
                    }
                    for sku_id in unique_sku_ids
                ],
                "removeLicenses": [],
            },
        )
        status = result.get("_status_code", 0)
        if status == 200:
            return {"success": True}
        return {
            "success": False,
            "error": result.get("error", {}).get("message", f"HTTP {status}"),
        }

    async def add_user_to_group(self, group_id: str, user_id: str) -> Dict[str, Any]:
        """Grant access by adding the user to a security or M365 group."""
        result = await self._graph_request(
            "post",
            f"/groups/{quote(group_id)}/members/$ref",
            {
                "@odata.id": f"{self.GRAPH_BASE}/directoryObjects/{user_id}",
            },
        )
        status = result.get("_status_code", 0)
        if status == 204:
            return {"success": True}

        error_message = result.get("error", {}).get("message", f"HTTP {status}")
        if status == 400 and "already" in error_message.lower():
            return {"success": True, "already_member": True}

        return {"success": False, "error": error_message}

    async def send_mail(
        self,
        sender_email: str,
        to_recipients: List[str],
        subject: str,
        body: str,
        cc_recipients: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Send an email from a specific mailbox using Graph."""
        cc_recipients = cc_recipients or []
        result = await self._graph_request(
            "post",
            f"/users/{quote(sender_email)}/sendMail",
            {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "Text",
                        "content": body,
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": email}}
                        for email in to_recipients
                    ],
                    "ccRecipients": [
                        {"emailAddress": {"address": email}}
                        for email in cc_recipients
                    ],
                },
                "saveToSentItems": True,
            },
        )
        status = result.get("_status_code", 0)
        if status == 202:
            return {"success": True}
        return {
            "success": False,
            "error": result.get("error", {}).get("message", f"HTTP {status}"),
        }


class UserCreationManager:
    """Orchestrates end-to-end automation for User Creation tickets."""

    ACTIVE_TICKET_STATUSES = {"New", "Bot Assisted", "In Progress", "Awaiting User", "Awaiting IT"}

    def __init__(self):
        self.graph_client = UserCreationGraphClient()
        self.quickbase = QuickBaseManager()
        self.email_store = ScheduledOnboardingEmailStore()
        self.approval_store = UserCreationApprovalStore()
        self.executor = ThreadPoolExecutor(max_workers=2)
        self.standard_license_sku = os.environ.get("USER_CREATION_STANDARD_LICENSE_SKU", "")
        self.intune_license_sku = os.environ.get("USER_CREATION_INTUNE_LICENSE_SKU", "")
        self.sharepoint_group_id = os.environ.get("USER_CREATION_SHAREPOINT_GROUP_ID", "")
        self.email_sender = os.environ.get("USER_CREATION_EMAIL_SENDER", "")
        self.email_cc = [
            email.strip()
            for email in os.environ.get("USER_CREATION_EMAIL_CC", "").split(",")
            if email.strip()
        ]
        self.email_timezone = os.environ.get(
            "USER_CREATION_EMAIL_TIMEZONE",
            DEFAULT_EMAIL_TIMEZONE,
        )
        self.usage_location = os.environ.get(
            "USER_CREATION_USAGE_LOCATION",
            DEFAULT_USAGE_LOCATION,
        )
        self.quickbase_participant_app_id = os.environ.get(
            "QB_USER_CREATION_APP_ID",
            DEFAULT_QB_APP_ID,
        )
        self.quickbase_participant_role_name = os.environ.get(
            "QB_USER_CREATION_ROLE_NAME",
            DEFAULT_QB_ROLE_NAME,
        )
        self.openai_admin_api_key = os.environ.get("OPENAI_ADMIN_API_KEY", "")
        self.openai_invite_role = os.environ.get(
            "OPENAI_USER_CREATION_ROLE",
            DEFAULT_OPENAI_ROLE,
        )
        self.openai_projects = self._parse_json_list(
            os.environ.get("OPENAI_USER_CREATION_PROJECTS", "")
        )
        self.anthropic_admin_api_key = os.environ.get("ANTHROPIC_ADMIN_API_KEY", "")
        self.anthropic_invite_role = os.environ.get(
            "ANTHROPIC_USER_CREATION_ROLE",
            DEFAULT_ANTHROPIC_ROLE,
        )

    @staticmethod
    def _parse_json_list(raw_value: str) -> List[Dict[str, Any]]:
        if not raw_value.strip():
            return []
        try:
            parsed = json.loads(raw_value)
            return parsed if isinstance(parsed, list) else []
        except Exception:
            logging.warning("Failed to parse JSON list from environment value")
            return []

    def validate_configuration(self) -> List[str]:
        """Return missing configuration items required for onboarding automation."""
        missing = []
        if not self.graph_client.client_id:
            missing.append("M365_GRAPH_CLIENT_ID")
        if not self.graph_client.client_secret:
            missing.append("M365_GRAPH_CLIENT_SECRET")
        if not self.graph_client.tenant_id:
            missing.append("M365_GRAPH_TENANT_ID or TEAMS_TENANT_ID")
        if not self.graph_client.domain:
            missing.append("M365_GRAPH_DOMAIN")
        if not self.standard_license_sku:
            missing.append("USER_CREATION_STANDARD_LICENSE_SKU")
        if not self.intune_license_sku:
            missing.append("USER_CREATION_INTUNE_LICENSE_SKU")
        if not self.sharepoint_group_id:
            missing.append("USER_CREATION_SHAREPOINT_GROUP_ID")
        if not self.email_sender:
            missing.append("USER_CREATION_EMAIL_SENDER")
        if not os.environ.get("AUTOMATION_ADMIN_EMAIL", "").strip():
            missing.append("AUTOMATION_ADMIN_EMAIL")
        if not self.email_store.is_configured():
            missing.append("AzureWebJobsStorage/TEAMS_STORAGE_CONNECTION_STRING")
        if not self.quickbase.realm:
            missing.append("QB_REALM")
        if not self.quickbase.user_token:
            missing.append("QB_USER_TOKEN")
        return missing

    async def prepare_approval_request(self, ticket_data: Dict[str, Any]) -> Dict[str, Any]:
        """Create a pending approval for user creation instead of executing immediately."""
        ticket_number = ticket_data.get("ticket_number", "")
        record_id = str(ticket_data.get("record_id", "") or "")
        if not is_user_creation_category(ticket_data.get("category", "")):
            return {"success": False, "skipped": True, "reason": "not a User Creation ticket"}

        status = (ticket_data.get("status") or "").strip()
        if status and status not in {"New"}:
            return {"success": False, "skipped": True, "reason": f"ticket status is {status}"}

        missing_config = self.validate_configuration()
        if missing_config:
            return {
                "success": False,
                "error": "Missing configuration: " + ", ".join(missing_config),
            }

        if not ticket_number and not record_id:
            return {"success": False, "error": "ticket_number or record_id is required"}

        ticket_identifier = ticket_number or f"RID-{record_id}"

        if await self.email_store.get_job(ticket_identifier):
            return {
                "success": False,
                "skipped": True,
                "reason": "credential email already queued",
            }

        request_id = ticket_identifier
        existing_request = await self.approval_store.get_request(request_id)
        if existing_request:
            return {
                "success": True,
                "pending_approval": True,
                "request": existing_request,
                "skipped": True,
                "reason": "approval already pending",
            }

        source_text = get_user_creation_source_text(ticket_data)
        first_name, last_name, personal_email = extract_user_creation_details(source_text)
        display_name = f"{first_name} {last_name}".strip()
        username_local = build_username_local_part(first_name, last_name)
        user_principal_name = f"{username_local}@{self.graph_client.domain}"
        request = {
            "request_id": request_id,
            "ticket_number": ticket_number,
            "record_id": record_id,
            "ticket_data": ticket_data,
            "first_name": first_name,
            "last_name": last_name,
            "display_name": display_name,
            "personal_email": personal_email or "",
            "predicted_username_local": username_local,
            "predicted_user_principal_name": user_principal_name,
            "created_at_utc": datetime.now(timezone.utc).isoformat(),
            "status": "pending_approval",
        }
        saved = await self.approval_store.save_request(request)
        if not saved:
            return {"success": False, "error": "Failed to save pending approval request"}

        await self.quickbase.append_ticket_resolution_note(
            ticket_number=ticket_number,
            record_id=record_id,
            note=(
                "[USER_CREATION_PENDING_APPROVAL]\n"
                f"Awaiting admin approval for username {user_principal_name} "
                f"and credential email to {personal_email or 'N/A'}."
            ),
            status="Awaiting IT",
        )

        return {
            "success": True,
            "pending_approval": True,
            "request": request,
        }

    async def get_approval_request(self, request_id: str) -> Optional[Dict[str, Any]]:
        return await self.approval_store.get_request(request_id)

    async def deny_approval_request(self, request_id: str, denial_reason: str = "") -> Dict[str, Any]:
        request = await self.approval_store.get_request(request_id)
        if not request:
            return {"success": False, "error": "Approval request not found"}

        ticket_number = request.get("ticket_number", "")
        record_id = request.get("record_id", "")
        await self.approval_store.delete_request(request_id)
        await self.quickbase.append_ticket_resolution_note(
            ticket_number=ticket_number,
            record_id=record_id,
            note=(
                "[USER_CREATION_DENIED]\n"
                f"Admin denied automated provisioning. {denial_reason}".strip()
            ),
            status="Awaiting IT",
        )
        return {"success": True, "ticket_number": ticket_number, "record_id": record_id}

    async def execute_approved_request(
        self,
        request_id: str,
        username_value: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Execute a previously approved user-creation request."""
        request = await self.approval_store.get_request(request_id)
        if not request:
            return {"success": False, "error": "Approval request not found"}

        username_local = (username_value or request.get("predicted_username_local") or "").strip()
        if "@" in username_local:
            username_local = username_local.split("@", 1)[0].strip()
        username_local = re.sub(r"[^a-z0-9]", "", username_local.lower())
        if not username_local:
            return {"success": False, "error": "A confirmed username is required"}

        result = await self._execute_ticket_data(
            ticket_data=request.get("ticket_data", {}),
            username_local_override=username_local,
        )
        if not result.get("success"):
            await self.quickbase.append_ticket_resolution_note(
                ticket_number=request.get("ticket_number", ""),
                record_id=request.get("record_id", ""),
                note=(
                    "[USER_CREATION_ERROR]\n"
                    f"{result.get('error', 'Unknown automation error')}"
                ),
                status="Awaiting IT",
            )
        await self.approval_store.delete_request(request_id)
        return result

    async def process_ticket(self, ticket_data: Dict[str, Any]) -> Dict[str, Any]:
        """Provision the user and queue the due-date email."""
        return await self._execute_ticket_data(ticket_data=ticket_data)

    async def _execute_ticket_data(
        self,
        ticket_data: Dict[str, Any],
        username_local_override: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Provision the user and queue the due-date email."""
        ticket_number = ticket_data.get("ticket_number", "")
        record_id = str(ticket_data.get("record_id", "") or "")
        if not is_user_creation_category(ticket_data.get("category", "")):
            return {"success": False, "skipped": True, "reason": "not a User Creation ticket"}

        status = (ticket_data.get("status") or "").strip()
        if status in {"Resolved", "Closed", "Cancelled", "Bot Assisted"}:
            return {"success": False, "skipped": True, "reason": f"ticket status is {status}"}

        missing_config = self.validate_configuration()
        if missing_config:
            return {
                "success": False,
                "error": "Missing configuration: " + ", ".join(missing_config),
            }

        if not ticket_number and not record_id:
            return {"success": False, "error": "ticket_number or record_id is required"}

        ticket_identifier = ticket_number or f"RID-{record_id}"

        existing_job = await self.email_store.get_job(ticket_identifier)
        if existing_job:
            return {
                "success": True,
                "queued": True,
                "skipped": True,
                "reason": "credential email already queued",
                "user_principal_name": existing_job.get("username"),
            }

        source_text = get_user_creation_source_text(ticket_data)
        first_name, last_name, personal_email = extract_user_creation_details(source_text)
        display_name = f"{first_name} {last_name}".strip()
        username_local = username_local_override or build_username_local_part(first_name, last_name)
        user_principal_name = f"{username_local}@{self.graph_client.domain}"
        temporary_password = generate_temporary_password()
        send_at_utc = calculate_due_date_send_at(
            ticket_data.get("due_date", ""),
            self.email_timezone,
        )
        recipient_email = (personal_email or "").strip().lower()
        if not recipient_email:
            return {"success": False, "error": "A personal email is required in the ticket description"}

        user_result = await self.graph_client.create_or_reset_user(
            first_name=first_name,
            last_name=last_name,
            user_principal_name=user_principal_name,
            password=temporary_password,
            usage_location=self.usage_location,
        )
        if not user_result.get("success"):
            return user_result

        standard_sku_id, standard_sku_error = await self.graph_client.resolve_license_sku_id(
            self.standard_license_sku
        )
        if standard_sku_error or not standard_sku_id:
            return {
                "success": False,
                "error": standard_sku_error or "Unable to resolve standard license SKU",
            }

        intune_sku_id, intune_sku_error = await self.graph_client.resolve_license_sku_id(
            self.intune_license_sku
        )
        if intune_sku_error or not intune_sku_id:
            return {
                "success": False,
                "error": intune_sku_error or "Unable to resolve Intune license SKU",
            }

        license_result = await self.graph_client.assign_licenses(
            user_result["user_id"],
            [standard_sku_id, intune_sku_id],
        )
        if not license_result.get("success"):
            return license_result

        sharepoint_result = await self.graph_client.add_user_to_group(
            self.sharepoint_group_id,
            user_result["user_id"],
        )
        if not sharepoint_result.get("success"):
            return sharepoint_result

        qb_result = await self.quickbase.ensure_app_user_in_role(
            email=user_principal_name,
            first_name=first_name,
            last_name=last_name,
            app_id=self.quickbase_participant_app_id,
            role_name=self.quickbase_participant_role_name,
        )
        if not qb_result.get("success"):
            return qb_result

        openai_result = await self.invite_openai_user(user_principal_name)
        if not openai_result.get("success"):
            return openai_result

        anthropic_result = await self.invite_anthropic_user(user_principal_name)
        if not anthropic_result.get("success"):
            return anthropic_result

        job = {
            "ticket_number": ticket_identifier,
            "source_ticket_number": ticket_number,
            "source_record_id": record_id,
            "display_name": display_name,
            "username": user_principal_name,
            "temporary_password": temporary_password,
            "to": [recipient_email],
            "cc": self.email_cc,
            "sender": self.email_sender,
            "send_at_utc": send_at_utc.isoformat(),
            "subject": build_onboarding_email_subject(display_name),
            "body": build_onboarding_email_body(
                display_name=display_name,
                username=user_principal_name,
                temporary_password=temporary_password,
                ticket_number=ticket_identifier,
            ),
            "created_at_utc": datetime.now(timezone.utc).isoformat(),
            "attempts": 0,
        }
        queued = await self.email_store.queue_job(job)
        if not queued:
            return {"success": False, "error": "Failed to queue the onboarding email"}

        resolution = build_initial_ticket_resolution(
            display_name=display_name,
            username=user_principal_name,
            recipient_email=recipient_email,
            send_at_utc=send_at_utc,
        )
        await self.quickbase.append_ticket_resolution_note(
            ticket_number=ticket_number,
            record_id=record_id,
            note=resolution,
            status="In Progress",
        )

        return {
            "success": True,
            "ticket_number": ticket_identifier,
            "source_ticket_number": ticket_number,
            "record_id": record_id,
            "display_name": display_name,
            "user_principal_name": user_principal_name,
            "recipient_email": recipient_email,
            "email_queued_for": send_at_utc.isoformat(),
            "openai_invite": openai_result,
            "anthropic_invite": anthropic_result,
        }

    async def dispatch_due_emails(self, now_utc: Optional[datetime] = None) -> Dict[str, Any]:
        """Send due onboarding emails and mark completed tickets as Bot Assisted."""
        now = now_utc or datetime.now(timezone.utc)
        missing_config = self.validate_configuration()
        if missing_config:
            return {
                "success": False,
                "error": "Missing configuration: " + ", ".join(missing_config),
                "processed": 0,
            }

        due_jobs = await self.email_store.list_due_jobs(now)
        processed = 0
        sent = 0
        failed = 0

        for job in due_jobs:
            processed += 1
            ticket_number = job.get("ticket_number", "")
            send_result = await self.graph_client.send_mail(
                sender_email=job.get("sender", self.email_sender),
                to_recipients=job.get("to", []),
                cc_recipients=job.get("cc", []),
                subject=job.get("subject", ""),
                body=job.get("body", ""),
            )
            if send_result.get("success"):
                sent += 1
                await self.email_store.delete_job(ticket_number)
                resolution = build_completion_ticket_resolution(
                    display_name=job.get("display_name", "New User"),
                    username=job.get("username", ""),
                    recipient_email=", ".join(job.get("to", [])),
                    sent_at_utc=now,
                )
                await self.quickbase.append_ticket_resolution_note(
                    ticket_number=job.get("source_ticket_number", ""),
                    record_id=job.get("source_record_id", ""),
                    note=resolution,
                    status="Bot Assisted",
                )
                continue

            failed += 1
            job["attempts"] = int(job.get("attempts", 0)) + 1
            job["last_error"] = send_result.get("error", "Unknown error")
            await self.email_store.queue_job(job)
            logging.error(
                f"Failed to send onboarding email for ticket {ticket_number}: "
                f"{job['last_error']}"
            )

        return {
            "success": True,
            "processed": processed,
            "sent": sent,
            "failed": failed,
        }

    async def invite_openai_user(self, email: str) -> Dict[str, Any]:
        """Invite a user to the OpenAI Platform organization when configured."""
        if not self.openai_admin_api_key:
            return {"success": True, "skipped": True, "reason": "OPENAI_ADMIN_API_KEY not configured"}

        loop = asyncio.get_event_loop()

        def send_invite():
            try:
                payload = {
                    "email": email,
                    "role": self.openai_invite_role,
                }
                if self.openai_projects:
                    payload["projects"] = self.openai_projects
                response = requests.post(
                    "https://api.openai.com/v1/organization/invites",
                    headers={
                        "Authorization": f"Bearer {self.openai_admin_api_key}",
                        "Content-Type": "application/json",
                    },
                    json=payload,
                    timeout=30,
                )
                result = response.json() if response.text else {}
                if response.status_code in (200, 201):
                    return {"success": True, "invite_id": result.get("id"), "status": result.get("status")}
                if response.status_code == 409:
                    return {"success": True, "skipped": True, "reason": "OpenAI invite already exists"}
                error = result.get("error", {}).get("message", response.text)
                return {"success": False, "error": f"OpenAI invite failed: {error}"}
            except Exception as exc:
                return {"success": False, "error": f"OpenAI invite failed: {exc}"}

        return await loop.run_in_executor(self.executor, send_invite)

    async def invite_anthropic_user(self, email: str) -> Dict[str, Any]:
        """Invite a user to Anthropic when configured."""
        if not self.anthropic_admin_api_key:
            return {"success": True, "skipped": True, "reason": "ANTHROPIC_ADMIN_API_KEY not configured"}

        loop = asyncio.get_event_loop()

        def send_invite():
            try:
                response = requests.post(
                    "https://api.anthropic.com/v1/organizations/invites",
                    headers={
                        "x-api-key": self.anthropic_admin_api_key,
                        "anthropic-version": "2023-06-01",
                        "content-type": "application/json",
                    },
                    json={
                        "email": email,
                        "role": self.anthropic_invite_role,
                    },
                    timeout=30,
                )
                result = response.json() if response.text else {}
                if response.status_code in (200, 201):
                    return {"success": True, "invite_id": result.get("id"), "status": result.get("status")}
                if response.status_code == 409:
                    return {"success": True, "skipped": True, "reason": "Anthropic invite already exists"}
                return {
                    "success": False,
                    "error": f"Anthropic invite failed: {result.get('error', {}).get('message', response.text)}",
                }
            except Exception as exc:
                return {"success": False, "error": f"Anthropic invite failed: {exc}"}

        return await loop.run_in_executor(self.executor, send_invite)
