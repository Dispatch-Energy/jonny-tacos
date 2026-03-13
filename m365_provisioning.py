"""
M365 Provisioning Handler - Automates creation of SharePoint sites,
Teams teams, and shared mailboxes via Microsoft Graph API.

Implements the AutomationHandler interface from automation_manager.py.
Env vars are prefixed with M365_GRAPH_ to keep them contained to this task.

Graph API permissions required (Application):
- Group.ReadWrite.All
- Directory.ReadWrite.All
- Sites.ReadWrite.All
"""

import os
import re
import logging
import asyncio
import requests
from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor

from automation_manager import AutomationHandler, AutomationRequest, AutomationStatus


# ============================================================================
# CONSTANTS
# ============================================================================

RESOURCE_TYPES = {
    "shared_mailbox": {
        "display": "Shared Mailbox",
        "icon": "📧",
        "description": "A shared email address (like support@ or info@) that multiple people can send and receive from.",
    },
    "teams_team": {
        "display": "Microsoft Team",
        "icon": "👥",
        "description": "A collaboration workspace with chat, video calls, shared files, and a group email — all in one.",
    },
    "sharepoint_site": {
        "display": "SharePoint Site",
        "icon": "📁",
        "description": "A site for document storage, collaboration, news, and lists.",
    },
}

# Keywords for detecting M365 provisioning requests
M365_KEYWORDS = {
    "shared_mailbox": [
        "shared mailbox", "shared email", "shared inbox", "group email",
        "distribution list", "email address for", "email for the",
        "mailbox for", "create an email", "create email", "new email address",
        "set up an email", "set up email", "need an email", "need a mailbox",
        "reporting@", "support@", "info@", "sales@",
    ],
    "teams_team": [
        "teams channel", "teams team", "microsoft team", "team channel",
        "set up a team", "create a team", "new team for", "team for the",
        "teams workspace", "collaboration channel",
    ],
    "sharepoint_site": [
        "sharepoint site", "sharepoint", "document library", "document site",
        "collaboration site", "file sharing site", "intranet site",
        "create a site", "new site for",
    ],
    "general": [
        "shared workspace", "workspace for", "collaboration space",
        "group workspace", "team workspace", "set up a workspace",
    ],
}


# ============================================================================
# M365 GRAPH API CLIENT
# ============================================================================

class M365GraphClient:
    """Microsoft Graph API client for M365 resource provisioning.

    Uses its own app registration credentials (M365_GRAPH_*) separate from
    the Teams bot credentials to follow principle of least privilege.
    """

    GRAPH_BASE = "https://graph.microsoft.com/v1.0"

    def __init__(self):
        self.client_id = os.getenv("M365_GRAPH_CLIENT_ID", "")
        self.client_secret = os.getenv("M365_GRAPH_CLIENT_SECRET", "")
        self.tenant_id = os.getenv(
            "M365_GRAPH_TENANT_ID",
            os.getenv("TEAMS_TENANT_ID", "")
        )
        self.domain = os.getenv("M365_GRAPH_DOMAIN", "")
        self.executor = ThreadPoolExecutor(max_workers=2)
        self._token = None
        self._token_expiry = None

    async def _get_token(self) -> str:
        """Get or refresh Graph API access token."""
        if self._token and self._token_expiry and datetime.now() < self._token_expiry:
            return self._token

        token_url = (
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        )
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }

        loop = asyncio.get_event_loop()

        def fetch_token():
            resp = requests.post(token_url, data=data, timeout=30)
            if resp.status_code == 200:
                token_data = resp.json()
                return (
                    token_data.get("access_token"),
                    token_data.get("expires_in", 3600),
                )
            logging.error(f"Graph token request failed: {resp.status_code} {resp.text}")
            return None, None

        token, expires_in = await loop.run_in_executor(self.executor, fetch_token)

        if token:
            self._token = token
            self._token_expiry = datetime.now() + timedelta(seconds=expires_in - 60)
            return token

        return ""

    async def _graph_request(self, method: str, endpoint: str,
                             json_body: Dict = None) -> Dict[str, Any]:
        """Make an authenticated Graph API request."""
        token = await self._get_token()
        if not token:
            return {"error": "Failed to get Graph API token"}

        url = f"{self.GRAPH_BASE}{endpoint}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        loop = asyncio.get_event_loop()

        def do_request():
            resp = getattr(requests, method.lower())(
                url, headers=headers, json=json_body, timeout=30
            )
            try:
                result = resp.json()
            except Exception:
                result = {}
            result["_status_code"] = resp.status_code
            return result

        return await loop.run_in_executor(self.executor, do_request)

    async def create_m365_group(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """Create an M365 Unified Group (foundation for Teams/SharePoint/Mailbox).

        An M365 Group automatically provisions:
        - A shared mailbox (mailEnabled=true)
        - A SharePoint team site
        - A Planner plan
        Adding a Team on top gives the full Teams experience.
        """
        mail_nickname = re.sub(r'[^a-zA-Z0-9]', '', config.get("mail_nickname", ""))

        owners = config.get("owners", [])
        members = config.get("members", [])

        body = {
            "displayName": config.get("display_name", ""),
            "description": config.get("description", ""),
            "mailEnabled": True,
            "mailNickname": mail_nickname,
            "securityEnabled": False,
            "groupTypes": ["Unified"],
            "visibility": config.get("visibility", "Private"),
        }

        # Add owners
        if owners:
            body["owners@odata.bind"] = [
                f"{self.GRAPH_BASE}/users/{email}" for email in owners
            ]

        # Add members (owners are auto-added as members)
        if members:
            body["members@odata.bind"] = [
                f"{self.GRAPH_BASE}/users/{email}" for email in members
            ]

        result = await self._graph_request("post", "/groups", body)
        status = result.get("_status_code", 0)

        if status in (200, 201):
            logging.info(f"M365 Group created: {result.get('id')} ({mail_nickname})")
            return {
                "success": True,
                "group_id": result.get("id"),
                "mail": result.get("mail", f"{mail_nickname}@{self.domain}"),
                "display_name": result.get("displayName"),
            }

        logging.error(f"Failed to create M365 Group: {status} {result}")
        return {
            "success": False,
            "error": result.get("error", {}).get("message", f"HTTP {status}"),
        }

    async def teamify_group(self, group_id: str) -> Dict[str, Any]:
        """Add Teams capabilities to an existing M365 Group."""
        body = {
            "memberSettings": {"allowCreateUpdateChannels": True},
            "messagingSettings": {
                "allowUserEditMessages": True,
                "allowUserDeleteMessages": True,
            },
            "funSettings": {
                "allowGiphy": True,
                "giphyContentRating": "moderate",
            },
        }

        result = await self._graph_request("put", f"/groups/{group_id}/team", body)
        status = result.get("_status_code", 0)

        if status in (200, 201):
            logging.info(f"Team created for group {group_id}")
            return {"success": True, "group_id": group_id}

        logging.error(f"Failed to teamify group {group_id}: {status} {result}")
        return {
            "success": False,
            "error": result.get("error", {}).get("message", f"HTTP {status}"),
        }

    async def get_group_sharepoint_url(self, group_id: str) -> str:
        """Get the SharePoint site URL for an M365 Group."""
        result = await self._graph_request("get", f"/groups/{group_id}/sites/root")
        return result.get("webUrl", "")

    async def provision_shared_mailbox(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """Provision a shared mailbox by creating an M365 Group with mail enabled."""
        group_result = await self.create_m365_group(config)
        if not group_result.get("success"):
            return group_result

        return {
            "success": True,
            "type": "shared_mailbox",
            "group_id": group_result["group_id"],
            "email": group_result.get("mail", ""),
            "display_name": group_result.get("display_name", ""),
        }

    async def provision_teams_team(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """Provision a Microsoft Team (M365 Group + Teams layer)."""
        group_result = await self.create_m365_group(config)
        if not group_result.get("success"):
            return group_result

        group_id = group_result["group_id"]

        # Graph API needs a short delay before the group can be teamified
        await asyncio.sleep(3)

        team_result = await self.teamify_group(group_id)
        if not team_result.get("success"):
            return {
                "success": False,
                "error": f"Group created but Team creation failed: {team_result.get('error')}",
                "group_id": group_id,
                "partial": True,
            }

        sharepoint_url = await self.get_group_sharepoint_url(group_id)

        return {
            "success": True,
            "type": "teams_team",
            "group_id": group_id,
            "email": group_result.get("mail", ""),
            "display_name": group_result.get("display_name", ""),
            "sharepoint_url": sharepoint_url,
        }

    async def provision_sharepoint_site(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """Provision a SharePoint site via M365 Group (creates a team site)."""
        group_result = await self.create_m365_group(config)
        if not group_result.get("success"):
            return group_result

        group_id = group_result["group_id"]

        # Give it a moment to provision the SharePoint site
        await asyncio.sleep(2)

        sharepoint_url = await self.get_group_sharepoint_url(group_id)

        return {
            "success": True,
            "type": "sharepoint_site",
            "group_id": group_id,
            "email": group_result.get("mail", ""),
            "display_name": group_result.get("display_name", ""),
            "sharepoint_url": sharepoint_url,
        }


# ============================================================================
# M365 PROVISIONING HANDLER
# ============================================================================

class M365ProvisioningHandler(AutomationHandler):
    """Handles M365 resource provisioning requests."""

    @property
    def automation_type(self) -> str:
        return "m365_provisioning"

    @property
    def display_name(self) -> str:
        return "M365 Resource Provisioning"

    def __init__(self):
        self.graph_client = M365GraphClient()

    def detect_intent(self, message: str) -> Optional[Dict[str, Any]]:
        """Detect if message is an M365 provisioning request and extract details."""
        msg_lower = message.lower()

        suggested_type = None
        match_score = 0

        for resource_type, keywords in M365_KEYWORDS.items():
            for kw in keywords:
                if kw in msg_lower:
                    if resource_type == "general":
                        if not suggested_type:
                            suggested_type = "unclear"
                            match_score = max(match_score, 1)
                    else:
                        suggested_type = resource_type
                        match_score = max(match_score, 2)

        if not suggested_type:
            return None

        # Try to extract a name/alias from the message
        extracted_name = None
        # Look for patterns like "called X", "named X", "for X team"
        name_patterns = [
            r'(?:called|named)\s+"([^"]+)"',
            r'(?:called|named)\s+(\S+)',
            r'for\s+(?:the\s+)?(.+?)(?:\s+team|\s+group|\s+department)?\s*$',
        ]
        for pattern in name_patterns:
            match = re.search(pattern, message, re.IGNORECASE)
            if match:
                extracted_name = match.group(1).strip().rstrip('?!.')
                break

        # Extract email alias if mentioned (e.g., "reporting@" or "support@company.com")
        extracted_alias = None
        alias_match = re.search(r'(\w+)@(?:\w+\.\w+)?', message)
        if alias_match:
            extracted_alias = alias_match.group(1)

        return {
            "suggested_type": suggested_type,
            "extracted_name": extracted_name,
            "extracted_alias": extracted_alias,
            "original_message": message,
        }

    def create_routing_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create the initial resource selection card with smart suggestions."""
        suggested = request.ai_extracted.get("suggested_type", "unclear")

        # Build suggestion text
        if suggested == "shared_mailbox":
            intro = (
                "It sounds like you need a shared email address. Before we set one up, "
                "consider whether a **Microsoft Team** might be a better fit — it includes "
                "a shared email, document storage, and chat all in one."
            )
        elif suggested == "teams_team":
            intro = (
                "Sounds like you need a collaboration workspace! "
                "A Microsoft Team gives you chat, meetings, shared files, and a group email."
            )
        elif suggested == "sharepoint_site":
            intro = (
                "It sounds like you need a document/collaboration site. "
                "If your team also needs chat and meetings, a Microsoft Team includes "
                "a SharePoint site automatically."
            )
        else:
            intro = (
                "I can help set up M365 resources for your team. "
                "Which type of resource best fits your needs?"
            )

        # Resource option descriptions
        options_body = []
        for rtype, info in RESOURCE_TYPES.items():
            is_suggested = rtype == suggested
            label = f"{info['icon']} **{info['display']}**"
            if is_suggested:
                label += " (Suggested)"
            options_body.append({
                "type": "TextBlock",
                "text": label,
                "weight": "Bolder",
                "spacing": "Medium",
            })
            options_body.append({
                "type": "TextBlock",
                "text": info["description"],
                "wrap": True,
                "isSubtle": True,
                "spacing": "None",
            })

        body = [
            {
                "type": "TextBlock",
                "text": "🔧 M365 Resource Setup",
                "weight": "Bolder",
                "size": "Large",
            },
            {
                "type": "TextBlock",
                "text": intro,
                "wrap": True,
                "spacing": "Medium",
            },
            {
                "type": "Container",
                "separator": True,
                "spacing": "Medium",
                "items": options_body,
            },
        ]

        # Tip about Teams including everything
        if suggested != "teams_team":
            body.append({
                "type": "TextBlock",
                "text": (
                    "💡 **Tip:** A Microsoft Team includes a shared email, "
                    "SharePoint document library, and chat — all in one!"
                ),
                "wrap": True,
                "isSubtle": True,
                "spacing": "Medium",
                "size": "Small",
            })

        actions = []
        for rtype, info in RESOURCE_TYPES.items():
            actions.append({
                "type": "Action.Submit",
                "title": f"{info['icon']} {info['display']}",
                "data": {
                    "action": "provisioning_select_type",
                    "request_id": request.request_id,
                    "resource_type": rtype,
                },
            })
        actions.append({
            "type": "Action.Submit",
            "title": "❌ Cancel",
            "data": {
                "action": "provisioning_cancel",
                "request_id": request.request_id,
            },
        })

        return {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": body,
            "actions": actions,
        }

    def create_config_form(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create configuration form based on selected resource type."""
        rtype = request.resource_type
        info = RESOURCE_TYPES.get(rtype, {})
        extracted = request.ai_extracted

        # Pre-fill values from AI extraction
        prefill_name = extracted.get("extracted_name", "")
        prefill_alias = extracted.get("extracted_alias", "")

        domain = os.getenv("M365_GRAPH_DOMAIN", "yourcompany.com")

        if rtype == "shared_mailbox":
            return self._mailbox_config_form(
                request.request_id, prefill_alias, prefill_name, domain
            )
        elif rtype == "teams_team":
            return self._teams_config_form(
                request.request_id, prefill_name
            )
        elif rtype == "sharepoint_site":
            return self._sharepoint_config_form(
                request.request_id, prefill_name
            )

        return self._teams_config_form(request.request_id, prefill_name)

    def _mailbox_config_form(self, request_id: str, prefill_alias: str,
                             prefill_name: str, domain: str) -> Dict[str, Any]:
        return {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "📧 Set Up Shared Mailbox",
                    "weight": "Bolder",
                    "size": "Large",
                },
                {
                    "type": "TextBlock",
                    "text": "Fill in the details for the new shared mailbox.",
                    "wrap": True,
                    "isSubtle": True,
                },
                {
                    "type": "Container",
                    "separator": True,
                    "spacing": "Medium",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": f"**Email Address** * (the part before @{domain})",
                            "weight": "Bolder",
                        },
                        {
                            "type": "Input.Text",
                            "id": "mail_nickname",
                            "placeholder": "e.g. support, reporting, info",
                            "value": prefill_alias or "",
                            "isRequired": True,
                            "errorMessage": "Email address is required",
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Display Name** *",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "display_name",
                            "placeholder": "e.g. Customer Support, Reporting Team",
                            "value": prefill_name or "",
                            "isRequired": True,
                            "errorMessage": "Display name is required",
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Members** * (comma-separated email addresses)",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "members",
                            "placeholder": "user1@company.com, user2@company.com",
                            "isRequired": True,
                            "isMultiline": True,
                            "errorMessage": "At least one member is required",
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Allow external senders?**",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.ChoiceSet",
                            "id": "external_senders",
                            "style": "expanded",
                            "value": "yes",
                            "choices": [
                                {"title": "Yes - anyone can email this address", "value": "yes"},
                                {"title": "No - internal only", "value": "no"},
                            ],
                        },
                    ],
                },
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit Request",
                    "style": "positive",
                    "data": {
                        "action": "provisioning_submit_config",
                        "request_id": request_id,
                        "resource_type": "shared_mailbox",
                    },
                },
                {
                    "type": "Action.Submit",
                    "title": "Cancel",
                    "data": {
                        "action": "provisioning_cancel",
                        "request_id": request_id,
                    },
                },
            ],
        }

    def _teams_config_form(self, request_id: str,
                           prefill_name: str) -> Dict[str, Any]:
        return {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "👥 Set Up Microsoft Team",
                    "weight": "Bolder",
                    "size": "Large",
                },
                {
                    "type": "TextBlock",
                    "text": (
                        "Fill in the details for the new Team. This will create "
                        "a Team with chat, a shared email, and a SharePoint document library."
                    ),
                    "wrap": True,
                    "isSubtle": True,
                },
                {
                    "type": "Container",
                    "separator": True,
                    "spacing": "Medium",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "**Team Name** *",
                            "weight": "Bolder",
                        },
                        {
                            "type": "Input.Text",
                            "id": "display_name",
                            "placeholder": "e.g. Marketing Team, Project Alpha",
                            "value": prefill_name or "",
                            "isRequired": True,
                            "errorMessage": "Team name is required",
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Description**",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "description",
                            "placeholder": "What is this team for?",
                            "isMultiline": True,
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Privacy** *",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.ChoiceSet",
                            "id": "visibility",
                            "style": "expanded",
                            "value": "Private",
                            "choices": [
                                {"title": "Private - only members can access", "value": "Private"},
                                {"title": "Public - anyone in the org can join", "value": "Public"},
                            ],
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Owners** * (comma-separated emails, at least 1)",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "owners",
                            "placeholder": "owner@company.com",
                            "isRequired": True,
                            "errorMessage": "At least one owner is required",
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Members** (comma-separated emails)",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "members",
                            "placeholder": "user1@company.com, user2@company.com",
                            "isMultiline": True,
                        },
                    ],
                },
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit Request",
                    "style": "positive",
                    "data": {
                        "action": "provisioning_submit_config",
                        "request_id": request_id,
                        "resource_type": "teams_team",
                    },
                },
                {
                    "type": "Action.Submit",
                    "title": "Cancel",
                    "data": {
                        "action": "provisioning_cancel",
                        "request_id": request_id,
                    },
                },
            ],
        }

    def _sharepoint_config_form(self, request_id: str,
                                prefill_name: str) -> Dict[str, Any]:
        return {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "📁 Set Up SharePoint Site",
                    "weight": "Bolder",
                    "size": "Large",
                },
                {
                    "type": "TextBlock",
                    "text": "Fill in the details for the new SharePoint site.",
                    "wrap": True,
                    "isSubtle": True,
                },
                {
                    "type": "Container",
                    "separator": True,
                    "spacing": "Medium",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "**Site Name** *",
                            "weight": "Bolder",
                        },
                        {
                            "type": "Input.Text",
                            "id": "display_name",
                            "placeholder": "e.g. Project Documentation, HR Policies",
                            "value": prefill_name or "",
                            "isRequired": True,
                            "errorMessage": "Site name is required",
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Description**",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "description",
                            "placeholder": "What is this site for?",
                            "isMultiline": True,
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Privacy** *",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.ChoiceSet",
                            "id": "visibility",
                            "style": "expanded",
                            "value": "Private",
                            "choices": [
                                {"title": "Private - only members can access", "value": "Private"},
                                {"title": "Public - anyone in the org can access", "value": "Public"},
                            ],
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Owners** * (comma-separated emails, at least 1)",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "owners",
                            "placeholder": "owner@company.com",
                            "isRequired": True,
                            "errorMessage": "At least one owner is required",
                        },
                        {
                            "type": "TextBlock",
                            "text": "**Members** (comma-separated emails)",
                            "weight": "Bolder",
                            "spacing": "Medium",
                        },
                        {
                            "type": "Input.Text",
                            "id": "members",
                            "placeholder": "user1@company.com, user2@company.com",
                            "isMultiline": True,
                        },
                    ],
                },
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit Request",
                    "style": "positive",
                    "data": {
                        "action": "provisioning_submit_config",
                        "request_id": request_id,
                        "resource_type": "sharepoint_site",
                    },
                },
                {
                    "type": "Action.Submit",
                    "title": "Cancel",
                    "data": {
                        "action": "provisioning_cancel",
                        "request_id": request_id,
                    },
                },
            ],
        }

    def create_summary_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create summary card for the requester after config submission."""
        rtype = request.resource_type
        info = RESOURCE_TYPES.get(rtype, {})
        config = request.config
        domain = os.getenv("M365_GRAPH_DOMAIN", "yourcompany.com")

        facts = [
            {"title": "Resource:", "value": f"{info.get('icon', '')} {info.get('display', rtype)}"},
            {"title": "Name:", "value": config.get("display_name", "N/A")},
        ]

        if rtype == "shared_mailbox":
            facts.append({"title": "Email:", "value": f"{config.get('mail_nickname', '')}@{domain}"})
            facts.append({
                "title": "External senders:",
                "value": "Yes" if config.get("external_senders") == "yes" else "No",
            })

        if config.get("owners"):
            facts.append({"title": "Owners:", "value": ", ".join(config["owners"])})
        if config.get("members"):
            facts.append({"title": "Members:", "value": ", ".join(config["members"])})
        if config.get("visibility"):
            facts.append({"title": "Privacy:", "value": config["visibility"]})

        return {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "Container",
                    "style": "emphasis",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "📋 Request Submitted for Approval",
                            "weight": "Bolder",
                            "size": "Large",
                            "color": "Accent",
                        },
                        {
                            "type": "TextBlock",
                            "text": f"Request #{request.request_id}",
                            "isSubtle": True,
                        },
                    ],
                },
                {
                    "type": "Container",
                    "separator": True,
                    "spacing": "Medium",
                    "items": [
                        {"type": "FactSet", "facts": facts},
                    ],
                },
                {
                    "type": "TextBlock",
                    "text": (
                        "Your request has been sent to an IT Admin for approval. "
                        "You'll be notified once it's been reviewed."
                    ),
                    "wrap": True,
                    "spacing": "Large",
                    "isSubtle": True,
                    "size": "Small",
                },
            ],
        }

    def create_approval_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create approval card for the admin."""
        rtype = request.resource_type
        info = RESOURCE_TYPES.get(rtype, {})
        config = request.config
        domain = os.getenv("M365_GRAPH_DOMAIN", "yourcompany.com")

        facts = [
            {"title": "Requested by:", "value": f"{request.requester_name} ({request.requester_email})"},
            {"title": "Resource:", "value": f"{info.get('icon', '')} {info.get('display', rtype)}"},
            {"title": "Name:", "value": config.get("display_name", "N/A")},
        ]

        if rtype == "shared_mailbox":
            facts.append({"title": "Email:", "value": f"{config.get('mail_nickname', '')}@{domain}"})
            facts.append({
                "title": "External senders:",
                "value": "Yes" if config.get("external_senders") == "yes" else "No",
            })

        if config.get("owners"):
            facts.append({"title": "Owners:", "value": ", ".join(config["owners"])})
        if config.get("members"):
            facts.append({"title": "Members:", "value": ", ".join(config["members"])})
        if config.get("visibility"):
            facts.append({"title": "Privacy:", "value": config["visibility"]})
        if config.get("description"):
            facts.append({"title": "Description:", "value": config["description"][:100]})

        return {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "Container",
                    "style": "emphasis",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "🔔 M365 Provisioning Request",
                            "weight": "Bolder",
                            "size": "Large",
                        },
                        {
                            "type": "TextBlock",
                            "text": f"Request #{request.request_id}",
                            "isSubtle": True,
                        },
                    ],
                },
                {
                    "type": "Container",
                    "separator": True,
                    "spacing": "Medium",
                    "items": [
                        {"type": "FactSet", "facts": facts},
                    ],
                },
                {
                    "type": "TextBlock",
                    "text": f"**Original request:** \"{request.original_message[:200]}\"",
                    "wrap": True,
                    "isSubtle": True,
                    "spacing": "Medium",
                    "size": "Small",
                },
                {
                    "type": "TextBlock",
                    "text": "**Denial reason** (optional, only used if denied):",
                    "weight": "Bolder",
                    "spacing": "Medium",
                    "size": "Small",
                },
                {
                    "type": "Input.Text",
                    "id": "denial_reason",
                    "placeholder": "Reason for denial (optional)",
                },
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "✅ Approve",
                    "style": "positive",
                    "data": {
                        "action": "provisioning_approve",
                        "request_id": request.request_id,
                    },
                },
                {
                    "type": "Action.Submit",
                    "title": "❌ Deny",
                    "data": {
                        "action": "provisioning_deny",
                        "request_id": request.request_id,
                    },
                },
            ],
        }

    def create_result_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create result card showing provisioning outcome."""
        result = request.result or {}
        success = result.get("success", False)
        rtype = request.resource_type
        info = RESOURCE_TYPES.get(rtype, {})
        domain = os.getenv("M365_GRAPH_DOMAIN", "yourcompany.com")

        if success:
            facts = [
                {"title": "Resource:", "value": f"{info.get('icon', '')} {info.get('display', rtype)}"},
                {"title": "Name:", "value": result.get("display_name", request.config.get("display_name", "N/A"))},
                {"title": "Status:", "value": "✅ Created Successfully"},
            ]

            email = result.get("email", "")
            if email:
                facts.append({"title": "Email:", "value": email})

            sharepoint_url = result.get("sharepoint_url", "")
            if sharepoint_url:
                facts.append({"title": "SharePoint:", "value": sharepoint_url})

            return {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [
                    {
                        "type": "Container",
                        "style": "emphasis",
                        "items": [
                            {
                                "type": "ColumnSet",
                                "columns": [
                                    {
                                        "type": "Column",
                                        "width": "auto",
                                        "items": [{"type": "TextBlock", "text": "✅", "size": "ExtraLarge"}],
                                    },
                                    {
                                        "type": "Column",
                                        "width": "stretch",
                                        "items": [
                                            {
                                                "type": "TextBlock",
                                                "text": "Resource Created!",
                                                "weight": "Bolder",
                                                "size": "Large",
                                                "color": "Good",
                                            },
                                            {
                                                "type": "TextBlock",
                                                "text": f"Request #{request.request_id} has been provisioned",
                                                "isSubtle": True,
                                                "wrap": True,
                                            },
                                        ],
                                    },
                                ],
                            },
                        ],
                    },
                    {
                        "type": "Container",
                        "separator": True,
                        "spacing": "Medium",
                        "items": [{"type": "FactSet", "facts": facts}],
                    },
                    {
                        "type": "TextBlock",
                        "text": (
                            "The resource is now available. It may take a few minutes "
                            "for all features to become fully accessible."
                        ),
                        "wrap": True,
                        "isSubtle": True,
                        "spacing": "Large",
                        "size": "Small",
                    },
                ],
            }
        else:
            error_msg = result.get("error", "Unknown error occurred")
            return {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": "❌ Provisioning Failed",
                        "weight": "Bolder",
                        "size": "Large",
                        "color": "Attention",
                    },
                    {
                        "type": "TextBlock",
                        "text": f"Request #{request.request_id}",
                        "isSubtle": True,
                    },
                    {
                        "type": "TextBlock",
                        "text": f"**Error:** {error_msg}",
                        "wrap": True,
                        "spacing": "Medium",
                    },
                    {
                        "type": "TextBlock",
                        "text": "An IT Admin has been notified. They may need to provision this manually.",
                        "wrap": True,
                        "isSubtle": True,
                        "spacing": "Medium",
                        "size": "Small",
                    },
                ],
            }

    def create_denied_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create card notifying user their request was denied."""
        reason = request.denial_reason or "No reason provided."
        info = RESOURCE_TYPES.get(request.resource_type, {})

        return {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "❌ Request Denied",
                    "weight": "Bolder",
                    "size": "Large",
                    "color": "Attention",
                },
                {
                    "type": "TextBlock",
                    "text": f"Request #{request.request_id}",
                    "isSubtle": True,
                },
                {
                    "type": "FactSet",
                    "spacing": "Medium",
                    "facts": [
                        {"title": "Resource:", "value": f"{info.get('icon', '')} {info.get('display', '')}"},
                        {"title": "Name:", "value": request.config.get("display_name", "N/A")},
                        {"title": "Reason:", "value": reason},
                    ],
                },
                {
                    "type": "TextBlock",
                    "text": "If you have questions about this decision, please reach out to your IT Admin.",
                    "wrap": True,
                    "isSubtle": True,
                    "spacing": "Medium",
                    "size": "Small",
                },
            ],
        }

    async def execute(self, request: AutomationRequest) -> Dict[str, Any]:
        """Execute the provisioning via Graph API."""
        rtype = request.resource_type
        config = request.config

        if not self.graph_client.client_id or not self.graph_client.client_secret:
            return {
                "success": False,
                "error": (
                    "M365 Graph API credentials not configured. "
                    "Set M365_GRAPH_CLIENT_ID and M365_GRAPH_CLIENT_SECRET."
                ),
            }

        logging.info(
            f"Executing M365 provisioning: type={rtype}, "
            f"name={config.get('display_name')}, request={request.request_id}"
        )

        if rtype == "shared_mailbox":
            return await self.graph_client.provision_shared_mailbox(config)
        elif rtype == "teams_team":
            return await self.graph_client.provision_teams_team(config)
        elif rtype == "sharepoint_site":
            return await self.graph_client.provision_sharepoint_site(config)
        else:
            return {"success": False, "error": f"Unknown resource type: {rtype}"}


def parse_email_list(raw: str) -> List[str]:
    """Parse a comma/semicolon/space-separated list of emails."""
    if not raw:
        return []
    # Split on commas, semicolons, or whitespace
    parts = re.split(r'[,;\s]+', raw.strip())
    emails = [p.strip() for p in parts if '@' in p]
    return emails


def build_config_from_form(resource_type: str, form_data: Dict[str, Any]) -> Dict[str, Any]:
    """Build a clean config dict from submitted form data."""
    config = {
        "display_name": form_data.get("display_name", "").strip(),
        "description": form_data.get("description", "").strip(),
        "visibility": form_data.get("visibility", "Private"),
        "members": parse_email_list(form_data.get("members", "")),
        "owners": parse_email_list(form_data.get("owners", "")),
    }

    if resource_type == "shared_mailbox":
        nickname = form_data.get("mail_nickname", "").strip()
        # Clean the nickname - alphanumeric only
        config["mail_nickname"] = re.sub(r'[^a-zA-Z0-9]', '', nickname)
        config["external_senders"] = form_data.get("external_senders", "yes")
        # For shared mailbox, members are also owners by default
        if not config["owners"]:
            config["owners"] = config["members"]
    else:
        # Generate mail nickname from display name
        config["mail_nickname"] = re.sub(
            r'[^a-zA-Z0-9]', '',
            config["display_name"].replace(" ", "")
        ).lower()

    return config
