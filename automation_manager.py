"""
IT Automation Manager - Generic framework for IT automations with admin approval.

Future-proofed for adding more automation types beyond M365 provisioning.
To add a new automation:
1. Create a class implementing AutomationHandler
2. Register it with AutomationManager.register_handler()
3. Update the router prompt in support_chain.py to detect the new automation type
4. Add invoke action handling in function_app.py
"""

import os
import logging
import time
import uuid
from typing import Dict, Any, Optional, List
from abc import ABC, abstractmethod
from enum import Enum


class AutomationStatus(str, Enum):
    GATHERING_INFO = "gathering_info"
    PENDING_APPROVAL = "pending_approval"
    APPROVED = "approved"
    DENIED = "denied"
    EXECUTING = "executing"
    COMPLETED = "completed"
    FAILED = "failed"


class AutomationRequest:
    """Tracks state for a single automation request through its lifecycle."""

    def __init__(self, request_id: str, automation_type: str,
                 requester_email: str, requester_name: str = ""):
        self.request_id = request_id
        self.automation_type = automation_type
        self.requester_email = requester_email
        self.requester_name = requester_name
        self.status = AutomationStatus.GATHERING_INFO
        self.config: Dict[str, Any] = {}
        self.resource_type: Optional[str] = None
        self.ai_extracted: Dict[str, Any] = {}
        self.original_message: str = ""
        self.created_at = time.time()
        self.updated_at = time.time()
        self.result: Optional[Dict[str, Any]] = None
        self.denial_reason: Optional[str] = None


class AutomationHandler(ABC):
    """Base class for specific automation implementations."""

    @property
    @abstractmethod
    def automation_type(self) -> str:
        """Unique identifier, e.g. 'm365_provisioning'"""
        pass

    @property
    @abstractmethod
    def display_name(self) -> str:
        """Human-readable name, e.g. 'M365 Resource Provisioning'"""
        pass

    @abstractmethod
    def detect_intent(self, message: str) -> Optional[Dict[str, Any]]:
        """Check if message matches this automation. Returns extracted details or None."""
        pass

    @abstractmethod
    def create_routing_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create the initial routing/question card for the user."""
        pass

    @abstractmethod
    def create_config_form(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create configuration form based on selected resource type."""
        pass

    @abstractmethod
    def create_approval_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create approval card for the admin."""
        pass

    @abstractmethod
    def create_result_card(self, request: AutomationRequest) -> Dict[str, Any]:
        """Create result card showing success/failure."""
        pass

    @abstractmethod
    async def execute(self, request: AutomationRequest) -> Dict[str, Any]:
        """Execute the automation after approval. Returns result dict."""
        pass


class AutomationManager:
    """Manages automation requests, state tracking, and admin approval flow.

    Usage:
        manager = AutomationManager()
        manager.register_handler(M365ProvisioningHandler())

        # When router classifies as automation_request:
        detection = manager.detect_automation(message)
        if detection:
            request = manager.create_request(...)
            card = handler.create_routing_card(request)

        # When admin approves:
        result = await manager.approve_and_execute(request_id)
    """

    REQUEST_TTL_SECONDS = 7200  # 2 hours

    def __init__(self):
        self._handlers: Dict[str, AutomationHandler] = {}
        self._requests: Dict[str, AutomationRequest] = {}
        self.admin_email = os.getenv("AUTOMATION_ADMIN_EMAIL", "")

    def register_handler(self, handler: AutomationHandler) -> None:
        self._handlers[handler.automation_type] = handler
        logging.info(f"Registered automation handler: {handler.automation_type}")

    def detect_automation(self, message: str) -> Optional[Dict[str, Any]]:
        """Check all handlers to find which automation matches this message.
        Returns {automation_type, extracted} or None.
        """
        self._prune_old_requests()

        for handler_type, handler in self._handlers.items():
            try:
                result = handler.detect_intent(message)
                if result:
                    return {
                        "automation_type": handler_type,
                        "extracted": result
                    }
            except Exception as e:
                logging.error(f"Error in {handler_type}.detect_intent: {e}")

        return None

    def create_request(self, automation_type: str, requester_email: str,
                       requester_name: str = "", extracted: Dict = None,
                       original_message: str = "") -> AutomationRequest:
        request_id = str(uuid.uuid4())[:8]
        request = AutomationRequest(
            request_id=request_id,
            automation_type=automation_type,
            requester_email=requester_email,
            requester_name=requester_name
        )
        if extracted:
            request.ai_extracted = extracted
        request.original_message = original_message
        self._requests[request_id] = request
        logging.info(
            f"Created automation request {request_id} "
            f"type={automation_type} for {requester_email}"
        )
        return request

    def get_request(self, request_id: str) -> Optional[AutomationRequest]:
        return self._requests.get(request_id)

    def get_handler(self, automation_type: str) -> Optional[AutomationHandler]:
        return self._handlers.get(automation_type)

    def get_active_request(self, user_email: str) -> Optional[AutomationRequest]:
        """Get the most recent active request for a user."""
        for req in reversed(list(self._requests.values())):
            if (req.requester_email.lower() == user_email.lower() and
                    req.status in (AutomationStatus.GATHERING_INFO,
                                   AutomationStatus.PENDING_APPROVAL)):
                return req
        return None

    async def approve_and_execute(self, request_id: str) -> Dict[str, Any]:
        """Approve a request and execute the automation."""
        request = self._requests.get(request_id)
        if not request:
            return {"success": False, "error": "Request not found"}

        handler = self._handlers.get(request.automation_type)
        if not handler:
            return {"success": False, "error": f"No handler for {request.automation_type}"}

        request.status = AutomationStatus.EXECUTING
        request.updated_at = time.time()

        try:
            result = await handler.execute(request)
            request.result = result
            request.status = (
                AutomationStatus.COMPLETED if result.get("success")
                else AutomationStatus.FAILED
            )
        except Exception as e:
            logging.error(f"Automation execution failed for {request_id}: {e}")
            request.result = {"success": False, "error": str(e)}
            request.status = AutomationStatus.FAILED

        request.updated_at = time.time()
        return request.result

    def deny_request(self, request_id: str, reason: str = "") -> Optional[AutomationRequest]:
        request = self._requests.get(request_id)
        if request:
            request.status = AutomationStatus.DENIED
            request.denial_reason = reason
            request.updated_at = time.time()
        return request

    def _prune_old_requests(self):
        cutoff = time.time() - self.REQUEST_TTL_SECONDS
        expired = [
            rid for rid, req in self._requests.items()
            if req.created_at < cutoff
        ]
        for rid in expired:
            del self._requests[rid]
