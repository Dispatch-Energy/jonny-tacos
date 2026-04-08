import unittest
from datetime import datetime, timedelta, timezone

from function_app import (
    create_user_creation_approval_card,
    create_user_creation_confirmation_card,
    create_solution_card,
    extract_webhook_ticket_data,
    get_follow_up_candidate_tickets,
    is_explicit_ticket_request,
    normalize_webhook_ticket_data,
    should_auto_create_ticket,
)


class FunctionAppTests(unittest.TestCase):
    def test_quick_fix_still_creates_ticket(self):
        self.assertTrue(
            should_auto_create_ticket(
                question="My VPN is being weird",
                needs_human=False,
                confidence=0.85,
            )
        )

    def test_human_action_request_auto_creates_ticket(self):
        self.assertTrue(
            should_auto_create_ticket(
                question="Please create a shared mailbox",
                needs_human=True,
                confidence=0.85,
            )
        )

    def test_explicit_ticket_request_auto_creates_ticket(self):
        self.assertTrue(
            should_auto_create_ticket(
                question="Please open a ticket for my laptop issue",
                needs_human=False,
                confidence=0.85,
            )
        )

    def test_i_do_want_ticket_counts_as_explicit_ticket_request(self):
        self.assertTrue(is_explicit_ticket_request("I do want a ticket for this"))
        self.assertTrue(is_explicit_ticket_request("I want ticket"))

    def test_follow_up_candidates_ignore_old_or_closed_tickets(self):
        recent_ticket = {
            "ticket_number": "IT-1001",
            "status": "In Progress",
            "submitted_date": (datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=1)).isoformat(),
        }
        old_ticket = {
            "ticket_number": "IT-1002",
            "status": "In Progress",
            "submitted_date": (datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=30)).isoformat(),
        }
        closed_ticket = {
            "ticket_number": "IT-1003",
            "status": "Closed",
            "submitted_date": (datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=1)).isoformat(),
        }

        candidates = get_follow_up_candidate_tickets([recent_ticket, old_ticket, closed_ticket])

        self.assertEqual([ticket["ticket_number"] for ticket in candidates], ["IT-1001"])

    def test_solution_card_shows_ticket_context_when_ticket_created(self):
        card = create_solution_card(
            solution="Try restarting Teams and clearing the cache.",
            question="Teams keeps freezing",
            category="Teams/Office 365",
            confidence=0.8,
            offer_escalate=True,
            ticket_number="IT-1234",
            ticket_state="created",
        )

        body_text = " ".join(
            item.get("text", "")
            for item in card["body"]
            if item.get("type") == "TextBlock"
        )
        actions = [action["title"] for action in card["actions"]]

        self.assertIn("📋 Check Ticket", actions)
        self.assertIn("🛠️ Need IT follow-up", actions)
        self.assertIn("Here's what I'd try next", body_text)

        emphasis_items = [
            item for item in card["body"]
            if item.get("type") == "Container" and item.get("style") == "emphasis"
        ]
        self.assertTrue(any("Ticket opened: IT-1234" in str(item) for item in emphasis_items))

    def test_solution_card_shows_no_ticket_yet_state_on_failure_path(self):
        card = create_solution_card(
            solution="Restart Outlook and test again.",
            question="Outlook is slow",
            category="Email Issues",
            confidence=0.9,
            offer_escalate=True,
            ticket_state="required_failed",
        )

        actions = [action["title"] for action in card["actions"]]
        attention_items = [
            item for item in card["body"]
            if item.get("type") == "Container" and item.get("style") == "attention"
        ]

        self.assertIn("🎫 Still need help", actions)
        self.assertTrue(any("couldn't open the ticket automatically" in str(item) for item in attention_items))

    def test_extract_and_normalize_webhook_payload(self):
        payload = {
            "data": [{
                "Ticket Number": "IT-5555",
                "Submitted By": "user@example.com",
                "Previous Status": "New",
                "Status": "In Progress",
                "Description": "Diana,Boylan;diana.boylan@gmail.com",
                "Due Date": "2025-01-15",
            }]
        }

        extracted = extract_webhook_ticket_data(payload)
        normalized = normalize_webhook_ticket_data(extracted)

        self.assertEqual(normalized["ticket_number"], "IT-5555")
        self.assertEqual(normalized["submitted_by"], "user@example.com")
        self.assertEqual(normalized["old_status"], "New")
        self.assertEqual(normalized["status"], "In Progress")
        self.assertEqual(normalized["description"], "Diana,Boylan;diana.boylan@gmail.com")
        self.assertEqual(normalized["due_date"], "2025-01-15")

    def test_user_creation_confirmation_card_includes_ticket_details(self):
        card = create_user_creation_confirmation_card(
            {
                "ticket_number": "IT-5555",
                "submitted_by": "requester@example.com",
                "quickbase_url": "https://example.quickbase.com/db/abc?a=dr&rid=1",
            },
            {
                "display_name": "Diana Boylan",
                "user_principal_name": "dboylan@example.com",
                "recipient_email": "diana.boylan@gmail.com",
                "email_queued_for": "2025-01-15T10:00:00+00:00",
                "openai_invite": {"success": True},
                "anthropic_invite": {"success": True},
            }
        )

        facts = []
        for item in card["body"]:
            if item.get("type") == "Container":
                for child in item.get("items", []):
                    if child.get("type") == "FactSet":
                        facts.extend(child.get("facts", []))

        self.assertIn("User Creation Completed", str(card))
        self.assertTrue(any(fact["value"] == "IT-5555" for fact in facts))
        self.assertTrue(any(fact["value"] == "dboylan@example.com" for fact in facts))
        self.assertTrue(any(fact["title"] == "Credential Recipient:" and fact["value"] == "diana.boylan@gmail.com" for fact in facts))
        self.assertTrue(any(fact["title"] == "OpenAI:" and fact["value"] == "Invited" for fact in facts))
        self.assertTrue(any(fact["title"] == "Claude:" and fact["value"] == "Invited" for fact in facts))
        self.assertEqual(card["actions"][0]["title"], "View in QuickBase")

    def test_user_creation_approval_card_prefills_predicted_username(self):
        card = create_user_creation_approval_card(
            {
                "ticket_number": "IT-7777",
                "submitted_by": "requester@example.com",
                "due_date": "2025-01-15",
            },
            {
                "request_id": "IT-7777",
                "display_name": "Diana Boylan",
                "personal_email": "diana.boylan@gmail.com",
                "predicted_username_local": "dboylan",
                "predicted_user_principal_name": "dboylan@example.com",
            }
        )

        self.assertIn("Confirm User Creation", str(card))
        self.assertEqual(card["body"][1]["items"][2]["value"], "dboylan")
        self.assertEqual(card["actions"][0]["data"]["action"], "user_creation_approve")


if __name__ == "__main__":
    unittest.main()
