import unittest

from function_app import (
    create_solution_card,
    extract_webhook_ticket_data,
    normalize_webhook_ticket_data,
    should_auto_create_ticket,
)


class FunctionAppTests(unittest.TestCase):
    def test_quick_fix_does_not_auto_create_ticket(self):
        self.assertFalse(
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

    def test_solution_card_shows_ticket_context_when_ticket_created(self):
        card = create_solution_card(
            solution="Try restarting Teams and clearing the cache.",
            question="Teams keeps freezing",
            category="Teams/Office 365",
            confidence=0.8,
            offer_escalate=False,
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
        self.assertNotIn("🎫 Still need help", actions)
        self.assertIn("Here's what I'd try next", body_text)

        emphasis_items = [
            item for item in card["body"]
            if item.get("type") == "Container" and item.get("style") == "emphasis"
        ]
        self.assertTrue(any("Ticket opened: IT-1234" in str(item) for item in emphasis_items))

    def test_solution_card_shows_no_ticket_yet_state(self):
        card = create_solution_card(
            solution="Restart Outlook and test again.",
            question="Outlook is slow",
            category="Email Issues",
            confidence=0.9,
            offer_escalate=True,
            ticket_state="not_created",
        )

        actions = [action["title"] for action in card["actions"]]
        emphasis_items = [
            item for item in card["body"]
            if item.get("type") == "Container" and item.get("style") == "emphasis"
        ]

        self.assertIn("🎫 Still need help", actions)
        self.assertTrue(any("No ticket yet" in str(item) for item in emphasis_items))

    def test_extract_and_normalize_webhook_payload(self):
        payload = {
            "data": [{
                "Ticket Number": "IT-5555",
                "Submitted By": "user@example.com",
                "Previous Status": "New",
                "Status": "In Progress",
            }]
        }

        extracted = extract_webhook_ticket_data(payload)
        normalized = normalize_webhook_ticket_data(extracted)

        self.assertEqual(normalized["ticket_number"], "IT-5555")
        self.assertEqual(normalized["submitted_by"], "user@example.com")
        self.assertEqual(normalized["old_status"], "New")
        self.assertEqual(normalized["status"], "In Progress")


if __name__ == "__main__":
    unittest.main()
