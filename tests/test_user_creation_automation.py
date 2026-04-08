import unittest
from datetime import timezone

from user_creation_automation import (
    build_username_local_part,
    calculate_due_date_send_at,
    extract_personal_email,
    extract_first_last_name,
    extract_user_creation_details,
    get_user_creation_source_text,
    is_user_creation_category,
)


class UserCreationAutomationTests(unittest.TestCase):
    def test_extract_first_and_last_name_from_labeled_fields(self):
        first_name, last_name = extract_first_last_name(
            "First Name: Diana\nLast Name: Boylan\nDepartment: Operations"
        )

        self.assertEqual(first_name, "Diana")
        self.assertEqual(last_name, "Boylan")

    def test_extract_first_and_last_name_from_onboarding_phrase(self):
        first_name, last_name = extract_first_last_name(
            '"on boarding for:" Michael Zayas'
        )

        self.assertEqual(first_name, "Michael")
        self.assertEqual(last_name, "Zayas")

    def test_extract_first_and_last_name_from_camel_case(self):
        first_name, last_name = extract_first_last_name("on boarding for: DianaBoylan")

        self.assertEqual(first_name, "Diana")
        self.assertEqual(last_name, "Boylan")

    def test_extract_user_creation_details_from_compact_quickbase_format(self):
        first_name, last_name, personal_email = extract_user_creation_details(
            "Diana,Boylan;diana.boylan@gmail.com"
        )

        self.assertEqual(first_name, "Diana")
        self.assertEqual(last_name, "Boylan")
        self.assertEqual(personal_email, "diana.boylan@gmail.com")

    def test_extract_personal_email_from_labeled_format(self):
        self.assertEqual(
            extract_personal_email("First Name: Diana\nLast Name: Boylan\nEmail: diana@example.com"),
            "diana@example.com",
        )

    def test_get_user_creation_source_text_prefers_subject_when_it_has_email(self):
        source_text = get_user_creation_source_text({
            "subject": "\"on boarding for:\" Aleksandar,Aleksic;alek.aleksich@gmail.com",
            "description": "AleksandarAleksic",
        })

        self.assertEqual(
            source_text,
            "\"on boarding for:\" Aleksandar,Aleksic;alek.aleksich@gmail.com",
        )

    def test_build_username_local_part_uses_first_initial_and_last_name(self):
        self.assertEqual(build_username_local_part("Diana", "Boylan"), "dboylan")

    def test_build_username_local_part_normalizes_accents(self):
        self.assertEqual(build_username_local_part("Jose", "Nuñez"), "jnunez")

    def test_calculate_due_date_send_at_uses_5am_local_time(self):
        send_at = calculate_due_date_send_at("2025-01-15", "America/New_York")

        self.assertEqual(send_at.tzinfo, timezone.utc)
        self.assertEqual(send_at.isoformat(), "2025-01-15T10:00:00+00:00")

    def test_is_user_creation_category_is_case_insensitive(self):
        self.assertTrue(is_user_creation_category("User Creation"))
        self.assertTrue(is_user_creation_category("user creation"))
        self.assertFalse(is_user_creation_category("Teams/Office 365"))


if __name__ == "__main__":
    unittest.main()
