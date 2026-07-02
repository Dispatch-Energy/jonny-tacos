"""
Unit tests for QuickBase integration
"""
import asyncio
import unittest
from unittest.mock import AsyncMock, MagicMock, patch


class TestQuickBaseManager(unittest.TestCase):
    """Test QuickBase manager functionality"""

    def _make_manager(self):
        with patch.dict('os.environ', {
            'QB_REALM': 'test.quickbase.com',
            'QB_USER_TOKEN': 'token',
            'QB_APP_ID': 'app1',
            'QB_APP_TOKEN': 'tok',
            'QB_TICKETS_TABLE_ID': 'tbl1',
        }):
            from quickbase_manager import QuickBaseManager
            return QuickBaseManager()

    def _run(self, coro):
        return asyncio.get_event_loop().run_until_complete(coro)

    def test_create_ticket_blocks_empty_subject(self):
        mgr = self._make_manager()
        result = self._run(mgr.create_ticket({'subject': '', 'description': 'Some description'}))
        self.assertIsNone(result)

    def test_create_ticket_blocks_empty_description(self):
        mgr = self._make_manager()
        result = self._run(mgr.create_ticket({'subject': 'My issue', 'description': ''}))
        self.assertIsNone(result)

    def test_create_ticket_blocks_whitespace_only(self):
        mgr = self._make_manager()
        result = self._run(mgr.create_ticket({'subject': '   ', 'description': '\t\n'}))
        self.assertIsNone(result)

    def test_create_ticket_blocks_missing_fields(self):
        mgr = self._make_manager()
        result = self._run(mgr.create_ticket({}))
        self.assertIsNone(result)

    def test_record_id_field_is_selected(self):
        """Field 3 (Record ID#) must be mapped so queries select it and updates
        resolve a real key instead of None."""
        mgr = self._make_manager()
        self.assertEqual(mgr.field_mapping.get('record_id'), 3)

    def test_update_ticket_refuses_when_no_record_id(self):
        """Regression: an update with no resolved record_id must NOT POST —
        QuickBase's /records upsert would insert a blank ticket."""
        mgr = self._make_manager()
        mgr.get_ticket_by_reference = AsyncMock(
            return_value={'ticket_number': 'IT-1', 'record_id': None}
        )
        mgr.execute_request = AsyncMock()

        result = self._run(mgr.update_ticket({'ticket_id': 'IT-1', 'status': 'Closed'}))

        self.assertFalse(result)
        mgr.execute_request.assert_not_awaited()

    def test_update_ticket_posts_real_record_id(self):
        """A resolved record_id is sent as the field-3 key, producing an update."""
        mgr = self._make_manager()
        mgr.get_ticket_by_reference = AsyncMock(
            return_value={'ticket_number': 'IT-1', 'record_id': 42}
        )
        mgr.execute_request = AsyncMock(return_value={'metadata': {'updatedRecordIds': [42]}})

        result = self._run(mgr.update_ticket({'ticket_id': 'IT-1', 'status': 'Closed'}))

        self.assertTrue(result)
        _, url, payload = mgr.execute_request.await_args.args
        self.assertTrue(url.endswith('/records'))
        self.assertEqual(payload['data'][0]['3'], {'value': 42})


if __name__ == '__main__':
    unittest.main()
