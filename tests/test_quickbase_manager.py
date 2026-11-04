import unittest
from unittest.mock import AsyncMock
from xml.etree import ElementTree

from quickbase_manager import QuickBaseManager


class QuickBaseManagerTests(unittest.IsolatedAsyncioTestCase):
    async def test_get_quickbase_user_id_reads_user_attribute(self):
        qb = QuickBaseManager()
        root = ElementTree.fromstring(
            """
            <qdbapi>
                <errcode>0</errcode>
                <errtext>No error</errtext>
                <user id="68707250.dhey">
                    <login>rdesai@dispatchenergy.com</login>
                </user>
            </qdbapi>
            """
        )
        qb.execute_legacy_api = AsyncMock(return_value=root)

        user_id = await qb.get_quickbase_user_id("rdesai@dispatchenergy.com")

        self.assertEqual(user_id, "68707250.dhey")

    async def test_ensure_app_user_in_role_treats_existing_role_as_success(self):
        qb = QuickBaseManager()
        qb.get_app_role_id = AsyncMock(return_value="10")
        qb.get_quickbase_user_id = AsyncMock(return_value="68707250.dhey")
        qb.execute_legacy_api = AsyncMock(
            return_value=ElementTree.fromstring(
                """
                <qdbapi>
                    <errcode>113</errcode>
                    <errtext>User is already in specified role.</errtext>
                </qdbapi>
                """
            )
        )

        result = await qb.ensure_app_user_in_role(
            email="rdesai@dispatchenergy.com",
            first_name="Roshni",
            last_name="Desai",
            app_id="btq7amgsn",
            role_name="Participant",
        )

        self.assertTrue(result["success"])
        self.assertEqual(result["user_id"], "68707250.dhey")
        self.assertFalse(result["provisioned"])

    async def test_ensure_ticket_requester_access_adds_user_without_role(self):
        qb = QuickBaseManager()
        qb.app_id = "bvajz826s"
        qb.get_quickbase_user_id = AsyncMock(return_value="68028209.dmbd")
        qb.get_app_role_id = AsyncMock(return_value="13")
        qb.execute_legacy_api = AsyncMock(
            side_effect=[
                ElementTree.fromstring(
                    """
                    <qdbapi>
                        <errcode>0</errcode>
                        <user id="68028209.dmbd"><name>Hank Fleury</name><roles></roles></user>
                    </qdbapi>
                    """
                ),
                ElementTree.fromstring("<qdbapi><errcode>0</errcode></qdbapi>"),
            ]
        )

        result = await qb.ensure_ticket_requester_access("HFleury@dispatchenergy.com")

        self.assertTrue(result["success"])
        self.assertEqual(result["role_id"], "13")
        add_call = qb.execute_legacy_api.await_args_list[1]
        self.assertEqual(add_call.args[0], "API_AddUserToRole")
        self.assertEqual(add_call.args[2], {"userid": "68028209.dmbd", "roleid": "13"})

        # Second ticket from the same user doesn't hit QuickBase again
        await qb.ensure_ticket_requester_access("hfleury@dispatchenergy.com")
        self.assertEqual(qb.execute_legacy_api.await_count, 2)

    async def test_ensure_ticket_requester_access_leaves_existing_roles_alone(self):
        qb = QuickBaseManager()
        qb.app_id = "bvajz826s"
        qb.get_quickbase_user_id = AsyncMock(return_value="67761938.vpje")
        qb.get_app_role_id = AsyncMock(return_value="13")
        qb.execute_legacy_api = AsyncMock(
            return_value=ElementTree.fromstring(
                """
                <qdbapi>
                    <errcode>0</errcode>
                    <user id="67761938.vpje">
                        <roles><role id="12"><name>Administrator</name></role></roles>
                    </user>
                </qdbapi>
                """
            )
        )

        result = await qb.ensure_ticket_requester_access("jheisler@dispatchenergy.com")

        self.assertTrue(result["success"])
        self.assertEqual(result["skipped"], "has role")
        self.assertEqual(qb.execute_legacy_api.await_count, 1)
        qb.get_app_role_id.assert_not_awaited()

    async def test_ensure_ticket_requester_access_skips_unknown_user(self):
        qb = QuickBaseManager()
        qb.get_quickbase_user_id = AsyncMock(return_value=None)
        qb.execute_legacy_api = AsyncMock()

        result = await qb.ensure_ticket_requester_access("contractor@example.com")

        self.assertFalse(result["success"])
        qb.execute_legacy_api.assert_not_awaited()


if __name__ == "__main__":
    unittest.main()
