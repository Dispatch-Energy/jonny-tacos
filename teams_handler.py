"""
Teams Handler - Manages Teams bot interactions and messaging
"""

import os
import json
import logging
import requests
import asyncio
import re
from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
import uuid

try:
    from azure.storage.blob import BlobServiceClient
except ImportError:
    BlobServiceClient = None

class TeamsHandler:
    def __init__(self):
        """Initialize Teams bot handler"""
        self.app_id = os.environ.get('TEAMS_APP_ID', '')
        self.app_secret = os.environ.get('TEAMS_APP_SECRET', '')
        self.tenant_id = os.environ.get('TEAMS_TENANT_ID', '')
        self.service_url = "https://smba.trafficmanager.net/amer/"
        
        self.executor = ThreadPoolExecutor(max_workers=3)
        self._token = None
        self._token_expiry = None
        self._graph_token = None
        self._graph_token_expiry = None
        self.storage_connection_string = (
            os.environ.get('TEAMS_STORAGE_CONNECTION_STRING', '')
            or os.environ.get('AzureWebJobsStorage', '')
            or os.environ.get('AZURE_STORAGE_CONNECTION_STRING', '')
        )
        self.conversation_container = os.environ.get(
            'TEAMS_CONVERSATION_CONTAINER',
            'teams-conversation-references'
        )
        self.conversation_fallback_dir = os.environ.get(
            'TEAMS_CONVERSATION_FALLBACK_DIR',
            os.path.join(os.getcwd(), '.local', 'teams-conversation-references')
        )
        self._blob_service_client = None
        self._conversation_container_ready = False

    async def get_auth_token(self) -> str:
        """
        Get or refresh authentication token for Teams bot
        """
        try:
            # Check if we have a valid token
            if self._token and self._token_expiry and datetime.now() < self._token_expiry:
                return self._token
            
            # Get new token
            token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            
            data = {
                'grant_type': 'client_credentials',
                'client_id': self.app_id,
                'client_secret': self.app_secret,
                'scope': 'https://api.botframework.com/.default'
            }
            
            loop = asyncio.get_event_loop()
            
            def get_token():
                response = requests.post(token_url, data=data)
                if response.status_code == 200:
                    token_data = response.json()
                    return token_data.get('access_token'), token_data.get('expires_in', 3600)
                return None, None
            
            token, expires_in = await loop.run_in_executor(self.executor, get_token)
            
            if token:
                self._token = token
                self._token_expiry = datetime.now() + timedelta(seconds=expires_in - 60)
                return token
            
            logging.error("Failed to get Teams auth token")
            return ""
            
        except Exception as e:
            logging.error(f"Error getting auth token: {str(e)}")
            return ""

    async def get_graph_token(self) -> str:
        """Get or refresh Microsoft Graph API token for user lookups."""
        try:
            if self._graph_token and self._graph_token_expiry and datetime.now() < self._graph_token_expiry:
                return self._graph_token

            token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            data = {
                'grant_type': 'client_credentials',
                'client_id': self.app_id,
                'client_secret': self.app_secret,
                'scope': 'https://graph.microsoft.com/.default'
            }

            loop = asyncio.get_event_loop()

            def get_token():
                response = requests.post(token_url, data=data)
                if response.status_code == 200:
                    token_data = response.json()
                    return token_data.get('access_token'), token_data.get('expires_in', 3600)
                logging.error(f"Failed to get Graph token: {response.status_code} - {response.text}")
                return None, None

            token, expires_in = await loop.run_in_executor(self.executor, get_token)

            if token:
                self._graph_token = token
                self._graph_token_expiry = datetime.now() + timedelta(seconds=expires_in - 60)
                return token

            logging.error("Failed to get Graph API token")
            return ""

        except Exception as e:
            logging.error(f"Error getting Graph token: {str(e)}")
            return ""

    async def get_user_aad_id(self, user_email: str) -> Optional[str]:
        """Resolve a user's email (UPN) to their Azure AD Object ID via Graph API."""
        try:
            graph_token = await self.get_graph_token()
            if not graph_token:
                return None

            headers = {
                'Authorization': f'Bearer {graph_token}',
                'Content-Type': 'application/json'
            }

            # Use /users/{upn} endpoint to get user's AAD ID
            graph_url = f"https://graph.microsoft.com/v1.0/users/{user_email}?$select=id"

            loop = asyncio.get_event_loop()

            def lookup_user():
                response = requests.get(graph_url, headers=headers, timeout=10)
                if response.status_code == 200:
                    return response.json().get('id')
                logging.error(f"Graph user lookup failed for {user_email}: {response.status_code} - {response.text}")
                return None

            aad_id = await loop.run_in_executor(self.executor, lookup_user)
            if aad_id:
                logging.info(f"Resolved {user_email} to AAD ID: {aad_id}")
            return aad_id

        except Exception as e:
            logging.error(f"Error looking up AAD ID for {user_email}: {str(e)}")
            return None

    async def send_message(self, activity: Dict[str, Any], message: str) -> bool:
        """
        Send a text message as a reply
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return False
            
            service_url = activity.get('serviceUrl', self.service_url)
            conversation_id = activity.get('conversation', {}).get('id')
            activity_id = activity.get('id')
            
            reply_url = f"{service_url}v3/conversations/{conversation_id}/activities/{activity_id}"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            reply_activity = {
                'type': 'message',
                'from': activity.get('recipient'),
                'conversation': activity.get('conversation'),
                'recipient': activity.get('from'),
                'text': message,
                'replyToId': activity_id
            }
            
            loop = asyncio.get_event_loop()
            
            def send():
                response = requests.post(reply_url, headers=headers, json=reply_activity)
                return response.status_code in [200, 201, 202]
            
            return await loop.run_in_executor(self.executor, send)
            
        except Exception as e:
            logging.error(f"Error sending message: {str(e)}")
            return False

    async def send_card(self, activity: Dict[str, Any], card: Dict[str, Any]) -> bool:
        """
        Send an Adaptive Card as a reply
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return False
            
            service_url = activity.get('serviceUrl', self.service_url)
            conversation_id = activity.get('conversation', {}).get('id')
            activity_id = activity.get('id')
            
            reply_url = f"{service_url}v3/conversations/{conversation_id}/activities/{activity_id}"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            reply_activity = {
                'type': 'message',
                'from': activity.get('recipient'),
                'conversation': activity.get('conversation'),
                'recipient': activity.get('from'),
                'attachments': [{
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': card
                }],
                'replyToId': activity_id
            }
            
            loop = asyncio.get_event_loop()
            
            def send():
                response = requests.post(reply_url, headers=headers, json=reply_activity)
                return response.status_code in [200, 201, 202]
            
            return await loop.run_in_executor(self.executor, send)
            
        except Exception as e:
            logging.error(f"Error sending card: {str(e)}")
            return False

    async def update_card(self, activity: Dict[str, Any], card: Dict[str, Any]) -> bool:
        """
        Update an existing Adaptive Card
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return False
            
            service_url = activity.get('serviceUrl', self.service_url)
            conversation_id = activity.get('conversation', {}).get('id')
            activity_id = activity.get('replyToId', activity.get('id'))
            
            update_url = f"{service_url}v3/conversations/{conversation_id}/activities/{activity_id}"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            updated_activity = {
                'type': 'message',
                'attachments': [{
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': card
                }]
            }
            
            loop = asyncio.get_event_loop()
            
            def update():
                response = requests.put(update_url, headers=headers, json=updated_activity)
                return response.status_code in [200, 201, 202]
            
            return await loop.run_in_executor(self.executor, update)
            
        except Exception as e:
            logging.error(f"Error updating card: {str(e)}")
            return False

    async def send_to_channel(self, channel_id: str, card: Dict[str, Any]) -> bool:
        """
        Send a proactive message to a Teams channel
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return False
            
            # Create conversation reference for proactive messaging
            service_url = self.service_url
            
            # First, create a conversation
            create_conv_url = f"{service_url}v3/conversations"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            conversation_params = {
                'bot': {
                    'id': self.app_id,
                    'name': 'IT Support Bot'
                },
                'isGroup': True,
                'channelData': {
                    'teamsChannelId': channel_id,
                    'tenant': {
                        'id': self.tenant_id
                    }
                },
                'activity': {
                    'type': 'message',
                    'attachments': [{
                        'contentType': 'application/vnd.microsoft.card.adaptive',
                        'content': card
                    }]
                }
            }
            
            loop = asyncio.get_event_loop()
            
            def send():
                response = requests.post(create_conv_url, headers=headers, json=conversation_params)
                return response.status_code in [200, 201, 202]
            
            return await loop.run_in_executor(self.executor, send)
            
        except Exception as e:
            logging.error(f"Error sending to channel: {str(e)}")
            return False

    async def send_typing_indicator(self, activity: Dict[str, Any]) -> bool:
        """
        Send typing indicator to show bot is processing
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return False
            
            service_url = activity.get('serviceUrl', self.service_url)
            conversation_id = activity.get('conversation', {}).get('id')
            
            typing_url = f"{service_url}v3/conversations/{conversation_id}/activities"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            typing_activity = {
                'type': 'typing',
                'from': activity.get('recipient')
            }
            
            loop = asyncio.get_event_loop()
            
            def send():
                response = requests.post(typing_url, headers=headers, json=typing_activity)
                return response.status_code in [200, 201, 202]
            
            return await loop.run_in_executor(self.executor, send)
            
        except Exception as e:
            logging.error(f"Error sending typing indicator: {str(e)}")
            return False

    def remove_mentions(self, text: str) -> str:
        """
        Remove bot mentions from message text
        """
        try:
            # Remove <at>bot name</at> mentions
            import re
            cleaned = re.sub(r'<at>.*?</at>', '', text)
            return cleaned.strip()
        except Exception as e:
            logging.error(f"Error removing mentions: {str(e)}")
            return text

    async def get_user_info(self, activity: Dict[str, Any], user_id: str) -> Optional[Dict[str, Any]]:
        """
        Get user information from Teams
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return None
            
            service_url = activity.get('serviceUrl', self.service_url)
            conversation_id = activity.get('conversation', {}).get('id')
            
            members_url = f"{service_url}v3/conversations/{conversation_id}/members/{user_id}"
            
            headers = {
                'Authorization': f'Bearer {token}'
            }
            
            loop = asyncio.get_event_loop()
            
            def get_user():
                response = requests.get(members_url, headers=headers)
                if response.status_code == 200:
                    return response.json()
                return None
            
            return await loop.run_in_executor(self.executor, get_user)
            
        except Exception as e:
            logging.error(f"Error getting user info: {str(e)}")
            return None

    async def get_channel_members(self, activity: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Get all members of a Teams channel
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return []
            
            service_url = activity.get('serviceUrl', self.service_url)
            conversation_id = activity.get('conversation', {}).get('id')
            
            members_url = f"{service_url}v3/conversations/{conversation_id}/members"
            
            headers = {
                'Authorization': f'Bearer {token}'
            }
            
            loop = asyncio.get_event_loop()
            
            def get_members():
                response = requests.get(members_url, headers=headers)
                if response.status_code == 200:
                    return response.json()
                return []
            
            return await loop.run_in_executor(self.executor, get_members)
            
        except Exception as e:
            logging.error(f"Error getting channel members: {str(e)}")
            return []

    def create_conversation_reference(self, activity: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create a conversation reference for proactive messaging
        """
        return {
            'activityId': activity.get('id'),
            'user': activity.get('from'),
            'bot': activity.get('recipient'),
            'conversation': activity.get('conversation'),
            'channelId': 'msteams',
            'serviceUrl': activity.get('serviceUrl', self.service_url)
        }

    def is_personal_conversation(self, activity: Dict[str, Any]) -> bool:
        """Return True when the activity is from a 1:1 chat with the bot."""
        conversation = activity.get('conversation', {})
        conversation_type = (conversation.get('conversationType') or '').lower()
        return conversation_type == 'personal'

    def _conversation_blob_name(self, user_email: str) -> str:
        safe_email = re.sub(r'[^a-z0-9._-]+', '_', user_email.lower())
        return f"{safe_email}.json"

    def _conversation_fallback_path(self, user_email: str) -> str:
        return os.path.join(
            self.conversation_fallback_dir,
            self._conversation_blob_name(user_email)
        )

    async def _get_conversation_container_client(self):
        """Get a blob container client for persisted conversation references."""
        if not BlobServiceClient or not self.storage_connection_string:
            return None

        if self._blob_service_client is None:
            try:
                self._blob_service_client = BlobServiceClient.from_connection_string(
                    self.storage_connection_string
                )
            except Exception as e:
                logging.error(f"Failed to initialize blob storage client: {str(e)}")
                return None

        container_client = self._blob_service_client.get_container_client(
            self.conversation_container
        )

        if not self._conversation_container_ready:
            loop = asyncio.get_event_loop()

            def ensure_container():
                try:
                    container_client.create_container()
                except Exception:
                    pass
                return True

            await loop.run_in_executor(self.executor, ensure_container)
            self._conversation_container_ready = True

        return container_client

    async def store_conversation_reference(self, activity: Dict[str, Any], user_email: str) -> bool:
        """
        Persist a personal-chat conversation reference so webhook notifications can
        reliably message the user later without creating a new Teams chat.
        """
        if not user_email or not self.is_personal_conversation(activity):
            return False

        reference = self.create_conversation_reference(activity)
        reference['user_email'] = user_email
        reference['stored_at'] = datetime.utcnow().isoformat() + 'Z'

        loop = asyncio.get_event_loop()

        container_client = await self._get_conversation_container_client()
        if not container_client:
            fallback_path = self._conversation_fallback_path(user_email)

            def write_local():
                try:
                    os.makedirs(self.conversation_fallback_dir, exist_ok=True)
                    with open(fallback_path, 'w', encoding='utf-8') as f:
                        json.dump(reference, f)
                    return True
                except Exception as e:
                    logging.error(
                        f"Failed to store local conversation reference for {user_email}: {str(e)}"
                    )
                    return False

            success = await loop.run_in_executor(self.executor, write_local)
            if success:
                logging.info(f"Stored personal conversation reference locally for {user_email}")
            return success

        blob_name = self._conversation_blob_name(user_email)

        def upload():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                blob_client.upload_blob(json.dumps(reference), overwrite=True)
                return True
            except Exception as e:
                logging.error(f"Failed to store conversation reference for {user_email}: {str(e)}")
                return False

        success = await loop.run_in_executor(self.executor, upload)
        if success:
            logging.info(f"Stored personal conversation reference for {user_email}")
        return success

    async def get_conversation_reference(self, user_email: str) -> Optional[Dict[str, Any]]:
        """Load a persisted conversation reference for a user if available."""
        if not user_email:
            return None

        loop = asyncio.get_event_loop()

        container_client = await self._get_conversation_container_client()
        if not container_client:
            fallback_path = self._conversation_fallback_path(user_email)

            def read_local():
                try:
                    with open(fallback_path, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except Exception:
                    return None

            reference = await loop.run_in_executor(self.executor, read_local)
            if reference:
                logging.info(f"Loaded local conversation reference for {user_email}")
            return reference

        blob_name = self._conversation_blob_name(user_email)

        def download():
            try:
                blob_client = container_client.get_blob_client(blob_name)
                raw = blob_client.download_blob().readall()
                return json.loads(raw)
            except Exception:
                return None

        reference = await loop.run_in_executor(self.executor, download)
        if reference:
            logging.info(f"Loaded stored conversation reference for {user_email}")
        return reference

    async def send_proactive_card(self, conversation_reference: Dict[str, Any], card: Dict[str, Any]) -> bool:
        """Send an adaptive card using a stored conversation reference."""
        try:
            token = await self.get_auth_token()
            if not token:
                return False

            service_url = conversation_reference.get('serviceUrl', self.service_url)
            conversation_id = conversation_reference.get('conversation', {}).get('id')
            if not conversation_id:
                logging.error("Conversation reference missing conversation id")
                return False

            message_url = f"{service_url}v3/conversations/{conversation_id}/activities"
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            message_activity = {
                'type': 'message',
                'from': conversation_reference.get('bot', {'id': self.app_id, 'name': 'IT Support Bot'}),
                'conversation': conversation_reference.get('conversation'),
                'recipient': conversation_reference.get('user'),
                'attachments': [{
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': card
                }]
            }

            loop = asyncio.get_event_loop()

            def send():
                response = requests.post(
                    message_url,
                    headers=headers,
                    json=message_activity,
                    timeout=30
                )
                if response.status_code not in [200, 201, 202]:
                    logging.error(
                        f"Failed proactive send to stored conversation {conversation_id}: "
                        f"{response.status_code} - {response.text}"
                    )
                return response.status_code in [200, 201, 202]

            return await loop.run_in_executor(self.executor, send)

        except Exception as e:
            logging.error(f"Error sending proactive card: {str(e)}")
            return False

    async def _create_personal_conversation(self, token: str, user_email: str, user_aad_id: str) -> Optional[Dict[str, Any]]:
        """Create a 1:1 Teams conversation with the user using Bot Framework APIs."""
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        create_conv_url = f"{self.service_url}v3/conversations"

        payloads = [
            {
                'bot': {
                    'id': self.app_id,
                    'name': 'IT Support Bot'
                },
                'members': [
                    {
                        'aadObjectId': user_aad_id,
                        'name': user_email
                    }
                ],
                'channelData': {
                    'tenant': {
                        'id': self.tenant_id
                    }
                },
                'tenantId': self.tenant_id,
                'isGroup': False
            },
            {
                'bot': {
                    'id': self.app_id,
                    'name': 'IT Support Bot'
                },
                'members': [
                    {
                        'id': f'29:{user_aad_id}',
                        'name': user_email
                    }
                ],
                'channelData': {
                    'tenant': {
                        'id': self.tenant_id
                    }
                },
                'tenantId': self.tenant_id,
                'isGroup': False
            }
        ]

        loop = asyncio.get_event_loop()

        def create(payload):
            response = requests.post(
                create_conv_url,
                headers=headers,
                json=payload,
                timeout=30
            )
            return response

        for payload in payloads:
            response = await loop.run_in_executor(self.executor, create, payload)
            if response.status_code in [200, 201, 202]:
                return response.json()
            logging.error(
                f"Failed to create conversation for {user_email}: "
                f"{response.status_code} - {response.text}"
            )

        return None

    async def send_proactive_message(self, conversation_reference: Dict[str, Any], message: str) -> bool:
        """
        Send a proactive message using a stored conversation reference
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return False
            
            service_url = conversation_reference.get('serviceUrl', self.service_url)
            conversation_id = conversation_reference.get('conversation', {}).get('id')
            
            message_url = f"{service_url}v3/conversations/{conversation_id}/activities"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            message_activity = {
                'type': 'message',
                'from': conversation_reference.get('bot'),
                'conversation': conversation_reference.get('conversation'),
                'recipient': conversation_reference.get('user'),
                'text': message
            }
            
            loop = asyncio.get_event_loop()
            
            def send():
                response = requests.post(message_url, headers=headers, json=message_activity)
                return response.status_code in [200, 201, 202]
            
            return await loop.run_in_executor(self.executor, send)
            
        except Exception as e:
            logging.error(f"Error sending proactive message: {str(e)}")
            return False

    def validate_auth_header(self, auth_header: str) -> bool:
        """
        Validate authentication header from Teams
        """
        try:
            # In production, implement proper JWT validation
            # This should verify the token with Microsoft's public keys

            if not auth_header or not auth_header.startswith('Bearer '):
                return False

            # Add proper JWT validation here
            # For now, basic check
            return True

        except Exception as e:
            logging.error(f"Error validating auth header: {str(e)}")
            return False

    async def send_notification_to_user(self, user_email: str, card: Dict[str, Any]) -> bool:
        """
        Send a proactive notification card to a user by their email address.

        This creates a new 1:1 conversation with the user and sends the card.
        Requires the bot to be installed for the user in Teams.

        Args:
            user_email: The user's email address (UPN)
            card: The Adaptive Card to send

        Returns:
            bool: True if notification was sent successfully
        """
        try:
            token = await self.get_auth_token()
            if not token:
                logging.error("Failed to get auth token for proactive message")
                return False

            stored_reference = await self.get_conversation_reference(user_email)
            if stored_reference:
                sent = await self.send_proactive_card(stored_reference, card)
                if sent:
                    logging.info(f"Successfully sent proactive notification via stored reference to {user_email}")
                    return True
                logging.warning(
                    f"Stored conversation reference send failed for {user_email}; "
                    "falling back to direct conversation creation"
                )

            # Resolve email to AAD Object ID for Bot Framework conversation creation
            user_aad_id = await self.get_user_aad_id(user_email)
            if not user_aad_id:
                logging.error(f"Could not resolve AAD ID for {user_email}, cannot create conversation")
                return False
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            loop = asyncio.get_event_loop()
            conv_response = await self._create_personal_conversation(token, user_email, user_aad_id)

            if not conv_response:
                logging.error(f"Could not create conversation with user {user_email}")
                return False

            conversation_id = conv_response.get('id')
            if not conversation_id:
                logging.error("No conversation ID in response")
                return False

            # Step 2: Send the card to the conversation
            message_url = f"{self.service_url}v3/conversations/{conversation_id}/activities"

            message_activity = {
                'type': 'message',
                'from': {
                    'id': self.app_id,
                    'name': 'IT Support Bot'
                },
                'attachments': [{
                    'contentType': 'application/vnd.microsoft.card.adaptive',
                    'content': card
                }]
            }

            def send_message():
                response = requests.post(
                    message_url,
                    headers=headers,
                    json=message_activity,
                    timeout=30
                )
                if response.status_code not in [200, 201, 202]:
                    logging.error(
                        f"Failed to send card to conversation {conversation_id}: "
                        f"{response.status_code} - {response.text}"
                    )
                return response.status_code in [200, 201, 202]

            success = await loop.run_in_executor(self.executor, send_message)

            if success:
                logging.info(f"Successfully sent proactive notification to {user_email}")
            else:
                logging.error(f"Failed to send message to conversation {conversation_id}")

            return success

        except Exception as e:
            logging.error(f"Error sending notification to user {user_email}: {str(e)}")
            return False

    async def send_notification_card_to_user(self, user_email: str, text: str) -> bool:
        """
        Send a simple text notification to a user by email.

        Args:
            user_email: The user's email address
            text: The message text to send

        Returns:
            bool: True if sent successfully
        """
        try:
            token = await self.get_auth_token()
            if not token:
                return False

            # Resolve email to AAD Object ID for Bot Framework conversation creation
            user_aad_id = await self.get_user_aad_id(user_email)
            if not user_aad_id:
                logging.error(f"Could not resolve AAD ID for {user_email}")
                return False

            member_id = f"29:{user_aad_id}"

            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }

            create_conv_url = f"{self.service_url}v3/conversations"

            conversation_params = {
                'bot': {
                    'id': self.app_id,
                    'name': 'IT Support Bot'
                },
                'members': [
                    {
                        'id': member_id,
                        'name': user_email
                    }
                ],
                'channelData': {
                    'tenant': {
                        'id': self.tenant_id
                    }
                },
                'isGroup': False,
                'activity': {
                    'type': 'message',
                    'text': text
                }
            }

            loop = asyncio.get_event_loop()

            def send():
                response = requests.post(
                    create_conv_url,
                    headers=headers,
                    json=conversation_params,
                    timeout=30
                )
                return response.status_code in [200, 201, 202]

            return await loop.run_in_executor(self.executor, send)

        except Exception as e:
            logging.error(f"Error sending text notification to {user_email}: {str(e)}")
            return False
