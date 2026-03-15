"""
Teams Handler - Manages Teams bot interactions and messaging
"""

import os
import json
import logging
import requests
import asyncio
from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
import uuid

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

            # Resolve email to AAD Object ID for Bot Framework conversation creation
            user_aad_id = await self.get_user_aad_id(user_email)
            if not user_aad_id:
                logging.error(f"Could not resolve AAD ID for {user_email}, cannot create conversation")
                return False

            member_id = f"29:{user_aad_id}"

            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }

            # Step 1: Create a conversation with the user
            # The bot needs to create a 1:1 chat to send proactive messages
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
                'isGroup': False
            }

            loop = asyncio.get_event_loop()

            def create_conversation():
                response = requests.post(
                    create_conv_url,
                    headers=headers,
                    json=conversation_params,
                    timeout=30
                )
                if response.status_code in [200, 201, 202]:
                    return response.json()
                else:
                    logging.error(f"Failed to create conversation: {response.status_code} - {response.text}")
                    return None

            conv_response = await loop.run_in_executor(self.executor, create_conversation)

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