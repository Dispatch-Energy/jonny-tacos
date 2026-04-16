"""
QuickBase Manager - Handles all QuickBase ticket operations
"""

import os
import json
import logging
import requests
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List
import asyncio
from concurrent.futures import ThreadPoolExecutor
from xml.etree import ElementTree

class QuickBaseManager:
    def __init__(self):
        """Initialize QuickBase connection settings"""
        self.realm = os.environ.get('QB_REALM', '')
        self.user_token = os.environ.get('QB_USER_TOKEN', '')
        self.app_id = os.environ.get('QB_APP_ID', '')
        self.app_token = os.environ.get('QB_APP_TOKEN', '')
        self.user_creation_app_id = os.environ.get('QB_USER_CREATION_APP_ID', '')
        self.user_creation_app_token = os.environ.get('QB_USER_CREATION_APP_TOKEN', '')
        self.table_id = os.environ.get('QB_TICKETS_TABLE_ID', '')
        self.user_creation_user_token = os.environ.get(
            'QB_USER_CREATION_USER_TOKEN',
            self.user_token
        )
        
        self.base_url = f"https://api.quickbase.com/v1"
        self.headers = {
            'QB-Realm-Hostname': self.realm,
            'Authorization': f'QB-USER-TOKEN {self.user_token}',
            'Content-Type': 'application/json'
        }
        
        # Field mappings from your QuickBase schema
        self.field_mapping = {
            'ticket_number': 6,      # __Ticket Number__
            'subject': 7,             # __Subject__
            'description': 8,         # __Description__
            'priority': 9,            # __Priority__
            'category': 10,           # __Category__
            'status': 11,             # __Status__
            'submitted_date': 12,     # __Submitted Date__
            'due_date': 13,           # __Due Date__
            'resolved_date': 14,      # __Resolved Date__
            'resolution': 15,         # __Resolution__
            'time_spent': 16,         # __Time Spent (hours)__
            'record_links': 17,       # __Record records__
            'add_record': 18,         # __Add Record__
            'submitted_by': 19        # __Submitted By__ (email field)
        }
        
        # Priority mappings
        self.priority_values = {
            'Low': 1,
            'Medium': 2,
            'High': 3,
            'Critical': 4
        }
        
        # Category options
        self.categories = [
            'Password Reset',
            'Software Installation',
            'Hardware Issue',
            'Network Connectivity',
            'Email Issues',
            'Teams/Office 365',
            'VPN Access',
            'Printer Problems',
            'File Access',
            'Security Concern',
            'New User Setup',
            'User Creation',
            'General Support',
            'Other'
        ]
        
        # Status options
        self.status_values = [
            'New',
            'Bot Assisted',
            'In Progress',
            'Awaiting User',
            'Awaiting IT',
            'Resolved',
            'Closed',
            'Cancelled'
        ]
        
        self.executor = ThreadPoolExecutor(max_workers=5)
        self._role_id_cache: Dict[str, Dict[str, str]] = {}

    async def create_ticket(self, ticket_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """
        Create a new ticket in QuickBase
        """
        subject = (ticket_data.get('subject') or '').strip()
        description = (ticket_data.get('description') or '').strip()
        if not subject or not description:
            logging.warning(
                "Blocked empty ticket creation: subject=%r description=%r",
                subject, description
            )
            return None

        try:
            # Generate unique ticket number using timestamp (guaranteed unique)
            ticket_number = f"IT-{datetime.now().strftime('%y%m%d%H%M%S')}"
            print(f"   [DEBUG] Generated ticket number: {ticket_number}")
            
            # Calculate due date based on priority
            submitted_date = datetime.now()
            due_date = self.calculate_due_date(ticket_data.get('priority', 'Medium'))
            
            # Format dates for QuickBase
            # Field 12 (timestamp) needs ISO format with Z suffix
            # Field 13 (date) needs just YYYY-MM-DD
            submitted_str = submitted_date.strftime('%Y-%m-%dT%H:%M:%SZ')
            due_date_str = due_date.strftime('%Y-%m-%d')
            
            # Build description - include user name if provided
            if ticket_data.get('user_name'):
                description += f"\n\n---\nTeams User: {ticket_data.get('user_name')}"

            # Prepare the record data
            record_data = {
                "to": self.table_id,
                "data": [{
                    self.field_mapping['ticket_number']: {"value": ticket_number},
                    self.field_mapping['subject']: {"value": subject},
                    self.field_mapping['description']: {"value": description},
                    self.field_mapping['priority']: {"value": ticket_data.get('priority', 'Medium')},
                    self.field_mapping['category']: {"value": ticket_data.get('category', 'General Support')},
                    self.field_mapping['status']: {"value": ticket_data.get('status', 'New')},
                    self.field_mapping['submitted_date']: {"value": submitted_str},
                    self.field_mapping['due_date']: {"value": due_date_str}
                }],
                "fieldsToReturn": list(self.field_mapping.values())
            }
            
            # Add submitted_by email field (ID 19)
            user_email = ticket_data.get('user_email', '')
            if user_email:
                record_data["data"][0][self.field_mapping['submitted_by']] = {
                    "value": user_email
                }
                print(f"   [DEBUG] Setting field 19 (submitted_by) = {user_email}")
            else:
                print(f"   [DEBUG] WARNING: No user_email in ticket_data, field 19 will be empty")
                logging.warning("No user_email provided for ticket - field 19 (submitted_by) will be empty")

            print(f"   [DEBUG] Sending to QuickBase...")
            
            # Make the API call
            response = await self.execute_request(
                'POST',
                f"{self.base_url}/records",
                record_data
            )
            
            print(f"   [DEBUG] Response: {response}")
            
            if response:
                # QuickBase returns created record ID in metadata, not data
                created_ids = response.get('metadata', {}).get('createdRecordIds', [])
                print(f"   [DEBUG] Created IDs: {created_ids}")
                
                if created_ids:
                    record_id = created_ids[0]
                    
                    # Format the ticket response
                    ticket = {
                        'ticket_number': ticket_number,
                        'record_id': record_id,
                        'subject': ticket_data.get('subject'),
                        'status': ticket_data.get('status', 'New'),
                        'priority': ticket_data.get('priority'),
                        'category': ticket_data.get('category'),
                        'submitted_date': submitted_str,
                        'due_date': due_date_str,
                        'quickbase_url': self.get_ticket_url(record_id)
                    }
                    
                    logging.info(f"Ticket created successfully: {ticket_number}")
                    return ticket
                else:
                    # Check for line errors
                    line_errors = response.get('metadata', {}).get('lineErrors', {})
                    if line_errors:
                        print(f"   [DEBUG] Line errors: {line_errors}")
                    print(f"   [DEBUG] No createdRecordIds in response!")
            else:
                print(f"   [DEBUG] Response was None!")
            
            return None
            
        except Exception as e:
            logging.error(f"Error creating ticket in QuickBase: {str(e)}")
            return None

    async def get_ticket(self, ticket_number: str) -> Optional[Dict[str, Any]]:
        """
        Get ticket details by ticket number
        """
        try:
            query_data = {
                "from": self.table_id,
                "select": list(self.field_mapping.values()),
                "where": f"{{{self.field_mapping['ticket_number']}.EX.'{ticket_number}'}}"
            }
            
            response = await self.execute_request(
                'POST',
                f"{self.base_url}/records/query",
                query_data
            )
            
            if response and response.get('data') and len(response['data']) > 0:
                return self.format_ticket_response(response['data'][0])
            
            return None
            
        except Exception as e:
            logging.error(f"Error getting ticket {ticket_number}: {str(e)}")
            return None

    async def get_ticket_by_record_id(self, record_id: str) -> Optional[Dict[str, Any]]:
        """
        Get ticket details by QuickBase record ID.
        """
        try:
            query_data = {
                "from": self.table_id,
                "select": list(self.field_mapping.values()),
                "where": f"{{3.EX.'{record_id}'}}"
            }

            response = await self.execute_request(
                'POST',
                f"{self.base_url}/records/query",
                query_data
            )

            if response and response.get('data') and len(response['data']) > 0:
                return self.format_ticket_response(response['data'][0])

            return None

        except Exception as e:
            logging.error(f"Error getting ticket by record_id {record_id}: {str(e)}")
            return None

    async def get_ticket_by_reference(
        self,
        ticket_number: Optional[str] = None,
        record_id: Optional[str] = None,
    ) -> Optional[Dict[str, Any]]:
        """
        Resolve a ticket by ticket number or record ID.
        """
        if ticket_number:
            ticket = await self.get_ticket(ticket_number)
            if ticket:
                return ticket
        if record_id:
            return await self.get_ticket_by_record_id(record_id)
        return None

    async def get_user_tickets(self, user_email: str, status_filter: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        """
        Get all tickets for a specific user by their email in Submitted By field (ID 19)
        """
        try:
            # Build the query - search by submitted_by field (ID 19)
            where_clause = f"{{{self.field_mapping['submitted_by']}.EX.'{user_email}'}}"
            
            if status_filter:
                status_conditions = " OR ".join([
                    f"{{{self.field_mapping['status']}.EX.'{status}'}}"
                    for status in status_filter
                ])
                where_clause = f"({where_clause}) AND ({status_conditions})"
            else:
                # Default to active tickets only
                active_statuses = ['New', 'Bot Assisted', 'In Progress', 'Awaiting User', 'Awaiting IT']
                active_conditions = " OR ".join([
                    f"{{{self.field_mapping['status']}.EX.'{status}'}}"
                    for status in active_statuses
                ])
                where_clause = f"({where_clause}) AND ({active_conditions})"
            
            query_data = {
                "from": self.table_id,
                "select": list(self.field_mapping.values()),
                "where": where_clause,
                "sortBy": [
                    {"fieldId": self.field_mapping['submitted_date'], "order": "DESC"}
                ],
                "options": {
                    "top": 10  # Limit to 10 most recent tickets
                }
            }
            
            response = await self.execute_request(
                'POST',
                f"{self.base_url}/records/query",
                query_data
            )
            
            if response and response.get('data'):
                return [self.format_ticket_response(record) for record in response['data']]
            
            return []
            
        except Exception as e:
            logging.error(f"Error getting user tickets for {user_email}: {str(e)}")
            return []

    async def update_ticket(self, ticket_update: Dict[str, Any]) -> bool:
        """
        Update an existing ticket
        """
        try:
            # First get the record ID for the ticket
            ticket = await self.get_ticket_by_reference(
                ticket_number=ticket_update.get('ticket_id'),
                record_id=ticket_update.get('record_id')
            )
            if not ticket:
                return False
            
            record_id = ticket.get('record_id')
            
            # Prepare update data
            update_data = {
                "to": self.table_id,
                "data": [{
                    "3": {"value": record_id}  # Record ID field
                }]
            }
            
            # Add fields to update
            if 'status' in ticket_update:
                update_data["data"][0][self.field_mapping['status']] = {"value": ticket_update['status']}
                
                # If resolved, set resolved date
                if ticket_update['status'] in ['Resolved', 'Closed']:
                    update_data["data"][0][self.field_mapping['resolved_date']] = {
                        "value": datetime.now().isoformat()
                    }
            
            if 'resolution' in ticket_update:
                update_data["data"][0][self.field_mapping['resolution']] = {"value": ticket_update['resolution']}
            
            if 'time_spent' in ticket_update:
                update_data["data"][0][self.field_mapping['time_spent']] = {"value": float(ticket_update['time_spent'])}
            
            response = await self.execute_request(
                'POST',
                f"{self.base_url}/records",
                update_data
            )
            
            return response is not None
            
        except Exception as e:
            logging.error(f"Error updating ticket: {str(e)}")
            return False

    async def append_ticket_resolution_note(
        self,
        ticket_number: Optional[str],
        note: str,
        status: Optional[str] = None,
        record_id: Optional[str] = None,
    ) -> bool:
        """
        Append a note to the resolution field without discarding existing history.
        """
        try:
            ticket = await self.get_ticket_by_reference(
                ticket_number=ticket_number,
                record_id=record_id
            )
            if not ticket:
                return False

            existing_resolution = (ticket.get('resolution') or '').strip()
            new_note = (note or '').strip()
            if not new_note:
                return False

            resolution = (
                f"{existing_resolution}\n\n{new_note}"
                if existing_resolution else new_note
            )
            update_data = {
                'ticket_id': ticket.get('ticket_number'),
                'record_id': ticket.get('record_id'),
                'resolution': resolution
            }
            if status:
                update_data['status'] = status

            return await self.update_ticket(update_data)
        except Exception as e:
            logging.error(
                f"Error appending resolution note for {ticket_number or record_id}: {str(e)}"
            )
            return False

    async def resolve_ticket(self, ticket_number: str, resolution: str, resolved_by: str) -> bool:
        """
        Resolve a ticket with resolution details
        """
        try:
            ticket_update = {
                'ticket_id': ticket_number,
                'status': 'Resolved',
                'resolution': f"{resolution}\n\nResolved by: {resolved_by} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            }
            
            return await self.update_ticket(ticket_update)
            
        except Exception as e:
            logging.error(f"Error resolving ticket {ticket_number}: {str(e)}")
            return False

    async def get_ticket_statistics(self) -> Dict[str, Any]:
        """
        Get ticket statistics for dashboard
        """
        try:
            # Query for various statistics
            stats = {
                'total_open': 0,
                'total_resolved_today': 0,
                'by_priority': {},
                'by_category': {},
                'avg_resolution_time': 0,
                'sla_compliance': 0
            }
            
            # Get open tickets count
            open_query = {
                "from": self.table_id,
                "select": [3],  # Just get record ID for counting
                "where": f"{{{self.field_mapping['status']}.NE.'Closed'}} AND {{{self.field_mapping['status']}.NE.'Resolved'}}",
                "options": {"count": True}
            }
            
            response = await self.execute_request('POST', f"{self.base_url}/records/query", open_query)
            if response:
                stats['total_open'] = response.get('metadata', {}).get('totalRecords', 0)
            
            # Get tickets resolved today
            today_start = datetime.now().replace(hour=0, minute=0, second=0).isoformat()
            resolved_today_query = {
                "from": self.table_id,
                "select": [3],
                "where": f"{{{self.field_mapping['resolved_date']}.GTE.'{today_start}'}}",
                "options": {"count": True}
            }
            
            response = await self.execute_request('POST', f"{self.base_url}/records/query", resolved_today_query)
            if response:
                stats['total_resolved_today'] = response.get('metadata', {}).get('totalRecords', 0)
            
            # Get breakdown by priority
            for priority in self.priority_values.keys():
                priority_query = {
                    "from": self.table_id,
                    "select": [3],
                    "where": f"{{{self.field_mapping['priority']}.EX.'{priority}'}} AND {{{self.field_mapping['status']}.NE.'Closed'}}",
                    "options": {"count": True}
                }
                
                response = await self.execute_request('POST', f"{self.base_url}/records/query", priority_query)
                if response:
                    stats['by_priority'][priority] = response.get('metadata', {}).get('totalRecords', 0)
            
            return stats
            
        except Exception as e:
            logging.error(f"Error getting ticket statistics: {str(e)}")
            return {}

    async def generate_ticket_number(self) -> str:
        """
        Generate a unique ticket number
        """
        try:
            # Get the last ticket number
            query_data = {
                "from": self.table_id,
                "select": [self.field_mapping['ticket_number']],
                "sortBy": [{"fieldId": self.field_mapping['ticket_number'], "order": "DESC"}],
                "options": {"top": 1}
            }
            
            response = await self.execute_request('POST', f"{self.base_url}/records/query", query_data)
            
            if response and response.get('data') and len(response['data']) > 0:
                last_ticket = response['data'][0].get(str(self.field_mapping['ticket_number']), {}).get('value', 'IT-0000')
                # Extract number and increment
                last_num = int(last_ticket.split('-')[1]) if '-' in last_ticket else 0
                new_num = last_num + 1
            else:
                new_num = 1
            
            return f"IT-{new_num:04d}"
            
        except Exception as e:
            logging.error(f"Error generating ticket number: {str(e)}")
            # Fallback to timestamp-based number
            return f"IT-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    def calculate_due_date(self, priority: str) -> datetime:
        """
        Calculate due date based on priority SLA
        """
        sla_hours = {
            'Critical': 4,
            'High': 8,
            'Medium': 24,
            'Low': 48
        }
        
        hours = sla_hours.get(priority, 24)
        
        # Calculate business hours (simple version, you might want to enhance this)
        due_date = datetime.now() + timedelta(hours=hours)
        
        # Skip weekends (simple implementation)
        while due_date.weekday() in [5, 6]:  # Saturday = 5, Sunday = 6
            due_date += timedelta(days=1)
        
        return due_date

    def format_ticket_response(self, record: Dict[str, Any]) -> Dict[str, Any]:
        """
        Format QuickBase record into ticket response
        """
        try:
            return {
                'ticket_number': record.get(str(self.field_mapping['ticket_number']), {}).get('value'),
                'record_id': record.get('3', {}).get('value'),
                'subject': record.get(str(self.field_mapping['subject']), {}).get('value'),
                'description': record.get(str(self.field_mapping['description']), {}).get('value'),
                'priority': record.get(str(self.field_mapping['priority']), {}).get('value'),
                'category': record.get(str(self.field_mapping['category']), {}).get('value'),
                'status': record.get(str(self.field_mapping['status']), {}).get('value'),
                'submitted_date': record.get(str(self.field_mapping['submitted_date']), {}).get('value'),
                'due_date': record.get(str(self.field_mapping['due_date']), {}).get('value'),
                'resolved_date': record.get(str(self.field_mapping['resolved_date']), {}).get('value'),
                'resolution': record.get(str(self.field_mapping['resolution']), {}).get('value'),
                'time_spent': record.get(str(self.field_mapping['time_spent']), {}).get('value'),
                'submitted_by': record.get(str(self.field_mapping['submitted_by']), {}).get('value'),
                'quickbase_url': self.get_ticket_url(record.get('3', {}).get('value'))
            }
        except Exception as e:
            logging.error(f"Error formatting ticket response: {str(e)}")
            return {}

    def get_ticket_url(self, record_id: str) -> str:
        """
        Generate QuickBase URL for ticket
        """
        return f"https://{self.realm}/db/{self.table_id}?a=dr&rid={record_id}"

    async def execute_request(self, method: str, url: str, data: Optional[Dict] = None) -> Optional[Dict]:
        """
        Execute QuickBase API request with error handling
        """
        loop = asyncio.get_event_loop()
        
        def make_request():
            try:
                if method == 'GET':
                    response = requests.get(url, headers=self.headers)
                elif method == 'POST':
                    response = requests.post(url, headers=self.headers, json=data)
                elif method == 'DELETE':
                    response = requests.delete(url, headers=self.headers)
                else:
                    return None
                
                response.raise_for_status()
                return response.json() if response.text else {}
                
            except requests.exceptions.RequestException as e:
                logging.error(f"QuickBase API request failed: {str(e)}")
                return None
        
        return await loop.run_in_executor(self.executor, make_request)

    async def execute_legacy_api(
        self,
        action: str,
        dbid: str,
        xml_fields: Dict[str, Any],
        user_token: Optional[str] = None,
    ) -> Optional[ElementTree.Element]:
        """
        Execute a QuickBase legacy XML API call.
        """
        loop = asyncio.get_event_loop()
        token = user_token or self.user_token

        def make_request():
            try:
                root = ElementTree.Element('qdbapi')
                for key, value in xml_fields.items():
                    if value in (None, ''):
                        continue
                    child = ElementTree.SubElement(root, key)
                    child.text = str(value)

                if token:
                    usertoken_el = ElementTree.SubElement(root, 'usertoken')
                    usertoken_el.text = token
                target_app_token = ''
                if dbid == self.app_id and self.app_token:
                    target_app_token = self.app_token
                elif dbid == self.user_creation_app_id and self.user_creation_app_token:
                    target_app_token = self.user_creation_app_token
                if target_app_token:
                    apptoken_el = ElementTree.SubElement(root, 'apptoken')
                    apptoken_el.text = target_app_token

                payload = ElementTree.tostring(root, encoding='utf-8')
                response = requests.post(
                    f"https://{self.realm}/db/{dbid}",
                    headers={
                        'QUICKBASE-ACTION': action,
                        'Content-Type': 'application/xml'
                    },
                    data=payload,
                    timeout=30,
                )
                response.raise_for_status()
                return ElementTree.fromstring(response.text)
            except Exception as e:
                logging.error(f"QuickBase legacy API request failed ({action}): {str(e)}")
                return None

        return await loop.run_in_executor(self.executor, make_request)

    @staticmethod
    def _legacy_text(root: Optional[ElementTree.Element], tag: str) -> str:
        if root is None:
            return ''
        node = root.find(tag)
        return node.text.strip() if node is not None and node.text else ''

    async def get_quickbase_user_id(self, email: str) -> Optional[str]:
        """
        Look up a QuickBase user ID by email address.
        """
        root = await self.execute_legacy_api(
            'API_GetUserInfo',
            'main',
            {'email': email},
            user_token=self.user_creation_user_token,
        )
        if root is None:
            return None

        errcode = self._legacy_text(root, 'errcode')
        if errcode == '0':
            return self._legacy_text(root, 'userid') or self._legacy_text(root, 'uid')

        logging.info(
            f"QuickBase user lookup for {email} returned errcode={errcode} "
            f"errtext={self._legacy_text(root, 'errtext')}"
        )
        return None

    async def get_app_role_id(self, app_id: str, role_name: str) -> Optional[str]:
        """
        Resolve a role name to its QuickBase role ID for an application.
        """
        cached = self._role_id_cache.get(app_id, {}).get(role_name.lower())
        if cached:
            return cached

        root = await self.execute_legacy_api(
            'API_GetRoleInfo',
            app_id,
            {},
            user_token=self.user_creation_user_token,
        )
        if root is None:
            return None

        errcode = self._legacy_text(root, 'errcode')
        if errcode != '0':
            logging.error(
                f"Failed to load QuickBase role info for {app_id}: "
                f"{self._legacy_text(root, 'errtext')}"
            )
            return None

        role_map = self._role_id_cache.setdefault(app_id, {})
        for role in root.findall('.//role'):
            role_id = role.get('id') or self._legacy_text(role, 'id')
            name = (role.get('name') or self._legacy_text(role, 'name')).strip()
            if role_id and name:
                role_map[name.lower()] = role_id

        return role_map.get(role_name.lower())

    async def ensure_app_user_in_role(
        self,
        email: str,
        first_name: str,
        last_name: str,
        app_id: str,
        role_name: str,
    ) -> Dict[str, Any]:
        """
        Ensure a user exists in QuickBase and is assigned to the target app role.
        """
        role_id = await self.get_app_role_id(app_id, role_name)
        if not role_id:
            return {
                'success': False,
                'error': f"Could not resolve QuickBase role '{role_name}' in app {app_id}",
            }

        user_id = await self.get_quickbase_user_id(email)
        if user_id:
            root = await self.execute_legacy_api(
                'API_AddUserToRole',
                app_id,
                {
                    'userid': user_id,
                    'roleid': role_id,
                },
                user_token=self.user_creation_user_token,
            )
            if root is None:
                return {'success': False, 'error': 'QuickBase API_AddUserToRole failed'}

            errcode = self._legacy_text(root, 'errcode')
            errtext = self._legacy_text(root, 'errtext')
            if errcode == '0' or 'already' in errtext.lower():
                return {
                    'success': True,
                    'user_id': user_id,
                    'provisioned': False,
                }

            return {'success': False, 'error': errtext or f'QuickBase errcode {errcode}'}

        root = await self.execute_legacy_api(
            'API_ProvisionUser',
            app_id,
            {
                'email': email,
                'fname': first_name,
                'lname': last_name,
                'roleid': role_id,
            },
            user_token=self.user_creation_user_token,
        )
        if root is None:
            return {'success': False, 'error': 'QuickBase API_ProvisionUser failed'}

        errcode = self._legacy_text(root, 'errcode')
        errtext = self._legacy_text(root, 'errtext')
        if errcode == '0':
            return {
                'success': True,
                'user_id': self._legacy_text(root, 'userid') or self._legacy_text(root, 'uid'),
                'provisioned': True,
            }

        return {'success': False, 'error': errtext or f'QuickBase errcode {errcode}'}
