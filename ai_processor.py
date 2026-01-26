"""
AI Processor - Handles GPT-5 integration for IT support responses
"""

import os
import json
import logging
import requests
import asyncio
from typing import Dict, Any, Optional, List
from datetime import datetime
import re
from concurrent.futures import ThreadPoolExecutor

class AIProcessor:
    def __init__(self):
        """Initialize AI processor with GPT-5 endpoint"""
        # Using your internal GPT-5 endpoint
        self.endpoint = os.environ.get('GPT5_ENDPOINT', '')
        self.api_key = os.environ.get('GPT5_API_KEY', '')
        self.model = os.environ.get('GPT5_MODEL', 'gpt-5')
        
        # Fallback to Azure OpenAI if GPT-5 not available
        self.azure_endpoint = os.environ.get('AZURE_OPENAI_ENDPOINT', '')
        self.azure_api_key = os.environ.get('AZURE_OPENAI_KEY', '')
        self.deployment_name = os.environ.get('AZURE_OPENAI_DEPLOYMENT', 'gpt-4')
        
        self.executor = ThreadPoolExecutor(max_workers=3)
        
        # Knowledge base for common IT issues
        self.knowledge_base = self.load_knowledge_base()
        
        # Categories that typically need tickets
        self.ticket_required_categories = [
            'hardware replacement',
            'new user setup',
            'software license',
            'security incident',
            'data recovery',
            'server access',
            'database issue'
        ]

    def load_knowledge_base(self) -> Dict[str, Any]:
        """
        Load IT knowledge base for quick responses
        """
        return {
            "password_reset": {
                "keywords": ["password", "reset password", "forgot password", "can't login", "locked out"],
                "solution": """To reset your password:
1. Go to https://passwordreset.microsoftonline.com
2. Enter your work email address
3. Complete the verification process
4. Create a new password following these requirements:
   - At least 12 characters
   - Include uppercase, lowercase, numbers, and special characters
   - Cannot be one of your last 5 passwords

If you're still having issues, I can create a ticket for IT to assist you directly.""",
                "category": "Password Reset",
                "needs_ticket": False
            },
            "vpn_issues": {
                "keywords": ["vpn", "remote access", "work from home", "can't connect", "connection failed"],
                "solution": """To troubleshoot VPN issues:
1. Ensure you're connected to the internet
2. Open the VPN client and check for updates
3. Try these steps:
   - Disconnect and reconnect
   - Restart the VPN client
   - Clear saved credentials and re-enter them
4. Check if your VPN credentials have expired

Common fixes:
- Windows: Run 'ipconfig /flushdns' in Command Prompt
- Mac: Run 'sudo dscacheutil -flushcache' in Terminal
- Verify firewall isn't blocking VPN ports (UDP 500, 4500)

Still having issues? I'll create a ticket for deeper investigation.""",
                "category": "VPN Access",
                "needs_ticket": False
            },
            "teams_issues": {
                "keywords": ["teams", "microsoft teams", "can't join meeting", "no audio", "no video", "teams crashed"],
                "solution": """To fix Microsoft Teams issues:

**For Audio/Video Problems:**
1. Click your profile picture → Settings → Devices
2. Test your speaker and microphone
3. Ensure correct devices are selected
4. Check Windows Settings → Privacy → Microphone/Camera permissions

**For Performance Issues:**
1. Clear Teams cache:
   - Close Teams completely
   - Press Win+R, type %appdata%\\Microsoft\\Teams
   - Delete contents of these folders: Cache, Blob_storage, Databases, GPUCache, IndexedDB, Local Storage, tmp
2. Restart Teams

**For Login Issues:**
1. Sign out completely: Profile → Sign out
2. Clear credentials: Windows Credential Manager → Remove Teams entries
3. Sign in with your full email address

Need more help? I can create a ticket for you.""",
                "category": "Teams/Office 365",
                "needs_ticket": False
            },
            "printer_issues": {
                "keywords": ["printer", "can't print", "print queue", "printer offline", "print job stuck"],
                "solution": """To resolve printer issues:

**Printer Offline:**
1. Open Settings → Devices → Printers & scanners
2. Select your printer → Open queue
3. Click Printer → Use Printer Online
4. Try printing a test page

**Stuck Print Jobs:**
1. Open print queue (as above)
2. Right-click stuck jobs → Cancel
3. Restart the Print Spooler service:
   - Press Win+R, type services.msc
   - Find 'Print Spooler' → Right-click → Restart

**Driver Issues:**
1. Go to Settings → Devices → Printers & scanners
2. Remove the printer
3. Click 'Add a printer' to reinstall
4. Windows will download latest drivers

Still not working? I'll create a ticket for on-site assistance.""",
                "category": "Printer Problems",
                "needs_ticket": False
            },
            "email_issues": {
                "keywords": ["email", "outlook", "can't send", "can't receive", "email stuck", "mailbox full"],
                "solution": """To troubleshoot email issues:

**Outlook Not Syncing:**
1. Check your internet connection
2. In Outlook: File → Office Account → Update Options → Update Now
3. Restart Outlook

**Mailbox Full:**
1. Archive old emails: File → Info → Cleanup Tools → Archive
2. Empty Deleted Items: Right-click folder → Empty Folder
3. Check mailbox size: File → Info → Mailbox Settings

**Can't Send/Receive:**
1. Click Send/Receive → Send/Receive Groups → Define Send/Receive Groups
2. Ensure 'Include this group in send/receive' is checked
3. Check account settings: File → Account Settings → Repair

**Outlook Profile Issues:**
1. Control Panel → Mail → Show Profiles
2. Add new profile and set as default
3. Re-add your email account

Need additional help? I can create a priority ticket.""",
                "category": "Email Issues",
                "needs_ticket": False
            },
            "software_installation": {
                "keywords": ["install", "software", "application", "download", "need access", "license"],
                "solution": """For software installation:

**Company-Approved Software:**
1. Open Software Center (Windows) or Self Service (Mac)
2. Browse available applications
3. Click Install for needed software

**Software Not in Catalog:**
- I'll need to create a ticket for software approval
- Please provide: Software name, version, business justification

**License Issues:**
- Requires IT approval and procurement
- I'll create a ticket with high priority

Which software do you need? I can help create a detailed request.""",
                "category": "Software Installation",
                "needs_ticket": True
            },
            "slow_computer": {
                "keywords": ["slow", "performance", "freezing", "lagging", "not responding"],
                "solution": """To improve computer performance:

**Quick Fixes:**
1. Restart your computer (if you haven't in a while)
2. Check for Windows Updates: Settings → Update & Security
3. Close unnecessary programs and browser tabs

**Clear Temporary Files:**
1. Press Win+R, type 'cleanmgr'
2. Select your drive → OK
3. Check all boxes → OK → Delete Files

**Check Resources:**
1. Open Task Manager (Ctrl+Shift+Esc)
2. Click 'More details' if needed
3. Check CPU, Memory, and Disk usage
4. Sort by usage to find problem programs

**Disable Startup Programs:**
1. Task Manager → Startup tab
2. Disable unnecessary programs

If issues persist after these steps, hardware upgrade might be needed. Should I create a ticket for IT assessment?""",
                "category": "Hardware Issue",
                "needs_ticket": False
            }
        }

    async def get_support_response(self, question: str) -> Dict[str, Any]:
        """
        Get AI-powered support response for user question
        """
        try:
            # First, check knowledge base for quick response
            kb_response = self.check_knowledge_base(question)
            if kb_response:
                return kb_response
            
            # Use GPT-5 for complex queries
            ai_response = await self.query_ai(question)
            
            # Analyze if ticket is needed
            needs_ticket = self.analyze_ticket_requirement(question, ai_response)
            
            # Suggest category and priority
            category = self.suggest_category(question)
            priority = self.suggest_priority(question)
            
            return {
                'solution': ai_response,
                'needs_ticket': needs_ticket,
                'suggested_category': category,
                'suggested_priority': priority,
                'suggested_subject': self.generate_subject(question),
                'confidence': 0.85  # Adjust based on response quality
            }
            
        except Exception as e:
            logging.error(f"Error getting support response: {str(e)}")
            return {
                'solution': "I'm having trouble processing your request. Please create a ticket for IT assistance.",
                'needs_ticket': True,
                'suggested_category': 'General Support',
                'suggested_priority': 'Medium',
                'confidence': 0
            }

    def check_knowledge_base(self, question: str) -> Optional[Dict[str, Any]]:
        """
        Check if question matches known issues in knowledge base
        """
        question_lower = question.lower()
        
        for issue_key, issue_data in self.knowledge_base.items():
            # Check if any keywords match
            if any(keyword in question_lower for keyword in issue_data['keywords']):
                return {
                    'solution': issue_data['solution'],
                    'needs_ticket': issue_data['needs_ticket'],
                    'suggested_category': issue_data['category'],
                    'suggested_priority': 'Medium',
                    'suggested_subject': self.generate_subject(question),
                    'confidence': 0.95
                }
        
        return None

    async def query_ai(self, question: str) -> str:
        """
        Query GPT-5 or Azure OpenAI for IT support response
        """
        try:
            # Prepare the prompt
            system_prompt = """You are an expert IT support assistant for a corporate environment.
            Provide clear, step-by-step solutions for technical issues.
            Focus on Microsoft products, Windows, Office 365, Teams, and common enterprise software.
            Be concise but thorough. Include troubleshooting steps when appropriate.
            IMPORTANT: Users do not have admin access to their laptops and cannot install programs themselves.
            Do not suggest solutions that require admin rights or software installation.
            If the issue requires administrative access, software installation, or hardware replacement, mention that IT assistance is needed."""
            
            user_prompt = f"""User's IT Issue: {question}

Please provide:
1. Immediate troubleshooting steps the user can try
2. Likely cause of the issue
3. Whether IT ticket is required (if admin access, hardware, or complex configuration needed)

Format the response in a clear, friendly manner suitable for non-technical users."""

            # Try GPT-5 first
            if self.endpoint and self.api_key:
                response = await self.call_gpt5(system_prompt, user_prompt)
                if response:
                    return response
            
            # Fallback to Azure OpenAI
            if self.azure_endpoint and self.azure_api_key:
                response = await self.call_azure_openai(system_prompt, user_prompt)
                if response:
                    return response
            
            # Fallback response
            return self.get_fallback_response(question)
            
        except Exception as e:
            logging.error(f"Error querying AI: {str(e)}")
            return self.get_fallback_response(question)

    async def call_gpt5(self, system_prompt: str, user_prompt: str) -> Optional[str]:
        """
        Call internal GPT-5 endpoint
        """
        loop = asyncio.get_event_loop()
        
        def make_request():
            try:
                headers = {
                    'Authorization': f'Bearer {self.api_key}',
                    'Content-Type': 'application/json'
                }
                
                data = {
                    'model': self.model,
                    'messages': [
                        {'role': 'system', 'content': system_prompt},
                        {'role': 'user', 'content': user_prompt}
                    ],
                    'max_tokens': 500,
                    'temperature': 0.7
                }
                
                response = requests.post(
                    self.endpoint,
                    headers=headers,
                    json=data,
                    timeout=10
                )
                
                if response.status_code == 200:
                    result = response.json()
                    return result.get('choices', [{}])[0].get('message', {}).get('content', '')
                
                return None
                
            except Exception as e:
                logging.error(f"GPT-5 call failed: {str(e)}")
                return None
        
        return await loop.run_in_executor(self.executor, make_request)

    async def call_azure_openai(self, system_prompt: str, user_prompt: str) -> Optional[str]:
        """
        Call Azure OpenAI as fallback
        """
        loop = asyncio.get_event_loop()
        
        def make_request():
            try:
                url = f"{self.azure_endpoint}/openai/deployments/{self.deployment_name}/chat/completions?api-version=2024-02-15-preview"
                
                headers = {
                    'api-key': self.azure_api_key,
                    'Content-Type': 'application/json'
                }
                
                data = {
                    'messages': [
                        {'role': 'system', 'content': system_prompt},
                        {'role': 'user', 'content': user_prompt}
                    ],
                    'max_tokens': 500,
                    'temperature': 0.7
                }
                
                response = requests.post(
                    url,
                    headers=headers,
                    json=data,
                    timeout=10
                )
                
                if response.status_code == 200:
                    result = response.json()
                    return result.get('choices', [{}])[0].get('message', {}).get('content', '')
                
                return None
                
            except Exception as e:
                logging.error(f"Azure OpenAI call failed: {str(e)}")
                return None
        
        return await loop.run_in_executor(self.executor, make_request)

    def analyze_ticket_requirement(self, question: str, ai_response: str) -> bool:
        """
        Determine if the issue requires a ticket
        """
        question_lower = question.lower()
        response_lower = ai_response.lower() if ai_response else ""
        
        # Keywords that indicate ticket is needed
        ticket_indicators = [
            'hardware replacement',
            'admin access',
            'administrative privileges',
            'new equipment',
            'software license',
            'security incident',
            'data loss',
            'server',
            'database',
            'network configuration',
            'firewall',
            'cannot be resolved',
            'requires it support',
            'contact it',
            'on-site assistance'
        ]
        
        # Check if any indicators are present
        for indicator in ticket_indicators:
            if indicator in question_lower or indicator in response_lower:
                return True
        
        # Check categories that need tickets
        for category in self.ticket_required_categories:
            if category in question_lower:
                return True
        
        return False

    def suggest_category(self, question: str) -> str:
        """
        Suggest appropriate ticket category
        """
        question_lower = question.lower()
        
        category_keywords = {
            'Password Reset': ['password', 'reset', 'locked out', 'can\'t login'],
            'Software Installation': ['install', 'software', 'application', 'license'],
            'Hardware Issue': ['computer', 'laptop', 'monitor', 'keyboard', 'mouse', 'slow', 'broken'],
            'Network Connectivity': ['network', 'internet', 'wifi', 'connection', 'can\'t connect'],
            'Email Issues': ['email', 'outlook', 'mailbox', 'calendar'],
            'Teams/Office 365': ['teams', 'office', 'word', 'excel', 'powerpoint', 'onedrive', 'sharepoint'],
            'VPN Access': ['vpn', 'remote', 'work from home'],
            'Printer Problems': ['printer', 'print', 'scanner', 'scan'],
            'File Access': ['file', 'folder', 'share', 'drive', 'permission', 'access'],
            'Security Concern': ['security', 'virus', 'malware', 'phishing', 'suspicious'],
            'New User Setup': ['new user', 'onboarding', 'new employee', 'new hire']
        }
        
        for category, keywords in category_keywords.items():
            if any(keyword in question_lower for keyword in keywords):
                return category
        
        return 'General Support'

    def suggest_priority(self, question: str) -> str:
        """
        Suggest ticket priority based on keywords and urgency
        """
        question_lower = question.lower()
        
        # Critical priority indicators
        if any(word in question_lower for word in ['urgent', 'emergency', 'critical', 'asap', 'down', 'outage', 'security incident', 'data loss', 'ransomware']):
            return 'Critical'
        
        # High priority indicators
        if any(word in question_lower for word in ['important', 'deadline', 'can\'t work', 'blocking', 'multiple users', 'department']):
            return 'High'
        
        # Low priority indicators
        if any(word in question_lower for word in ['when you can', 'not urgent', 'nice to have', 'question', 'how to']):
            return 'Low'
        
        # Default to medium
        return 'Medium'

    def generate_subject(self, question: str) -> str:
        """
        Generate concise ticket subject from question
        """
        # Remove common words and limit length
        words_to_remove = ['the', 'a', 'an', 'is', 'are', 'was', 'were', 'been', 'have', 'has', 'had', 'i', 'my', 'me']
        
        words = question.lower().split()
        filtered_words = [word for word in words if word not in words_to_remove]
        
        # Take first 5-7 significant words
        subject_words = filtered_words[:7]
        subject = ' '.join(subject_words).title()
        
        # Limit to 50 characters
        if len(subject) > 50:
            subject = subject[:47] + '...'
        
        return subject

    def get_fallback_response(self, question: str) -> str:
        """
        Provide fallback response when AI is unavailable
        """
        return f"""I understand you're experiencing an IT issue. While I'm having trouble accessing the AI service right now, here are some general steps:

1. Try restarting the affected application or your computer
2. Check if other users are experiencing the same issue
3. Ensure you have a stable network connection
4. Check for any recent system updates

For immediate assistance, I recommend creating a ticket with the details of your issue. This will ensure IT support can help you promptly.

Your issue: {question[:100]}..."""

    async def process_feedback(self, feedback_data: Dict[str, Any]) -> None:
        """
        Process user feedback to improve responses
        """
        try:
            # Log feedback for analysis
            logging.info(f"AI Feedback: {json.dumps(feedback_data)}")
            
            # In production, store this in a database for analysis
            # and continuous improvement of the knowledge base
            
        except Exception as e:
            logging.error(f"Error processing feedback: {str(e)}")