"""
Local Test - Full interactive CLI for IT Support Bot
Tests: LangChain routing, GPT responses, QuickBase ticket creation
"""

import json
import os
import asyncio

# ============================================================================
# Load environment variables from local.settings.json
# ============================================================================

def load_local_settings():
    """Load env vars from Azure Functions local.settings.json"""
    settings_path = "local.settings.json"
    
    if not os.path.exists(settings_path):
        print(f"⚠️  {settings_path} not found!")
        print("Create one with this structure:")
        print("""
{
  "IsEncrypted": false,
  "Values": {
    "GPT5_ENDPOINT": "https://your-endpoint.com",
    "GPT5_API_KEY": "your-api-key",
    "GPT5_MODEL": "gpt-5",
    "QB_REALM": "your-realm.quickbase.com",
    "QB_USER_TOKEN": "your-token",
    "QB_APP_ID": "your-app-id",
    "QB_TICKETS_TABLE_ID": "your-table-id",
    "TEAMS_APP_ID": "your-teams-app-id",
    "TEAMS_APP_SECRET": "your-teams-secret",
    "TEAMS_TENANT_ID": "your-tenant-id"
  }
}
""")
        return False
    
    with open(settings_path, "r") as f:
        settings = json.load(f)
    
    # Load all values into environment
    values = settings.get("Values", {})
    for key, value in values.items():
        os.environ[key] = str(value)
    
    print(f"📦 Loaded {len(values)} env vars from {settings_path}")
    return True


# Load settings BEFORE importing modules
if not load_local_settings():
    print("❌ Cannot proceed without local.settings.json")
    exit(1)


# ============================================================================
# Import after env vars loaded
# ============================================================================

from support_chain import ITSupportChain
from quickbase_manager import QuickBaseManager


# ============================================================================
# Interactive CLI
# ============================================================================

class ITBotCLI:
    def __init__(self):
        print("\n🔧 Initializing components...")
        
        # Initialize chain
        try:
            self.chain = ITSupportChain()
            print("  ✅ LangChain support chain")
        except Exception as e:
            print(f"  ❌ Chain failed: {e}")
            self.chain = None
        
        # Initialize QuickBase
        try:
            self.qb = QuickBaseManager()
            print("  ✅ QuickBase manager")
        except Exception as e:
            print(f"  ❌ QuickBase failed: {e}")
            self.qb = None
        
        # Test user info (simulating Teams user)
        self.test_user = {
            "email": "test.user@dispatchenergy.com",
            "name": "Test User (CLI)"
        }
        
        print("\n" + "="*60)
    
    def print_help(self):
        print("""
📚 Commands:
  [any message]     Ask IT support question
  /ticket           Create a ticket directly
  /status [ID]      Check ticket status
  /my-tickets       List your open tickets
  /stats            Show ticket statistics
  /test-qb          Test QuickBase connection
  /test-create      Test creating a ticket directly (bypasses chain)
  /help             Show this help
  /quit             Exit

💡 Examples:
  "I can't reset my password"
  "Install Adobe Photoshop"
  "VPN keeps disconnecting"
  /status IT-0042
""")
    
    async def process_message(self, message: str) -> None:
        """Process user message and show response"""
        
        message = message.strip()
        if not message:
            return
        
        # Handle commands
        if message.lower() == "/help":
            self.print_help()
            return
        
        if message.lower() == "/quit":
            print("\n👋 Goodbye!")
            exit(0)
        
        if message.lower() == "/test-qb":
            await self.test_quickbase()
            return
        
        if message.lower() == "/test-create":
            await self.test_create_ticket()
            return
        
        if message.lower() == "/stats":
            await self.show_stats()
            return
        
        if message.lower() == "/my-tickets":
            await self.show_my_tickets()
            return
        
        if message.lower().startswith("/status"):
            parts = message.split()
            ticket_id = parts[1] if len(parts) > 1 else None
            await self.check_status(ticket_id)
            return
        
        if message.lower() == "/ticket":
            await self.create_ticket_interactive()
            return
        
        # Regular message - process through chain
        await self.handle_support_question(message)
    
    async def handle_support_question(self, question: str) -> None:
        """Process IT support question through full pipeline"""
        
        print("\n⏳ Processing...")
        
        if not self.chain:
            print("❌ Chain not initialized")
            return
        
        try:
            # Get chain response
            result = self.chain.process(question)
            
            print("\n" + "-"*60)
            print(f"🎯 Intent: {result.get('type')}")
            print("-"*60)
            
            if result.get('type') == 'solution':
                # Bot has a solution
                print(f"\n💡 Solution (confidence: {result.get('confidence', 0):.0%}):\n")
                print(result.get('solution', 'No solution provided'))
                
                # Always create ticket for tracking
                print("\n📝 Creating ticket for tracking...")
                ticket = await self.create_ticket_from_result(
                    question, 
                    result,
                    status="Bot Assisted",
                    priority="Low"
                )
                
                if ticket:
                    print(f"\n✅ Ticket created: {ticket.get('ticket_number')}")
                    print(f"   Status: Bot Assisted (auto-closes in 24h if resolved)")
                    print(f"   URL: {ticket.get('quickbase_url')}")
                
                if result.get('offer_ticket'):
                    print("\n⚠️  If this doesn't work, reply 'escalate' to raise priority")
            
            elif result.get('type') in ['ticket_needed', 'troubleshooting_needed']:
                # Needs IT help
                rec = result.get('recommendation', {})
                
                print(f"\n🎫 This requires IT assistance:\n")
                print(f"   Subject:  {rec.get('subject')}")
                print(f"   Category: {rec.get('category')}")
                print(f"   Priority: {rec.get('priority')}")
                print(f"\n   Reason: {rec.get('reasoning', 'Complex issue')}")
                
                # Create ticket
                print("\n📝 Creating ticket...")
                ticket = await self.create_ticket_from_result(
                    question,
                    result,
                    status="New",
                    priority=rec.get('priority', 'Medium')
                )
                
                if ticket:
                    print(f"\n✅ Ticket created: {ticket.get('ticket_number')}")
                    print(f"   Assigned to IT queue")
                    print(f"   URL: {ticket.get('quickbase_url')}")
            
            elif result.get('type') == 'status_check':
                ticket_num = result.get('ticket_number')
                if ticket_num:
                    await self.check_status(ticket_num)
                else:
                    await self.show_my_tickets()
            
            else:
                print(f"\n🤷 Unknown response type: {result}")
            
            print()
            
        except Exception as e:
            print(f"\n❌ Error: {e}")
            import traceback
            traceback.print_exc()
    
    async def create_ticket_from_result(
        self, 
        question: str, 
        result: dict,
        status: str = "New",
        priority: str = "Medium"
    ) -> dict:
        """Create ticket in QuickBase from chain result"""
        
        if not self.qb:
            print("❌ QuickBase not initialized")
            return None
        
        rec = result.get('recommendation', {})
        
        # Build description
        description = question
        if result.get('solution'):
            description += f"\n\n--- Bot Response ---\n{result.get('solution')}"
        
        ticket_data = {
            'subject': rec.get('subject', question[:50]),
            'description': description,
            'priority': priority,
            'category': rec.get('category', result.get('category', 'General Support')),
            'user_email': self.test_user['email'],
            'user_name': self.test_user['name'],
        }
        
        print(f"\n   Debug - Ticket data:")
        print(f"     Subject: {ticket_data['subject']}")
        print(f"     Category: {ticket_data['category']}")
        print(f"     Priority: {ticket_data['priority']}")
        
        try:
            ticket = await self.qb.create_ticket(ticket_data)
            
            if ticket:
                return ticket
            else:
                print("❌ QuickBase returned None - check QB logs")
                return None
                
        except Exception as e:
            print(f"❌ Failed to create ticket: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    async def create_ticket_interactive(self) -> None:
        """Interactive ticket creation"""
        
        print("\n🎫 Create New Ticket\n")
        
        subject = input("Subject: ").strip()
        if not subject:
            print("❌ Subject required")
            return
        
        description = input("Description: ").strip()
        
        print("\nCategories: Password Reset, Software Installation, Hardware Issue,")
        print("           Network Connectivity, Email Issues, Teams/Office 365,")
        print("           VPN Access, Printer Problems, General Support")
        category = input("Category [General Support]: ").strip() or "General Support"
        
        print("\nPriority: Low, Medium, High, Critical")
        priority = input("Priority [Medium]: ").strip() or "Medium"
        
        ticket_data = {
            'subject': subject,
            'description': description,
            'priority': priority,
            'category': category,
            'user_email': self.test_user['email'],
            'user_name': self.test_user['name'],
        }
        
        print("\n⏳ Creating ticket...")
        
        try:
            ticket = await self.qb.create_ticket(ticket_data)
            if ticket:
                print(f"\n✅ Ticket created: {ticket.get('ticket_number')}")
                print(f"   URL: {ticket.get('quickbase_url')}")
            else:
                print("❌ Failed to create ticket")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    async def check_status(self, ticket_id: str = None) -> None:
        """Check ticket status"""
        
        if not self.qb:
            print("❌ QuickBase not initialized")
            return
        
        if not ticket_id:
            ticket_id = input("Ticket number: ").strip()
        
        if not ticket_id:
            print("❌ Ticket number required")
            return
        
        print(f"\n⏳ Looking up {ticket_id}...")
        
        try:
            ticket = await self.qb.get_ticket(ticket_id)
            if ticket:
                print(f"\n📋 Ticket: {ticket.get('ticket_number')}")
                print(f"   Subject:  {ticket.get('subject')}")
                print(f"   Status:   {ticket.get('status')}")
                print(f"   Priority: {ticket.get('priority')}")
                print(f"   Category: {ticket.get('category')}")
                print(f"   Created:  {ticket.get('submitted_date')}")
                if ticket.get('resolution'):
                    print(f"   Resolution: {ticket.get('resolution')}")
                print(f"   URL: {ticket.get('quickbase_url')}")
            else:
                print(f"❌ Ticket {ticket_id} not found")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    async def show_my_tickets(self) -> None:
        """Show user's open tickets"""
        
        if not self.qb:
            print("❌ QuickBase not initialized")
            return
        
        print(f"\n⏳ Fetching tickets for {self.test_user['email']}...")
        
        try:
            tickets = await self.qb.get_user_tickets(self.test_user['email'])
            
            if tickets:
                print(f"\n📋 Your Open Tickets ({len(tickets)}):\n")
                for t in tickets:
                    print(f"  {t.get('ticket_number'):10} | {t.get('status'):15} | {t.get('subject', '')[:40]}")
            else:
                print("\n✅ No open tickets")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    async def show_stats(self) -> None:
        """Show ticket statistics"""
        
        if not self.qb:
            print("❌ QuickBase not initialized")
            return
        
        print("\n⏳ Fetching statistics...")
        
        try:
            stats = await self.qb.get_ticket_statistics()
            
            print(f"\n📊 Ticket Statistics:\n")
            print(f"   Open tickets:      {stats.get('total_open', 0)}")
            print(f"   Resolved today:    {stats.get('total_resolved_today', 0)}")
            print(f"\n   By Priority:")
            for priority, count in stats.get('by_priority', {}).items():
                print(f"     {priority}: {count}")
        except Exception as e:
            print(f"❌ Error: {e}")
    
    async def test_quickbase(self) -> None:
        """Test QuickBase connection"""
        
        if not self.qb:
            print("❌ QuickBase not initialized")
            return
        
        print("\n⏳ Testing QuickBase connection...")
        
        try:
            stats = await self.qb.get_ticket_statistics()
            print(f"✅ QuickBase connected!")
            print(f"   Realm: {self.qb.realm}")
            print(f"   Table: {self.qb.table_id}")
            print(f"   Open tickets: {stats.get('total_open', 'unknown')}")
        except Exception as e:
            print(f"❌ Connection failed: {e}")
            import traceback
            traceback.print_exc()
    
    async def test_create_ticket(self) -> None:
        """Direct test of QuickBase ticket creation - bypasses chain"""
        
        if not self.qb:
            print("❌ QuickBase not initialized")
            return
        
        print("\n⏳ Testing direct ticket creation...")
        
        ticket_data = {
            'subject': 'TEST - CLI Bot Test Ticket',
            'description': 'This is a test ticket from local_test.py CLI. Please delete.',
            'priority': 'Low',
            'category': 'General Support',
            'user_email': self.test_user['email'],
            'user_name': self.test_user['name'],
        }
        
        print(f"\n   Ticket data: {ticket_data}")
        print(f"\n   QB Realm: {self.qb.realm}")
        print(f"   QB Table: {self.qb.table_id}")
        print(f"   QB App: {self.qb.app_id}")
        
        try:
            print("\n   Calling qb.create_ticket()...")
            ticket = await self.qb.create_ticket(ticket_data)
            
            if ticket:
                print(f"\n✅ SUCCESS! Ticket created:")
                print(f"   Number: {ticket.get('ticket_number')}")
                print(f"   Record ID: {ticket.get('record_id')}")
                print(f"   URL: {ticket.get('quickbase_url')}")
            else:
                print("\n❌ FAILED - create_ticket returned None")
                print("   Check QuickBase field mappings in quickbase_manager.py")
                
        except Exception as e:
            print(f"\n❌ Exception: {e}")
            import traceback
            traceback.print_exc()
    
    async def run(self) -> None:
        """Main CLI loop"""
        
        print("\n" + "="*60)
        print("🤖 IT Support Bot - Interactive CLI")
        print("="*60)
        print("Type a message or /help for commands")
        print("="*60 + "\n")
        
        while True:
            try:
                message = input("You: ").strip()
                await self.process_message(message)
            except KeyboardInterrupt:
                print("\n\n👋 Goodbye!")
                break
            except EOFError:
                break


# ============================================================================
# Main
# ============================================================================

if __name__ == "__main__":
    cli = ITBotCLI()
    asyncio.run(cli.run())