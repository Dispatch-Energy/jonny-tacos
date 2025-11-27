"""
Azure Function - IT Support Bot for Teams
Uses LangChain ITSupportChain for routing and response generation.

Flow:
1. Route via LangChain → Classify intent
2. Search vector store (or KB) → Get relevant context  
3. GPT generates response using context
4. ALWAYS respond to user with solution
5. ALWAYS create a ticket for tracking
"""

import azure.functions as func
import logging
import json
import os
from datetime import datetime
from typing import Dict, Any

app = func.FunctionApp()

# Lazy initialization of components
_support_chain = None
_teams_handler = None
_qb_manager = None
_card_builder = None


def get_support_chain():
    """Get or initialize the LangChain support chain"""
    global _support_chain
    if _support_chain is None:
        from support_chain import ITSupportChain
        _support_chain = ITSupportChain()
    return _support_chain


def get_teams_handler():
    global _teams_handler
    if _teams_handler is None:
        from teams_handler import TeamsHandler
        _teams_handler = TeamsHandler()
    return _teams_handler


def get_qb_manager():
    global _qb_manager
    if _qb_manager is None:
        from quickbase_manager import QuickBaseManager
        _qb_manager = QuickBaseManager()
    return _qb_manager


def get_card_builder():
    global _card_builder
    if _card_builder is None:
        from adaptive_cards import AdaptiveCardBuilder
        _card_builder = AdaptiveCardBuilder()
    return _card_builder


# =============================================================================
# Main Messages Endpoint
# =============================================================================

@app.route(route="messages", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
async def messages(req: func.HttpRequest) -> func.HttpResponse:
    """Main Teams bot endpoint"""
    logging.info("Teams bot message received")
    
    try:
        body = req.get_json()
        activity_type = body.get('type')
        
        if activity_type == 'message':
            return await handle_message(body)
        elif activity_type == 'invoke':
            return await handle_invoke(body)
        elif activity_type == 'conversationUpdate':
            return await handle_conversation_update(body)
        else:
            return func.HttpResponse(status_code=200)
            
    except Exception as e:
        logging.error(f"Error processing message: {str(e)}")
        return func.HttpResponse(
            json.dumps({"error": str(e)}),
            status_code=500,
            mimetype="application/json"
        )


# =============================================================================
# Message Handler
# =============================================================================

async def handle_message(activity: Dict[str, Any]) -> func.HttpResponse:
    """Handle incoming text messages"""
    try:
        teams = get_teams_handler()
        
        user_message = activity.get('text', '').strip()
        user_info = activity.get('from', {})
        
        # Remove bot @mentions from message
        user_message = teams.remove_mentions(user_message)
        
        if not user_message:
            return func.HttpResponse(status_code=200)
        
        # Handle slash commands directly (fast path)
        if user_message.startswith('/'):
            return await handle_command(user_message, user_info, activity)
        
        # Show typing indicator while processing
        await teams.send_typing_indicator(activity)
        
        # Process through LangChain support chain
        return await handle_support_question(user_message, user_info, activity)
        
    except Exception as e:
        logging.error(f"Error handling message: {str(e)}")
        teams = get_teams_handler()
        cards = get_card_builder()
        error_card = cards.create_error_card(
            "Something went wrong. Please try /ticket to create a support request."
        )
        await teams.send_card(activity, error_card)
        return func.HttpResponse(status_code=200)


# =============================================================================
# Support Question Handler - Uses LangChain
# =============================================================================

async def handle_support_question(
    question: str, 
    user_info: Dict, 
    activity: Dict
) -> func.HttpResponse:
    """
    Process IT support question through LangChain.
    
    ALWAYS:
    1. Generate a response (from vector store context or GPT directly)
    2. Send solution to user
    3. Create a ticket for tracking
    """
    
    teams = get_teams_handler()
    chain = get_support_chain()
    qb = get_qb_manager()
    cards = get_card_builder()
    
    user_email = user_info.get('email') or user_info.get('userPrincipalName', '')
    user_name = user_info.get('name', 'Unknown User')
    
    try:
        # Process through LangChain - this handles routing and response generation
        result = chain.process(question)
        logging.info(f"Chain result type: {result.get('type')}")
        
    except Exception as e:
        logging.error(f"Chain processing error: {str(e)}")
        # Fallback - still respond and create ticket
        result = {
            "type": "error",
            "solution": "I encountered an issue processing your request, but I've logged it for IT to review.",
            "category": "General Support",
            "priority": "Medium"
        }
    
    # Extract response data based on result type
    response_type = result.get('type')
    
    if response_type == 'solution':
        # Quick fix - we have a solution from KB/vector store + GPT
        solution = result.get('solution', '')
        confidence = result.get('confidence', 0.8)
        category = result.get('category', 'General Support')
        priority = result.get('priority', 'Low')
        offer_escalate = result.get('offer_ticket', True)
        ticket_status = 'Bot Assisted'
        
    elif response_type == 'ticket_needed':
        # Router determined this needs IT - but we still respond with context
        rec = result.get('recommendation', {})
        solution = f"This issue requires IT assistance.\n\nI've analyzed your request:\n• Category: {rec.get('category', 'General Support')}\n• Priority: {rec.get('priority', 'Medium')}\n• Reason: {rec.get('reasoning', 'Complex issue')}\n\nA ticket has been created for you."
        confidence = 0.5
        category = rec.get('category', 'General Support')
        priority = rec.get('priority', 'Medium')
        offer_escalate = False  # Already creating a real ticket
        ticket_status = 'New'
        
    elif response_type == 'troubleshooting_needed':
        # Complex issue - respond with what we know, create ticket
        rec = result.get('recommendation', {})
        solution = f"This looks like a complex issue that may need investigation.\n\nWhile IT reviews this, you can try:\n1. Restart the affected application\n2. Check if others are experiencing the same issue\n3. Note any error messages you see\n\nI've created a ticket for IT to follow up."
        confidence = 0.4
        category = rec.get('category', 'General Support')
        priority = rec.get('priority', 'Medium')
        offer_escalate = False
        ticket_status = 'New'
        
    elif response_type == 'status_check':
        # User asking about ticket status - handle separately
        ticket_num = result.get('ticket_number')
        if ticket_num:
            ticket = await qb.get_ticket(ticket_num)
            if ticket:
                status_card = create_ticket_status_card(ticket)
                await teams.send_card(activity, status_card)
            else:
                await teams.send_message(activity, f"Ticket {ticket_num} not found.")
        else:
            # Show user's tickets
            tickets = await qb.get_user_tickets(user_email)
            if tickets:
                list_card = create_ticket_list_card(tickets)
                await teams.send_card(activity, list_card)
            else:
                await teams.send_message(activity, "You have no open tickets. Type your issue and I'll help!")
        return func.HttpResponse(status_code=200)
        
    else:
        # Fallback for any other type
        solution = "I'll help you with that. Let me create a ticket to track this request."
        confidence = 0.3
        category = 'General Support'
        priority = 'Medium'
        offer_escalate = False
        ticket_status = 'New'
    
    # ALWAYS send solution card to user
    solution_card = create_solution_card(
        solution=solution,
        question=question,
        category=category,
        confidence=confidence,
        offer_escalate=offer_escalate
    )
    await teams.send_card(activity, solution_card)
    
    # ALWAYS create ticket for tracking
    ticket_data = {
        'subject': generate_subject(question),
        'description': f"User question: {question}\n\n---\nBot response:\n{solution[:500]}",
        'priority': priority if ticket_status == 'New' else 'Low',
        'category': category,
        'status': ticket_status,
        'user_email': user_email,
        'user_name': user_name
    }
    
    ticket = await qb.create_ticket(ticket_data)
    if ticket:
        logging.info(f"Ticket created: {ticket.get('ticket_number')} (status: {ticket_status})")
    else:
        logging.error("Failed to create tracking ticket")
    
    return func.HttpResponse(status_code=200)


def generate_subject(question: str) -> str:
    """Generate concise ticket subject from question"""
    # Remove common words and limit length
    words_to_remove = ['the', 'a', 'an', 'is', 'are', 'was', 'were', 'been', 'have', 'has', 'had', 'i', 'my', 'me', "can't", "cannot", "won't"]
    
    words = question.lower().split()
    filtered_words = [word for word in words if word not in words_to_remove]
    
    subject = ' '.join(filtered_words[:7]).title()
    
    if len(subject) > 50:
        subject = subject[:47] + '...'
    
    return subject or "IT Support Request"


def create_solution_card(solution: str, question: str, category: str, confidence: float = 0.8, offer_escalate: bool = True) -> Dict:
    """Create adaptive card for bot solution"""
    body = [
        {
            "type": "TextBlock",
            "text": "💡 Here's what I found:",
            "weight": "Bolder",
            "size": "Medium"
        },
        {
            "type": "TextBlock",
            "text": solution,
            "wrap": True,
            "spacing": "Medium"
        }
    ]
    
    # Add confidence note if lower confidence
    if confidence < 0.7:
        body.append({
            "type": "TextBlock",
            "text": "ℹ️ A ticket has been created. IT will follow up if needed.",
            "wrap": True,
            "isSubtle": True,
            "spacing": "Medium",
            "size": "Small"
        })
    
    actions = [
        {
            "type": "Action.Submit",
            "title": "✅ This helped",
            "style": "positive",
            "data": {
                "action": "solution_feedback",
                "helpful": True,
                "question": question[:200]
            }
        }
    ]
    
    if offer_escalate:
        actions.append({
            "type": "Action.Submit",
            "title": "🎫 Still need help",
            "data": {
                "action": "escalate_ticket",
                "question": question[:200],
                "category": category
            }
        })
    
    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": body,
        "actions": actions
    }


# =============================================================================
# Command Handler
# =============================================================================

async def handle_command(
    command: str, 
    user_info: Dict, 
    activity: Dict
) -> func.HttpResponse:
    """Handle /slash commands - fast path, no LangChain needed"""
    
    teams = get_teams_handler()
    cards = get_card_builder()
    qb = get_qb_manager()
    
    parts = command.split()
    cmd = parts[0].lower()
    
    if cmd == '/help':
        help_card = cards.create_help_card()
        await teams.send_card(activity, help_card)
        
    elif cmd == '/ticket':
        ticket_form = cards.create_ticket_form()
        await teams.send_card(activity, ticket_form)
        
    elif cmd == '/status':
        ticket_num = parts[1] if len(parts) > 1 else None
        user_email = user_info.get('email') or user_info.get('userPrincipalName', '')
        
        if ticket_num:
            ticket = await qb.get_ticket(ticket_num)
            if ticket:
                status_card = create_ticket_status_card(ticket)
                await teams.send_card(activity, status_card)
            else:
                await teams.send_message(activity, f"Ticket {ticket_num} not found.")
        else:
            tickets = await qb.get_user_tickets(user_email)
            if tickets:
                list_card = create_ticket_list_card(tickets)
                await teams.send_card(activity, list_card)
            else:
                await teams.send_message(activity, "You have no open tickets.")
                
    elif cmd == '/stats':
        stats = await qb.get_ticket_statistics()
        if hasattr(cards, 'create_statistics_card'):
            stats_card = cards.create_statistics_card(stats)
            await teams.send_card(activity, stats_card)
        else:
            by_priority = stats.get('by_priority', {})
            stats_text = f"📊 **Ticket Stats**\n• Open: {stats.get('total_open', 0)}\n• Resolved today: {stats.get('total_resolved_today', 0)}\n• Critical: {by_priority.get('Critical', 0)} | High: {by_priority.get('High', 0)}"
            await teams.send_message(activity, stats_text)
        
    else:
        await teams.send_message(activity, f"Unknown command: {cmd}. Try /help")
    
    return func.HttpResponse(status_code=200)


# =============================================================================
# Invoke Handler (Adaptive Card Actions)
# =============================================================================

async def handle_invoke(activity: Dict[str, Any]) -> func.HttpResponse:
    """Handle adaptive card button clicks"""
    try:
        action_data = activity.get('value', {})
        action_type = action_data.get('action')
        user_info = activity.get('from', {})
        
        teams = get_teams_handler()
        qb = get_qb_manager()
        cards = get_card_builder()
        
        if action_type == 'create_ticket':
            # User submitted ticket form
            user_email = user_info.get('email') or user_info.get('userPrincipalName', '')
            user_name = user_info.get('name', 'Unknown User')
            
            ticket_data = {
                'subject': action_data.get('subject', 'No Subject'),
                'description': action_data.get('description', ''),
                'priority': action_data.get('priority', 'Medium'),
                'category': action_data.get('category', 'General Support'),
                'status': 'New',
                'user_email': user_email,
                'user_name': user_name
            }
            
            if action_data.get('additional_info'):
                ticket_data['description'] += f"\n\nAdditional info: {action_data['additional_info']}"
            
            ticket = await qb.create_ticket(ticket_data)
            
            if ticket:
                confirmation_card = cards.create_ticket_confirmation_card(ticket)
                await teams.update_card(activity, confirmation_card)
                await notify_it_channel(ticket)
            else:
                await teams.send_message(activity, "❌ Failed to create ticket. Please try again.")
        
        elif action_type == 'escalate_ticket':
            # User wants to escalate after bot solution didn't help
            question = action_data.get('question', 'Issue not resolved')
            category = action_data.get('category', 'General Support')
            
            ticket_form = cards.create_ticket_form(
                subject=generate_subject(question),
                description=f"{question}\n\n[User tried self-service but still needs help]",
                category=category,
                priority='Medium'
            )
            await teams.update_card(activity, ticket_form)
        
        elif action_type == 'solution_feedback':
            helpful = action_data.get('helpful', False)
            question = action_data.get('question', '')
            logging.info(f"Solution feedback: helpful={helpful}, question={question[:50]}")
            
            thanks_card = {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [{
                    "type": "TextBlock",
                    "text": "✅ Thanks for the feedback!",
                    "weight": "Bolder",
                    "color": "Good"
                }]
            }
            await teams.update_card(activity, thanks_card)
        
        elif action_type == 'check_status':
            ticket_num = action_data.get('ticket_number')
            if ticket_num:
                ticket = await qb.get_ticket(ticket_num)
                if ticket:
                    status_card = create_ticket_status_card(ticket)
                    await teams.send_card(activity, status_card)
        
        elif action_type == 'help':
            help_card = cards.create_help_card()
            await teams.send_card(activity, help_card)
        
        elif action_type == 'cancel':
            cancel_card = {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [{
                    "type": "TextBlock",
                    "text": "Cancelled. Let me know if you need anything else!",
                    "wrap": True
                }]
            }
            await teams.update_card(activity, cancel_card)
        
        return func.HttpResponse(
            json.dumps({"status": "ok"}),
            mimetype="application/json",
            status_code=200
        )
        
    except Exception as e:
        logging.error(f"Error handling invoke: {str(e)}")
        return func.HttpResponse(
            json.dumps({"error": str(e)}),
            mimetype="application/json",
            status_code=500
        )


# =============================================================================
# Conversation Update Handler
# =============================================================================

async def handle_conversation_update(activity: Dict[str, Any]) -> func.HttpResponse:
    """Handle bot being added to channel/chat"""
    try:
        members_added = activity.get('membersAdded', [])
        bot_id = activity.get('recipient', {}).get('id')
        
        for member in members_added:
            if member.get('id') == bot_id:
                teams = get_teams_handler()
                cards = get_card_builder()
                welcome_card = cards.create_welcome_card()
                await teams.send_card(activity, welcome_card)
                break
                
    except Exception as e:
        logging.error(f"Error handling conversation update: {str(e)}")
    
    return func.HttpResponse(status_code=200)


# =============================================================================
# Helper Functions
# =============================================================================

async def notify_it_channel(ticket: Dict) -> None:
    """Send notification to IT support channel"""
    it_channel_id = os.environ.get('IT_CHANNEL_ID', '')
    if not it_channel_id:
        return
    
    try:
        teams = get_teams_handler()
        cards = get_card_builder()
        
        if hasattr(cards, 'create_it_notification_card'):
            notification_card = cards.create_it_notification_card(ticket)
            await teams.send_to_channel(it_channel_id, notification_card)
        else:
            logging.info(f"New ticket notification: {ticket.get('ticket_number')}")
    except Exception as e:
        logging.error(f"Error notifying IT channel: {str(e)}")


def create_ticket_status_card(ticket: Dict) -> Dict:
    """Create status card for a ticket"""
    priority_icons = {'Critical': '🔴', 'High': '🟠', 'Medium': '🟡', 'Low': '🟢'}
    
    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "text": f"📋 Ticket {ticket.get('ticket_number', 'N/A')}",
                "weight": "Bolder",
                "size": "Medium"
            },
            {
                "type": "FactSet",
                "facts": [
                    {"title": "Subject:", "value": ticket.get('subject', 'N/A')},
                    {"title": "Status:", "value": ticket.get('status', 'N/A')},
                    {"title": "Priority:", "value": f"{priority_icons.get(ticket.get('priority', ''), '⚪')} {ticket.get('priority', 'N/A')}"},
                    {"title": "Category:", "value": ticket.get('category', 'N/A')},
                    {"title": "Created:", "value": ticket.get('submitted_date', 'N/A')[:10] if ticket.get('submitted_date') else 'N/A'}
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.OpenUrl",
                "title": "View in QuickBase",
                "url": ticket.get('quickbase_url', '#')
            }
        ]
    }


def create_ticket_list_card(tickets: list) -> Dict:
    """Create card listing multiple tickets"""
    items = []
    
    for t in tickets[:5]:
        items.append({
            "type": "TextBlock",
            "text": f"**{t.get('ticket_number')}** - {t.get('status')} - {t.get('subject', '')[:40]}",
            "wrap": True,
            "spacing": "Small"
        })
    
    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "text": f"📋 Your Open Tickets ({len(tickets)})",
                "weight": "Bolder",
                "size": "Medium"
            },
            {
                "type": "Container",
                "items": items
            }
        ]
    }


# =============================================================================
# Health Check
# =============================================================================

@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
async def health_check(req: func.HttpRequest) -> func.HttpResponse:
    """Health check endpoint"""
    return func.HttpResponse(
        json.dumps({
            "status": "healthy",
            "timestamp": datetime.utcnow().isoformat(),
            "version": "1.0.0",
            "architecture": "langchain",
            "modules": ["support_chain", "teams_handler", "quickbase_manager", "adaptive_cards"]
        }),
        mimetype="application/json",
        status_code=200
    )
