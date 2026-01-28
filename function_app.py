"""
Azure Function - IT Support Bot for Teams
Uses LangChain ITSupportChain for routing and response generation.

Flow:
1. Route via LangChain → Classify intent
2. Search vector store + static KB → Get relevant context
3. GPT ALWAYS generates response (with or without context)
4. ALWAYS respond to user with solution
5. Create ticket (AI checks if follow-up to existing ticket to avoid duplicates)
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


async def get_user_email(activity: Dict[str, Any]) -> str:
    """
    Extract user email from Teams activity.

    The activity 'from' object often doesn't include email or userPrincipalName.
    Falls back to the Teams Bot connector API to fetch the member profile which
    contains the email address.
    """
    user_info = activity.get('from', {})

    # Try direct fields first (sometimes present depending on tenant config)
    email = user_info.get('email') or user_info.get('userPrincipalName', '')
    if email:
        logging.info(f"Got email from activity.from: {email}")
        return email

    # Fall back to Teams API to get full user profile with email
    user_id = user_info.get('id', '')
    if user_id:
        try:
            teams = get_teams_handler()
            member_info = await teams.get_user_info(activity, user_id)
            if member_info:
                email = member_info.get('email') or member_info.get('userPrincipalName', '')
                if email:
                    logging.info(f"Got email from Teams API for user {user_id}: {email}")
                    return email
                else:
                    logging.warning(f"Teams API returned member info but no email for user {user_id}")
        except Exception as e:
            logging.warning(f"Could not fetch user email from Teams API: {e}")

    logging.warning(f"Could not resolve email for user (from.id={user_id})")
    return ''


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

        # Resolve user email from Teams API (from object often lacks email)
        user_email = await get_user_email(activity)
        if user_email:
            user_info['email'] = user_email

        # Remove bot @mentions from message
        user_message = teams.remove_mentions(user_message)

        if not user_message:
            return func.HttpResponse(status_code=200)

        # Handle slash commands directly (fast path)
        if user_message.startswith('/'):
            return await handle_command(user_message, user_info, activity)

        # Show typing indicator while processing
        await teams.send_typing_indicator(activity)

        # Check if this is a follow-up to an existing ticket using AI
        skip_ticket = False

        if user_email:
            try:
                qb = get_qb_manager()
                chain = get_support_chain()

                # Get user's recent open tickets
                recent_tickets = await qb.get_user_tickets(user_email)

                if recent_tickets:
                    # Use AI to determine if this is a follow-up
                    followup_result = chain.is_follow_up(user_message, recent_tickets)
                    skip_ticket = followup_result.get('is_follow_up', False)

                    if skip_ticket:
                        related = followup_result.get('related_ticket', 'existing ticket')
                        logging.info(f"AI detected follow-up to {related}: {followup_result.get('reasoning')}")
                    else:
                        logging.info(f"AI detected new issue: {followup_result.get('reasoning')}")
            except Exception as e:
                logging.warning(f"Follow-up check failed, will create ticket: {e}")
                skip_ticket = False

        # Process through LangChain support chain
        return await handle_support_question(user_message, user_info, activity, skip_ticket=skip_ticket)
        
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
    activity: Dict,
    skip_ticket: bool = False
) -> func.HttpResponse:
    """
    Process IT support question through LangChain.

    ALWAYS:
    1. Generate a response (from vector store context or GPT directly)
    2. Send solution to user
    3. Create a ticket for tracking (unless skip_ticket=True for thread replies)
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
        logging.info(f"Chain result type: {result.get('type')}, confidence: {result.get('confidence')}")
        
    except Exception as e:
        logging.error(f"Chain processing error: {str(e)}")
        # Fallback - still respond and create ticket
        result = {
            "type": "error",
            "solution": get_fallback_response(question),
            "category": "General Support",
            "priority": "Medium",
            "confidence": 0.3,
            "needs_human": True
        }
    
    # Handle different response types
    response_type = result.get('type')
    
    if response_type == 'status_check':
        # User asking about ticket status
        ticket_num = result.get('ticket_number')
        if ticket_num:
            ticket = await qb.get_ticket(ticket_num)
            if ticket:
                status_card = create_ticket_status_card(ticket)
                await teams.send_card(activity, status_card)
            else:
                await teams.send_message(activity, f"Ticket {ticket_num} not found. Use /status to see your open tickets.")
        else:
            # Show user's tickets
            tickets = await qb.get_user_tickets(user_email)
            if tickets:
                list_card = create_ticket_list_card(tickets)
                await teams.send_card(activity, list_card)
            else:
                await teams.send_message(activity, "You have no open tickets. Type your issue and I'll help!")
        return func.HttpResponse(status_code=200)
    
    elif response_type == 'command':
        # Shouldn't hit this (commands handled separately) but just in case
        return await handle_command(question, user_info, activity)
    
    else:
        # 'solution' or 'error' - ALWAYS provide solution
        solution = result.get('solution', '')
        confidence = result.get('confidence', 0.5)
        category = result.get('category', 'General Support')
        priority = result.get('priority', 'Medium')
        needs_human = result.get('needs_human', False)
        sources = result.get('sources', [])
        
        # Ensure we always have a solution
        if not solution or len(solution.strip()) < 10:
            solution = get_fallback_response(question)
            confidence = 0.3
            needs_human = True
        
        # Determine ticket status based on confidence and needs_human flag
        if needs_human or confidence < 0.5:
            ticket_status = 'New'  # IT will review
            ticket_priority = priority
            offer_escalate = False  # Already getting IT attention
        else:
            ticket_status = 'Bot Assisted'  # Logged but low priority
            ticket_priority = 'Low'
            offer_escalate = True  # User can escalate if needed
        
        # ALWAYS send solution card to user
        solution_card = create_solution_card(
            solution=solution,
            question=question,
            category=category,
            confidence=confidence,
            offer_escalate=offer_escalate,
            sources=sources
        )
        await teams.send_card(activity, solution_card)

        # Create ticket for tracking (skip for thread replies to avoid duplicates)
        if skip_ticket:
            logging.info("Skipping ticket creation - this is a reply in an existing thread")
        else:
            ticket_data = {
                'subject': generate_subject(question),
                'description': build_ticket_description(question, solution, sources, confidence),
                'priority': ticket_priority,
                'category': category,
                'status': ticket_status,
                'user_email': user_email,
                'user_name': user_name
            }

            ticket = await qb.create_ticket(ticket_data)
            if ticket:
                logging.info(f"Ticket created: {ticket.get('ticket_number')} (status: {ticket_status}, priority: {ticket_priority})")
            else:
                logging.error("Failed to create tracking ticket")

        return func.HttpResponse(status_code=200)


def get_fallback_response(question: str) -> str:
    """Fallback response when everything else fails"""
    return f"""I'm having trouble processing your request, but here are some general steps:

1. **Restart** the affected application or your computer
2. **Check** if others are experiencing the same issue
3. **Note** any error messages you see
4. **Try** the web version if using a desktop app

Your issue has been logged and IT will follow up: "{question[:80]}..."

In the meantime, try /help for common solutions or /ticket to submit detailed information."""


def build_ticket_description(question: str, solution: str, sources: list, confidence: float) -> str:
    """Build comprehensive ticket description"""
    sources_str = ", ".join(sources) if sources else "GPT General Knowledge"
    
    return f"""**User Question:**
{question}

---
**Bot Response (Confidence: {confidence:.0%}):**
{solution[:500]}{'...' if len(solution) > 500 else ''}

---
**Sources Used:** {sources_str}

---
*Auto-generated by IT Support Bot*"""


def generate_subject(question: str) -> str:
    """Generate concise ticket subject from question"""
    words_to_remove = ['the', 'a', 'an', 'is', 'are', 'was', 'were', 'been', 
                       'have', 'has', 'had', 'i', 'my', 'me', "can't", "cannot", 
                       "won't", "please", "help", "need"]
    
    words = question.lower().split()
    filtered_words = [word for word in words if word not in words_to_remove]
    
    subject = ' '.join(filtered_words[:7]).title()
    
    if len(subject) > 50:
        subject = subject[:47] + '...'
    
    return subject or "IT Support Request"


def create_solution_card(
    solution: str, 
    question: str, 
    category: str, 
    confidence: float = 0.8, 
    offer_escalate: bool = True,
    sources: list = None
) -> Dict:
    """Create adaptive card for bot solution"""
    
    # Header based on confidence
    if confidence >= 0.8:
        header_text = "💡 Here's what I found:"
        header_color = "good"
    elif confidence >= 0.6:
        header_text = "💡 This might help:"
        header_color = "accent"
    else:
        header_text = "💡 While IT reviews this, try:"
        header_color = "warning"
    
    body = [
        {
            "type": "TextBlock",
            "text": header_text,
            "weight": "Bolder",
            "size": "Medium",
            "color": header_color
        },
        {
            "type": "TextBlock",
            "text": solution,
            "wrap": True,
            "spacing": "Medium"
        }
    ]
    
    # Add confidence/status note
    if confidence < 0.7:
        body.append({
            "type": "TextBlock",
            "text": "📋 A ticket has been created. IT will follow up if needed.",
            "wrap": True,
            "isSubtle": True,
            "spacing": "Medium",
            "size": "Small"
        })
    
    # Add sources if available (subtle)
    if sources:
        body.append({
            "type": "TextBlock",
            "text": f"_Sources: {', '.join(sources)}_",
            "wrap": True,
            "isSubtle": True,
            "spacing": "Small",
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
            # User submitted ticket form - resolve email via Teams API
            user_email = await get_user_email(activity)
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
                    "text": "✅ Thanks for the feedback!" if helpful else "📝 Feedback noted. A ticket was created for IT follow-up.",
                    "weight": "Bolder",
                    "color": "Good" if helpful else "Accent"
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
# QuickBase Webhook - Closed Ticket Notification
# =============================================================================

@app.route(route="webhook/ticket-closed", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
async def webhook_ticket_closed(req: func.HttpRequest) -> func.HttpResponse:
    """
    Webhook endpoint for QuickBase to notify when a ticket is closed.

    QuickBase Webhook Configuration:
    1. Go to your QuickBase app settings
    2. Navigate to Webhooks
    3. Create a new webhook with:
       - URL: https://your-function-app.azurewebsites.net/api/webhook/ticket-closed
       - Method: POST
       - Trigger: When record is modified
       - Condition: Status field changes to "Closed"
       - Include fields: ticket_number, subject, status, resolution, submitted_by (email)

    Expected payload format:
    {
        "ticket_number": "IT-240101123456",
        "subject": "Issue subject",
        "status": "Closed",
        "resolution": "Resolution details",
        "submitted_by": "user@company.com",
        "category": "General Support",
        "priority": "Medium"
    }
    """
    logging.info("Received QuickBase webhook for closed ticket")

    try:
        # Validate webhook secret if configured
        webhook_secret = os.environ.get('QB_WEBHOOK_SECRET', '')
        if webhook_secret:
            provided_secret = req.headers.get('X-QB-Webhook-Secret', '')
            if provided_secret != webhook_secret:
                logging.warning("Invalid webhook secret provided")
                return func.HttpResponse(
                    json.dumps({"error": "Unauthorized"}),
                    status_code=401,
                    mimetype="application/json"
                )

        body = req.get_json()
        logging.info(f"Webhook payload: {json.dumps(body)}")

        # Handle QuickBase webhook format (may wrap data differently)
        ticket_data = body
        if 'data' in body:
            ticket_data = body.get('data', {})
        if isinstance(ticket_data, list) and len(ticket_data) > 0:
            ticket_data = ticket_data[0]

        # Extract ticket information
        ticket_number = ticket_data.get('ticket_number', '')
        status = ticket_data.get('status', '')
        user_email = ticket_data.get('submitted_by', '')

        # Only process if status is "Closed"
        if status != 'Closed':
            logging.info(f"Ticket {ticket_number} status is '{status}', not 'Closed'. Skipping notification.")
            return func.HttpResponse(
                json.dumps({"status": "skipped", "reason": "status not Closed"}),
                status_code=200,
                mimetype="application/json"
            )

        if not ticket_number:
            logging.warning("No ticket_number in webhook payload")
            return func.HttpResponse(
                json.dumps({"error": "Missing ticket_number"}),
                status_code=400,
                mimetype="application/json"
            )

        if not user_email:
            logging.warning(f"No user email for ticket {ticket_number}, cannot send notification")
            return func.HttpResponse(
                json.dumps({"status": "skipped", "reason": "no user email"}),
                status_code=200,
                mimetype="application/json"
            )

        # Send notification to user
        notification_sent = await send_closed_ticket_notification(ticket_data, user_email)

        if notification_sent:
            logging.info(f"Closed ticket notification sent for {ticket_number} to {user_email}")
            return func.HttpResponse(
                json.dumps({"status": "success", "ticket_number": ticket_number, "notified": user_email}),
                status_code=200,
                mimetype="application/json"
            )
        else:
            logging.warning(f"Failed to send notification for {ticket_number}")
            return func.HttpResponse(
                json.dumps({"status": "partial", "ticket_number": ticket_number, "notification_sent": False}),
                status_code=200,
                mimetype="application/json"
            )

    except ValueError as e:
        logging.error(f"Invalid JSON in webhook payload: {str(e)}")
        return func.HttpResponse(
            json.dumps({"error": "Invalid JSON payload"}),
            status_code=400,
            mimetype="application/json"
        )
    except Exception as e:
        logging.error(f"Error processing webhook: {str(e)}")
        return func.HttpResponse(
            json.dumps({"error": str(e)}),
            status_code=500,
            mimetype="application/json"
        )


async def send_closed_ticket_notification(ticket_data: Dict[str, Any], user_email: str) -> bool:
    """
    Send a Teams notification to the user that their ticket has been closed.

    Uses proactive messaging to reach the user directly.
    """
    try:
        teams = get_teams_handler()
        cards = get_card_builder()

        # Create the closed ticket notification card
        notification_card = create_closed_ticket_card(ticket_data)

        # Send proactive message to user
        # Note: For proactive messaging to work, the bot must have had a prior conversation with the user
        success = await teams.send_notification_to_user(user_email, notification_card)

        return success

    except Exception as e:
        logging.error(f"Error sending closed ticket notification: {str(e)}")
        return False


def create_closed_ticket_card(ticket_data: Dict[str, Any]) -> Dict:
    """Create adaptive card for closed ticket notification"""
    ticket_number = ticket_data.get('ticket_number', 'N/A')
    subject = ticket_data.get('subject', 'No Subject')
    resolution = ticket_data.get('resolution', 'No resolution details provided.')
    category = ticket_data.get('category', 'General Support')
    priority = ticket_data.get('priority', 'Medium')

    priority_icons = {'Critical': '🔴', 'High': '🟠', 'Medium': '🟡', 'Low': '🟢'}
    priority_icon = priority_icons.get(priority, '⚪')

    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "style": "emphasis",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "✅",
                                        "size": "ExtraLarge"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Ticket Closed",
                                        "weight": "Bolder",
                                        "size": "Large",
                                        "color": "Good"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"Your ticket #{ticket_number} has been resolved",
                                        "size": "Medium",
                                        "isSubtle": True,
                                        "wrap": True
                                    }
                                ]
                            }
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "separator": True,
                "spacing": "Medium",
                "items": [
                    {
                        "type": "FactSet",
                        "facts": [
                            {"title": "Subject:", "value": subject[:50] + ('...' if len(subject) > 50 else '')},
                            {"title": "Category:", "value": category},
                            {"title": "Priority:", "value": f"{priority_icon} {priority}"},
                            {"title": "Status:", "value": "✅ Closed"}
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "separator": True,
                "spacing": "Medium",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "**Resolution:**",
                        "weight": "Bolder"
                    },
                    {
                        "type": "TextBlock",
                        "text": resolution[:500] + ('...' if len(resolution) > 500 else ''),
                        "wrap": True,
                        "spacing": "Small"
                    }
                ]
            },
            {
                "type": "TextBlock",
                "text": "If you have any further questions or the issue persists, please create a new ticket.",
                "wrap": True,
                "isSubtle": True,
                "spacing": "Large",
                "size": "Small"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Create New Ticket",
                "data": {
                    "action": "create_ticket_form"
                }
            },
            {
                "type": "Action.OpenUrl",
                "title": "View in QuickBase",
                "url": ticket_data.get('quickbase_url', '#')
            }
        ]
    }


# =============================================================================
# Health Check
# =============================================================================

@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
async def health_check(req: func.HttpRequest) -> func.HttpResponse:
    """Health check endpoint"""
    
    # Check if chain can be initialized
    chain_status = "unknown"
    try:
        chain = get_support_chain()
        chain_status = "ok"
    except Exception as e:
        chain_status = f"error: {str(e)}"
    
    return func.HttpResponse(
        json.dumps({
            "status": "healthy",
            "timestamp": datetime.utcnow().isoformat(),
            "version": "2.0.0",
            "architecture": "langchain-gpt",
            "chain_status": chain_status,
            "modules": ["support_chain", "teams_handler", "quickbase_manager", "adaptive_cards"]
        }),
        mimetype="application/json",
        status_code=200
    )