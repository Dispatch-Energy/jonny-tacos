"""
Azure Function - IT Support Bot for Teams
Uses LangChain ITSupportChain for routing and response generation.

Flow:
1. Route via LangChain → Classify intent
2. Streaming topic detection → Check if follow-up to existing conversation
   (conversation stream tracking + AI semantic analysis to prevent duplicate tickets)
3. Search static KB → Get relevant context
4. GPT ALWAYS generates response (with or without context)
5. ALWAYS respond to user with solution (with clear IT Admin action disclaimers)
6. Create a ticket only when IT action is needed or the user explicitly asks for one
"""

import azure.functions as func
import logging
import json
import os
import re
from datetime import datetime
from typing import Dict, Any, Optional, Tuple

app = func.FunctionApp()

# Lazy initialization of components
_support_chain = None
_teams_handler = None
_qb_manager = None
_card_builder = None
_automation_manager = None


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


def get_automation_manager():
    """Get or initialize the automation manager with registered handlers."""
    global _automation_manager
    if _automation_manager is None:
        from automation_manager import AutomationManager
        from m365_provisioning import M365ProvisioningHandler
        _automation_manager = AutomationManager()
        _automation_manager.register_handler(M365ProvisioningHandler())
        # Register additional automation handlers here as they are built:
        # _automation_manager.register_handler(SomeOtherHandler())
    return _automation_manager


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


def extract_on_behalf_of_email(message: str, sender_email: str) -> Tuple[Optional[str], str]:
    """
    Extract an email address from the message text to file a ticket on behalf of someone else.

    If the message contains an email address that differs from the sender's,
    it's treated as an "on behalf of" request. The email is stripped from the
    message so the remaining text is processed as the issue description.

    Returns:
        (target_email, cleaned_message) - target_email is None if not on-behalf-of
    """
    # Find email addresses in the message
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    found_emails = re.findall(email_pattern, message)

    if not found_emails:
        return None, message

    for email in found_emails:
        if email.lower() != sender_email.lower():
            # Found an email that isn't the sender's - this is an on-behalf-of request
            cleaned = message.replace(email, '').strip()
            # Clean up any extra whitespace left behind
            cleaned = re.sub(r'\s{2,}', ' ', cleaned).strip()
            # Strip leading "ticket" keyword and punctuation/quotes since the
            # intent is already clear (filing on behalf of someone)
            cleaned = re.sub(r'^ticket\s*[,:\-]?\s*', '', cleaned, flags=re.IGNORECASE).strip()
            cleaned = cleaned.strip('"\'').strip()
            logging.info(f"On-behalf-of detected: filing ticket for {email} (submitted by {sender_email})")
            return email, cleaned

    return None, message


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
            # Check if this is a card Action.Submit that Teams sent as a message
            # (Teams sends Action.Submit with input fields as type "message",
            # not "invoke", so card actions arrive here with empty text but
            # action data in activity.value)
            value_data = activity.get('value', {})
            if isinstance(value_data, dict) and value_data.get('action'):
                logging.info(f"Card action received as message type: {value_data.get('action')}")
                return await handle_invoke(activity)
            return func.HttpResponse(status_code=200)

        # Handle slash commands directly (fast path)
        if user_message.startswith('/'):
            return await handle_command(user_message, user_info, activity)

        # Show typing indicator while processing
        await teams.send_typing_indicator(activity)

        # Check if filing on behalf of someone else (message contains another user's email)
        on_behalf_of_email, cleaned_message = extract_on_behalf_of_email(user_message, user_email)
        if on_behalf_of_email:
            user_message = cleaned_message

        # Streaming topic detection: check if this is a follow-up to an existing
        # ticket using conversation stream tracking + AI analysis.
        # This prevents duplicate tickets from chatty back-and-forth conversations
        # (e.g. "create a reporting@ email" → "did you create it?" → "any update?")
        skip_ticket = False
        related_ticket = None

        if user_email:
            try:
                qb = get_qb_manager()
                chain = get_support_chain()

                # Get user's recent open tickets
                recent_tickets = await qb.get_user_tickets(user_email)

                if recent_tickets:
                    # Use streaming topic detection (conversation stream + AI)
                    followup_result = chain.is_follow_up(
                        user_message, recent_tickets, user_email=user_email
                    )
                    skip_ticket = followup_result.get('is_follow_up', False)
                    related_ticket = followup_result.get('related_ticket')

                    if skip_ticket:
                        logging.info(
                            f"Stream detection: follow-up to {related_ticket or 'existing ticket'} "
                            f"- {followup_result.get('reasoning')}"
                        )
                    else:
                        logging.info(f"Stream detection: new issue - {followup_result.get('reasoning')}")
            except Exception as e:
                logging.warning(f"Follow-up check failed, will create ticket: {e}")
                skip_ticket = False

        # Process through LangChain support chain
        return await handle_support_question(
            user_message, user_info, activity,
            skip_ticket=skip_ticket, related_ticket=related_ticket,
            on_behalf_of=on_behalf_of_email
        )
        
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
    skip_ticket: bool = False,
    related_ticket: str = None,
    on_behalf_of: str = None
) -> func.HttpResponse:
    """
    Process IT support question through LangChain.

    ALWAYS:
    1. Generate a response (from vector store context or GPT directly)
    2. Send solution to user
    3. Create a ticket only when the request needs IT follow-up, the user
       explicitly asks for one, or the bot has low confidence

    If on_behalf_of is set, the ticket's "Submitted By" field is set to
    that email so the ticket shows up under the target user in QuickBase.
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

    if response_type == 'automation_request':
        # Route to automation flow (M365 provisioning, etc.)
        return await start_automation_flow(question, user_info, activity)

    if response_type == 'status_check' and not on_behalf_of:
        # User asking about ticket status (skip when filing on behalf of someone)
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
            ticket_status = 'New'  # Only used when user explicitly asks for a ticket
            ticket_priority = 'Low'
            offer_escalate = True  # User can escalate if needed

        explicit_ticket_request = is_explicit_ticket_request(question)
        should_create_ticket = should_auto_create_ticket(
            question=question,
            needs_human=needs_human,
            confidence=confidence,
            on_behalf_of=on_behalf_of
        )

        # Create ticket only when warranted (skip for follow-ups to avoid duplicates)
        ticket_number = None
        ticket_state = "not_created"
        if skip_ticket:
            ticket_number = related_ticket
            ticket_state = "existing" if related_ticket else "not_created"
            logging.info(
                f"Skipping ticket creation - follow-up to {related_ticket or 'existing conversation'}"
            )
        elif should_create_ticket:
            # When filing on behalf of someone else, the ticket's "Submitted By"
            # is set to the target user so it appears under their name in QuickBase.
            ticket_email = on_behalf_of if on_behalf_of else user_email
            description = build_ticket_description(
                question, solution, sources, confidence,
                on_behalf_of=on_behalf_of, filed_by=user_email if on_behalf_of else None
            )

            ticket_data = {
                'subject': generate_subject(question),
                'description': description,
                'priority': ticket_priority,
                'category': category,
                'status': ticket_status,
                'user_email': ticket_email,
                'user_name': user_name
            }

            ticket = await qb.create_ticket(ticket_data)
            if ticket:
                ticket_number = ticket.get('ticket_number')
                ticket_state = "created"
                logging.info(f"Ticket created: {ticket_number} (status: {ticket_status}, priority: {ticket_priority})")
            else:
                logging.error("Failed to create tracking ticket")
                ticket_state = "required_failed"
                offer_escalate = True
        else:
            logging.info(
                "No ticket created - continuing conversational support "
                f"(confidence={confidence:.0%}, needs_human={needs_human}, explicit_ticket_request={explicit_ticket_request})"
            )

        # Send solution card to user (after ticket creation so we can include the ticket number)
        solution_card = create_solution_card(
            solution=solution,
            question=question,
            category=category,
            confidence=confidence,
            offer_escalate=offer_escalate,
            sources=sources,
            needs_human=needs_human,
            ticket_number=ticket_number,
            ticket_state=ticket_state
        )
        await teams.send_card(activity, solution_card)

        # Record this message in the conversation stream for future follow-up detection
        if user_email:
            chain.conversation_stream.record_message(user_email, question, ticket_number)

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


def build_ticket_description(
    question: str, solution: str, sources: list, confidence: float,
    on_behalf_of: str = None, filed_by: str = None
) -> str:
    """Build comprehensive ticket description"""
    sources_str = ", ".join(sources) if sources else "GPT General Knowledge"

    behalf_section = ""
    if on_behalf_of and filed_by:
        behalf_section = f"""
---
**Filed on behalf of:** {on_behalf_of}
**Filed by:** {filed_by}
"""

    return f"""**User Question:**
{question}
{behalf_section}
---
**Bot Response (Confidence: {confidence:.0%}):**
{solution[:500]}{'...' if len(solution) > 500 else ''}

---
**Sources Used:** {sources_str}

---
*Auto-generated by IT Support Bot*"""


def generate_subject(question: str) -> str:
    """Generate concise ticket subject from question"""
    # Clean up the question and use it directly as the subject
    subject = question.strip().rstrip('?!.').strip()

    # Capitalize the first letter
    if subject:
        subject = subject[0].upper() + subject[1:]

    if len(subject) > 50:
        subject = subject[:47] + '...'

    return subject or "IT Support Request"


def is_explicit_ticket_request(question: str) -> bool:
    """Return True when the user clearly asks to open a ticket."""
    patterns = [
        r'\b(create|open|submit|file|log)\s+(a\s+)?ticket\b',
        r'\bescalate\b',
        r'\bneed\s+(it|human)\s+help\b'
    ]
    question_lower = question.lower()
    return any(re.search(pattern, question_lower) for pattern in patterns)


def should_auto_create_ticket(
    question: str,
    needs_human: bool,
    confidence: float,
    on_behalf_of: Optional[str] = None
) -> bool:
    """
    Decide when the bot should open a QuickBase ticket automatically.

    High-confidence self-service answers stay conversational by default.
    Tickets are opened when a human needs to act, confidence is low, the
    user explicitly requests escalation, or the message is filed on behalf
    of another person.
    """
    return any([
        needs_human,
        confidence < 0.5,
        bool(on_behalf_of),
        is_explicit_ticket_request(question)
    ])


def create_solution_card(
    solution: str,
    question: str,
    category: str,
    confidence: float = 0.8,
    offer_escalate: bool = True,
    sources: list = None,
    needs_human: bool = False,
    ticket_number: str = None,
    ticket_state: str = "not_created"
) -> Dict:
    """Create adaptive card for bot solution"""

    # Header based on confidence and whether IT Admin action is needed
    if needs_human:
        header_text = "📋 Here's the plan"
        header_color = "accent"
    elif confidence >= 0.8:
        header_text = "💡 Here's what I'd try next"
        header_color = "good"
    elif confidence >= 0.6:
        header_text = "💡 Try this next"
        header_color = "accent"
    else:
        header_text = "💡 A couple things to try"
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

    if ticket_state == "created" and ticket_number:
        body.append({
            "type": "Container",
            "style": "emphasis",
            "spacing": "Medium",
            "items": [
                {
                    "type": "TextBlock",
                    "text": f"🎫 **Ticket opened: {ticket_number}**",
                    "wrap": True,
                    "weight": "Bolder",
                    "size": "Small"
                },
                {
                    "type": "TextBlock",
                    "text": (
                        "I opened this in QuickBase and will keep follow-up replies tied to the same request. "
                        "This needs an IT Admin to make the change, and they will review it during business hours."
                        if needs_human else
                        "I logged this in QuickBase so there is a clear handoff if you still need IT involved. "
                        "Reply below and I'll keep follow-ups tied to this ticket."
                    ),
                    "wrap": True,
                    "isSubtle": True,
                    "size": "Small",
                    "spacing": "Small"
                }
            ]
        })
    elif ticket_state == "existing" and ticket_number:
        body.append({
            "type": "Container",
            "style": "emphasis",
            "spacing": "Medium",
            "items": [
                {
                    "type": "TextBlock",
                    "text": f"🧵 **Continuing ticket: {ticket_number}**",
                    "wrap": True,
                    "weight": "Bolder",
                    "size": "Small"
                },
                {
                    "type": "TextBlock",
                    "text": "This looks like the same request, so I'm keeping the conversation on the existing ticket instead of opening a new one.",
                    "wrap": True,
                    "isSubtle": True,
                    "size": "Small",
                    "spacing": "Small"
                }
            ]
        })
    elif ticket_state == "required_failed":
        body.append({
            "type": "Container",
            "style": "attention",
            "spacing": "Medium",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "⚠️ **I couldn't open the ticket automatically**",
                    "wrap": True,
                    "weight": "Bolder",
                    "size": "Small"
                },
                {
                    "type": "TextBlock",
                    "text": "This needs IT follow-up, but QuickBase ticket creation failed. Use **Still need help** or `/ticket` so the request does not get lost.",
                    "wrap": True,
                    "isSubtle": True,
                    "size": "Small",
                    "spacing": "Small"
                }
            ]
        })
    elif offer_escalate:
        body.append({
            "type": "Container",
            "style": "emphasis",
            "spacing": "Medium",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "💬 **No ticket yet**",
                    "wrap": True,
                    "weight": "Bolder",
                    "size": "Small"
                },
                {
                    "type": "TextBlock",
                    "text": "I'm keeping this conversational for now. Reply below if you want to keep troubleshooting, or use **Still need help** if you want me to open a ticket.",
                    "wrap": True,
                    "isSubtle": True,
                    "size": "Small",
                    "spacing": "Small"
                }
            ]
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

    # Add inline reply chat bar to encourage continued conversation
    body.append({
        "type": "Container",
        "separator": True,
        "spacing": "Medium",
        "items": [
            {
                "type": "TextBlock",
                "text": "💬 **Reply here to keep this on the same thread:**",
                "wrap": True,
                "size": "Small",
                "weight": "Bolder"
            },
            {
                "type": "Input.Text",
                "id": "reply_message",
                "placeholder": "Ask a follow-up question, describe what you tried, or tell me what's still not working...",
                "isMultiline": True,
                "maxLength": 500
            }
        ]
    })

    actions = [
        {
            "type": "Action.Submit",
            "title": "💬 Send Reply",
            "style": "positive",
            "data": {
                "action": "reply_to_solution",
                "original_question": question[:200],
                "category": category,
                "ticket_number": ticket_number or ""
            }
        }
    ]

    if ticket_number:
        actions.append({
            "type": "Action.Submit",
            "title": "📋 Check Ticket",
            "data": {
                "action": "check_status",
                "ticket_number": ticket_number
            }
        })

    actions.append({
        "type": "Action.Submit",
        "title": "✅ This helped",
        "data": {
            "action": "solution_feedback",
            "helpful": True,
            "question": question[:200]
        }
    })

    if offer_escalate and ticket_state not in ("created", "existing"):
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

        # Handle provisioning/automation actions first
        if action_type and action_type.startswith('provisioning_'):
            handled = await handle_provisioning_invoke(
                action_type, action_data, activity, user_info
            )
            if handled:
                return func.HttpResponse(
                    json.dumps({"status": "ok"}),
                    mimetype="application/json",
                    status_code=200
                )

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
        
        elif action_type == 'reply_to_solution':
            # User sent a follow-up reply from the inline chat bar on the solution card
            reply_text = action_data.get('reply_message', '').strip()
            original_question = action_data.get('original_question', '')
            reply_category = action_data.get('category', 'General Support')
            reply_ticket = action_data.get('ticket_number', '')

            if not reply_text:
                # Empty reply - nudge user to type something
                await teams.send_message(
                    activity,
                    "Please type your follow-up question in the reply box and click **Send Reply**."
                )
            else:
                # Resolve user email
                user_email = await get_user_email(activity)
                user_name = user_info.get('name', 'Unknown User')

                # Show typing indicator while processing
                await teams.send_typing_indicator(activity)

                # Process the reply through the support chain as a follow-up
                chain = get_support_chain()
                try:
                    result = chain.process(reply_text)
                except Exception as e:
                    logging.error(f"Chain error on reply: {e}")
                    result = {
                        "solution": get_fallback_response(reply_text),
                        "category": reply_category,
                        "priority": "Medium",
                        "confidence": 0.3,
                        "needs_human": True,
                        "sources": []
                    }

                reply_solution = result.get('solution', get_fallback_response(reply_text))
                reply_confidence = result.get('confidence', 0.5)
                reply_category = result.get('category', reply_category)
                reply_priority = result.get('priority', 'Medium')
                reply_sources = result.get('sources', [])
                reply_needs_human = result.get('needs_human', False)
                reply_ticket_state = "existing" if reply_ticket else "not_created"

                if not reply_ticket and should_auto_create_ticket(
                    question=reply_text,
                    needs_human=reply_needs_human,
                    confidence=reply_confidence
                ):
                    ticket_description = build_ticket_description(
                        question=(
                            f"Original request: {original_question}\n\n"
                            f"Follow-up: {reply_text}"
                            if original_question else reply_text
                        ),
                        solution=reply_solution,
                        sources=reply_sources,
                        confidence=reply_confidence
                    )
                    ticket_data = {
                        'subject': generate_subject(original_question or reply_text),
                        'description': ticket_description,
                        'priority': reply_priority if reply_needs_human or reply_confidence < 0.5 else 'Low',
                        'category': reply_category,
                        'status': 'New',
                        'user_email': user_email,
                        'user_name': user_name
                    }
                    created_ticket = await qb.create_ticket(ticket_data)
                    if created_ticket:
                        reply_ticket = created_ticket.get('ticket_number', '')
                        reply_ticket_state = "created"
                        await notify_it_channel(created_ticket)
                    else:
                        reply_ticket_state = "required_failed"
                offer_reply_escalate = (not reply_needs_human) or reply_ticket_state == "required_failed"

                # Send a new solution card as a reply (keeping the conversation threaded)
                follow_up_card = create_solution_card(
                    solution=reply_solution,
                    question=reply_text,
                    category=reply_category,
                    confidence=reply_confidence,
                    offer_escalate=offer_reply_escalate,
                    sources=reply_sources,
                    needs_human=reply_needs_human,
                    ticket_number=reply_ticket,
                    ticket_state=reply_ticket_state
                )
                await teams.send_card(activity, follow_up_card)

                # Record in conversation stream for follow-up detection
                if user_email:
                    chain.conversation_stream.record_message(
                        user_email, reply_text, reply_ticket or None
                    )

                logging.info(
                    f"Reply processed for ticket {reply_ticket or 'N/A'}: "
                    f"'{reply_text[:50]}' (confidence: {reply_confidence:.0%})"
                )

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
        
        elif action_type == 'create_ticket_form':
            # User clicked "Create New Ticket" from a notification card
            ticket_form = cards.create_ticket_form()
            await teams.send_card(activity, ticket_form)

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


def parse_webhook_body(req: func.HttpRequest) -> Dict[str, Any]:
    """Parse webhook request body, handling QuickBase's non-standard JSON format.

    QuickBase field reference tokens (e.g. [Ticket Number]) get substituted with
    raw values that may not be properly quoted, producing invalid JSON. This
    function falls back to fixing unquoted values when standard JSON parsing fails.
    """
    # Try standard JSON first
    try:
        return req.get_json()
    except ValueError:
        pass

    # Fall back to raw body - attempt to fix unquoted string values
    raw = req.get_body().decode('utf-8').strip()
    logging.info(f"Raw webhook body (JSON parse failed): {raw}")

    if raw.startswith('{'):
        # Quote unquoted string values: match "key": value where value isn't
        # already quoted, a number, a boolean, null, or a nested structure
        fixed = re.sub(
            r'("[\w]+")\s*:\s*(?!"|-?\d|true|false|null|\[|\{)([^,}\n]+)',
            lambda m: f'{m.group(1)}: "{m.group(2).strip()}"',
            raw
        )
        try:
            return json.loads(fixed)
        except json.JSONDecodeError as e:
            logging.error(f"Failed to parse fixed JSON: {e}\nFixed body: {fixed}")
            raise ValueError(f"Cannot parse webhook body as JSON: {e}")

    # Try form-encoded as last resort
    from urllib.parse import parse_qs
    if '=' in raw:
        params = parse_qs(raw)
        if params:
            return {k: v[0] if len(v) == 1 else v for k, v in params.items()}

    raise ValueError(f"Unrecognized webhook body format")


def extract_webhook_ticket_data(body: Dict[str, Any]) -> Dict[str, Any]:
    """Unwrap QuickBase webhook payloads into a single ticket dictionary."""
    ticket_data = body
    if isinstance(ticket_data, dict) and 'data' in ticket_data:
        ticket_data = ticket_data.get('data', {})
    if isinstance(ticket_data, list):
        ticket_data = ticket_data[0] if ticket_data else {}
    return ticket_data if isinstance(ticket_data, dict) else {}


def normalize_webhook_key(key: str) -> str:
    """Normalize payload keys so QuickBase labels and JSON keys map consistently."""
    normalized = re.sub(r'[^a-z0-9]+', '_', str(key).strip().lower())
    return re.sub(r'_+', '_', normalized).strip('_')


def normalize_webhook_ticket_data(ticket_data: Dict[str, Any]) -> Dict[str, Any]:
    """Normalize webhook payload field names into the keys the app expects."""
    alias_map = {
        'ticket_number': 'ticket_number',
        'ticket_no': 'ticket_number',
        'ticket_id': 'ticket_number',
        'subject': 'subject',
        'status': 'status',
        'old_status': 'old_status',
        'previous_status': 'old_status',
        'prior_status': 'old_status',
        'submitted_by': 'submitted_by',
        'submitted_by_email': 'submitted_by',
        'user_email': 'submitted_by',
        'email': 'submitted_by',
        'category': 'category',
        'priority': 'priority',
        'resolution': 'resolution',
        'record_id': 'record_id',
        'rid': 'record_id',
        'quickbase_url': 'quickbase_url'
    }

    normalized: Dict[str, Any] = {}
    for raw_key, value in ticket_data.items():
        key = normalize_webhook_key(raw_key)
        target_key = alias_map.get(key, key)
        normalized[target_key] = value

    return normalized


async def enrich_webhook_ticket_data(ticket_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Fill in missing QuickBase metadata so notification cards have working links
    and complete ticket details even when the webhook body is minimal.
    """
    if ticket_data.get('quickbase_url'):
        return ticket_data

    ticket_number = ticket_data.get('ticket_number')
    if not ticket_number:
        return ticket_data

    try:
        qb = get_qb_manager()
        existing_ticket = await qb.get_ticket(ticket_number)
        if not existing_ticket:
            return ticket_data

        enriched = dict(existing_ticket)
        for key, value in ticket_data.items():
            if value not in (None, ''):
                enriched[key] = value
        return enriched
    except Exception as e:
        logging.warning(f"Failed to enrich webhook ticket data for {ticket_number}: {e}")
        return ticket_data


# =============================================================================
# Automation Flow - M365 Provisioning & Future Automations
# =============================================================================

async def start_automation_flow(
    question: str, user_info: Dict, activity: Dict
) -> func.HttpResponse:
    """Start an automation flow when the router detects an automation request.

    Detects which automation handler matches, creates a request, and sends
    the initial routing card to the user.
    """
    teams = get_teams_handler()
    auto_mgr = get_automation_manager()

    user_email = user_info.get('email') or user_info.get('userPrincipalName', '')
    user_name = user_info.get('name', 'Unknown User')

    # Detect which automation handler matches this message
    detection = auto_mgr.detect_automation(question)

    if not detection:
        # No automation handler matched - fall back to normal needs_human flow
        logging.info("Automation request detected by router but no handler matched, falling back")
        await teams.send_message(
            activity,
            "This request requires an IT Administrator. "
            "A ticket has been created for IT to review."
        )
        # Create a ticket for manual handling
        qb = get_qb_manager()
        chain = get_support_chain()
        ticket_data = {
            'subject': generate_subject(question),
            'description': f"**User Request:**\n{question}\n\n---\n*Auto-generated by IT Support Bot*",
            'priority': 'Medium',
            'category': 'General Support',
            'status': 'New',
            'user_email': user_email,
            'user_name': user_name
        }
        ticket = await qb.create_ticket(ticket_data)
        if ticket and user_email:
            chain.conversation_stream.record_message(
                user_email, question, ticket.get('ticket_number')
            )
        return func.HttpResponse(status_code=200)

    # Create the automation request
    request = auto_mgr.create_request(
        automation_type=detection["automation_type"],
        requester_email=user_email,
        requester_name=user_name,
        extracted=detection.get("extracted", {}),
        original_message=question
    )

    # Get the handler and send the routing card
    handler = auto_mgr.get_handler(detection["automation_type"])
    routing_card = handler.create_routing_card(request)
    await teams.send_card(activity, routing_card)

    # Record in conversation stream so follow-ups don't create tickets
    if user_email:
        chain = get_support_chain()
        chain.conversation_stream.record_message(user_email, question)

    logging.info(
        f"Automation flow started: request={request.request_id} "
        f"type={detection['automation_type']} for {user_email}"
    )

    return func.HttpResponse(status_code=200)


async def handle_provisioning_invoke(
    action_type: str, action_data: Dict, activity: Dict, user_info: Dict
) -> bool:
    """Handle all provisioning-related invoke actions from adaptive cards.

    Returns True if the action was handled, False if not a provisioning action.
    """
    if not action_type.startswith('provisioning_'):
        return False

    teams = get_teams_handler()
    auto_mgr = get_automation_manager()
    request_id = action_data.get('request_id', '')

    if action_type == 'provisioning_select_type':
        # User selected a resource type from the routing card
        request = auto_mgr.get_request(request_id)
        if not request:
            await teams.send_message(activity, "This request has expired. Please start a new one.")
            return True

        resource_type = action_data.get('resource_type', '')
        request.resource_type = resource_type
        request.updated_at = __import__('time').time()

        handler = auto_mgr.get_handler(request.automation_type)
        config_form = handler.create_config_form(request)
        await teams.update_card(activity, config_form)
        return True

    elif action_type == 'provisioning_submit_config':
        # User submitted the configuration form
        request = auto_mgr.get_request(request_id)
        if not request:
            await teams.send_message(activity, "This request has expired. Please start a new one.")
            return True

        from m365_provisioning import build_config_from_form
        resource_type = action_data.get('resource_type', request.resource_type)
        request.resource_type = resource_type
        request.config = build_config_from_form(resource_type, action_data)
        request.status = __import__('automation_manager').AutomationStatus.PENDING_APPROVAL
        request.updated_at = __import__('time').time()

        handler = auto_mgr.get_handler(request.automation_type)

        # Send summary card to the requester
        summary_card = handler.create_summary_card(request)
        await teams.update_card(activity, summary_card)

        # Send approval card to the admin
        admin_email = auto_mgr.admin_email
        if admin_email:
            approval_card = handler.create_approval_card(request)
            await teams.send_notification_to_user(admin_email, approval_card)
            logging.info(
                f"Approval card sent to {admin_email} for request {request_id}"
            )
        else:
            logging.warning(
                "AUTOMATION_ADMIN_EMAIL not set - approval card not sent. "
                "Set this env var to enable the approval workflow."
            )

        return True

    elif action_type == 'provisioning_approve':
        # Admin approved the request
        request = auto_mgr.get_request(request_id)
        if not request:
            await teams.send_message(activity, "This request has expired or was already processed.")
            return True

        handler = auto_mgr.get_handler(request.automation_type)

        # Update the admin's card to show it's being processed
        processing_card = {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [{
                "type": "TextBlock",
                "text": f"⏳ Approved! Provisioning request #{request_id}...",
                "weight": "Bolder",
                "color": "Good",
                "wrap": True
            }]
        }
        await teams.update_card(activity, processing_card)

        # Execute the automation
        result = await auto_mgr.approve_and_execute(request_id)

        # Send result card to the requester
        result_card = handler.create_result_card(request)
        await teams.send_notification_to_user(request.requester_email, result_card)

        # Update admin's card with the result
        if result.get("success"):
            admin_done_card = {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [{
                    "type": "TextBlock",
                    "text": f"✅ Request #{request_id} provisioned successfully.",
                    "weight": "Bolder",
                    "color": "Good",
                    "wrap": True
                }, {
                    "type": "TextBlock",
                    "text": f"Resource: {request.config.get('display_name', 'N/A')}",
                    "isSubtle": True,
                    "wrap": True
                }]
            }
        else:
            admin_done_card = {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [{
                    "type": "TextBlock",
                    "text": f"❌ Request #{request_id} failed: {result.get('error', 'Unknown error')}",
                    "weight": "Bolder",
                    "color": "Attention",
                    "wrap": True
                }]
            }
        await teams.update_card(activity, admin_done_card)

        logging.info(
            f"Provisioning {'succeeded' if result.get('success') else 'failed'} "
            f"for request {request_id}"
        )
        return True

    elif action_type == 'provisioning_deny':
        # Admin denied the request
        request = auto_mgr.get_request(request_id)
        if not request:
            await teams.send_message(activity, "This request has expired or was already processed.")
            return True

        denial_reason = action_data.get('denial_reason', '')
        auto_mgr.deny_request(request_id, denial_reason)

        handler = auto_mgr.get_handler(request.automation_type)

        # Notify the requester
        denied_card = handler.create_denied_card(request)
        await teams.send_notification_to_user(request.requester_email, denied_card)

        # Update admin's card
        admin_done_card = {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [{
                "type": "TextBlock",
                "text": f"❌ Request #{request_id} denied.",
                "weight": "Bolder",
                "wrap": True
            }, {
                "type": "TextBlock",
                "text": f"Requester ({request.requester_email}) has been notified.",
                "isSubtle": True,
                "wrap": True
            }]
        }
        await teams.update_card(activity, admin_done_card)

        logging.info(f"Request {request_id} denied by admin. Reason: {denial_reason}")
        return True

    elif action_type == 'provisioning_cancel':
        # User cancelled the provisioning flow
        request = auto_mgr.get_request(request_id)
        if request:
            request.status = __import__('automation_manager').AutomationStatus.DENIED
            request.denial_reason = "Cancelled by user"

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
        return True

    return False


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

        body = parse_webhook_body(req)
        logging.info(f"Webhook payload: {json.dumps(body)}")

        ticket_data = normalize_webhook_ticket_data(extract_webhook_ticket_data(body))

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
        raw_body = req.get_body().decode('utf-8', errors='replace')[:500]
        logging.error(f"Invalid JSON in webhook payload: {str(e)} | Raw body: {raw_body}")
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

        ticket_data = await enrich_webhook_ticket_data(ticket_data)

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
# QuickBase Webhook - Ticket Status Update Notification
# =============================================================================

@app.route(route="webhook/ticket-update", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
async def webhook_ticket_update(req: func.HttpRequest) -> func.HttpResponse:
    """
    Webhook endpoint for QuickBase to notify users when a ticket status changes.

    QuickBase Webhook Configuration:
    1. Go to your QuickBase app settings
    2. Navigate to Webhooks
    3. Create a new webhook with:
       - URL: https://your-function-app.azurewebsites.net/api/webhook/ticket-update
       - Method: POST
       - Trigger: When record is modified
       - Condition: Status field changes
       - Include fields: ticket_number, subject, status, old_status, submitted_by (email),
                         category, priority, resolution (optional)

    Expected payload format:
    {
        "ticket_number": "IT-240101123456",
        "subject": "Issue subject",
        "status": "In Progress",
        "old_status": "New",
        "submitted_by": "user@company.com",
        "category": "General Support",
        "priority": "Medium",
        "resolution": ""
    }
    """
    logging.info("Received QuickBase webhook for ticket status update")

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

        body = parse_webhook_body(req)
        logging.info(f"Ticket update webhook payload: {json.dumps(body)}")

        ticket_data = normalize_webhook_ticket_data(extract_webhook_ticket_data(body))

        # Extract ticket information
        ticket_number = ticket_data.get('ticket_number', '')
        new_status = ticket_data.get('status', '')
        old_status = ticket_data.get('old_status', '')
        user_email = ticket_data.get('submitted_by', '')

        if not ticket_number:
            logging.warning("No ticket_number in webhook payload")
            return func.HttpResponse(
                json.dumps({"error": "Missing ticket_number"}),
                status_code=400,
                mimetype="application/json"
            )

        if not new_status:
            logging.warning(f"No status in webhook payload for ticket {ticket_number}")
            return func.HttpResponse(
                json.dumps({"error": "Missing status"}),
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

        # Skip if status hasn't actually changed
        if old_status and old_status == new_status:
            logging.info(f"Ticket {ticket_number} status unchanged ({new_status}). Skipping.")
            return func.HttpResponse(
                json.dumps({"status": "skipped", "reason": "status unchanged"}),
                status_code=200,
                mimetype="application/json"
            )

        # Send notification to the user who submitted the ticket
        notification_sent = await send_status_update_notification(ticket_data, user_email)

        if notification_sent:
            logging.info(f"Status update notification sent for {ticket_number} ({old_status} -> {new_status}) to {user_email}")
            return func.HttpResponse(
                json.dumps({
                    "status": "success",
                    "ticket_number": ticket_number,
                    "new_status": new_status,
                    "old_status": old_status,
                    "notified": user_email
                }),
                status_code=200,
                mimetype="application/json"
            )
        else:
            logging.warning(f"Failed to send status update notification for {ticket_number}")
            return func.HttpResponse(
                json.dumps({"status": "partial", "ticket_number": ticket_number, "notification_sent": False}),
                status_code=200,
                mimetype="application/json"
            )

    except ValueError as e:
        raw_body = req.get_body().decode('utf-8', errors='replace')[:500]
        logging.error(f"Invalid JSON in webhook payload: {str(e)} | Raw body: {raw_body}")
        return func.HttpResponse(
            json.dumps({"error": "Invalid JSON payload"}),
            status_code=400,
            mimetype="application/json"
        )
    except Exception as e:
        logging.error(f"Error processing ticket update webhook: {str(e)}")
        return func.HttpResponse(
            json.dumps({"error": str(e)}),
            status_code=500,
            mimetype="application/json"
        )


@app.route(route="webhook/ticket-updated", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
async def webhook_ticket_updated_alias(req: func.HttpRequest) -> func.HttpResponse:
    """Alias route for QuickBase configurations that use 'ticket-updated'."""
    return await webhook_ticket_update(req)


async def send_status_update_notification(ticket_data: Dict[str, Any], user_email: str) -> bool:
    """
    Send a Teams notification to the user when their ticket status changes.

    Uses proactive messaging to reach the user directly.
    """
    try:
        teams = get_teams_handler()

        ticket_data = await enrich_webhook_ticket_data(ticket_data)

        # Create the status update notification card
        notification_card = create_status_update_card(ticket_data)

        # Send proactive message to user
        success = await teams.send_notification_to_user(user_email, notification_card)

        return success

    except Exception as e:
        logging.error(f"Error sending status update notification: {str(e)}")
        return False


def create_status_update_card(ticket_data: Dict[str, Any]) -> Dict:
    """Create adaptive card for ticket status update notification"""
    ticket_number = ticket_data.get('ticket_number', 'N/A')
    subject = ticket_data.get('subject', 'No Subject')
    new_status = ticket_data.get('status', 'Unknown')
    old_status = ticket_data.get('old_status', '')
    category = ticket_data.get('category', 'General Support')
    priority = ticket_data.get('priority', 'Medium')
    resolution = ticket_data.get('resolution', '')

    priority_icons = {'Critical': '🔴', 'High': '🟠', 'Medium': '🟡', 'Low': '🟢'}
    priority_icon = priority_icons.get(priority, '⚪')

    # Status icons and colors for the timeline
    status_config = {
        'New': {'icon': '🆕', 'color': 'Default'},
        'In Progress': {'icon': '🔧', 'color': 'Accent'},
        'Awaiting User': {'icon': '⏳', 'color': 'Warning'},
        'Awaiting IT': {'icon': '👨‍💻', 'color': 'Warning'},
        'Resolved': {'icon': '✅', 'color': 'Good'},
        'Closed': {'icon': '📁', 'color': 'Good'},
        'Cancelled': {'icon': '❌', 'color': 'Attention'}
    }

    new_config = status_config.get(new_status, {'icon': '🔄', 'color': 'Default'})
    status_icon = new_config['icon']
    status_color = new_config['color']

    # Build the transition text
    if old_status:
        old_icon = status_config.get(old_status, {'icon': '🔄'})['icon']
        transition_text = f"{old_icon} {old_status}  →  {status_icon} **{new_status}**"
    else:
        transition_text = f"{status_icon} **{new_status}**"

    # Build card body
    card_body = [
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
                                    "text": "🔔",
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
                                    "text": "Ticket Status Update",
                                    "weight": "Bolder",
                                    "size": "Large",
                                    "color": status_color
                                },
                                {
                                    "type": "TextBlock",
                                    "text": f"Ticket #{ticket_number} has been updated",
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
                    "type": "TextBlock",
                    "text": transition_text,
                    "size": "Medium",
                    "wrap": True,
                    "horizontalAlignment": "Center",
                    "spacing": "Small"
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
                        {"title": "Ticket:", "value": f"#{ticket_number}"},
                        {"title": "Subject:", "value": subject[:50] + ('...' if len(subject) > 50 else '')},
                        {"title": "Category:", "value": category},
                        {"title": "Priority:", "value": f"{priority_icon} {priority}"}
                    ]
                }
            ]
        }
    ]

    # Add resolution details if ticket is Resolved or Closed and resolution exists
    if new_status in ('Resolved', 'Closed') and resolution:
        card_body.append({
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
        })

    # Add contextual message based on new status
    status_messages = {
        'In Progress': "Your ticket is now being worked on by the IT team.",
        'Awaiting User': "The IT team needs more information from you. Please check your ticket and respond.",
        'Awaiting IT': "Your response has been received. The IT team will follow up shortly.",
        'Resolved': "Your ticket has been resolved. If the issue persists, please create a new ticket.",
        'Closed': "Your ticket has been closed. If you need further help, please create a new ticket.",
        'Cancelled': "Your ticket has been cancelled. If this was in error, please create a new ticket."
    }

    status_message = status_messages.get(new_status, '')
    if status_message:
        card_body.append({
            "type": "TextBlock",
            "text": status_message,
            "wrap": True,
            "isSubtle": True,
            "spacing": "Large",
            "size": "Small"
        })

    # Build actions
    actions = [
        {
            "type": "Action.OpenUrl",
            "title": "View in QuickBase",
            "url": ticket_data.get('quickbase_url', '#')
        }
    ]

    # Add "Create New Ticket" for terminal states
    if new_status in ('Resolved', 'Closed', 'Cancelled'):
        actions.insert(0, {
            "type": "Action.Submit",
            "title": "Create New Ticket",
            "data": {
                "action": "create_ticket_form"
            }
        })

    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": card_body,
        "actions": actions
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
