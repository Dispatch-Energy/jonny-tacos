"""
IT Support Chain - Compatible with langchain==0.1.0, langchain-openai==0.0.5
Key Changes:
1. ALWAYS call GPT (with or without vector store context)
2. Uses openai_api_base (not base_url) for older langchain-openai
3. No vector store (removed FAISS dependency for stability)
4. Never returns "can't help" - always provides a solution + creates ticket
"""

import os
import logging
from typing import Literal, Optional, List, Dict, Any
from pydantic import BaseModel, Field

# LangChain imports - compatible with 0.1.0
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_openai import ChatOpenAI

# Pydantic output parser - location varies by version
try:
    from langchain.output_parsers import PydanticOutputParser
except ImportError:
    from langchain_core.output_parsers import PydanticOutputParser


# ============================================================================
# CONFIGURATION
# ============================================================================

def get_llm(temperature: float = 0.3) -> ChatOpenAI:
    """
    Get properly configured LLM for custom GPT endpoint.
    
    Compatible with langchain-openai==0.0.5:
    - Uses 'openai_api_base' (not 'base_url')
    - Uses 'openai_api_key' (not 'api_key')
    """
    endpoint = os.getenv("GPT5_ENDPOINT", "")
    api_key = os.getenv("GPT5_API_KEY", "")
    model = os.getenv("GPT5_MODEL", "gpt-4")
    
    if not endpoint or not api_key:
        logging.warning("GPT5_ENDPOINT or GPT5_API_KEY not set!")
    
    # Configure for custom endpoint - use older parameter names
    return ChatOpenAI(
        model=model,
        temperature=temperature,
        openai_api_base=endpoint,      # Old name, compatible with 0.0.5
        openai_api_key=api_key,        # Old name, compatible with 0.0.5
        request_timeout=30,
        max_retries=2,
    )


# ============================================================================
# STRUCTURED OUTPUTS
# ============================================================================

class SupportIntent(BaseModel):
    """Router decision - what kind of support interaction is this?"""
    intent_type: Literal[
        "quick_fix",              # Can attempt solution immediately
        "needs_human",            # Definitely needs IT intervention
        "status_check",           # Checking existing ticket
        "command"                 # Bot command like /help
    ] = Field(description="Type of support interaction")
    
    confidence: float = Field(ge=0.0, le=1.0, description="Confidence in classification")
    reasoning: str = Field(description="Brief reason for classification")
    
    # Extracted entities
    category: Optional[str] = Field(default=None, description="IT category")
    priority: Optional[Literal["Low", "Medium", "High", "Critical"]] = Field(
        default="Medium", 
        description="Suggested priority"
    )
    ticket_number: Optional[str] = Field(default=None, description="If status check")


class SupportResponse(BaseModel):
    """Structured support response - ALWAYS has a solution"""
    solution: str = Field(description="Helpful solution or guidance - NEVER empty")
    confidence: float = Field(ge=0.0, le=1.0, description="How confident we are this solves it")
    category: str = Field(default="General Support", description="IT category")
    priority: Literal["Low", "Medium", "High", "Critical"] = Field(default="Medium")
    needs_human: bool = Field(default=False, description="Requires IT follow-up")
    sources_used: List[str] = Field(default_factory=list, description="KB sources")


# ============================================================================
# KNOWLEDGE BASE (Static - simple and stable)
# ============================================================================

STATIC_KB = {
    "password_reset": {
        "keywords": ["password", "reset", "locked", "can't login", "login", "locked out", "forgot password"],
        "solution": """**Password Reset Steps:**
1. Go to https://passwordreset.microsoftonline.com
2. Enter your work email address
3. Complete verification (phone or email)
4. Create new password (12+ chars, mixed case, numbers, symbols)

If still locked out after reset, your account may need IT intervention.""",
        "category": "Password Reset"
    },
    "vpn": {
        "keywords": ["vpn", "remote", "connection", "work from home", "connect remotely", "remote access"],
        "solution": """**VPN Troubleshooting:**
1. Check your internet connection first
2. Restart the VPN client completely
3. Clear saved VPN credentials and re-enter
4. Run: `ipconfig /flushdns` (Windows) or `sudo dscacheutil -flushcache` (Mac)
5. Check if your VPN account has expired (HR change, etc.)

Common VPN ports that must be open: UDP 500, 4500.""",
        "category": "VPN Access"
    },
    "teams": {
        "keywords": ["teams", "meeting", "audio", "video", "microphone", "camera", "can't hear", "no sound"],
        "solution": """**Teams Audio/Video Fix:**
1. Settings → Devices → Test speaker & microphone
2. Verify correct devices are selected
3. Windows: Settings → Privacy → Microphone/Camera → Allow Teams
4. Clear Teams cache: Close Teams, delete contents of %appdata%\\Microsoft\\Teams\\Cache
5. Restart Teams

For persistent issues, try the web version at teams.microsoft.com.""",
        "category": "Teams/Office 365"
    },
    "email": {
        "keywords": ["email", "outlook", "send", "receive", "mailbox", "calendar", "can't send"],
        "solution": """**Outlook Troubleshooting:**
1. Check internet connection
2. File → Office Account → Update Options → Update Now
3. Send/Receive → Send/Receive All Folders
4. If stuck, try Outlook Web at outlook.office.com

For mailbox full: File → Info → Cleanup Tools → Archive old items.""",
        "category": "Email Issues"
    },
    "printer": {
        "keywords": ["printer", "print", "printing", "queue", "stuck", "offline"],
        "solution": """**Printer Troubleshooting:**
1. Settings → Devices → Printers → Select printer → Open queue
2. Cancel stuck jobs
3. Set printer to Online if showing Offline
4. Restart Print Spooler: services.msc → Print Spooler → Restart

For network printers, verify you're on the correct network/VPN.""",
        "category": "Printer Problems"
    },
    "slow": {
        "keywords": ["slow", "performance", "freezing", "frozen", "not responding", "lagging", "hang"],
        "solution": """**Performance Quick Fixes:**
1. Restart your computer (seriously, it helps!)
2. Close unnecessary browser tabs and applications
3. Check Task Manager (Ctrl+Shift+Esc) for high CPU/Memory usage
4. Run Disk Cleanup: Win+R → cleanmgr
5. Check for Windows Updates

If consistently slow, you may need a hardware assessment.""",
        "category": "Hardware Issue"
    },
    "software": {
        "keywords": ["install", "software", "application", "license", "download", "program"],
        "solution": """**Software Installation:**
1. Check Software Center (Windows) or Self Service (Mac) first
2. Many approved apps are available for self-install
3. If not in catalog, IT approval is required

For licensed software (Adobe, Microsoft, etc.), a ticket is needed for procurement.""",
        "category": "Software Installation"
    },
    "wifi": {
        "keywords": ["wifi", "wireless", "internet", "network", "can't connect", "no internet"],
        "solution": """**WiFi/Network Troubleshooting:**
1. Toggle WiFi off and on
2. Forget the network and reconnect
3. Run: `ipconfig /release` then `ipconfig /renew` (Windows)
4. Check if others on same network have issues
5. Try connecting to a different network to isolate the issue

If office network, ensure you're connected to the corporate SSID.""",
        "category": "Network Connectivity"
    }
}


def search_static_kb(question: str) -> tuple:
    """
    Search static KB for relevant context.
    Returns (context_string, category) or ("", "General Support") if no match.
    """
    question_lower = question.lower()
    
    for key, entry in STATIC_KB.items():
        if any(kw in question_lower for kw in entry["keywords"]):
            return entry["solution"], entry["category"]
    
    return "", "General Support"


# ============================================================================
# MAIN SUPPORT CHAIN
# ============================================================================

class ITSupportChain:
    """
    Main orchestrator - ALWAYS provides helpful response + ticket data.
    
    Flow:
    1. Search static KB for relevant context
    2. ALWAYS call GPT with whatever context we have
    3. ALWAYS return structured response for ticket creation
    """
    
    def __init__(self):
        logging.info("Initializing ITSupportChain...")
        
        try:
            self.llm = get_llm(temperature=0.3)
            self.router_llm = get_llm(temperature=0.1)  # Lower temp for routing
            logging.info("LLM initialized successfully")
        except Exception as e:
            logging.error(f"Failed to initialize LLM: {e}")
            raise
        
        # Router prompt
        self.router_parser = PydanticOutputParser(pydantic_object=SupportIntent)
        self.router_prompt = ChatPromptTemplate.from_messages([
            ("system", """You classify IT support requests. Be decisive.

Categories:
- quick_fix: Common issues (password, VPN, Teams, email, printer, slow computer, software questions, wifi)
- needs_human: Hardware replacement, new user setup, software licenses, admin access, security incidents, complex server issues
- status_check: User asking about existing ticket (look for ticket numbers like IT-1234, IT-0042)
- command: Bot commands starting with /

Most issues are quick_fix - err on the side of attempting to help first.
{format_instructions}"""),
            ("human", "{question}")
        ])
        
        # Solution generation prompt
        self.solution_prompt = ChatPromptTemplate.from_messages([
            ("system", """You are a helpful IT support assistant. Your goal is to SOLVE problems.

Company culture: Direct, efficient, action-oriented. Skip lengthy explanations.

RULES:
1. ALWAYS provide actionable steps - never say "contact IT" as your only answer
2. If you have context from the knowledge base, use it
3. If no specific context, use your general IT knowledge (you're good at this!)
4. Be specific with commands, paths, and steps
5. Keep responses concise but complete
6. If the issue genuinely needs human IT (hardware, licenses, admin), say so but still provide what the user CAN do
7. NEVER mention or reference any IT Service Portal, ticketing portal, or external URLs for logging tickets - there is no such portal. Users create tickets through this bot using /ticket command.

Context from knowledge base:
{kb_context}

Remember: Users want solutions, not apologies."""),
            ("human", """User's issue: {question}

Provide a helpful response. If this requires IT follow-up, still give the user something useful to try first.""")
        ])
        
        logging.info("ITSupportChain initialized")
    
    def process(self, question: str) -> Dict[str, Any]:
        """
        Main entry point - process user question.
        ALWAYS returns a helpful response + ticket data.
        """
        logging.info(f"Processing question: {question[:100]}...")
        
        # Step 1: Quick route check for commands
        if question.strip().startswith('/'):
            return {
                "type": "command",
                "solution": "",
                "category": "General Support",
                "priority": "Low",
                "confidence": 1.0,
                "needs_human": False,
                "offer_ticket": False,
                "sources": []
            }
        
        # Step 2: Route the question
        try:
            intent = self._route(question)
            logging.info(f"Routed as: {intent.intent_type} (confidence: {intent.confidence})")
        except Exception as e:
            logging.error(f"Router failed: {e}, defaulting to quick_fix")
            intent = SupportIntent(
                intent_type="quick_fix",
                confidence=0.5,
                reasoning="Router error - attempting solution anyway",
                category="General Support",
                priority="Medium"
            )
        
        # Step 3: Handle status checks separately
        if intent.intent_type == "status_check":
            return {
                "type": "status_check",
                "ticket_number": intent.ticket_number,
                "solution": "",
                "category": "General Support",
                "priority": "Low",
                "confidence": 1.0,
                "needs_human": False,
                "offer_ticket": False,
                "sources": []
            }
        
        # Step 4: Get context and generate solution
        response = self._generate_solution(question, intent)
        
        return {
            "type": "solution",
            "solution": response.solution,
            "category": response.category or intent.category or "General Support",
            "priority": response.priority or intent.priority or "Medium",
            "confidence": response.confidence,
            "needs_human": response.needs_human or (intent.intent_type == "needs_human"),
            "offer_ticket": True,  # Always offer escalation
            "sources": response.sources_used
        }
    
    def _route(self, question: str) -> SupportIntent:
        """Route question to determine handling strategy"""
        chain = self.router_prompt | self.router_llm | self.router_parser
        return chain.invoke({
            "question": question,
            "format_instructions": self.router_parser.get_format_instructions()
        })
    
    def _generate_solution(self, question: str, intent: SupportIntent) -> SupportResponse:
        """
        Generate solution using:
        1. Static KB context (if matched)
        2. GPT general knowledge (always)
        """
        sources = []
        
        # Search static KB
        kb_context, kb_category = search_static_kb(question)
        if kb_context:
            sources.append("Static KB")
            logging.info(f"Found static KB match: {kb_category}")
        else:
            kb_context = "No specific documentation found. Use your general IT knowledge to help."
            sources.append("GPT General Knowledge")
        
        # Generate solution - THIS ALWAYS RUNS
        try:
            chain = self.solution_prompt | self.llm | StrOutputParser()
            solution = chain.invoke({
                "question": question,
                "kb_context": kb_context
            })
            
            # Determine confidence based on sources
            if "Static KB" in sources:
                confidence = 0.85
            else:
                confidence = 0.7  # GPT general knowledge
            
            return SupportResponse(
                solution=solution,
                confidence=confidence,
                category=kb_category or intent.category or "General Support",
                priority=intent.priority or "Medium",
                needs_human=intent.intent_type == "needs_human",
                sources_used=sources
            )
            
        except Exception as e:
            logging.error(f"Solution generation failed: {e}")
            # NEVER return empty - provide fallback
            return SupportResponse(
                solution=self._get_fallback_response(question),
                confidence=0.3,
                category=intent.category or "General Support",
                priority="Medium",
                needs_human=True,
                sources_used=["Fallback"]
            )
    
    def _get_fallback_response(self, question: str) -> str:
        """Fallback when GPT fails - still helpful!"""
        return f"""I'm experiencing a temporary issue with my AI service, but here are some general steps that often help:

1. **Restart** the affected application or your computer
2. **Check** if others are experiencing the same issue
3. **Verify** your network connection
4. **Note** any error messages you see

A ticket has been created for IT to review your specific issue: "{question[:100]}..."

An IT team member will follow up shortly."""


# ============================================================================
# TESTING
# ============================================================================

if __name__ == "__main__":
    # Quick test
    logging.basicConfig(level=logging.INFO)
    
    print("Testing ITSupportChain...")
    print(f"GPT5_ENDPOINT: {os.getenv('GPT5_ENDPOINT', 'NOT SET')[:30]}...")
    print(f"GPT5_MODEL: {os.getenv('GPT5_MODEL', 'NOT SET')}")
    
    try:
        chain = ITSupportChain()
        print("✅ Chain initialized\n")
    except Exception as e:
        print(f"❌ Chain init failed: {e}")
        exit(1)
    
    test_questions = [
        "I can't connect to VPN",
        "My computer is really slow",
        "I need Adobe Creative Suite installed",
        "What's the status of ticket IT-1234?",
        "Teams keeps crashing during meetings",
    ]
    
    for q in test_questions:
        print(f"\n{'='*60}")
        print(f"Q: {q}")
        try:
            result = chain.process(q)
            print(f"Type: {result['type']}")
            print(f"Category: {result.get('category')}")
            print(f"Confidence: {result.get('confidence')}")
            print(f"Sources: {result.get('sources')}")
            print(f"Solution preview: {result.get('solution', '')[:200]}...")
        except Exception as e:
            print(f"❌ Error: {e}")