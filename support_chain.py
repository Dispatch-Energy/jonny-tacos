"""
IT Support Chain - Composable LangChain architecture for direct, efficient support
Follows patterns from Enterprise RAG Challenge winners:
- Router pattern for decision-making
- Structured outputs (Pydantic)
- Composable handlers
- Fast single-shot for common issues
- Tool-ready for future expansion
"""

import os
from typing import Literal, Optional, List
from pydantic import BaseModel, Field
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

# Try new location first, fall back to old
try:
    from langchain.output_parsers import PydanticOutputParser
except ImportError:
    from langchain_core.output_parsers import PydanticOutputParser


# ============================================================================
# STRUCTURED OUTPUTS - First-class typed decisions
# ============================================================================

class SupportIntent(BaseModel):
    """Router decision - what kind of support interaction is this?"""
    intent_type: Literal[
        "quick_fix",           # Can be solved immediately with KB
        "needs_troubleshooting",  # Needs multi-step diagnosis
        "needs_ticket",        # Requires IT intervention
        "status_check",        # Checking existing ticket
        "command"              # Bot command like /help
    ] = Field(description="Type of support interaction")
    
    confidence: float = Field(ge=0.0, le=1.0, description="Confidence in classification")
    reasoning: str = Field(description="Why this classification was chosen")
    
    # Extracted entities
    category: Optional[str] = Field(default=None, description="IT category if applicable")
    priority: Optional[Literal["Low", "Medium", "High", "Critical"]] = Field(
        default=None, 
        description="Priority if ticket needed"
    )
    ticket_number: Optional[str] = Field(default=None, description="Ticket number if status check")


class QuickFixResponse(BaseModel):
    """Response for issues that can be solved immediately"""
    solution: str = Field(description="Step-by-step solution")
    solved: bool = Field(description="Whether this completely solves the issue")
    confidence: float = Field(ge=0.0, le=1.0)
    offer_ticket: bool = Field(default=False, description="Offer to create ticket if this doesn't work")


class TicketRecommendation(BaseModel):
    """Structured recommendation for ticket creation"""
    should_create: bool
    subject: str = Field(max_length=100)
    description: str
    category: str
    priority: Literal["Low", "Medium", "High", "Critical"]
    reasoning: str = Field(description="Why these values were chosen")


# ============================================================================
# COMPOSABLE HANDLERS - Each does one thing well
# ============================================================================

class SupportRouter:
    """
    Router - decides what to do with incoming question
    Inspired by ERC winners' router pattern
    """
    def __init__(self, llm: Optional[ChatOpenAI] = None):
        self.llm = llm or self._get_default_llm()
        self.parser = PydanticOutputParser(pydantic_object=SupportIntent)
        
        self.prompt = ChatPromptTemplate.from_messages([
            ("system", """You are a routing assistant for IT support.
            Classify the user's message into the appropriate category.
            
            Guidelines:
            - quick_fix: Common issues with known solutions (password, VPN, Teams, printer)
            - needs_troubleshooting: Complex issues needing diagnosis
            - needs_ticket: Hardware, software licenses, admin access, new user setup
            - status_check: User asking about existing ticket
            - command: Bot commands like /help, /ticket, /status
            
            Be decisive - our company culture values direct, efficient responses.
            {format_instructions}"""),
            ("human", "{question}")
        ])
    
    def route(self, question: str) -> SupportIntent:
        """Route incoming question to appropriate handler"""
        chain = self.prompt | self.llm | self.parser
        return chain.invoke({
            "question": question,
            "format_instructions": self.parser.get_format_instructions()
        })
    
    @staticmethod
    def _get_default_llm():
        return ChatOpenAI(
            model=os.getenv("GPT5_MODEL", "gpt-4"),
            temperature=0.1,  # Low temp for consistent routing
            openai_api_base=os.getenv("GPT5_ENDPOINT"),
            openai_api_key=os.getenv("GPT5_API_KEY")
        )


class QuickFixHandler:
    """
    Handles common IT issues that can be solved in one shot
    Uses structured KB + LLM for dynamic responses
    """
    def __init__(self, llm: Optional[ChatOpenAI] = None, kb: Optional[dict] = None):
        self.llm = llm or self._get_generation_llm()
        self.kb = kb or self._load_knowledge_base()
        self.parser = PydanticOutputParser(pydantic_object=QuickFixResponse)
        
        self.prompt = ChatPromptTemplate.from_messages([
            ("system", """You are a direct, efficient IT support assistant.
            
            Your company culture values:
            - Brevity and clarity over lengthy explanations
            - Actionable steps over theory
            - Getting users working ASAP
            
            Given the relevant knowledge base context below, provide a direct solution.
            If you're highly confident this will work, set solved=True.
            If there's uncertainty, set solved=False and offer_ticket=True.
            
            Knowledge Base Context:
            {kb_context}
            
            {format_instructions}"""),
            ("human", "{question}")
        ])
    
    def handle(self, question: str, category: Optional[str] = None) -> QuickFixResponse:
        """Generate direct solution from KB + LLM"""
        # Get relevant KB context
        kb_context = self._get_kb_context(question, category)
        
        chain = self.prompt | self.llm | self.parser
        return chain.invoke({
            "question": question,
            "kb_context": kb_context,
            "format_instructions": self.parser.get_format_instructions()
        })
    
    def _get_kb_context(self, question: str, category: Optional[str]) -> str:
        """Retrieve relevant KB entries - simple keyword matching for now"""
        question_lower = question.lower()
        relevant_entries = []
        
        for key, entry in self.kb.items():
            # Check keywords
            if any(kw in question_lower for kw in entry.get("keywords", [])):
                relevant_entries.append(f"### {key.upper()}\n{entry['solution']}")
        
        if relevant_entries:
            return "\n\n".join(relevant_entries[:2])  # Top 2 matches
        return "No specific KB entry found - provide general guidance."
    
    @staticmethod
    def _get_generation_llm():
        return ChatOpenAI(
            model=os.getenv("GPT5_MODEL", "gpt-4"),
            temperature=0.3,
            openai_api_base=os.getenv("GPT5_ENDPOINT"),
            openai_api_key=os.getenv("GPT5_API_KEY")
        )
    
    @staticmethod
    def _load_knowledge_base() -> dict:
        """Load existing KB - will eventually be replaced by vector store"""
        return {
            "password_reset": {
                "keywords": ["password", "reset", "locked", "can't login"],
                "solution": """Go to https://passwordreset.microsoftonline.com
1. Enter work email
2. Complete verification
3. Create new password (12+ chars, mixed case, numbers, special chars)"""
            },
            "vpn_issues": {
                "keywords": ["vpn", "remote", "connection"],
                "solution": """VPN Troubleshooting:
1. Check internet connection
2. Restart VPN client
3. Clear credentials and re-enter
4. Run: ipconfig /flushdns (Windows) or sudo dscacheutil -flushcache (Mac)"""
            },
            "teams_audio": {
                "keywords": ["teams", "audio", "microphone", "can't hear"],
                "solution": """Teams Audio Fix:
1. Settings → Devices → Test audio
2. Check correct device selected
3. Windows Settings → Privacy → Allow microphone access
4. Clear cache: %appdata%\\Microsoft\\Teams\\Cache"""
            },
            "slow_computer": {
                "keywords": ["slow", "performance", "freezing"],
                "solution": """Performance Fix:
1. Restart computer
2. Check Windows Updates
3. Run Disk Cleanup (cleanmgr)
4. Task Manager → Disable startup programs"""
            }
        }


class TicketHandler:
    """
    Handles ticket creation with structured recommendations
    """
    def __init__(self, llm: Optional[ChatOpenAI] = None):
        self.llm = llm or ChatOpenAI(
            model=os.getenv("GPT5_MODEL", "gpt-4"),
            temperature=0.2,
            openai_api_base=os.getenv("GPT5_ENDPOINT"),
            openai_api_key=os.getenv("GPT5_API_KEY")
        )
        self.parser = PydanticOutputParser(pydantic_object=TicketRecommendation)
        
        self.prompt = ChatPromptTemplate.from_messages([
            ("system", """You recommend ticket parameters for IT issues.
            
            Priority Guidelines:
            - Critical: System down, multiple users affected, security incident
            - High: Can't work, single user blocked
            - Medium: Impacted but have workaround
            - Low: Questions, minor issues
            
            Categories: Password Reset, Software Installation, Hardware Issue, 
            Network Connectivity, Email Issues, Teams/Office 365, VPN Access, 
            Printer Problems, File Access, Security Concern, New User Setup, 
            General Support, Other
            
            {format_instructions}"""),
            ("human", "User issue: {question}\n\nContext: {context}")
        ])
    
    def recommend(self, question: str, context: Optional[str] = None) -> TicketRecommendation:
        """Generate structured ticket recommendation"""
        chain = self.prompt | self.llm | self.parser
        return chain.invoke({
            "question": question,
            "context": context or "No additional context",
            "format_instructions": self.parser.get_format_instructions()
        })


# ============================================================================
# MAIN ORCHESTRATOR - Compose handlers together
# ============================================================================

class ITSupportChain:
    """
    Main orchestrator - routes and delegates to handlers
    Simple, composable, testable
    """
    def __init__(self):
        self.router = SupportRouter()
        self.quick_fix = QuickFixHandler()
        self.ticket_handler = TicketHandler()
    
    def process(self, question: str) -> dict:
        """
        Main entry point - process user question
        Returns structured response ready for Teams card builder
        """
        # Step 1: Route
        intent = self.router.route(question)
        
        # Step 2: Delegate to appropriate handler
        if intent.intent_type == "quick_fix":
            response = self.quick_fix.handle(question, intent.category)
            return {
                "type": "solution",
                "solution": response.solution,
                "confidence": response.confidence,
                "offer_ticket": response.offer_ticket,
                "category": intent.category,
                "priority": intent.priority
            }
        
        elif intent.intent_type == "needs_ticket":
            recommendation = self.ticket_handler.recommend(question)
            return {
                "type": "ticket_needed",
                "recommendation": recommendation.dict(),
                "reasoning": intent.reasoning
            }
        
        elif intent.intent_type == "status_check":
            return {
                "type": "status_check",
                "ticket_number": intent.ticket_number
            }
        
        elif intent.intent_type == "command":
            return {
                "type": "command",
                "intent": intent
            }
        
        else:  # needs_troubleshooting
            # Future: multi-step troubleshooting workflow
            # For now, route to ticket
            recommendation = self.ticket_handler.recommend(
                question, 
                context="Complex issue requiring troubleshooting"
            )
            return {
                "type": "troubleshooting_needed",
                "recommendation": recommendation.dict()
            }


# ============================================================================
# TOOL-READY EXTENSIONS (Future)
# ============================================================================

class SupportTools:
    """
    Placeholder for future tool integration
    Tools can be added without changing core architecture
    """
    
    @staticmethod
    def check_ticket_status(ticket_number: str) -> dict:
        """Tool: Check ticket status in QuickBase"""
        # Will integrate with QuickBaseManager
        pass
    
    @staticmethod
    def search_knowledge_base(query: str) -> List[str]:
        """Tool: Search KB (will become vector store)"""
        # Future: RAG with session-based retrieval
        pass
    
    @staticmethod
    def run_diagnostics(system_type: str) -> dict:
        """Tool: Run automated diagnostics"""
        # Future: Could integrate with monitoring systems
        pass


# ============================================================================
# EXAMPLE USAGE
# ============================================================================

if __name__ == "__main__":
    # Initialize chain
    chain = ITSupportChain()
    
    # Example 1: Quick fix
    result = chain.process("I can't connect to VPN")
    print("Quick Fix Example:")
    print(result)
    print("\n" + "="*80 + "\n")
    
    # Example 2: Needs ticket
    result = chain.process("I need to install Adobe Creative Suite")
    print("Ticket Needed Example:")
    print(result)
    print("\n" + "="*80 + "\n")
    
    # Example 3: Status check
    result = chain.process("What's the status of ticket IT-1234?")
    print("Status Check Example:")
    print(result)