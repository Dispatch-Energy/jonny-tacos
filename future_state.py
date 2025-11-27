"""
Future Extensions - RAG and Tool Integration
Shows how to add capabilities without rewriting core architecture
"""

from typing import List, Optional, Any
from pydantic import BaseModel, Field
from langchain_core.tools import tool
from langchain_core.embeddings import Embeddings
from langchain_community.vectorstores import FAISS
from langchain_openai import OpenAIEmbeddings


# ============================================================================
# PHASE 2: ADD RAG (When ready)
# ============================================================================

class KnowledgeBaseRAG:
    """
    Enhanced knowledge base with vector search
    Can be swapped in without changing SupportRouter/QuickFixHandler
    """
    def __init__(self, embeddings: Optional[Embeddings] = None):
        self.embeddings = embeddings or OpenAIEmbeddings(
            model="text-embedding-3-small",
            openai_api_base=os.getenv("OPENAI_ENDPOINT"),
            openai_api_key=os.getenv("OPENAI_API_KEY")
        )
        self.vector_store = None
        self._initialize_vector_store()
    
    def _initialize_vector_store(self):
        """
        Initialize FAISS vector store with existing KB
        Later: load from documents, PDFs, SharePoint, etc.
        """
        from support_chain import QuickFixHandler
        
        kb = QuickFixHandler._load_knowledge_base()
        
        # Convert KB to documents
        documents = []
        metadatas = []
        for key, entry in kb.items():
            documents.append(entry['solution'])
            metadatas.append({
                'source': key,
                'keywords': ','.join(entry.get('keywords', []))
            })
        
        # Create vector store
        self.vector_store = FAISS.from_texts(
            documents,
            self.embeddings,
            metadatas=metadatas
        )
    
    def search(self, query: str, k: int = 3) -> List[str]:
        """Semantic search over knowledge base"""
        if not self.vector_store:
            return []
        
        docs = self.vector_store.similarity_search(query, k=k)
        return [doc.page_content for doc in docs]
    
    def add_document(self, text: str, metadata: dict):
        """Add new document to KB"""
        if not self.vector_store:
            self._initialize_vector_store()
        
        self.vector_store.add_texts([text], metadatas=[metadata])
    
    def save_local(self, path: str):
        """Persist vector store"""
        if self.vector_store:
            self.vector_store.save_local(path)
    
    @classmethod
    def load_local(cls, path: str, embeddings: Optional[Embeddings] = None):
        """Load persisted vector store"""
        instance = cls(embeddings=embeddings)
        instance.vector_store = FAISS.load_local(path, instance.embeddings)
        return instance


# ============================================================================
# PHASE 3: ADD TOOLS (For complex workflows)
# ============================================================================

@tool
def check_ticket_status(ticket_number: str) -> dict:
    """
    Check the status of an IT support ticket.
    
    Args:
        ticket_number: The ticket number (e.g., IT-1234)
        
    Returns:
        Dictionary with ticket details
    """
    # Import here to avoid circular dependency
    from quickbase_manager import QuickBaseManager
    import asyncio
    
    qb = QuickBaseManager()
    
    # Run async function in sync context
    loop = asyncio.get_event_loop()
    ticket = loop.run_until_complete(qb.get_ticket(ticket_number))
    
    if ticket:
        return {
            'found': True,
            'ticket_number': ticket['ticket_number'],
            'status': ticket['status'],
            'priority': ticket['priority'],
            'subject': ticket['subject'],
            'submitted_date': ticket['submitted_date']
        }
    return {'found': False, 'message': f'Ticket {ticket_number} not found'}


@tool
def search_company_docs(query: str, max_results: int = 5) -> List[dict]:
    """
    Search company documentation and policies.
    
    Args:
        query: Search query
        max_results: Maximum number of results
        
    Returns:
        List of relevant documents
    """
    # Future: Connect to SharePoint, Confluence, etc.
    # For now, return placeholder
    return [{
        'title': 'IT Security Policy',
        'url': 'https://company.sharepoint.com/policies/security',
        'snippet': 'Relevant policy excerpt...'
    }]


@tool
def check_system_status(system_name: str) -> dict:
    """
    Check if a company system is operational.
    
    Args:
        system_name: Name of system (e.g., 'VPN', 'Email', 'SharePoint')
        
    Returns:
        System status information
    """
    # Future: Connect to monitoring systems
    return {
        'system': system_name,
        'status': 'operational',
        'last_check': '2025-01-15 10:30:00'
    }


@tool  
def escalate_to_human(issue: str, urgency: str = 'normal') -> dict:
    """
    Escalate issue to human IT staff immediately.
    
    Args:
        issue: Description of the issue
        urgency: One of: 'normal', 'high', 'critical'
        
    Returns:
        Escalation confirmation
    """
    # Future: Send to IT Slack channel, create high-priority ticket, etc.
    return {
        'escalated': True,
        'message': 'Issue escalated to IT team',
        'urgency': urgency
    }


# ============================================================================
# AGENT WITH TOOLS (When needed for complex workflows)
# ============================================================================

class ITSupportAgent:
    """
    Tool-using agent for complex multi-step support
    Only use when simple router → handler pattern isn't enough
    """
    def __init__(self, llm: Optional[Any] = None):
        from langchain.agents import create_openai_functions_agent, AgentExecutor
        from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
        from langchain_openai import ChatOpenAI
        
        self.llm = llm or ChatOpenAI(
            model=os.getenv("GPT5_MODEL", "gpt-4"),
            temperature=0,
            openai_api_base=os.getenv("GPT5_ENDPOINT"),
            openai_api_key=os.getenv("GPT5_API_KEY")
        )
        
        # Define available tools
        self.tools = [
            check_ticket_status,
            search_company_docs,
            check_system_status,
            escalate_to_human
        ]
        
        # Create agent prompt
        prompt = ChatPromptTemplate.from_messages([
            ("system", """You are an IT support agent with access to tools.
            
            You can:
            - Check ticket statuses
            - Search company documentation
            - Check system status
            - Escalate to human IT staff
            
            Use tools when needed, but prefer direct answers when possible.
            Be concise and action-oriented."""),
            ("user", "{input}"),
            MessagesPlaceholder(variable_name="agent_scratchpad"),
        ])
        
        # Create agent
        agent = create_openai_functions_agent(self.llm, self.tools, prompt)
        self.executor = AgentExecutor(agent=agent, tools=self.tools, verbose=True)
    
    def run(self, question: str) -> str:
        """Execute agent with tools"""
        result = self.executor.invoke({"input": question})
        return result["output"]


# ============================================================================
# CONVERSATION HISTORY (For follow-ups)
# ============================================================================

class ConversationMemory:
    """
    Lightweight conversation memory for follow-up questions
    Inspired by your session LangGraph implementation
    """
    def __init__(self, max_turns: int = 5):
        self.max_turns = max_turns
        self.conversations = {}  # session_id -> messages
    
    def add_message(self, session_id: str, role: str, content: str):
        """Add message to conversation history"""
        if session_id not in self.conversations:
            self.conversations[session_id] = []
        
        self.conversations[session_id].append({
            'role': role,
            'content': content
        })
        
        # Keep only last N turns
        if len(self.conversations[session_id]) > self.max_turns * 2:
            self.conversations[session_id] = self.conversations[session_id][-(self.max_turns * 2):]
    
    def get_history(self, session_id: str) -> List[dict]:
        """Get conversation history"""
        return self.conversations.get(session_id, [])
    
    def clear(self, session_id: str):
        """Clear conversation history"""
        if session_id in self.conversations:
            del self.conversations[session_id]


# ============================================================================
# ENHANCED SUPPORT CHAIN WITH CONTEXT
# ============================================================================

class ITSupportChainV2:
    """
    Enhanced version with RAG, tools, and conversation memory
    Backward compatible with V1
    """
    def __init__(self, use_rag: bool = False, use_tools: bool = False):
        from support_chain import SupportRouter, QuickFixHandler, TicketHandler
        
        self.router = SupportRouter()
        self.quick_fix = QuickFixHandler()
        self.ticket_handler = TicketHandler()
        
        # Optional enhancements
        self.use_rag = use_rag
        self.use_tools = use_tools
        
        if use_rag:
            self.kb_rag = KnowledgeBaseRAG()
        
        if use_tools:
            self.agent = ITSupportAgent()
        
        self.memory = ConversationMemory()
    
    def process(
        self, 
        question: str, 
        session_id: Optional[str] = None,
        use_agent: bool = False
    ) -> dict:
        """
        Process with optional RAG, tools, and context
        
        Args:
            question: User question
            session_id: Optional session for conversation memory
            use_agent: Whether to use tool-using agent
        """
        # Add to conversation history
        if session_id:
            self.memory.add_message(session_id, 'user', question)
        
        # If complex issue and tools enabled, use agent
        if use_agent and self.use_tools:
            response = self.agent.run(question)
            
            if session_id:
                self.memory.add_message(session_id, 'assistant', response)
            
            return {
                'type': 'agent_response',
                'response': response
            }
        
        # Otherwise use original router pattern
        # (V1 behavior - fast and efficient)
        intent = self.router.route(question)
        
        if intent.intent_type == "quick_fix":
            # Optionally enhance with RAG
            if self.use_rag:
                rag_context = self.kb_rag.search(question)
                # Could inject rag_context into handler
            
            response = self.quick_fix.handle(question, intent.category)
            
            if session_id:
                self.memory.add_message(session_id, 'assistant', response.solution)
            
            return {
                "type": "solution",
                "solution": response.solution,
                "confidence": response.confidence,
                "offer_ticket": response.offer_ticket,
                "category": intent.category,
                "priority": intent.priority
            }
        
        # Rest same as V1...
        elif intent.intent_type == "needs_ticket":
            recommendation = self.ticket_handler.recommend(question)
            return {
                "type": "ticket_needed",
                "recommendation": recommendation.dict(),
                "reasoning": intent.reasoning
            }
        
        # etc.


# ============================================================================
# MIGRATION PATH
# ============================================================================

"""
MIGRATION GUIDE:

PHASE 1 (Current): Router + Handlers + Structured Outputs
- ✅ Deployed
- Fast single-shot responses
- No RAG needed yet
- Direct, efficient

PHASE 2 (Add RAG): When you have documents to index
```python
# Initialize with RAG
chain = ITSupportChainV2(use_rag=True)

# Add documents
chain.kb_rag.add_document(
    text="VPN Configuration Guide...",
    metadata={'source': 'sharepoint', 'doc_id': '123'}
)

# Saves to disk
chain.kb_rag.save_local('./kb_vectors')

# Later: load it
chain = ITSupportChainV2(use_rag=True)
chain.kb_rag = KnowledgeBaseRAG.load_local('./kb_vectors')
```

PHASE 3 (Add Tools): For complex multi-step workflows
```python
# Initialize with tools
chain = ITSupportChainV2(use_rag=True, use_tools=True)

# Use agent for complex issues
result = chain.process(
    "Check if VPN is down and escalate if needed",
    use_agent=True
)
```

PHASE 4 (Add Conversation): For follow-up questions
```python
# Track conversation
chain = ITSupportChainV2(use_rag=True)

result1 = chain.process("My VPN won't connect", session_id="user123")
result2 = chain.process("I tried that already", session_id="user123")
# Agent knows context from previous turn
```
"""