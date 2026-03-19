# Jonny Tacos - Teams IT Support Bot

An AI-powered Microsoft Teams bot that serves as an intelligent IT Service Desk, integrating with QuickBase for ticket management and GPT-5 (with Azure OpenAI fallback) for automated IT support. Built on Azure Functions (Python 3.11) with LangChain-based routing.

## Table of Contents

- [Architecture Overview](#architecture-overview)
- [Tech Stack](#tech-stack)
- [Project Structure](#project-structure)
- [Getting Started](#getting-started)
- [Environment Variables](#environment-variables)
- [API Endpoints](#api-endpoints)
- [Core Modules](#core-modules)
- [Bot Commands](#bot-commands)
- [QuickBase Schema](#quickbase-schema)
- [M365 Provisioning & Automation](#m365-provisioning--automation)
- [Adaptive Cards UI](#adaptive-cards-ui)
- [AI & LangChain Routing](#ai--langchain-routing)
- [Webhook Integrations](#webhook-integrations)
- [Local Development](#local-development)
- [Testing](#testing)
- [Deployment](#deployment)
- [Customization](#customization)
- [Troubleshooting](#troubleshooting)

---

## Architecture Overview

```
Teams Client
    |
Azure Function App (Python 3.11)
    |
    +-- /api/messages (main bot endpoint)
    |     +-> function_app.py          -- Request router
    |     +-> teams_handler.py         -- Teams Bot Framework API
    |     +-> support_chain.py         -- LangChain intent classification & routing
    |     |     \-> ai_processor.py    -- GPT-5 / Azure OpenAI + knowledge base
    |     +-> automation_manager.py    -- M365 automation lifecycle
    |     |     \-> m365_provisioning.py -- Microsoft Graph API provisioning
    |     +-> quickbase_manager.py     -- QuickBase ticket CRUD
    |     \-> adaptive_cards.py        -- Adaptive Card UI builder
    |
    +-- /api/webhook/ticket-closed     -- QuickBase webhook (ticket closed)
    +-- /api/webhook/ticket-update     -- QuickBase webhook (status changes)
    \-- /api/health                    -- Health check
    |
    v                          v
QuickBase (Tickets DB)   Microsoft Graph API (M365 resources)
```

**Request flow**: Teams activity hits `/api/messages` -> `function_app.py` routes by activity type (message, invoke, conversationUpdate) -> LangChain classifies intent -> response generated via AI or knowledge base -> Adaptive Card rendered back to user.

---

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Runtime | Azure Functions v4 (Python 3.11+) |
| AI/ML | GPT-5 (custom endpoint), Azure OpenAI fallback, LangChain 0.1.0 |
| Ticket Management | QuickBase API |
| Chat Platform | Microsoft Teams (Bot Framework) |
| UI | Adaptive Cards v1.5 |
| M365 Automation | Microsoft Graph API |
| Infrastructure | Azure (Consumption plan) |
| CI/CD | GitHub Actions |
| Auth | Bot Framework JWT, Azure AD |

---

## Project Structure

```
jonny-tacos/
|-- function_app.py            # Main Azure Function entry point & all route handlers (~1,950 lines)
|-- support_chain.py           # LangChain routing, intent classification, duplicate detection (~725 lines)
|-- ai_processor.py            # GPT-5/Azure OpenAI integration, knowledge base (~530 lines)
|-- quickbase_manager.py       # QuickBase ticket CRUD operations (~490 lines)
|-- teams_handler.py           # Teams Bot Framework API, proactive messaging (~690 lines)
|-- adaptive_cards.py          # Adaptive Card templates & builders (~740 lines)
|-- automation_manager.py      # M365 automation request lifecycle (~220 lines)
|-- m365_provisioning.py       # Microsoft Graph API provisioning (~1,260 lines)
|-- local_test.py              # Interactive CLI for local testing (~630 lines)
|-- qb_debug.py                # QuickBase debugging utility
|-- future_state.py            # RAG & tool integration roadmap (reference only)
|
|-- host.json                  # Azure Functions config (routing, logging)
|-- manifest.json              # Teams app manifest (commands, scopes, permissions)
|-- requirements.txt           # Python dependencies
|-- .env.example               # Environment variable template
|
|-- .github/
|   |-- workflows/deploy.yml   # GitHub Actions CI/CD pipeline
|   \-- SECRETS_TEMPLATE.md    # Required GitHub secrets documentation
|
|-- tests/
|   |-- __init__.py
|   \-- test_quickbase.py      # QuickBase integration tests
|
|-- docs/
|   \-- IT_Support_Bot_User_Guide.html  # End-user guide
|
|-- setup-repo.sh              # Repository setup script (venv + deps)
|-- run_local.sh               # Local dev launcher (Azure Functions Core Tools)
|-- DEPLOYMENT.md              # Step-by-step Azure deployment guide
|-- GIT_COMMANDS.md            # Git workflow reference
\-- jonny_tacos_logo.png       # Bot logo
```

---

## Getting Started

### Prerequisites

- Python 3.11+
- Azure CLI (`az`)
- Azure Functions Core Tools v4 (`func`)
- An Azure subscription
- Microsoft Teams admin access
- QuickBase account with API access
- GPT-5 endpoint or Azure OpenAI resource

### 1. Clone and Install

```bash
git clone <repo-url>
cd jonny-tacos

# Option A: Use the setup script
chmod +x setup-repo.sh
./setup-repo.sh

# Option B: Manual setup
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Install Azure Functions Core Tools
npm install -g azure-functions-core-tools@4
```

### 2. Configure Environment

Copy the environment template and fill in your values:

```bash
cp .env.example local.settings.json
```

Edit `local.settings.json` — see the [Environment Variables](#environment-variables) section below for details on every variable.

### 3. Run Locally

```bash
# Using the script
chmod +x run_local.sh
./run_local.sh

# Or directly
func start --python
```

Local endpoint: `http://localhost:7071/api/messages`

### 4. Test Interactively

```bash
python local_test.py
```

This launches a CLI that simulates Teams interactions — tests LangChain routing, QuickBase operations, and AI responses without needing a Teams connection.

---

## Environment Variables

All variables are set in `local.settings.json` for local dev or as Application Settings in Azure for production.

### Teams Bot (Required)

| Variable | Description |
|----------|------------|
| `TEAMS_APP_ID` | Azure Bot app registration ID (GUID) |
| `TEAMS_APP_SECRET` | Bot app registration secret |
| `TEAMS_TENANT_ID` | Azure AD tenant ID |

### QuickBase (Required)

| Variable | Description |
|----------|------------|
| `QB_REALM` | Your QuickBase realm (e.g., `yourcompany.quickbase.com`) |
| `QB_USER_TOKEN` | QuickBase API user token |
| `QB_APP_ID` | QuickBase application ID |
| `QB_TICKETS_TABLE_ID` | Table ID for the tickets table |
| `QB_WEBHOOK_SECRET` | (Optional) Secret for validating webhook payloads |

### AI Configuration (Required - at least one provider)

| Variable | Description |
|----------|------------|
| `GPT5_ENDPOINT` | Custom GPT-5 endpoint URL |
| `GPT5_API_KEY` | API key for GPT-5 |
| `GPT5_MODEL` | Model name (default: `gpt-5`) |
| `AZURE_OPENAI_ENDPOINT` | Fallback Azure OpenAI endpoint |
| `AZURE_OPENAI_KEY` | Azure OpenAI API key |
| `AZURE_OPENAI_DEPLOYMENT` | Deployment name (default: `gpt-4`) |

The bot tries GPT-5 first and falls back to Azure OpenAI if unavailable.

### M365 Automation (Optional - for provisioning features)

| Variable | Description |
|----------|------------|
| `AUTOMATION_ADMIN_EMAIL` | Admin email that receives provisioning approval requests |
| `M365_GRAPH_CLIENT_ID` | Graph API app registration client ID |
| `M365_GRAPH_CLIENT_SECRET` | Graph API app registration secret |
| `M365_GRAPH_TENANT_ID` | (Optional) Override tenant for Graph API calls |
| `M365_GRAPH_DOMAIN` | Company domain for mailbox/resource creation |

### Monitoring (Optional)

| Variable | Description |
|----------|------------|
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | Azure Application Insights connection string |

---

## API Endpoints

| Route | Method | Auth | Description |
|-------|--------|------|-------------|
| `/api/messages` | POST | Anonymous (Bot Framework JWT validated internally) | Main Teams bot message handler |
| `/api/webhook/ticket-closed` | POST | Webhook secret | QuickBase webhook — sends closure notification to user |
| `/api/webhook/ticket-update` | POST | Webhook secret | QuickBase webhook — sends status update notification |
| `/api/health` | GET | None | Health check (returns 200 OK) |

---

## Core Modules

### `function_app.py` — Entry Point & Router

The main Azure Function app. Handles all HTTP triggers, routes incoming Teams activities by type:

- **`message`** — User text messages. Routed through LangChain intent classification, then dispatched to AI/knowledge base or QuickBase operations.
- **`invoke`** — Adaptive Card button clicks and form submissions (ticket forms, approval cards, feedback).
- **`conversationUpdate`** — Fires when bot is added to a conversation; sends a welcome card.

Key responsibilities:
- Bot Framework authentication and activity parsing
- Conversation state management
- Command dispatching (`/help`, `/ticket`, `/status`, `/stats`, `/resolve`)
- On-behalf-of ticket filing (file tickets for other users)
- Inline reply handling for streaming chat conversations

### `support_chain.py` — LangChain Intent Classification

Uses LangChain 0.1.0 to classify and route user messages through a multi-step chain:

- **`SupportIntent`** — Classifies message as: `quick_fix`, `needs_human`, `automation_request`, `status_check`, or `command`
- **`SupportResponse`** — Generates IT support responses (always provides a solution, never says "can't help")
- **`FollowUpCheck`** — Streaming topic detection to prevent duplicate tickets when users provide follow-up details

### `ai_processor.py` — AI Engine & Knowledge Base

Dual-layer response system:

1. **Knowledge base** (fast, no API call) — ~10 predefined IT scenarios with keyword matching:
   - Password resets, VPN issues, Teams/Office 365 problems, email issues, network connectivity, hardware, software installation
2. **GPT fallback** (API call) — For queries not matched by the knowledge base

Additional AI capabilities:
- `analyze_ticket_requirement()` — Determines if a ticket should be created
- `suggest_category()` / `suggest_priority()` — Auto-classifies tickets
- `process_feedback()` — Learns from user thumbs-up/down on responses

### `quickbase_manager.py` — Ticket Operations

Full QuickBase CRUD via REST API:

- `create_ticket()` — Creates ticket with auto-generated number (`IT-YYYYMMDDHHMISS`), sets due date by SLA
- `get_ticket()` — Retrieves a ticket by number
- `get_user_tickets()` — Lists a user's tickets (optional status filter)
- `update_ticket()` — Updates arbitrary fields
- `resolve_ticket()` — Marks resolved with resolution text and timestamp
- `get_ticket_statistics()` — Aggregated metrics for admin dashboard

### `teams_handler.py` — Teams API Integration

Handles all Microsoft Teams communication:

- **Messaging**: `send_message()`, `send_card()`, `update_card()`, `send_to_channel()`
- **Typing indicators**: `send_typing_indicator()`
- **User info**: `get_user_info()`, `get_user_aad_id()`, `get_channel_members()`
- **Auth**: `get_auth_token()` (Bot Framework), `get_graph_token()` (Microsoft Graph), `validate_auth_header()` (JWT)
- **Proactive messaging**: `send_notification_to_user()`, `send_proactive_message()` — sends 1:1 messages to users (e.g., ticket closure notifications)

### `adaptive_cards.py` — UI Card Builder

Builds Adaptive Card v1.5 JSON for all bot interactions:

- `create_welcome_card()` — Bot introduction on first install
- `create_help_card()` — Command reference
- `create_ticket_form()` — Ticket submission form (subject, description, priority, category)
- `create_ticket_confirmation_card()` — Post-creation confirmation with ticket number and link
- `create_error_card()` — Error display
- Solution cards with thumbs-up/down feedback buttons
- Status update and closure notification cards

Branded with Azure Blue (`#0078D4`).

### `automation_manager.py` — M365 Automation Lifecycle

Manages automation requests through a stateful workflow:

```
GATHERING_INFO -> PENDING_APPROVAL -> APPROVED -> EXECUTING -> COMPLETED / FAILED
```

- 2-hour TTL for pending requests
- Admin receives an approval Adaptive Card
- Extensible via `AutomationHandler` abstract base class
- Built-in handler: `M365ProvisioningHandler` (shared mailboxes, Teams, SharePoint sites)

### `m365_provisioning.py` — Microsoft Graph Provisioning

Detects natural language requests for M365 resources and provisions them via Graph API:

- **Resource types**: Shared Mailbox, Microsoft Team, SharePoint Site
- **Keyword detection**: Recognizes phrases like "shared email", "new team", "sharepoint site"
- **Admin approval**: Creates approval card, sends to configured admin, waits for response
- **Graph API permissions required**: `Group.ReadWrite.All`, `Directory.ReadWrite.All`, `Sites.ReadWrite.All`

---

## Bot Commands

### User Commands

| Command | Description |
|---------|-------------|
| `/help` | Show available commands and features |
| `/ticket` | Open the ticket creation form |
| `/status [ticket#]` | Check a specific ticket's status |
| (natural language) | Ask any IT question — AI responds with troubleshooting steps |

### Admin Commands

| Command | Description |
|---------|-------------|
| `/resolve [ticket#] [resolution]` | Resolve a ticket with resolution notes |
| `/stats` | View IT dashboard statistics (open/closed counts, avg resolution time) |

### Usage Examples

**Natural language support:**
```
User: "My Outlook won't sync emails"
Bot:  [Provides step-by-step troubleshooting via AI]
      [Offers to create a ticket if the issue persists]
```

**On-behalf-of ticket filing:**
```
User: "Create a ticket for john@company.com - laptop won't boot"
Bot:  [Creates ticket with john@company.com as submitter]
      [Notifies both the filer and John]
```

**M365 provisioning:**
```
User: "I need a new shared mailbox for the marketing team"
Bot:  [Detects M365 provisioning intent]
      [Collects config: display name, email, members]
      [Sends approval card to admin]
      [Creates mailbox after approval, confirms to user]
```

---

## QuickBase Schema

The tickets table must have these fields with the specified field IDs:

| Field Name | Field ID | Type | Notes |
|------------|----------|------|-------|
| Ticket Number | 6 | Text | Auto-generated (`IT-YYYYMMDDHHMISS`) |
| Subject | 7 | Text | Issue title |
| Description | 8 | Text - Multi-line | Full issue details |
| Priority | 9 | Text - Multiple Choice | `Low`, `Medium`, `High`, `Critical` |
| Category | 10 | Text - Multiple Choice | See categories below |
| Status | 11 | Text - Multiple Choice | See statuses below |
| Submitted Date | 12 | Date/Time | Auto-set on creation |
| Due Date | 13 | Date | Calculated from priority SLA |
| Resolved Date | 14 | Date/Time | Set when resolved |
| Resolution | 15 | Text - Multi-line | Resolution details |
| Time Spent | 16 | Numeric | Hours spent |
| Submitted By | 19 | Email | Submitter's email |

**Ticket categories**: Password Reset, Software Installation, Hardware Issue, Network/Connectivity, Email/Outlook, VPN/Remote Access, Account Access, Microsoft Teams, Phone/Voicemail, Printer Issue, Security Concern, New Equipment Request, Other

**Ticket statuses**: New, In Progress, Awaiting User, Awaiting IT, Resolved, Closed, Cancelled

**Priority SLA** (due date calculation): Low = longest, Critical = shortest

---

## M365 Provisioning & Automation

The bot can provision M365 resources through natural language requests. This requires the Graph API environment variables to be configured.

### Supported Resources

| Resource | Keywords Detected | What Gets Created |
|----------|-------------------|-------------------|
| Shared Mailbox | "shared email", "shared mailbox", "team mailbox" | Exchange shared mailbox with specified members |
| Microsoft Team | "new team", "create a team", "teams channel" | Team with owners, members, and default channels |
| SharePoint Site | "sharepoint site", "document site", "collaboration site" | SharePoint site with permissions |

### Approval Workflow

1. User requests a resource in natural language
2. Bot detects intent and collects configuration via Adaptive Card form
3. Bot sends an approval card to the configured `AUTOMATION_ADMIN_EMAIL`
4. Admin approves or rejects
5. On approval, bot executes provisioning via Graph API
6. Bot notifies the requester with results

Pending requests expire after 2 hours.

---

## Webhook Integrations

QuickBase can send webhooks to notify the bot of ticket changes. The bot then proactively messages the affected user in Teams.

### Ticket Closed Webhook

**Endpoint**: `POST /api/webhook/ticket-closed`

When a ticket is closed in QuickBase, the bot sends a closure notification Adaptive Card to the user who submitted the ticket.

### Ticket Update Webhook

**Endpoint**: `POST /api/webhook/ticket-update`

When a ticket status changes, the bot sends an update notification to the submitter.

Both webhooks validate the `QB_WEBHOOK_SECRET` if configured.

---

## Local Development

### Running the Function App

```bash
source .venv/bin/activate
func start --python
```

This starts the Azure Functions host locally at `http://localhost:7071`.

### Interactive CLI Testing

```bash
python local_test.py
```

The CLI test tool:
- Loads config from `local.settings.json`
- Simulates user message routing through LangChain
- Tests QuickBase ticket creation and retrieval
- Tests AI responses without needing Teams

### Debugging QuickBase

```bash
python qb_debug.py
```

Utility for testing QuickBase API connectivity and field mappings.

---

## Testing

```bash
python -m pytest tests/ -v
```

Test files are in `tests/`. Current coverage focuses on QuickBase integration.

---

## Deployment

### CI/CD (GitHub Actions)

The pipeline (`.github/workflows/deploy.yml`) triggers on pushes to `main`:

1. Checks out code
2. Sets up Python 3.11 with pip caching
3. Installs dependencies to `.python_packages/lib/site-packages`
4. Creates deployment ZIP (excludes `.git`, `*.md`, `local.settings.json`)
5. Authenticates with Azure via service principal (`AZURE_CREDENTIALS` secret)
6. Deploys via `az functionapp deployment source config-zip`

**Required GitHub Secrets** (see `.github/SECRETS_TEMPLATE.md`):
- `AZURE_CREDENTIALS` — Service principal JSON for Azure login
- All environment variables listed in the [Environment Variables](#environment-variables) section

### Manual Deployment

See [DEPLOYMENT.md](DEPLOYMENT.md) for a full step-by-step guide covering:

1. Azure Bot Service setup (App ID + Secret)
2. Azure Function App creation (Linux, Python 3.11, Consumption plan)
3. Application settings configuration
4. Code deployment (VS Code / Azure CLI / ZIP)
5. Messaging endpoint configuration in Azure Bot
6. Teams app manifest packaging and sideloading

### Teams App Package

```bash
# Update manifest.json with your Bot App ID
# Then create the package:
zip -r teams-app.zip manifest.json color.png outline.png

# Upload to Teams Admin Center or sideload for testing
```

The manifest (`manifest.json`) defines:
- Bot scopes: `personal`, `team`, `groupchat`
- Commands: `/help`, `/ticket`, `/status`, `/stats`
- Permissions: `identity`, `messageTeamMembers`

---

## Customization

### Adding IT Categories

In `quickbase_manager.py`, extend the categories list:

```python
self.categories = [
    'Password Reset',
    'Software Installation',
    # Add your categories here
]
```

Also add the category to your QuickBase field choices.

### Adding Knowledge Base Entries

In `ai_processor.py`, add entries to the knowledge base dict:

```python
self.knowledge_base = {
    "your_issue": {
        "keywords": ["keyword1", "keyword2"],
        "solution": "Step-by-step solution text",
        "category": "Category Name",
        "needs_ticket": False
    }
}
```

### Adding Automation Handlers

Extend `AutomationHandler` in `automation_manager.py`:

```python
class MyHandler(AutomationHandler):
    def detect_intent(self, message: str) -> bool:
        # Return True if this handler should process the message
        ...

    def execute(self, request: AutomationRequest) -> dict:
        # Perform the automation
        ...
```

### Customizing Cards

Edit `adaptive_cards.py` to modify card layouts, colors, and actions. Cards use Adaptive Card schema v1.5.

---

## Troubleshooting

### Bot Not Responding in Teams

1. Check Azure Function logs in the Azure Portal (Monitor > Log Stream)
2. Verify the messaging endpoint is set correctly in Azure Bot configuration
3. Confirm Teams app permissions and that the bot is installed
4. Check `TEAMS_APP_ID` and `TEAMS_APP_SECRET` are correct

### QuickBase Integration Issues

1. Verify `QB_USER_TOKEN` has read/write permissions on the tickets table
2. Confirm field IDs match the [QuickBase Schema](#quickbase-schema) above
3. Check `QB_REALM` format (e.g., `yourcompany.quickbase.com`)
4. Run `python qb_debug.py` to test connectivity

### AI Responses Not Working

1. Test GPT-5 endpoint directly with curl
2. If GPT-5 is down, verify Azure OpenAI fallback is configured
3. Check API key validity and rate limits
4. Review Application Insights for error details

### Webhook Notifications Not Arriving

1. Verify webhook URLs are publicly accessible (not localhost)
2. Check `QB_WEBHOOK_SECRET` matches the value configured in QuickBase
3. Confirm the user has a conversation reference stored (they must have messaged the bot at least once)

### M365 Provisioning Failures

1. Verify Graph API app registration has required permissions (`Group.ReadWrite.All`, `Directory.ReadWrite.All`, `Sites.ReadWrite.All`)
2. Confirm admin consent has been granted for the permissions
3. Check `AUTOMATION_ADMIN_EMAIL` is a valid Teams user
4. Review Graph API error responses in function logs

---

## Future Roadmap

- **Phase 2 — RAG**: Vector store (FAISS) for semantic knowledge base search, SharePoint document ingestion, PDF processing
- **Phase 3 — Tool Integration**: LangChain tool calling for AD user management, system administration, and IT service automation

---

## Security Notes

- Store all secrets in Azure Key Vault (never commit `local.settings.json`)
- Bot Framework JWT tokens are validated on every request
- QuickBase webhooks are validated via shared secret
- Graph API uses client credential flow with scoped permissions
- All communications use HTTPS

---

**License**: Internal use only - Proprietary
