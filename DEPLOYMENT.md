# IT Support Bot - Deployment Guide

## Overview
Deploy the IT Support Bot to Azure Functions + Teams in 4 steps:
1. Create Azure Bot Service (get App ID + Secret)
2. Create Azure Function App (deploy code)
3. Create Teams App Package (run script to update manifest)
4. Install in Teams

---

## Step 1: Create Azure Bot Service

### 1a. Create the Bot Resource
1. Go to [Azure Portal](https://portal.azure.com)
2. Click **Create a resource**
3. Search for **"Azure Bot"**
4. Click **Create**
5. Fill in:
   - **Bot handle**: `dispatch-it-support-bot`
   - **Subscription**: Your subscription
   - **Resource group**: Create new or use existing (e.g., `rg-it-support-bot`)
   - **Pricing tier**: F0 (free) for testing, S1 for production
   - **Microsoft App ID**: **Create new Microsoft App ID**

6. Click **Review + Create** → **Create**

### 1b. Get App ID and Secret
1. After creation, go to the bot resource
2. Click **Configuration** in left menu
3. Copy the **Microsoft App ID** (save this - you'll need it for the manifest!)
4. Click **Manage Password** next to Microsoft App ID
5. Click **+ New client secret**
   - Description: `IT Support Bot Secret`
   - Expires: 24 months
6. **COPY THE SECRET VALUE NOW** (you can't see it again!)

### 1c. Enable Teams Channel
1. In your Azure Bot, click **Channels**
2. Click **Microsoft Teams**
3. Accept the terms
4. Click **Apply**

---

## Step 2: Create Azure Function App

### 2a. Create the Function App
1. In Azure Portal, click **Create a resource**
2. Search for **"Function App"**
3. Click **Create**
4. Fill in:
   - **Subscription**: Same as bot
   - **Resource group**: Same as bot (e.g., `rg-it-support-bot`)
   - **Function App name**: `dispatch-it-support-func` (must be globally unique - save this name!)
   - **Runtime stack**: Python
   - **Version**: 3.11
   - **Region**: Same as bot (e.g., East US)
   - **Operating System**: Linux
   - **Plan type**: Consumption (Serverless)

5. Click **Review + Create** → **Create**

### 2b. Configure App Settings
1. Go to your Function App
2. Click **Configuration** → **Application settings**
3. Add these settings (click **+ New application setting** for each):

```
GPT5_ENDPOINT       = <your GPT-5 endpoint>
GPT5_API_KEY        = <your GPT-5 API key>
GPT5_MODEL          = gpt-5

QB_REALM            = dispatchenergy.quickbase.com
QB_USER_TOKEN       = <your QuickBase user token>
QB_APP_ID           = bvajz826s
QB_TICKETS_TABLE_ID = bvajz9sqr

TEAMS_APP_ID        = <App ID from Step 1b>
TEAMS_APP_SECRET    = <Secret from Step 1b>
TEAMS_TENANT_ID     = <your tenant ID>
```

4. Click **Save**

### 2c. Deploy the Code

**Option A: VS Code (easiest)**
1. Install Azure Functions extension in VS Code
2. Open the `it-support-bot` folder
3. Press F1, type "Azure Functions: Deploy to Function App"
4. Select your function app
5. Confirm deployment

**Option B: Azure CLI**
```bash
cd it-support-bot
func azure functionapp publish dispatch-it-support-func
```

**Option C: ZIP Deploy**
```bash
cd it-support-bot
zip -r deploy.zip . -x "*.git*" -x "__pycache__/*" -x "*.pyc" -x "teams-app/*"
az functionapp deployment source config-zip \
  --resource-group rg-it-support-bot \
  --name dispatch-it-support-func \
  --src deploy.zip
```

### 2d. Configure Bot Messaging Endpoint
1. Go back to your **Azure Bot** resource
2. Click **Configuration**
3. Set **Messaging endpoint** to:
   ```
   https://dispatch-it-support-func.azurewebsites.net/api/messages
   ```
   (Replace `dispatch-it-support-func` with your actual function app name)
4. Click **Apply**

---

## Step 3: Create Teams App Package

### 3a. Prepare the Manifest

You have two values from the previous steps:
- **Bot App ID**: From Step 1b (e.g., `12345678-1234-1234-1234-123456789abc`)
- **Function App Name**: From Step 2a (e.g., `dispatch-it-support-func`)

**Option A: Use the script**
```bash
cd teams-app
python prepare_manifest.py "YOUR_BOT_APP_ID" "YOUR_FUNCTION_APP_NAME"
```

This updates the manifest and creates `it-support-bot.zip` automatically.

**Option B: Manual edit**
1. Open `teams-app/manifest.json`
2. Replace all `{{BOT_APP_ID}}` with your Bot App ID (appears 3 times)
3. Replace all `{{FUNCTION_APP_NAME}}` with your Function App name (appears 2 times)
4. Create the ZIP:
   ```bash
   cd teams-app
   zip ../it-support-bot.zip manifest.json color.png outline.png
   ```

### 3b. Verify the Manifest
The manifest should have:
- `"id": "your-actual-bot-app-id"` (a GUID)
- `"botId": "your-actual-bot-app-id"` (same GUID)
- `"validDomains"` including `"your-function-app.azurewebsites.net"`

---

## Step 4: Install in Teams

### Option A: Sideload for Testing
1. Open Microsoft Teams
2. Click **Apps** (left sidebar)
3. Click **Manage your apps** (bottom)
4. Click **Upload an app** → **Upload a custom app**
5. Select your `it-support-bot.zip`
6. Click **Add**

### Option B: Admin Deployment (org-wide)
1. Go to [Teams Admin Center](https://admin.teams.microsoft.com)
2. Click **Teams apps** → **Manage apps**
3. Click **Upload new app**
4. Select your `it-support-bot.zip`
5. Configure policies for who can use it

---

## Testing

### Test Health Endpoint
```bash
curl https://dispatch-it-support-func.azurewebsites.net/api/health
```
Should return: `{"status": "healthy", ...}`

### Test in Teams
1. Open Teams
2. Start a chat with "IT Support" bot
3. Try these:
   - "I can't reset my password" → Should get password reset steps
   - "My VPN keeps disconnecting" → Should get VPN troubleshooting
   - "I need Adobe Creative Suite" → Should offer ticket form
   - "/help" → Show commands
   - "/ticket" → Create ticket form
   - "/status" → Show your tickets

### Expected Behavior
The bot should:
1. **Always respond with a solution** (from KB or GPT)
2. **Offer "This helped" and "Still need help" buttons**
3. **Create a tracking ticket in QuickBase** (status: "Bot Assisted")
4. **Let user escalate to real ticket if needed**

---

## Troubleshooting

### Bot doesn't respond
1. Check Function App logs: Function App → **Log stream**
2. Verify messaging endpoint is correct (ends with `/api/messages`)
3. Verify TEAMS_APP_ID and TEAMS_APP_SECRET are correct

### "App not found" in Teams
1. Verify manifest.json has correct Bot App ID (must match Azure Bot)
2. Check that `validDomains` includes your function app domain
3. Re-create the ZIP and re-upload

### Authentication errors
1. Verify client secret hasn't expired
2. Regenerate secret if needed
3. Update TEAMS_APP_SECRET in Function App settings

### QuickBase errors
1. Verify QB_USER_TOKEN is valid
2. Check QuickBase field mappings match your table

### View Function Logs
1. Go to Function App
2. Click **Log stream** (left menu)
3. Watch real-time logs as you test

---

## Files in this Package

```
it-support-bot/
├── function_app.py           # Main Azure Function entry point
├── teams_handler.py          # Teams API integration
├── ai_processor.py           # GPT integration + built-in knowledge base
├── quickbase_manager.py      # QuickBase ticket operations
├── adaptive_cards.py         # Teams UI cards
├── requirements.txt          # Python dependencies
├── host.json                 # Azure Functions config
├── local.settings.json.template  # Env vars template
├── support_chain.py.future   # LangChain version (for later)
└── teams-app/
    ├── manifest.json         # Teams app manifest (update with your IDs!)
    ├── prepare_manifest.py   # Script to update manifest
    ├── color.png             # App icon (192x192)
    └── outline.png           # App icon (32x32)
```

## Flow Summary

```
User sends message in Teams
         ↓
Azure Function receives POST to /api/messages
         ↓
function_app.py routes to handle_support_question()
         ↓
ai_processor.get_support_response() called
    - First checks built-in knowledge base (keywords)
    - Falls back to GPT-5 for complex queries
         ↓
Bot sends solution card to user (ALWAYS responds!)
         ↓
Tracking ticket created in QuickBase (status: "Bot Assisted")
         ↓
User clicks "This helped" → Thanks message
User clicks "Still need help" → Ticket form opens
```
