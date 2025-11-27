# GitHub Secrets Configuration

Add these secrets to your GitHub repository (Settings > Secrets > Actions):

## Required Secrets

### AZURE_CREDENTIALS
```json
{
  "clientId": "<your-service-principal-app-id>",
  "clientSecret": "<your-service-principal-password>",
  "subscriptionId": "<your-subscription-id>",
  "tenantId": "<your-tenant-id>"
}
```

To create this:
```bash
az ad sp create-for-rbac --name "github-actions-sp" \
  --role contributor \
  --scopes /subscriptions/{subscription-id}/resourceGroups/{resource-group} \
  --sdk-auth
```

### AZURE_FUNCTIONAPP_PUBLISH_PROFILE
Get from Azure Portal:
1. Go to your Function App
2. Click "Get publish profile"
3. Copy entire XML content

### Other Secrets
- `QB_USER_TOKEN`: Your QuickBase user token
- `GPT5_API_KEY`: Your GPT-5 API key
- `TEAMS_APP_SECRET`: Teams app secret

## Setting Secrets via CLI
```bash
gh secret set AZURE_CREDENTIALS < azure-creds.json
gh secret set QB_USER_TOKEN
```
