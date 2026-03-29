#!/bin/bash
# Automates creating the Entra ID App Registration using the Azure CLI (az).
# Ensure you are logged in using `az login` before running this script.

set -e

APP_NAME="m365-agent-cli"

echo "Checking az login status..."
if ! az account show > /dev/null 2>&1; then
    echo "You are not logged in. Please run 'az login' first."
    exit 1
fi

echo "Creating Entra ID App Registration: $APP_NAME..."

# Create the application allowing Microsoft accounts and Organizational accounts
APP_ID=$(az ad app create --display-name "$APP_NAME" --sign-in-audience AzureADandPersonalMicrosoftAccount --query "appId" -o tsv)
OBJECT_ID=$(az ad app list --display-name "$APP_NAME" --query "[0].id" -o tsv)

if [ -z "$APP_ID" ]; then
    echo "Failed to create application."
    exit 1
fi

echo "Successfully created App! Client ID (App ID): $APP_ID"
echo "Object ID: $OBJECT_ID"

echo "Configuring public client flows (isFallbackPublicClient) and Redirect URI (http://localhost)..."
# Set as public client and configure Redirect URI
az ad app update \
    --id "$OBJECT_ID" \
    --set publicClient='{"redirectUris":["http://localhost"]}' isFallbackPublicClient=true

echo "Adding Required Resource Access (API Permissions) for Graph API and Exchange Online..."

# Construct the requiredResourceAccess JSON for Graph API and Exchange Online scopes
TEMP_JSON=$(mktemp)
cat <<EOF > "$TEMP_JSON"
[
  {
    "resourceAppId": "00000003-0000-0000-c000-000000000000",
    "resourceAccess": [
      { "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope" },
      { "id": "1ec239c2-d7c9-4623-a91a-a9775856bb36", "type": "Scope" },
      { "id": "024d486e-b451-40bb-833d-3e66d98c5c73", "type": "Scope" },
      { "id": "863451e7-0667-486c-a5d6-d135439485f0", "type": "Scope" },
      { "id": "89fe6a52-be36-487e-b7d8-d061c450a026", "type": "Scope" },
      { "id": "2219042f-cab5-40cc-b0d2-16b1540b4c5f", "type": "Scope" },
      { "id": "4e46008b-f24c-477d-8fff-7bb4ec7aafe0", "type": "Scope" },
      { "id": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", "type": "Scope" }
    ]
  },
  {
    "resourceAppId": "00000002-0000-0ff1-ce00-000000000000",
    "resourceAccess": [
      { "id": "3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5", "type": "Scope" }
    ]
  }
]
EOF

az ad app update --id "$OBJECT_ID" --required-resource-accesses @"$TEMP_JSON"
rm -f "$TEMP_JSON"

echo ""
echo "=================================================================================="
echo "Setup Complete!"
echo "Client ID (EWS_CLIENT_ID): $APP_ID"
echo "Tenant ID: Common (since audience is AzureADandPersonalMicrosoftAccount)"
echo ""

# Update or append EWS_CLIENT_ID to .env
if [ -f .env ] && grep -q "^EWS_CLIENT_ID=" .env; then
    sed -i.bak "s/^EWS_CLIENT_ID=.*/EWS_CLIENT_ID=$APP_ID/" .env && rm -f .env.bak
    echo "Updated EWS_CLIENT_ID in .env file in the current directory."
else
    echo "EWS_CLIENT_ID=$APP_ID" >> .env
    echo "Appended EWS_CLIENT_ID to .env file in the current directory."
fi

echo ""
echo "Next steps:"
echo "1. Go to the Azure Portal (https://entra.microsoft.com/) to grant admin consent"
echo "   for the scopes if required by your tenant."
echo "2. Run 'clippy login' to start the interactive login flow and get the"
echo "   refresh tokens to store in GRAPH_REFRESH_TOKEN and EWS_REFRESH_TOKEN."
echo "=================================================================================="
