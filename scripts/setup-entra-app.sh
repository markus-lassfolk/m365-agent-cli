#!/bin/bash
# Automates creating the Entra ID App Registration using the Azure CLI (az).
# Ensure you are logged in using `az login` before running this script.
#
# Usage:
#   ./scripts/setup-entra-app.sh
#   ./scripts/setup-entra-app.sh "m365-agent-cli-beta"
#   M365_ENTRA_APP_NAME="m365-agent-cli Beta" M365_ENTRA_ENV_FILE="$HOME/.config/m365-agent-cli/.env.beta" ./scripts/setup-entra-app.sh
#   M365_ENTRA_SKIP_ENV=1 ./scripts/setup-entra-app.sh   # create app only; print Client ID (does not modify .env)
#
# See docs/ENTRA_SETUP.md — "Second app (beta / testing)".

set -e

# App display name: first argument, or M365_ENTRA_APP_NAME, or default
APP_NAME="${1:-${M365_ENTRA_APP_NAME:-m365-agent-cli}}"

# Optional: path to a dedicated .env file (e.g. .env.beta) so production tokens are untouched.
# Precedence: M365_ENTRA_ENV_FILE > second CLI argument > default ~/.config/m365-agent-cli/.env
CONFIG_DIR="${HOME}/.config/m365-agent-cli"
if [ -n "${M365_ENTRA_ENV_FILE:-}" ]; then
  CONFIG_ENV="$M365_ENTRA_ENV_FILE"
elif [ -n "${2:-}" ]; then
  CONFIG_ENV="$2"
else
  CONFIG_ENV="${CONFIG_DIR}/.env"
fi

SKIP_ENV="${M365_ENTRA_SKIP_ENV:-0}"

echo "Checking az login status..."
if ! az account show > /dev/null 2>&1; then
  echo "You are not logged in. Please run 'az login' first."
  exit 1
fi

echo "Creating Entra ID App Registration: $APP_NAME..."
echo "(This always creates a NEW registration; it does not modify an existing app.)"

# Create the application — capture appId + object id from this response (do not list by name; names can duplicate)
CREATE_JSON=$(az ad app create --display-name "$APP_NAME" --sign-in-audience AzureADandPersonalMicrosoftAccount -o json)
if command -v jq >/dev/null 2>&1; then
  APP_ID=$(echo "$CREATE_JSON" | jq -r '.appId')
  OBJECT_ID=$(echo "$CREATE_JSON" | jq -r '.id')
elif command -v node >/dev/null 2>&1; then
  APP_ID=$(echo "$CREATE_JSON" | node -e "let s='';process.stdin.on('data',c=>s+=c);process.stdin.on('end',()=>{const j=JSON.parse(s);console.log(j.appId);});")
  OBJECT_ID=$(echo "$CREATE_JSON" | node -e "let s='';process.stdin.on('data',c=>s+=c);process.stdin.on('end',()=>{const j=JSON.parse(s);console.log(j.id);});")
else
  echo "Need jq or Node.js to parse az ad app create JSON output."
  exit 1
fi

if [ -z "$APP_ID" ] || [ -z "$OBJECT_ID" ] || [ "$APP_ID" = "null" ] || [ "$OBJECT_ID" = "null" ]; then
  echo "Failed to parse app create response. Raw output:"
  echo "$CREATE_JSON"
  exit 1
fi

echo "Successfully created App! Client ID (App ID): $APP_ID"
echo "Object ID: $OBJECT_ID"

echo "Configuring public client flows (isFallbackPublicClient) and Redirect URI (http://localhost)..."
az ad app update \
  --id "$OBJECT_ID" \
  --set publicClient='{"redirectUris":["http://localhost"]}' isFallbackPublicClient=true

echo "Adding Required Resource Access (API Permissions) for Graph API and Exchange Online..."
echo "(Graph delegated scopes align with GRAPH_DEVICE_CODE_LOGIN_SCOPES in src/lib/graph-oauth-scopes.ts)"

TEMP_JSON=$(mktemp)
cat <<EOF > "$TEMP_JSON"
[
  {
    "resourceAppId": "00000003-0000-0000-c000-000000000000",
    "resourceAccess": [
      { "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope" },
      { "id": "1ec239c2-d7c9-4623-a91a-a9775856bb36", "type": "Scope" },
      { "id": "2b9c4092-424d-4249-948d-b43879977640", "type": "Scope" },
      { "id": "12466101-c9b8-439a-8589-dd09ee67e8e9", "type": "Scope" },
      { "id": "e383f46e-2787-4529-855e-0e479a3ffac0", "type": "Scope" },
      { "id": "024d486e-b451-40bb-833d-3e66d98c5c73", "type": "Scope" },
      { "id": "7b9103a5-4610-446b-9670-80643382c1fa", "type": "Scope" },
      { "id": "5df07973-7d5d-46ed-9847-1271055cbd51", "type": "Scope" },
      { "id": "818c620a-27a9-40bd-a6a5-d96f7d610b4b", "type": "Scope" },
      { "id": "cb8f45a0-5c2e-4ea1-b803-84b870a7d7ec", "type": "Scope" },
      { "id": "ba47897c-39ec-4d83-8086-ee8256fa737d", "type": "Scope" },
      { "id": "a154be20-db9c-4678-8ab7-66f6cc099a59", "type": "Scope" },
      { "id": "863451e7-0667-486c-a5d6-d135439485f0", "type": "Scope" },
      { "id": "89fe6a52-be36-487e-b7d8-d061c450a026", "type": "Scope" },
      { "id": "2219042f-cab5-40cc-b0d2-16b1540b4c5f", "type": "Scope" },
      { "id": "4e46008b-f24c-477d-8fff-7bb4ec7aafe0", "type": "Scope" },
      { "id": "d56682ec-c09e-4743-aaf4-1a3aac4caa21", "type": "Scope" },
      { "id": "242b9d9e-ed24-4d09-9a52-f43769beb9d4", "type": "Scope" },
      { "id": "afb6c84b-06be-49af-80bb-8f3f77004eab", "type": "Scope" },
      { "id": "a65f2972-a4f8-4f5e-afd7-69ccb046d5dc", "type": "Scope" },
      { "id": "64ac0503-b4fa-45d9-b544-71a463f05da0", "type": "Scope" },
      { "id": "485be79e-c497-4b35-9400-0e3fa7f2a5d4", "type": "Scope" },
      { "id": "9d8982ae-4365-4f57-95e9-d6032a4c0b87", "type": "Scope" },
      { "id": "767156cb-16ae-4d10-8f8b-41b657c8c8c8", "type": "Scope" },
      { "id": "ebf0f66e-9fb1-49e4-a278-222f76911cf4", "type": "Scope" },
      { "id": "9c7a330d-35b3-4aa1-963d-cb2b9f927841", "type": "Scope" },
      { "id": "8d3c54a7-cf58-4773-bf81-c0cd6ad522bb", "type": "Scope" },
      { "id": "948eb538-f19d-4ec5-9ccc-f059e1ea4c72", "type": "Scope" },
      { "id": "9ff7295e-131b-4d94-90e1-69fde507ac11", "type": "Scope" },
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

mkdir -p "$(dirname "$CONFIG_ENV")"

if [ "$SKIP_ENV" = "1" ] || [ "$SKIP_ENV" = "true" ]; then
  echo "M365_ENTRA_SKIP_ENV set — not writing any .env file."
  echo "Export when testing: EWS_CLIENT_ID=$APP_ID"
else
  echo "Writing EWS_CLIENT_ID to: $CONFIG_ENV"
  if [ -f "$CONFIG_ENV" ] && grep -q "^EWS_CLIENT_ID=" "$CONFIG_ENV"; then
    sed -i.bak "s/^EWS_CLIENT_ID=.*/EWS_CLIENT_ID=$APP_ID/" "$CONFIG_ENV" && rm -f "$CONFIG_ENV.bak"
    echo "Updated EWS_CLIENT_ID in $CONFIG_ENV."
  else
    echo "EWS_CLIENT_ID=$APP_ID" >> "$CONFIG_ENV"
    echo "Appended EWS_CLIENT_ID to $CONFIG_ENV."
  fi
fi

echo ""
echo "App Name: $APP_NAME"
echo "Direct Link: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$APP_ID/isMSAApp~/false"
echo "(Note: It may take 1-3 minutes for a newly created app to fully propagate and appear in the Azure Portal)"
echo ""
echo "Next steps:"
echo "1. Go to the Azure Portal (https://entra.microsoft.com/) to grant admin consent"
echo "   for the scopes if required by your tenant."
echo "2. Point the CLI at this app: ensure EWS_CLIENT_ID in your active .env matches the Client ID above"
echo "   (use a separate .env.beta for beta — see docs/ENTRA_SETUP.md)."
echo "3. Run 'm365-agent-cli login' to obtain refresh tokens for THIS app."
echo "4. Run 'm365-agent-cli verify-token' to verify your granted scopes!"
echo "=================================================================================="
