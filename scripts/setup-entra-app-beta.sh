#!/bin/bash
# Convenience wrapper: new app named m365-agent-cli-beta and writes EWS_CLIENT_ID to .env.beta
# (leaves your default ~/.config/m365-agent-cli/.env untouched).
# Override with M365_ENTRA_APP_NAME / M365_ENTRA_ENV_FILE before running if needed.

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export M365_ENTRA_APP_NAME="${M365_ENTRA_APP_NAME:-m365-agent-cli-beta}"
export M365_ENTRA_ENV_FILE="${M365_ENTRA_ENV_FILE:-$HOME/.config/m365-agent-cli/.env.beta}"
exec "$SCRIPT_DIR/setup-entra-app.sh" "$@"
