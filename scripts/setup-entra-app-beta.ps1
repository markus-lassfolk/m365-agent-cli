<#
.SYNOPSIS
    Creates a separate beta Entra app and writes EWS_CLIENT_ID to .env.beta (not the default .env).
#>
if (-not $env:M365_ENTRA_APP_NAME) { $env:M365_ENTRA_APP_NAME = "m365-agent-cli-beta" }
if (-not $env:M365_ENTRA_ENV_FILE) {
    $env:M365_ENTRA_ENV_FILE = Join-Path $HOME ".config/m365-agent-cli/.env.beta"
}
& "$PSScriptRoot/setup-entra-app.ps1" @args
