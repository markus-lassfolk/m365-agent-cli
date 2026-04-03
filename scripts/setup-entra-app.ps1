<#
.SYNOPSIS
    Automates creating the Entra ID App Registration using the Microsoft Graph PowerShell SDK.
.DESCRIPTION
    Always creates a NEW registration; it does not modify an existing app.
    Requires Microsoft.Graph PowerShell module >= 2.12 (see script constant). If you are not
    signed in to Graph, the script runs `Connect-MgGraph` with Application.ReadWrite.All
    (browser or device-code sign-in may appear).

    App name: first positional arg, or env M365_ENTRA_APP_NAME, default m365-agent-cli.
    Env file: second positional arg, or env M365_ENTRA_ENV_FILE, default ~/.config/m365-agent-cli/.env
    Skip writing .env: -SkipEnv or M365_ENTRA_SKIP_ENV=1|true
    Device code sign-in: -UseDeviceCode or M365_ENTRA_USE_DEVICE_CODE=1 (recommended in VS Code/Cursor
    terminals when browser/WAM sign-in fails or appears to hang). Device code clears any prior Graph
    session first so stale tokens do not skip Connect-MgGraph.

    -ReconnectGraph or M365_ENTRA_RECONNECT_GRAPH=1: Disconnect-MgGraph then sign in again.

    PowerShell 7.5+ preview (e.g. 7.6): Microsoft.Graph often breaks after Connect-MgGraph; this script exits
    unless M365_ENTRA_ALLOW_PREVIEW_PS=1. Prefer PowerShell 7.4.x LTS or Azure CLI + setup-entra-app.sh.

    See docs/ENTRA_SETUP.md - Second app (beta / testing).
#>

param(
    [Parameter(Position = 0)]
    [string] $AppDisplayName,
    [Parameter(Position = 1)]
    [string] $EnvPath,
    [switch] $SkipEnv,
    [switch] $UseDeviceCode,
    [switch] $ReconnectGraph
)

$ErrorActionPreference = "Stop"

# Connect-MgGraph can succeed while New-MgApplication fails (DeviceCodeCredential null ref) on PS 7.5+ preview / .NET 10.
$psv = $PSVersionTable.PSVersion
$graphBrokenPwsh = ($PSVersionTable.PSEdition -eq 'Core' -and $psv.Major -eq 7 -and $psv.Minor -ge 5)
if ($graphBrokenPwsh -and $env:M365_ENTRA_ALLOW_PREVIEW_PS -notin @('1', 'true')) {
    Write-Host ""
    Write-Host "This script uses Microsoft.Graph, which is unreliable on PowerShell $psv (typical: Connect-MgGraph works," -ForegroundColor Yellow
    Write-Host "then New-MgApplication fails with DeviceCodeCredential / Object reference)." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Use one of:" -ForegroundColor Yellow
    Write-Host "  - PowerShell 7.4.x LTS: https://github.com/PowerShell/PowerShell/releases (install 7.4, then run this script with that pwsh)" -ForegroundColor Cyan
    Write-Host "  - Azure CLI + bash (no Graph module):  az login" -ForegroundColor Cyan
    Write-Host "    bash scripts/setup-entra-app-beta.sh" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To try anyway on this host:  `$env:M365_ENTRA_ALLOW_PREVIEW_PS='1'" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}

# Baseline SDK: v2 line, aligns with Connect-MgGraph -NoWelcome and current app-registration cmdlets.
$MinimumMicrosoftGraphVersion = [version]'2.12.0'
# Prefer the newest -ListAvailable copy. Import-Module -MinimumVersion alone can resolve an older
# Microsoft.Graph under Program Files before CurrentUser 2.x (PSModulePath order).
$graphCandidates = @(Get-Module -ListAvailable -Name Microsoft.Graph | Sort-Object Version -Descending)
$newestGraph = $graphCandidates | Select-Object -First 1

if (-not $newestGraph -or $newestGraph.Version -lt $MinimumMicrosoftGraphVersion) {
    Write-Host ""
    Write-Host "This script requires the Microsoft.Graph PowerShell module version $MinimumMicrosoftGraphVersion or newer." -ForegroundColor Yellow
    if ($graphCandidates.Count -eq 0) {
        Write-Host "No Microsoft.Graph module was found." -ForegroundColor Yellow
        Write-Host "Install:" -ForegroundColor Yellow
        Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -Force" -ForegroundColor Cyan
    } else {
        Write-Host "Latest installed Microsoft.Graph version: $($newestGraph.Version)" -ForegroundColor Yellow
        Write-Host "Update:" -ForegroundColor Yellow
        Write-Host "  Update-Module Microsoft.Graph -Force" -ForegroundColor Cyan
        Write-Host "If that fails, reinstall:" -ForegroundColor Yellow
        Write-Host "  Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber" -ForegroundColor Cyan
    }
    Write-Host "Docs: https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation" -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}

try {
    Import-Module -Name Microsoft.Graph -RequiredVersion $newestGraph.Version -Force -ErrorAction Stop | Out-Null
} catch {
    Write-Host ""
    Write-Host "Failed to import Microsoft.Graph $($newestGraph.Version): $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "Path: $($newestGraph.Path)" -ForegroundColor DarkGray
    Write-Host "If multiple versions are installed, uninstall the old one (e.g. 1.x under Program Files) or repair the module." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

$AppName = if ($AppDisplayName) { $AppDisplayName }
elseif ($env:M365_ENTRA_APP_NAME) { $env:M365_ENTRA_APP_NAME }
else { "m365-agent-cli" }

$defaultEnv = Join-Path -Path $HOME -ChildPath ".config/m365-agent-cli/.env"
$ConfigEnv = if ($env:M365_ENTRA_ENV_FILE) { $env:M365_ENTRA_ENV_FILE }
elseif ($EnvPath) { $EnvPath }
else { $defaultEnv }

$skipWriteEnv = $SkipEnv -or ($env:M365_ENTRA_SKIP_ENV -in @("1", "true"))
$useDeviceCode = $UseDeviceCode -or ($env:M365_ENTRA_USE_DEVICE_CODE -in @("1", "true"))
$forceReconnect = $ReconnectGraph -or ($env:M365_ENTRA_RECONNECT_GRAPH -in @("1", "true"))

Write-Host "Creating Entra ID App Registration: $AppName..."
Write-Host "(This always creates a NEW registration; it does not modify an existing app.)"

# Stale Get-MgContext skips Connect-MgGraph; later calls then fail with DeviceCodeCredential / null reference.
if ($useDeviceCode -or $forceReconnect) {
    if ($useDeviceCode) {
        Write-Host "Clearing any existing Microsoft Graph session (required for device-code sign-in)..." -ForegroundColor DarkGray
    } else {
        Write-Host "Disconnecting existing Microsoft Graph session (-ReconnectGraph)..." -ForegroundColor DarkGray
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    $connectParams = @{
        Scopes        = "Application.ReadWrite.All"
        NoWelcome     = $true
        ErrorAction   = "Stop"
    }
    if ($useDeviceCode) {
        $connectParams["UseDeviceCode"] = $true
        Write-Host "Using device code sign-in (Application.ReadWrite.All)..." -ForegroundColor Cyan
        Write-Host "A URL and code will appear below - open the URL in a browser, enter the code, then return here." -ForegroundColor Cyan
    } else {
        Write-Host "Not connected to Microsoft Graph - starting sign-in (Application.ReadWrite.All)..." -ForegroundColor Cyan
        Write-Host "If nothing happens (common in VS Code/Cursor), cancel and run with -UseDeviceCode or set M365_ENTRA_USE_DEVICE_CODE=1" -ForegroundColor DarkGray
    }

    try {
        Connect-MgGraph @connectParams
        Write-Host "Signed in to Microsoft Graph. Sign-in used Application.ReadWrite.All only (needed to create this registration)." -ForegroundColor DarkGray
        Write-Host "Mail/Calendar/EWS scopes are applied to the new app below; you approve them when you run 'm365-agent-cli login'." -ForegroundColor DarkGray
    } catch {
        Write-Host ""
        Write-Host "Connect-MgGraph failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "If you see 'System.ComponentModel.Primitives' / 'Version=10.0.0.0', your PowerShell host is likely .NET 10" -ForegroundColor Yellow
        Write-Host "(e.g. pwsh 7.5 preview). Microsoft.Graph auth often breaks there until modules catch up." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Try in order:" -ForegroundColor Yellow
        Write-Host "  1. Same script, no profile (avoids other modules stealing assemblies):" -ForegroundColor Cyan
        Write-Host "     pwsh -NoProfile -File `"$PSCommandPath`"   # add -UseDeviceCode, app name, etc. as needed" -ForegroundColor White
        Write-Host "  2. Install PowerShell 7.4.x LTS (stable, .NET 8), not 7.5 preview:" -ForegroundColor Cyan
        Write-Host "     https://github.com/PowerShell/PowerShell/releases" -ForegroundColor White
        Write-Host "  3. Windows PowerShell 5.1 (desktop):" -ForegroundColor Cyan
        Write-Host "     powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -ForegroundColor White
        Write-Host "  4. Skip Graph PowerShell - use Azure CLI + bash script from repo root:" -ForegroundColor Cyan
        Write-Host "     az login" -ForegroundColor White
        Write-Host "     bash scripts/setup-entra-app.sh   # or setup-entra-app-beta.sh" -ForegroundColor White
        Write-Host ""
        exit 1
    }
}

# Define the API permissions (Required Resource Access)
# Microsoft Graph delegated scopes: mirror GRAPH_DEVICE_CODE_LOGIN_SCOPES in src/lib/graph-oauth-scopes.ts
# (each Id is the oauth2PermissionScopes id on the Microsoft Graph service principal).
# offline_access is listed last in this block; Exchange Online (EWS) is a separate resource below.
# Microsoft Graph API (00000003-0000-0000-c000-000000000000)
$GraphResourceAccess = @(
    # User.Read
    @{ Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"; Type = "Scope" },
    # Calendars.ReadWrite
    @{ Id = "1ec239c2-d7c9-4623-a91a-a9775856bb36"; Type = "Scope" },
    # Calendars.Read.Shared
    @{ Id = "2b9c4092-424d-4249-948d-b43879977640"; Type = "Scope" },
    # Calendars.ReadWrite.Shared
    @{ Id = "12466101-c9b8-439a-8589-dd09ee67e8e9"; Type = "Scope" },
    # Mail.Send
    @{ Id = "e383f46e-2787-4529-855e-0e479a3ffac0"; Type = "Scope" },
    # Mail.ReadWrite
    @{ Id = "024d486e-b451-40bb-833d-3e66d98c5c73"; Type = "Scope" },
    # Mail.Read.Shared
    @{ Id = "7b9103a5-4610-446b-9670-80643382c1fa"; Type = "Scope" },
    # Mail.ReadWrite.Shared
    @{ Id = "5df07973-7d5d-46ed-9847-1271055cbd51"; Type = "Scope" },
    # MailboxSettings.ReadWrite
    @{ Id = "818c620a-27a9-40bd-a6a5-d96f7d610b4b"; Type = "Scope" },
    # Place.Read.All
    @{ Id = "cb8f45a0-5c2e-4ea1-b803-84b870a7d7ec"; Type = "Scope" },
    # People.Read
    @{ Id = "ba47897c-39ec-4d83-8086-ee8256fa737d"; Type = "Scope" },
    # User.Read.All
    @{ Id = "a154be20-db9c-4678-8ab7-66f6cc099a59"; Type = "Scope" },
    # Files.ReadWrite.All
    @{ Id = "863451e7-0667-486c-a5d6-d135439485f0"; Type = "Scope" },
    # Sites.ReadWrite.All
    @{ Id = "89fe6a52-be36-487e-b7d8-d061c450a026"; Type = "Scope" },
    # Tasks.ReadWrite
    @{ Id = "2219042f-cab5-40cc-b0d2-16b1540b4c5f"; Type = "Scope" },
    # Group.ReadWrite.All
    @{ Id = "4e46008b-f24c-477d-8fff-7bb4ec7aafe0"; Type = "Scope" },
    # Contacts.ReadWrite
    @{ Id = "d56682ec-c09e-4743-aaf4-1a3aac4caa21"; Type = "Scope" },
    # Contacts.Read.Shared
    @{ Id = "242b9d9e-ed24-4d09-9a52-f43769beb9d4"; Type = "Scope" },
    # Contacts.ReadWrite.Shared
    @{ Id = "afb6c84b-06be-49af-80bb-8f3f77004eab"; Type = "Scope" },
    # OnlineMeetings.ReadWrite
    @{ Id = "a65f2972-a4f8-4f5e-afd7-69ccb046d5dc"; Type = "Scope" },
    # Notes.ReadWrite.All
    @{ Id = "64ac0503-b4fa-45d9-b544-71a463f05da0"; Type = "Scope" },
    # Team.ReadBasic.All
    @{ Id = "485be79e-c497-4b35-9400-0e3fa7f2a5d4"; Type = "Scope" },
    # Channel.ReadBasic.All
    @{ Id = "9d8982ae-4365-4f57-95e9-d6032a4c0b87"; Type = "Scope" },
    # ChannelMessage.Read.All
    @{ Id = "767156cb-16ae-4d10-8f8b-41b657c8c8c8"; Type = "Scope" },
    # ChannelMessage.Send
    @{ Id = "ebf0f66e-9fb1-49e4-a278-222f76911cf4"; Type = "Scope" },
    # Presence.Read.All
    @{ Id = "9c7a330d-35b3-4aa1-963d-cb2b9f927841"; Type = "Scope" },
    # Presence.ReadWrite
    @{ Id = "8d3c54a7-cf58-4773-bf81-c0cd6ad522bb"; Type = "Scope" },
    # Bookings.ReadWrite.All
    @{ Id = "948eb538-f19d-4ec5-9ccc-f059e1ea4c72"; Type = "Scope" },
    # Chat.ReadWrite
    @{ Id = "9ff7295e-131b-4d94-90e1-69fde507ac11"; Type = "Scope" },
    # offline_access
    @{ Id = "7427e0e9-2fba-42fe-b0c0-848c9e6a8182"; Type = "Scope" }
)

# Office 365 Exchange Online (00000002-0000-0ff1-ce00-000000000000)
$ExchangeResourceAccess = @(
    # EWS.AccessAsUser.All
    @{ Id = "3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5"; Type = "Scope" }
)

$RequiredResourceAccess = @(
    @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000"
        ResourceAccess = $GraphResourceAccess
    },
    @{
        ResourceAppId = "00000002-0000-0ff1-ce00-000000000000"
        ResourceAccess = $ExchangeResourceAccess
    }
)

# Define public client settings
$PublicClient = @{
    RedirectUris = @("http://localhost")
}

try {
    # One-shot New-MgApplication with full RequiredResourceAccess can trigger DeviceCodeCredential null refs on
    # some hosts (e.g. PowerShell 7.6 preview + Microsoft.Graph). Create base app first, then PATCH permissions.
    $AppParamsMinimal = @{
        DisplayName            = $AppName
        SignInAudience         = "AzureADandPersonalMicrosoftAccount"
        PublicClient           = $PublicClient
        IsFallbackPublicClient = $true
    }

    Write-Host "Creating app registration (step 1 of 2)..." -ForegroundColor DarkGray
    $App = New-MgApplication @AppParamsMinimal -ErrorAction Stop
    if (-not $App -or [string]::IsNullOrWhiteSpace($App.AppId)) {
        throw "New-MgApplication did not return a Client ID. Check Connect-MgGraph and Application.ReadWrite.All."
    }

    Write-Host "Applying Graph + Exchange API permissions (step 2 of 2)..." -ForegroundColor DarkGray
    Update-MgApplication -ApplicationId $App.Id -RequiredResourceAccess $RequiredResourceAccess -ErrorAction Stop

    Write-Host "`n=================================================================================="
    Write-Host "Setup Complete!"
    Write-Host "Client ID (EWS_CLIENT_ID): $($App.AppId)"
    Write-Host "Object ID: $($App.Id)"
    Write-Host "Tenant ID: Common (since audience is AzureADandPersonalMicrosoftAccount)"
    Write-Host ""

    $envDir = Split-Path -Parent $ConfigEnv
    if ($envDir -and -not (Test-Path -LiteralPath $envDir)) {
        New-Item -ItemType Directory -Force -Path $envDir | Out-Null
    }

    if ($skipWriteEnv) {
        Write-Host "M365_ENTRA_SKIP_ENV set - not writing any .env file."
        Write-Host "Use when testing: `$env:EWS_CLIENT_ID = '$($App.AppId)'"
    }
    else {
        Write-Host "Writing EWS_CLIENT_ID to: $ConfigEnv"
        if (Test-Path -LiteralPath $ConfigEnv) {
            $envContent = Get-Content -Path $ConfigEnv -Raw
            if ($envContent -match "(?m)^EWS_CLIENT_ID=.*$") {
                $envContent = $envContent -replace "(?m)^EWS_CLIENT_ID=.*$", "EWS_CLIENT_ID=$($App.AppId)"
                Set-Content -Path $ConfigEnv -Value $envContent.TrimEnd()
                Write-Host "Updated EWS_CLIENT_ID in $ConfigEnv."
            }
            else {
                Add-Content -Path $ConfigEnv -Value "EWS_CLIENT_ID=$($App.AppId)"
                Write-Host "Appended EWS_CLIENT_ID to $ConfigEnv."
            }
        }
        else {
            Set-Content -Path $ConfigEnv -Value "EWS_CLIENT_ID=$($App.AppId)"
            Write-Host "Created .env file with EWS_CLIENT_ID in $ConfigEnv."
        }
    }

    Write-Host ""
    Write-Host "App Name: $AppName"
    Write-Host "Direct Link: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$($App.AppId)/isMSAApp~/false"
    Write-Host "(Note: It may take 1-3 minutes for a newly created app to fully propagate and appear in the Azure Portal)"
    Write-Host ""
    Write-Host "Next steps:"
    Write-Host "1. Go to the Azure Portal (https://entra.microsoft.com/) to grant admin consent"
    Write-Host "   for the scopes if required by your tenant."
    Write-Host "2. Point the CLI at this app: ensure EWS_CLIENT_ID in your active .env matches the Client ID above"
    Write-Host "   (use a separate .env.beta for beta - see docs/ENTRA_SETUP.md)."
    Write-Host "3. Run 'm365-agent-cli login' to obtain refresh tokens for THIS app."
    Write-Host "4. Run 'm365-agent-cli verify-token' to verify your granted scopes!"
    Write-Host "   - Missing 'EWS.AccessAsUser.All'? Calendar/Mail functions will fail."
    Write-Host "   - Missing 'Files.ReadWrite.All'? OneDrive/SharePoint functions will fail."
    Write-Host "   - Missing 'Tasks.ReadWrite'? Planner/To-Do functions will fail."
    Write-Host "   - Missing 'Sites.ReadWrite.All'? Site Pages functions will fail."
    Write-Host "=================================================================================="
}
catch {
    Write-Host ""
    $detail = $_.Exception.Message
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
        $detail = "$detail`n$($_.ErrorDetails.Message)"
    }

    $looksLikeStaleAuth = $detail -match 'DeviceCodeCredential|InteractiveBrowserCredential'

    if ($looksLikeStaleAuth) {
        Write-Host "Microsoft Graph authentication failed inside an API call (DeviceCodeCredential / browser credential bug)." -ForegroundColor Red
        Write-Host $detail -ForegroundColor Red
        Write-Host ""
        Write-Host "If you already saw 'Signed in to Microsoft Graph' above, the session is fine; this is usually the host + SDK" -ForegroundColor Yellow
        Write-Host "(e.g. PowerShell 7.5+ preview / 7.6 with Microsoft.Graph), not a stale login. Disconnect-MgGraph rarely fixes it." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Try in order:" -ForegroundColor Yellow
        Write-Host "  1. PowerShell 7.4.x LTS (stable): https://github.com/PowerShell/PowerShell/releases  (install and use that pwsh)" -ForegroundColor Cyan
        Write-Host "  2. Azure CLI + bash (no Graph PowerShell):  az login" -ForegroundColor Cyan
        Write-Host "     bash scripts/setup-entra-app-beta.sh" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Current PS version: $($PSVersionTable.PSVersion)" -ForegroundColor DarkGray
        Write-Host ""
        exit 1
    }

    Write-Host "Failed to create the app registration (New-MgApplication)." -ForegroundColor Red
    Write-Host $detail -ForegroundColor Red
    Write-Host ""
    Write-Host "Why sign-in only showed one permission: Connect-MgGraph requests Application.ReadWrite.All so this script can CREATE" -ForegroundColor Yellow
    Write-Host "the registration. The CLI's Mail/Calendar/EWS/Graph scopes are added to that new app here; you approve those when" -ForegroundColor Yellow
    Write-Host "you run 'm365-agent-cli login' (or an admin grants tenant-wide consent in Entra)." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Common causes of this failure:" -ForegroundColor Yellow
    Write-Host "  - Your account is not allowed to create app registrations. Roles: Application Administrator or Global Administrator;" -ForegroundColor Yellow
    Write-Host "    or Entra: Users can register applications = Yes (User settings)." -ForegroundColor Yellow
    Write-Host "  - Guest / wrong tenant: ensure you signed into the directory where apps may be created." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
