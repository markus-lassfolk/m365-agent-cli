<#
.SYNOPSIS
    Automates creating the Entra ID App Registration using the Microsoft Graph PowerShell SDK.
.DESCRIPTION
    Ensure you are logged in using `Connect-MgGraph` with Directory.AccessAsUser.All or Application.ReadWrite.All
    before running this script.
#>

$AppName = "m365-agent-cli"

Write-Host "Creating Entra ID App Registration: $AppName..."

# Define the API permissions (Required Resource Access)
# Microsoft Graph API (00000003-0000-0000-c000-000000000000)
$GraphResourceAccess = @(
    @{ Id = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"; Type = "Scope" }, # User.Read
    @{ Id = "1ec239c2-d7c9-4623-a91a-a9775856bb36"; Type = "Scope" }, # Calendars.ReadWrite
    @{ Id = "024d486e-b451-40bb-833d-3e66d98c5c73"; Type = "Scope" }, # Mail.ReadWrite
    @{ Id = "863451e7-0667-486c-a5d6-d135439485f0"; Type = "Scope" }, # Files.ReadWrite.All
    @{ Id = "89fe6a52-be36-487e-b7d8-d061c450a026"; Type = "Scope" }, # Sites.ReadWrite.All
    @{ Id = "2219042f-cab5-40cc-b0d2-16b1540b4c5f"; Type = "Scope" }, # Tasks.ReadWrite
    @{ Id = "4e46008b-f24c-477d-8fff-7bb4ec7aafe0"; Type = "Scope" }, # Group.ReadWrite.All
    @{ Id = "7427e0e9-2fba-42fe-b0c0-848c9e6a8182"; Type = "Scope" }  # offline_access
)

# Office 365 Exchange Online (00000002-0000-0ff1-ce00-000000000000)
$ExchangeResourceAccess = @(
    @{ Id = "3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5"; Type = "Scope" }  # EWS.AccessAsUser.All
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

# Create the Application
try {
    $AppParams = @{
        DisplayName = $AppName
        SignInAudience = "AzureADandPersonalMicrosoftAccount"
        PublicClient = $PublicClient
        IsFallbackPublicClient = $true
        RequiredResourceAccess = $RequiredResourceAccess
    }
    
    $App = New-MgApplication @AppParams
    
    Write-Host "`n=================================================================================="
    Write-Host "Setup Complete!"
    Write-Host "Client ID (EWS_CLIENT_ID): $($App.AppId)"
    Write-Host "Object ID: $($App.Id)"
    Write-Host "Tenant ID: Common (since audience is AzureADandPersonalMicrosoftAccount)"
    Write-Host ""
    
$ConfigDir = Join-Path -Path $env:USERPROFILE -ChildPath ".config\m365-agent-cli"
    if (-not (Test-Path -Path $ConfigDir)) {
        New-Item -ItemType Directory -Force -Path $ConfigDir | Out-Null
    }
    $ConfigEnv = Join-Path -Path $ConfigDir -ChildPath ".env"

    # Update or append EWS_CLIENT_ID to .env
    if (Test-Path $ConfigEnv) {
        $envContent = Get-Content -Path $ConfigEnv -Raw
        if ($envContent -match "(?m)^EWS_CLIENT_ID=.*$") {
            $envContent = $envContent -replace "(?m)^EWS_CLIENT_ID=.*$", "EWS_CLIENT_ID=$($App.AppId)"
            Set-Content -Path $ConfigEnv -Value $envContent.TrimEnd()
            Write-Host "Updated EWS_CLIENT_ID in $ConfigEnv."
        } else {
            Add-Content -Path $ConfigEnv -Value "EWS_CLIENT_ID=$($App.AppId)"
            Write-Host "Appended EWS_CLIENT_ID to $ConfigEnv."
        }
    } else {
        Set-Content -Path $ConfigEnv -Value "EWS_CLIENT_ID=$($App.AppId)"
        Write-Host "Created .env file with EWS_CLIENT_ID in $ConfigEnv."
    }

    Write-Host ""
    Write-Host "App Name: $AppName"
    Write-Host "Direct Link: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/$AppId/isMSAApp~/false"
    Write-Host "(Note: It may take 1-3 minutes for a newly created app to fully propagate and appear in the Azure Portal)"
    Write-Host ""
    Write-Host "Next steps:"
    Write-Host "1. Go to the Azure Portal (https://entra.microsoft.com/) to grant admin consent"
    Write-Host "   for the scopes if required by your tenant."
    Write-Host "2. Run 'm365-agent-cli login' to start the interactive login flow and get the"
    Write-Host "   refresh tokens to store in GRAPH_REFRESH_TOKEN and EWS_REFRESH_TOKEN."
    Write-Host "3. Run 'm365-agent-cli verify-token' to verify your granted scopes!"
    Write-Host "   - Missing 'EWS.AccessAsUser.All'? Calendar/Mail functions will fail."
    Write-Host "   - Missing 'Files.ReadWrite.All'? OneDrive/SharePoint functions will fail."
    Write-Host "   - Missing 'Tasks.ReadWrite'? Planner/To-Do functions will fail."
    Write-Host "   - Missing 'Sites.ReadWrite.All'? Site Pages functions will fail."
    Write-Host "=================================================================================="
} catch {
    Write-Error "Failed to create application. Ensure you are authenticated with Connect-MgGraph and have sufficient privileges."
    Write-Error $_.Exception.Message
}
