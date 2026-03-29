# Entra ID (Azure AD) App Registration Setup

This guide explains how to set up the Entra ID Application Registration required for this application, which needs specific permissions to access Microsoft Graph and Office 365 Exchange Online (EWS).

You can configure the application manually via the Azure Portal, or automatically using the provided scripts.

---

## Option 1: Automated Setup (Recommended)

We provide scripts to automatically create and configure the App Registration.

### Using Azure CLI (Bash)
1. Install the [Azure CLI](https://learn.microsoft.com/en-us/cli/azure/install-azure-cli).
2. Log in to your tenant:
   ```bash
   az login
   ```
3. Run the setup script:
   ```bash
   ./scripts/setup-entra-app.sh
   ```

### Using Microsoft Graph PowerShell
1. Install the [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation).
2. Log in with the necessary permissions:
   ```powershell
   Connect-MgGraph -Scopes "Application.ReadWrite.All"
   ```
3. Run the setup script:
   ```powershell
   ./scripts/setup-entra-app.ps1
   ```

Both scripts will output your **Client ID** (`EWS_CLIENT_ID`) upon success. Proceed to **Step 4** to configure your environment.

---

## Option 2: Manual Setup

### 1. Create the App Registration
1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Navigate to **Applications** > **App registrations** > **New registration**.
3. Name your application (e.g., "Clippy").
4. For **Supported account types**, choose **Accounts in any organizational directory and personal Microsoft accounts** (or the scope that fits your use case).
5. Leave the Redirect URI blank for now and click **Register**.

### 2. Configure Redirect URI (Public Client)
1. In your new App Registration, go to **Authentication** under the Manage menu.
2. Under **Platform configurations**, click **Add a platform** and select **Mobile and desktop applications**.
3. Check the box for `http://localhost` and click **Configure**.
4. Scroll down to **Advanced settings** and ensure **Allow public client flows** is set to **Yes**. This is required for the native mobile & desktop application flow.
5. Click **Save**.

### 3. Configure API Permissions
The application requires specific Delegated permissions for both Microsoft Graph and Office 365 Exchange Online.

#### Microsoft Graph Permissions
1. Go to **API permissions**.
2. Click **Add a permission** > **Microsoft Graph** > **Delegated permissions**.
3. Search for and select the following scopes:
   - `User.Read`
   - `Calendars.ReadWrite`
   - `Mail.ReadWrite`
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
   - `Tasks.ReadWrite`
   - `Group.ReadWrite.All`
   - `offline_access`
4. Click **Add permissions**.

#### Office 365 Exchange Online Permissions
1. Click **Add a permission** > **APIs my organization uses**.
2. Search for **Office 365 Exchange Online** and select it.
3. Choose **Delegated permissions**.
4. Check the box for `EWS.AccessAsUser.All`.
5. Click **Add permissions**.

*(Optional but recommended: Click **Grant admin consent** on the API permissions page to pre-approve these scopes for your tenant.)*

---

## 4. Update Your Environment Variables

After completing the setup (either manually or automatically), you need to capture your credentials for the application's `.env` file:

1. **`EWS_CLIENT_ID`**: If you used the manual setup, go to the **Overview** page of your App Registration and copy the **Application (client) ID**.
2. **Refresh Tokens**: To get your initial `GRAPH_REFRESH_TOKEN` and `EWS_REFRESH_TOKEN`, use the interactive login flow provided by the app. Run the application, and it will prompt you to authenticate.
3. Once you authenticate and accept the scopes, the application will provide or log the refresh tokens.
4. Save the Client ID and refresh tokens into your `.env` file.
