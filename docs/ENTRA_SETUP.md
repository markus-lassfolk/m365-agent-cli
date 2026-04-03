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
   **WSL:** If `./scripts/setup-entra-app-beta.sh` fails with `cannot execute: required file not found`, the script had Windows (CRLF) line endings; run `bash scripts/setup-entra-app-beta.sh` instead, or ensure `*.sh` files use LF (the repo stores them as LF). On `/mnt/f/` clones, `core.autocrlf` on Windows can still rewrite line endings until normalized.

### Using Microsoft Graph PowerShell
Use **PowerShell 7.4.x LTS** for this flow. **PowerShell 7.5+ preview** (e.g. **7.6**) often breaks **Microsoft.Graph**: `Connect-MgGraph` succeeds, then **`New-MgApplication` fails** with `DeviceCodeCredential` / null-reference errors. Either install [7.4.x from releases](https://github.com/PowerShell/PowerShell/releases), or use **Azure CLI** + the bash script below (no Graph module). The script can be forced on 7.5+ with `M365_ENTRA_ALLOW_PREVIEW_PS=1` (not recommended).

1. Install the [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation). The script **`setup-entra-app.ps1` requires `Microsoft.Graph` module version 2.12.0 or newer** (checked on startup). To see what you have:
   ```powershell
   Get-Module Microsoft.Graph -ListAvailable | Sort-Object Version
   ```
2. If the module is missing or too old, install or update:
   ```powershell
   # First-time install (CurrentUser scope)
   Install-Module Microsoft.Graph -Scope CurrentUser -Force

   # Upgrade an existing install
   Update-Module Microsoft.Graph -Force
   ```
   If `Update-Module` fails (conflicting versions), reinstall with:
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
   ```
   If you have both **1.x** (e.g. under `C:\Program Files\WindowsPowerShell\Modules`) and **2.x** under your user profile, remove or upgrade the old **AllUsers** copy so only a current **2.x** remains—otherwise the wrong version can load first.
3. Run the setup script (if you are not already signed in to Graph, the script runs `Connect-MgGraph` and prompts for sign-in):
   ```powershell
   ./scripts/setup-entra-app.ps1
   ```
   **Integrated terminal (VS Code / Cursor):** browser sign-in can look “stuck” or fail with WAM / assembly errors. Prefer **device code** so the URL and code print in the terminal:
   ```powershell
   ./scripts/setup-entra-app.ps1 -UseDeviceCode
   ```
   Or: `$env:M365_ENTRA_USE_DEVICE_CODE = "1"; ./scripts/setup-entra-app.ps1`
   Alternatively open **Windows Terminal** or **PowerShell** outside the editor and run the same script there.

   **If `Connect-MgGraph` fails with `System.ComponentModel.Primitives` / `Version=10.0.0.0`:**  
   That usually means **PowerShell 7.5+ preview** (**.NET 10**) plus **Microsoft.Graph** / **Azure.Identity** do not load cleanly in that host yet—**device code does not fix it**. Prefer one of:
   - **PowerShell 7.4.x LTS** (stable): [PowerShell releases](https://github.com/PowerShell/PowerShell/releases) — install the latest **7.4** patch, not **7.5 preview**.
   - **Fresh session without your profile** (other modules can cause assembly conflicts):  
     `pwsh -NoProfile -File .\scripts\setup-entra-app.ps1`
   - **Windows PowerShell 5.1**:  
     `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup-entra-app.ps1`
   - **Azure CLI** (no Graph PowerShell): `az login`, then run **`./scripts/setup-entra-app.sh`** or **`setup-entra-app-beta.sh`** from Git Bash or WSL (see the Bash section above).

   To sign in yourself first (optional):
   ```powershell
   Connect-MgGraph -Scopes "Application.ReadWrite.All"
   ```

Both scripts will output your **Client ID** (`EWS_CLIENT_ID`) upon success. Proceed to **Step 4** to configure your environment.

### Second app (beta / testing)

The automated scripts **always create a new** app registration; they do not update an existing one. To avoid overwriting `EWS_CLIENT_ID` in your main `.env` (and to keep production refresh tokens tied to the original app), use a **separate display name** and a **separate env file** for beta.

**Bash (Azure CLI)** — env vars:

```bash
M365_ENTRA_APP_NAME="m365-agent-cli-beta" \
M365_ENTRA_ENV_FILE="$HOME/.config/m365-agent-cli/.env.beta" \
./scripts/setup-entra-app.sh
```

Or use the wrapper (defaults shown above):

```bash
./scripts/setup-entra-app-beta.sh
```

**PowerShell (Graph)** — env vars:

```powershell
$env:M365_ENTRA_APP_NAME = "m365-agent-cli-beta"
$env:M365_ENTRA_ENV_FILE = "$HOME/.config/m365-agent-cli/.env.beta"
./scripts/setup-entra-app.ps1
```

Or:

```powershell
./scripts/setup-entra-app-beta.ps1
```

To **only** create the app in Entra and print the Client ID (no `.env` changes), set `M365_ENTRA_SKIP_ENV=1` (bash) or `$env:M365_ENTRA_SKIP_ENV = "1"` (PowerShell) before running the main script.

Then point the CLI at the beta file and run `login` / `verify-token`:

```bash
export M365_AGENT_ENV_FILE="$HOME/.config/m365-agent-cli/.env.beta"
m365-agent-cli login
m365-agent-cli verify-token
```

`M365_AGENT_ENV_FILE` must be set in the shell **before** starting the CLI (it is not read from inside `.env`).

---

## Option 2: Manual Setup

### 1. Create the App Registration
1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Navigate to **Applications** > **App registrations** > **New registration**.
3. Name your application (e.g., "m365-agent-cli").
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

**Full table (purpose of each scope):** see **[GRAPH_SCOPES.md](./GRAPH_SCOPES.md)** — keep the portal list aligned with that file and with [`src/lib/graph-oauth-scopes.ts`](../src/lib/graph-oauth-scopes.ts).

#### Microsoft Graph Permissions
1. Go to **API permissions**.
2. Click **Add a permission** > **Microsoft Graph** > **Delegated permissions**.
3. Search for and select the following scopes:
   - `User.Read`
   - `Calendars.ReadWrite`
   - `Calendars.Read.Shared` (delegate / shared calendars via Graph)
   - `Calendars.ReadWrite.Shared`
   - `Mail.Send` (`send` via Graph `sendMail`; add explicitly — some tenants require it alongside `Mail.ReadWrite`)
   - `Mail.ReadWrite`
   - `Mail.Read.Shared` (delegate / shared mailboxes — `mail` / `calendar` with `--mailbox` to another user)
   - `Mail.ReadWrite.Shared`
   - `MailboxSettings.ReadWrite` (automatic replies / `oof`, master categories, mailbox settings)
   - `Place.Read.All` (Places API — `rooms`, meeting rooms in `create-event`)
   - `People.Read` (`find` — `/me/people`)
   - `User.Read.All` (`find` — `/users` search; **often requires admin consent**)
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
   - `Tasks.ReadWrite`
   - `Group.ReadWrite.All` (Planner groups, group-related calls; also covers `find` group search)
   - `Contacts.ReadWrite` (`contacts`; `outlook-graph` contact APIs)
   - `OnlineMeetings.ReadWrite` (`meeting` — standalone Teams online meetings)
   - `Notes.ReadWrite.All` (`onenote`)
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

After completing the setup (either manually or automatically), you need to capture your credentials for the global `~/.config/m365-agent-cli/.env` file:

1. **`EWS_CLIENT_ID`**: If you used the automated setup scripts, this is already appended to your `~/.config/m365-agent-cli/.env` file. If you used the manual setup, go to the **Overview** page of your App Registration, copy the **Application (client) ID**, and add it to your `~/.config/m365-agent-cli/.env` file as `EWS_CLIENT_ID=<id>`.
2. **Refresh Tokens**: Run the `login` command — it saves **`M365_REFRESH_TOKEN`** (preferred) and legacy `GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN` (same value):
   ```bash
   m365-agent-cli login
   ```
   It will prompt you to authenticate via the Device Code flow and will automatically save the refresh tokens into your `~/.config/m365-agent-cli/.env` file upon successful authentication.
