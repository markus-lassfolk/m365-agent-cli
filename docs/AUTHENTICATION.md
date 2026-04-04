# Authentication

How **m365-agent-cli** signs in to Microsoft 365: OAuth2 device code, refresh tokens, cache files, shared mailboxes, and Graph vs EWS behavior.

- [README (overview)](../README.md)
- [Entra app setup](./ENTRA_SETUP.md) — scripts and portal steps
- [Graph scopes](./GRAPH_SCOPES.md) — what each delegated permission is for

---

**Need help setting up the Azure AD App?** Follow our [Automated Entra ID App Setup Guide](./ENTRA_SETUP.md) for bash and PowerShell scripts that configure the exact permissions you need in seconds. **Delegated Graph scopes (what each is for):** [GRAPH_SCOPES.md](./GRAPH_SCOPES.md).

**EWS retirement:** Microsoft is phasing out EWS for Exchange Online in favor of Microsoft Graph. Track migration work in [EWS_TO_GRAPH_MIGRATION_EPIC.md](./EWS_TO_GRAPH_MIGRATION_EPIC.md) (phased plan, inventory, Graph-primary + EWS-fallback strategy).

**Optional error reporting:** To receive CLI crashes and unhandled errors in [GlitchTip](https://glitchtip.com/) (Sentry-compatible), set **`GLITCHTIP_DSN`** (or **`SENTRY_DSN`**) in your environment. See [GLITCHTIP.md](./GLITCHTIP.md).

m365-agent-cli uses OAuth2 with a refresh token to authenticate against Microsoft 365. You need an Azure AD app registration.

### Setup

If you used the setup scripts from [ENTRA_SETUP.md](./ENTRA_SETUP.md), your `EWS_CLIENT_ID` is already appended to your `~/.config/m365-agent-cli/.env` file.

The easiest way to obtain your refresh tokens is to run the interactive login command:

```bash
m365-agent-cli login
```

This will initiate the Microsoft Device Code flow and automatically save **`M365_REFRESH_TOKEN`** (preferred single name) plus legacy `EWS_REFRESH_TOKEN` and `GRAPH_REFRESH_TOKEN` (same value) into your `~/.config/m365-agent-cli/.env` file upon successful authentication.

Alternatively, you can manually create a `~/.config/m365-agent-cli/.env` file (or set environment variables):

```bash
EWS_CLIENT_ID=your-azure-app-client-id
M365_REFRESH_TOKEN=your-refresh-token
EWS_USERNAME=your@email.com
EWS_ENDPOINT=https://outlook.office365.com/EWS/Exchange.asmx
EWS_TENANT_ID=common  # or your tenant ID
```

### Shared and delegated mailboxes (`--mailbox`)

To send from or access another mailbox, set the default in your env:

```bash
EWS_TARGET_MAILBOX=shared@company.com
```

Or pass `--mailbox <email>` per command.

**Microsoft Graph vs EWS:** By default **`M365_EXCHANGE_BACKEND=auto`** tries Microsoft Graph first and falls back to Exchange Web Services when Graph cannot satisfy the request (see [GRAPH_V2_STATUS.md](./GRAPH_V2_STATUS.md)). Set **`M365_EXCHANGE_BACKEND=graph`** to force Graph only, or **`ews`** for EWS only. Exchange delegation / shared access in Outlook does **not** automatically grant the same rights to Graph API calls. When using Graph (including **`auto`**), reading or updating **another user’s** mail or calendar requires **delegated Graph permissions** `Mail.Read.Shared`, `Mail.ReadWrite.Shared`, `Calendars.Read.Shared`, and `Calendars.ReadWrite.Shared` on your Entra app, in addition to `Mail.ReadWrite` / `Calendars.ReadWrite`. For **contacts** in another mailbox, add **`Contacts.Read.Shared`** / **`Contacts.ReadWrite.Shared`** and use **`contacts --user <email>`** (Graph path). Add those in the Azure Portal (see [ENTRA_SETUP.md](./ENTRA_SETUP.md)), then run **`m365-agent-cli login`** again so the refresh token includes the new scopes. If you see **Access is denied** only when using `--mailbox` for another user, missing **\*.Shared** scopes is the usual cause.

### How It Works

1. m365-agent-cli uses the refresh token to obtain a short-lived access token via Microsoft's OAuth2 endpoint
2. Access tokens are cached under `~/.config/m365-agent-cli/`:
   - **Unified OAuth cache:** `token-cache-{identity}.json` (default identity: `default`) holds separate **EWS** and **Microsoft Graph** access tokens and refresh token metadata. Legacy `graph-token-cache-{identity}.json` is merged on load and removed after save.
   Tokens are refreshed automatically when expired.
3. Microsoft may rotate the refresh token on each use — the latest one is cached automatically in the same directory

### Verify Authentication

```bash
# Check who you're logged in as
m365-agent-cli whoami

# Verify your Graph API token scopes (raw scp) or feature coverage matrix
m365-agent-cli verify-token
m365-agent-cli verify-token --capabilities
```
