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

**Tenant ID precedence** (for the OAuth endpoint path, `--tenant`, login device-code flow): `M365_TENANT_ID` > `MICROSOFT_TENANT_ID` > `EWS_TENANT_ID` (legacy) > `common`. The legacy `EWS_TENANT_ID` name remains supported for backwards compatibility.

### Driving `login` from a script or agent

For a full walkthrough of automating the interactive step end-to-end (headless browser + TOTP) so an agent can re-authenticate on its own, see **[UNATTENDED_LOGIN.md](./UNATTENDED_LOGIN.md)** and the adaptable reference scripts in [`examples/unattended-login/`](../examples/unattended-login/).

- **Machine-readable output:** `m365-agent-cli login --json` emits newline-delimited JSON events (`device_code`, `authenticated`, `complete`, `error`) to stdout so a wrapper can capture the `user_code` / `verification_uri` without scraping log text; human-readable messages go to stderr. Requires `EWS_CLIENT_ID` to be preset (it does not prompt in `--json` mode).
- **`login` is a synchronous, foreground poll.** It calls Microsoft's token endpoint on an interval until sign-in completes in a browser or the device code expires (`expires_in`, ~15 min). A wrapper that backgrounds `login` **must keep the process alive** until sign-in finishes — killing it early discards the pending device-code session even if the browser page already shows "signed in".
- **Verify from the CLI, not the wrapper.** After any login, confirm with `m365-agent-cli whoami` and `m365-agent-cli verify-token --capabilities`; don't rely on a wrapper's own success message.
- **No manual lock cleanup.** Refresh-token exchange is serialized per identity via `.refresh-{identity}.lock` in the config dir, which **auto-heals** stale locks (holder PID gone, or older than 120 s). You never need to delete a lock file by hand between runs.

### Shared and delegated mailboxes (`--mailbox`)

To send from or access another mailbox, set the default in your env:

```bash
EWS_TARGET_MAILBOX=shared@company.com
```

Or pass `--mailbox <email>` per command.

**Microsoft Graph vs EWS:** By default **`M365_EXCHANGE_BACKEND=auto`** tries Microsoft Graph first and falls back to Exchange Web Services when Graph cannot satisfy the request (see [GRAPH_V2_STATUS.md](./GRAPH_V2_STATUS.md)). Set **`M365_EXCHANGE_BACKEND=graph`** to force Graph only, or **`ews`** for EWS only. Exchange delegation / shared access in Outlook does **not** automatically grant the same rights to Graph API calls. When using Graph (including **`auto`**), reading or updating **another user’s** mail or calendar requires **delegated Graph permissions** `Mail.Read.Shared`, `Mail.ReadWrite.Shared`, `Calendars.Read.Shared`, and `Calendars.ReadWrite.Shared` on your Entra app, in addition to `Mail.ReadWrite` / `Calendars.ReadWrite`. For **Microsoft To Do** shared/delegated scenarios, add **`Tasks.Read.Shared`** / **`Tasks.ReadWrite.Shared`**; note that Graph To Do `/users/{other}/todo/...` does not behave like mailbox/calendar delegation for every target user and can still return `Invalid request` after scopes are present if the target To Do service/sharing state is not available through that endpoint. For **contacts** in another mailbox, add **`Contacts.Read.Shared`** / **`Contacts.ReadWrite.Shared`** and use **`contacts --user <email>`** (Graph path). Add those in the Azure Portal (see [ENTRA_SETUP.md](./ENTRA_SETUP.md)), then run **`m365-agent-cli login`** again so the refresh token includes the new scopes. If you see **Access is denied** only when using `--mailbox` for another user, missing **\*.Shared** scopes is the usual cause.

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
