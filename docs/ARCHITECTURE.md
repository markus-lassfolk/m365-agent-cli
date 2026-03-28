# Clippy Architecture

## Principles

### 1. Single Authentication Token

Clippy authenticates once using **Microsoft OAuth2 (Azure AD)**. A single refresh token is cached and used to obtain access tokens for all APIs.

**Current state:** EWS and Graph use separate token caches. This is being consolidated.

**Target state:**
- One Azure AD app registration
- One refresh token
- One token cache file (`~/.config/clippy/token-cache.json`)
- Incremental consent: new API scopes are added to the existing app without requiring re-authentication

**API priority:**
1. **Microsoft Graph REST** — preferred for new features (cleaner, modern)
2. **EWS SOAP** — for operations not available in Graph (delegate management, MailTips, full inbox rules, sharing)
3. **PowerShell** — NOT used. If an operation requires PowerShell remoting, it is out of scope.

### 2. Dynamic Settings

Clippy must not hardcode user-specific settings. These must always be read from the user's actual Microsoft 365 profile:

| Setting | Source | API |
|---------|--------|-----|
| Timezone | `mailboxSettings.timeZone` | Graph: `GET /me/mailboxSettings` |
| Locale / language | `mailboxSettings.automaticRepliesSetting locale` | Graph: `GET /me/mailboxSettings` |
| Date/time format | Derived from locale | Graph: `GET /me/mailboxSettings` |
| Working hours | `mailboxSettings.workingHours` | Graph: `GET /me/mailboxSettings` |
| Display name | `userPrincipalName`, `displayName` | Graph: `GET /me` |
| Email timezone offset | `GetUserAvailability` response | EWS: `GetUserAvailability` |

**Rules:**
- Never hardcode `CET` or any fixed UTC offset as a default
- Never assume `en-US` or any fixed locale
- Always read from the authenticated user's settings when presenting dates, times, or timezones
- Fallback: `Intl.DateTimeFormat().resolvedOptions().timeZone` (system timezone) if API call fails

### 3. Token Cache Security

The token cache file is the most sensitive file on disk.

- Directory: `~/.config/clippy/` — created with `0o700` (owner-only)
- Token file: `token-cache.json` — written with `0o600` (owner-only read/write)
- Cache path uses `homedir()` — never a configurable path that could redirect to arbitrary locations
- Refresh token failures are silently tolerated — Clippy fails gracefully with an auth error rather than crashing

### 4. Error Handling

- Network errors: retry once, then fail gracefully
- API errors: surface the actual error message from the API (not a generic "failed")
- Auth errors: clear, actionable messages pointing to re-authentication
- Rate limits: respect `Retry-After` headers; back off and report

## Auth Flow

```
User sets env vars:
  EWS_CLIENT_ID         — Azure AD app client ID
  EWS_REFRESH_TOKEN     — OAuth refresh token (EWS + Graph combined)
  EWS_USERNAME          — user's email address
  EWS_ENDPOINT          — Exchange Online EWS endpoint (default: outlook.office365.com)
  GRAPH_SCOPES          — optional: additional Graph scopes (incremental consent)

Token cache:
  ~/.config/clippy/token-cache.json
  — holds EWS access token + refresh token + expiry
  — both EWS and Graph reuse this same cached token
  — on expiry: refresh token is used to obtain a new access token
```

### Scope Strategy

Scopes are requested incrementally. The base token covers:

**EWS (required):**
- `https://outlook.office365.com/EWS.AccessAsUser.All`

**Graph base (required):**
- `User.Read`
- `Files.ReadWrite`
- `OfflineAccess`

**Graph additions (added as features require):**
| Feature | Additional Scope |
|---------|----------------|
| Calendar read/write | `Calendars.ReadWrite` |
| Mail read/write | `Mail.ReadWrite` |
| Room discovery | `Place.Read.All` |
| People/GAL lookup | `People.Read` |
| To-Do tasks | `Tasks.ReadWrite` |
| OOF / mailbox settings | `MailboxSettings.ReadWrite` |
| Delegate management | (EWS SOAP — same token) |

## Directory Structure

```
src/
  cli.ts              — entry point, argument parsing
  lib/
    auth.ts           — EWS OAuth2 (token cache, refresh, validation)
    graph-auth.ts     — Graph OAuth2 (reuses EWS token via incremental consent)
    ews-client.ts     — all EWS SOAP operations
    graph-client.ts   — all Microsoft Graph REST calls
    jwt-utils.ts       — JWT parsing (expiry, structure validation)
    xml-utils.ts      — XML escape, SOAP envelope builders
    date-utils.ts     — date parsing, formatting (locale-aware)
    url-utils.ts      — URL sanitization, safe filename handling
  commands/
    whoami.ts
    calendar.ts
    create-event.ts
    update-event.ts
    delete-event.ts
    respond.ts
    findtime.ts
    find.ts
    mail.ts
    folders.ts
    send.ts
    drafts.ts
    files.ts
```

## API Coverage

### EWS SOAP (via `ews-client.ts`)

Used for operations with no Graph equivalent:

| Operation | EWS SOAP | Notes |
|-----------|----------|-------|
| Delegate management | `AddDelegate`, `GetDelegate`, `UpdateDelegate`, `RemoveDelegate` | Same token |
| Inbox rules (full) | `GetInboxRules`, `UpdateInboxRules` | Full condition/action set |
| MailTips | `GetMailTips` | No Graph equivalent |
| Sharing | `GetSharingMetadata`, `AcceptSharingInvitation` | |
| Conversation actions | `ApplyConversationAction` | |
| Room lists | `GetRoomLists`, `GetRooms` | Graph Places API is preferred |
| Free/busy | `GetUserAvailability` | Graph getSchedule is preferred |
| People search | `ResolveNames` | Graph People API is preferred |

### Microsoft Graph REST (via `graph-client.ts`)

Preferred for new features:

| Resource | Endpoints | Notes |
|----------|-----------|-------|
| Calendar | `GET/POST/PATCH/DELETE /me/events` | Full CRUD + recurrence |
| Free/busy | `POST /me/calendar/getSchedule` | Preferred over EWS |
| Room discovery | `GET /places`, `/roomLists`, `/rooms` | Richer than EWS |
| Mail | `GET/POST/PATCH/DELETE /me/mailFolders` | Full CRUD |
| Message rules | `GET/POST/PATCH/DELETE /me/mailFolders/inbox/rules` | Partial — no template replies |
| Contacts | `GET/POST/PATCH/DELETE /me/contacts` | Personal contacts |
| People | `GET /me/people` | Relevance-ranked, not true GAL |
| Directory | `GET /users`, `/groups/{id}/members` | Requires `Directory.Read.All` |
| To-Do | `GET/POST/PATCH/DELETE /me/todo/lists/{id}/tasks` | Use To-Do API, NOT Outlook Tasks (deprecated) |
| OOF | `PATCH /me/mailboxSettings` | `automaticRepliesSetting` |
| Mailbox settings | `GET/PATCH /me/mailboxSettings` | timezone, working hours, language |
| Subscriptions | `POST /subscriptions` | Webhook push notifications |

## Out of Scope

The following are explicitly NOT part of Clippy's roadmap:

- **Exchange PowerShell remoting** — requires WinRM/RDP or separate credential management
- **SendAs / SendOnBehalf permission granting** — requires Exchange Admin role; Clippy can USE an already-granted SendAs permission but cannot grant it
- **eDiscovery / compliance** — admin-only APIs
- **SharePoint / OneDrive business** — separate auth domain
- **Azure AD B2C guest accounts** — different auth surface

## Security Considerations

- Tokens are live secrets — never log or print token contents
- `EWS_ENDPOINT` and `GRAPH_BASE_URL` are validated at startup — custom URLs only allowed in non-production environments
- File attachments: paths are validated against allowed directories, no symlink traversal
- URLs in email content: only safe schemes (`http`, `https`, `mailto`) are allowed; `javascript:`, `data:`, `file:` are stripped
- SQL injection: not applicable (SQLite is local-only)
- XML injection: all user strings in SOAP requests go through `xmlEscape()`
