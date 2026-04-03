# m365-agent-cli Architecture

## Principles

### 1. Single Authentication Token

m365-agent-cli authenticates using **Microsoft OAuth2 (Azure AD)**. One refresh token (env: **`M365_REFRESH_TOKEN`**, or legacy names) obtains **separate** short-lived access tokens for **EWS** and **Graph** (different resource audiences), stored in one file.

**Implementation:**

- One Azure AD app registration (`EWS_CLIENT_ID`)
- One refresh token in env (prefer **`M365_REFRESH_TOKEN`**)
- One on-disk cache per identity: **`token-cache-{identity}.json`** (v1: `ews` + `graph` access slots + optional refresh metadata)
- Incremental consent: new API scopes are added to the app; user re-runs `login` or consent when needed

Legacy **`graph-token-cache-*.json`** files are merged into the unified cache on load and removed after save.

**API priority:**

1. **Microsoft Graph REST** — preferred for new features (cleaner, modern)
2. **EWS SOAP** — for operations not available in Graph (delegate management, MailTips, full inbox rules, sharing)
3. **PowerShell** — NOT used. If an operation requires PowerShell remoting, it is out of scope.

**EWS → Graph migration:** Phased plan, command inventory, and GitHub tracking notes live in [EWS_TO_GRAPH_MIGRATION_EPIC.md](./EWS_TO_GRAPH_MIGRATION_EPIC.md) (Graph primary, EWS fallback until each slice is done).

### 2. Dynamic Settings

m365-agent-cli must not hardcode user-specific settings. These must always be read from the user's actual Microsoft 365 profile:

| Setting | Source | API |
| --- | --- | --- |
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

- Directory: `~/.config/m365-agent-cli/` — created with `0o700` (owner-only)
- Token file: `token-cache-{identity}.json` (unified EWS + Graph slots) — written with `0o600` (owner-only read/write)
- Cache path uses `homedir()` — never a configurable path that could redirect to arbitrary locations
- Refresh token failures are silently tolerated — m365-agent-cli fails gracefully with an auth error rather than crashing

### 4. Error Handling

- Network errors: retry once, then fail gracefully
- API errors: surface the actual error message from the API (not a generic "failed")
- Auth errors: clear, actionable messages pointing to re-authentication
- Rate limits: respect `Retry-After` headers; back off and report

## Auth Flow

```text
User sets env vars:
  EWS_CLIENT_ID         — Azure AD app client ID
  M365_REFRESH_TOKEN    — preferred single OAuth refresh token (Graph + EWS flows after `login`)
  GRAPH_REFRESH_TOKEN   — legacy alias (same value as `login` output)
  EWS_REFRESH_TOKEN     — legacy alias (same value as `login` output)
  EWS_USERNAME          — user's email address
  EWS_ENDPOINT          — Exchange Online EWS endpoint (default: outlook.office365.com)
  GRAPH_SCOPES          — optional: additional Graph scopes (incremental consent)

Token cache:
  ~/.config/m365-agent-cli/
  — `token-cache-{identity}.json` — unified v1 cache: EWS access slot, Graph access slot, optional refresh token field
  — on expiry: the appropriate slot is refreshed via OAuth using the env refresh token (or cached rotation)
```

### Scope Strategy

**Authoritative delegated Graph list and feature mapping:** [GRAPH_SCOPES.md](./GRAPH_SCOPES.md). **Code:** [`src/lib/graph-oauth-scopes.ts`](../src/lib/graph-oauth-scopes.ts) (`GRAPH_DEVICE_CODE_LOGIN_SCOPES`, `GRAPH_REFRESH_SCOPE_CANDIDATES`).

**EWS (Exchange Online SOAP):** `https://outlook.office365.com/EWS.AccessAsUser.All` — see `auth.ts`.

**Summary:** `login` requests mail/calendar **\*.Shared** scopes for delegated **`--mailbox`**, **`Place.Read.All`** (Places / `rooms`), **`People.Read`** and **`User.Read.All`** (`find`), plus Files, Sites, Tasks, Groups, MailboxSettings. **User.Read.All** and **Place.Read.All** often require **admin consent**. Graph token refresh tries a **full** scope string and a **fallback without `User.Read.All`** when tenants restrict directory read consent.

## Directory Structure

```text
src/
  cli.ts              — entry point, argument parsing
  lib/
    auth.ts           — EWS OAuth2 (token cache, refresh, validation)
    graph-auth.ts     — Graph OAuth2 (refresh; scope strings from graph-oauth-scopes.ts)
    graph-oauth-scopes.ts — canonical Graph delegated scope strings for login + refresh
    ews-client.ts     — all EWS SOAP operations
    graph-client.ts   — all Microsoft Graph REST calls
    jwt-utils.ts       — JWT parsing (expiry, structure validation)
    xml-utils.ts      — XML escape, SOAP envelope builders
    date-utils.ts     — date parsing, formatting (locale-aware)
    dates.ts          — shared date parsing (`parseDay`, weekday-relative resolution)
    calendar-range.ts — calendar window helpers (`--days`, `--business-days`, `clipCalendarRangeStartToNow` for `--now`, etc. on `calendar`)
    outlook-master-categories.ts — Graph `GET .../outlook/masterCategories`
    planner-client.ts — Planner tasks, plans, buckets, plan details (label names)
    graph-advanced-client.ts — raw Graph paths + JSON `$batch` (`graph invoke`, `graph batch`)
    graph-bookings-client.ts — Microsoft Bookings (businesses, appointments, services, staff, calendarView, …)
    graph-calendar-client.ts — Graph `GET .../calendars`, `calendarView`, `GET .../events/{id}`
    graph-calendar-recurrence.ts — Graph recurring series truncation (`delete-event --scope future`)
    graph-excel-client.ts — Excel worksheets, range, tables, used range on drive items
    graph-presence-client.ts — `GET /me/presence`, `GET /users/{id}/presence`
    graph-teams-client.ts — Teams channels (incl. all/incoming), messages, chats, tabs, members, apps
    onenote-graph-client.ts — OneNote notebooks, sections, pages, copy operations
    todo-client.ts    — Microsoft To Do lists/tasks (including `categories`)
    url-utils.ts      — URL sanitization, safe filename handling
  commands/
    whoami.ts
    calendar.ts
    create-event.ts
    update-event.ts
    delete-event.ts
    respond.ts
    graph-calendar.ts
    findtime.ts
    find.ts
    mail.ts
    folders.ts
    send.ts
    drafts.ts
    outlook-categories.ts
    planner.ts
    todo.ts
    files.ts
    graph.ts
    graph-search.ts
    graph-calendar.ts
    teams.ts
    bookings.ts
    excel.ts
    presence.ts
```

## API Coverage

### EWS SOAP (via `ews-client.ts`)

Used for operations with no Graph equivalent:

| Operation | EWS SOAP | Notes |
| --- | --- | --- |
| Delegate management | `AddDelegate`, `GetDelegate`, `UpdateDelegate`, `RemoveDelegate` | Same token |
| Inbox rules (full) | `GetInboxRules`, `UpdateInboxRules` | Full condition/action set |
| MailTips | `GetMailTips` | No Graph equivalent |
| Sharing | `GetSharingMetadata`, `AcceptSharingInvitation` | |
| Conversation actions | `ApplyConversationAction` | |
| Room lists | `GetRoomLists`, `GetRooms` | Graph Places API is preferred |
| Free/busy | `GetUserAvailability` | Graph getSchedule is preferred |
| People search | `ResolveNames` | Graph People API is preferred |

**Write operations:** `ews-client.ts` resolves targets with **`GetItem`** (messages) or **`getCalendarEvent`** (calendar) before mutating SOAP calls and includes **ChangeKey** on `ItemId`, `ReferenceItemId`, `ParentItemId`, and related shapes where Exchange requires it (notably delegated/shared mailbox scenarios). Callers continue to pass only **item IDs** from list/read commands. `updateEvent` may prefetch **ChangeKey** when the caller does not supply one; a failed prefetch returns that error instead of sending an invalid **UpdateItem**.

**Categories:** Mail and calendar items expose **`item:Categories`** (string list). Colors for those names are defined by the mailbox **master category list** (Graph: `outlook/masterCategories`, CLI: `outlook-categories list`), not as per-item color fields in EWS.

### Microsoft Graph REST (via `graph-client.ts`)

Preferred for new features:

| Resource | Endpoints | Notes |
| --- | --- | --- |
| Calendar | `GET .../calendars`, `GET .../calendar/calendarView`, `GET/POST/PATCH/DELETE .../events`, invitation responses `POST .../events/{id}/accept`, `.../decline`, `.../tentativelyAccept` | CLI **`graph-calendar`** for list/view/get + invitation responses; EWS **`calendar`** / **`create-event`** still primary for many flows |
| Free/busy | `POST /me/calendar/getSchedule` | Preferred over EWS |
| Room discovery | `GET /places`, `/roomLists`, `/rooms` | Richer than EWS |
| Mail | `GET/POST/PATCH/DELETE /me/mailFolders`, folder + `/me/messages`, `POST /sendMail`, `POST .../move`, `copy`, `send`, `POST .../createReply`, `createReplyAll`, `createForward`, attachments | CLI **`outlook-graph`** (Graph REST); EWS **`mail`** and **`folders`** remain primary for many flows |
| Message rules | `GET/POST/PATCH/DELETE /me/mailFolders/inbox/rules` | Partial — no template replies |
| Contacts | `GET/POST/PATCH/DELETE /me/contacts`, contactFolders, photo, attachments, delta | Primary CLI **`contacts`**; parallel **`outlook-graph`** contact helpers |
| OneNote | `GET/POST/PATCH/DELETE …/onenote/...` | CLI **`onenote`**; no EWS equivalent |
| Online meetings | `GET/POST/PATCH/DELETE /me/onlineMeetings` | CLI **`meeting`** |
| People | `GET /me/people` | Relevance-ranked, not true GAL |
| Directory | `GET /users`, `/groups/{id}/members` | Requires `Directory.Read.All` |
| To-Do | `GET/POST/PATCH/DELETE /me/todo/lists/{id}/tasks`, checklistItems CRUD, `GET .../attachments/{id}/$value` for file bytes | Use To-Do API, NOT Outlook Tasks (deprecated) |
| OOF | `PATCH /me/mailboxSettings` | `automaticRepliesSetting` |
| Mailbox settings | `GET/PATCH /me/mailboxSettings` | timezone, working hours, language |
| Subscriptions | `POST /subscriptions` | Webhook push notifications |
| Outlook categories | `GET/POST/PATCH/DELETE .../outlook/masterCategories` | Master list CRUD (names + `preset0`..`preset24`); CLI `outlook-categories` subcommands: list, create, update, delete |
| Planner | `GET/PATCH /planner/tasks`, `GET /users/{id}/planner/tasks` (when permitted), `GET /planner/plans/{id}/details`, beta `planner/rosters`, `POST /planner/plans` with roster container | Task `appliedCategories` (six slots); plan `categoryDescriptions` for labels; roster-backed plans use beta Graph |

## Out of Scope

The following are explicitly NOT part of m365-agent-cli's roadmap:

- **Exchange PowerShell remoting** — requires WinRM/RDP or separate credential management
- **SendAs / SendOnBehalf permission granting** — requires Exchange Admin role; m365-agent-cli can USE an already-granted SendAs permission but cannot grant it
- **eDiscovery / compliance** — admin-only APIs
- **SharePoint / OneDrive business** — separate auth domain
- **Azure AD B2C guest accounts** — different auth surface

## Optional error reporting (GlitchTip)

When **`GLITCHTIP_DSN`** or **`SENTRY_DSN`** is set, the CLI may initialize **`@sentry/node`** (Sentry-compatible ingest) to report **uncaught exceptions**, **unhandled rejections**, and **Commander parse errors** — only when the install matches the **latest npm release** and the **embedded commit** matches the GitHub **`v{version}`** tag (unless overridden by env); see [GLITCHTIP.md](./GLITCHTIP.md). Events are **scrubbed** (no argv text, user, or breadcrumbs; redacted paths and messages) so reports do not include mail content or usernames. Default `beforeSend` also filters common **network errno** values and **OAuth token** failure messages. No reporting when DSN is unset.

## Security Considerations

- Tokens are live secrets — never log or print token contents
- `EWS_ENDPOINT` and `GRAPH_BASE_URL` are validated at startup — custom URLs only allowed in non-production environments
- File attachments: paths are validated against allowed directories, no symlink traversal
- URLs in email content: only safe schemes (`http`, `https`, `mailto`) are allowed; `javascript:`, `data:`, `file:` are stripped
- SQL injection: not applicable (SQLite is local-only)
- XML injection: all user strings in SOAP requests go through `xmlEscape()`
