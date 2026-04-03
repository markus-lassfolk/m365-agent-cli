# Microsoft Graph OAuth scopes (m365-agent-cli)

This document lists **delegated** permissions the CLI is designed to use. **Source of truth in code:** [`src/lib/graph-oauth-scopes.ts`](../src/lib/graph-oauth-scopes.ts) (`GRAPH_DEVICE_CODE_LOGIN_SCOPES` for `login`, `GRAPH_REFRESH_SCOPE_CANDIDATES` for token refresh in [`graph-auth`](../src/lib/graph-auth.ts)).

Configure the same permissions on your **Entra ID app registration** (API permissions → Microsoft Graph → Delegated). Then run **`m365-agent-cli login`** so the refresh token includes them. Use **`m365-agent-cli verify-token`** to inspect granted `scp` claims.

**Office 365 Exchange Online:** add **`EWS.AccessAsUser.All`** (delegated) for EWS-backed commands when `M365_EXCHANGE_BACKEND` is `ews` or `auto` (see [`ENTRA_SETUP.md`](./ENTRA_SETUP.md)).

---

## Full scope set (recommended)

| Scope | Purpose in this CLI |
| --- | --- |
| `offline_access` | Refresh tokens |
| `User.Read` | Sign-in profile; `/me` |
| `Calendars.ReadWrite` | Own calendar read/write |
| `Calendars.Read.Shared` | Delegated / shared calendars (`/users/{upn}/calendar/...`) |
| `Calendars.ReadWrite.Shared` | Same, with write |
| `Mail.Send` | **`POST /me/sendMail`** and sending mail via Graph (explicit; use with `Mail.ReadWrite` for full mail UX) |
| `Mail.ReadWrite` | Own mailbox mail APIs |
| `Mail.Read.Shared` | Mail in mailboxes the user can access (delegated/shared) |
| `Mail.ReadWrite.Shared` | Same, with write (where applicable) |
| `MailboxSettings.ReadWrite` | Mailbox settings, OOF, categories, rules-related settings |
| `Place.Read.All` | Places API — `rooms`, room resolution in `create-event` / Places |
| `People.Read` | `GET /me/people` — `find` (people/relevant contacts) |
| `User.Read.All` | `GET /users` directory search — `find` (user query); **often requires admin consent** |
| `Files.ReadWrite.All` | OneDrive / **`files`**; **`excel`** workbook — worksheets **get/add/update/delete**, **range** read + **range-patch**, **used-range**, **tables** / **table-get** / **table-rows** / **table-rows-add**, **names**, **charts** |
| `Sites.ReadWrite.All` | SharePoint / site pages |
| `Tasks.ReadWrite` | Microsoft To Do |
| `Group.ReadWrite.All` | Planner (groups), group-related Graph calls; **`teams members`**, **`teams channel-members`**, **`teams apps`**, **`teams tabs`** (narrower: `TeamMember.Read.*`, `ChannelMember.Read.All`, `TeamsAppInstallation.ReadForTeam`, `TeamsTab.Read.All`) |
| `Contacts.ReadWrite` | **`contacts`** — `/me/contacts`, `/me/contactFolders`, photo, file + **reference (link)** attachments, delta, `$search`, `$filter` on list |
| `Contacts.Read.Shared` | Read contacts in **shared / delegated** mailboxes (`--user` on `contacts`) |
| `Contacts.ReadWrite.Shared` | Create/update/delete contacts for mailboxes you have delegate access to |
| `OnlineMeetings.ReadWrite` | **`meeting`** — `POST/PATCH/DELETE/GET /me/onlineMeetings` (standalone Teams meeting; **`meeting create --json-file`** for full Graph body). **`Calendars.ReadWrite`** + **`create-event … --teams`** — calendar invitations with Teams; parse **`--json`** → `event.teamsMeeting` / `event.onlineMeeting` for assistants. |
| `Notes.ReadWrite.All` | **`onenote`** — notebooks / section groups / sections (CRUD), pages (list, get, HTML, export, create, **delete**, **patch-page-content**, **copy-page**), async **operation** poll for copy |
| `Team.ReadBasic.All` | **`teams`** — joined teams, team metadata (`GET /me/joinedTeams`, `GET /teams/{id}`) |
| `Channel.ReadBasic.All` | **`teams channels`**, **`teams all-channels`**, **`teams incoming-channels`**, **`teams primary-channel`**, **`teams channel-get`** — list/get channel (`/channels`, `/allChannels`, `/incomingChannels`, `primaryChannel`, `channels/{id}`) |
| `ChannelMessage.Read.All` | **`teams messages`**, **`teams channel-message-get`**, **`teams message-replies`** — channel messages and thread replies; **delegated admin consent** often required |
| `ChannelMessage.Send` | **`teams channel-message-send`**, **`teams channel-message-reply`** — `POST …/channels/{id}/messages` and `…/messages/{id}/replies` |
| `Presence.Read.All` | **`presence me`**, **`presence user`**, **`presence bulk`** (`POST /communications/getPresencesByUserId`) |
| `Presence.ReadWrite` | **`presence set-me`**, **`presence set-user`**, **`presence clear-me`**, **`presence clear-user`** — set/clear presence session |
| `Bookings.ReadWrite.All` | **`bookings`** — delegated read/write as listed; **`staff-availability`** is **not** delegated per Microsoft (use **app-only** token) |
| `Chat.ReadWrite` | **`teams chats`**, **`teams chat-get`**, **`teams chat-pinned`**, **`teams chat-messages`**, **`teams chat-message-get`**, **`teams chat-message-replies`**, **`teams chat-members`**, **`teams chat-message-send`**, **`teams chat-message-reply`** |
| *(entity-specific)* | **`graph-search`** — Microsoft Graph Search (`POST /search/query`) uses the least-privilege permission for each entity type (e.g. mail → Mail.Read, files → Files.Read.All); see Graph Search API docs |
| — | **`graph invoke`** / **`graph batch`** — arbitrary JSON Graph paths and `$batch` (see command help); use for APIs not wrapped as dedicated subcommands |

**Note:** `Group.ReadWrite.All` implies broad group read/write. For **`find`** group listing, this is sufficient; a narrower `Group.Read.All` is **not** requested separately to avoid redundant consent alongside `Group.ReadWrite.All`.

---

## Admin consent

These commonly require **admin consent** in tenant consent policies (especially **`User.Read.All`**, **`Place.Read.All`**). If refresh fails after login, check the Entra **Enterprise applications** → your app → **Permissions** and use **Grant admin consent**, or ask an admin to approve.

---

## Refresh fallback behavior

[`graph-auth`](../src/lib/graph-auth.ts) tries several scope strings when refreshing. It includes a candidate **without** `User.Read.All` so users who cannot obtain admin consent for directory read may still refresh tokens for mail/calendar/files-heavy operations.

---

## Related docs

- [`ENTRA_SETUP.md`](./ENTRA_SETUP.md) — portal steps and automated scripts  
- [README](../README.md) — authentication overview  
- [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md) — Graph vs EWS, `--mailbox` behavior  

*Last updated: 2026-04-03 — **`channel-message-get`**, **`chat-message-get`**, **`chat-message-replies`**, **`chat-message-reply`**; aligned with `graph-oauth-scopes.ts`.*
