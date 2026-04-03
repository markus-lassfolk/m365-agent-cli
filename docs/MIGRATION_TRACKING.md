# EWS → Microsoft Graph migration tracking

**Purpose:** Single place to see **🟢 migrated**, **🟡 partial**, and **🔴 no Graph path** (or no 1:1 parity) for Exchange-related CLI behavior when `M365_EXCHANGE_BACKEND=graph` (default in [`exchange-backend.ts`](../src/lib/exchange-backend.ts); was introduced on `dev_v2`).

**Related:** [`GRAPH_V2_STATUS.md`](./GRAPH_V2_STATUS.md) (branch status log), [`EWS_TO_GRAPH_MIGRATION_EPIC.md`](./EWS_TO_GRAPH_MIGRATION_EPIC.md) (epic).

### Graph-first policy and `M365_EXCHANGE_BACKEND=auto`

| Value | Behavior |
| --- | --- |
| **`graph`** (default) | **Graph only.** No EWS fallback. Fails fast if Graph auth or the API call fails. |
| **`auto`** | **Graph first** for every command that implements a Graph path. **EWS only when** Graph authentication fails, the Graph request fails, or the operation has **no Microsoft Graph equivalent** in this CLI (see 🔴 rows and per-command notes). A **successful** Graph result — including an empty list — is **authoritative**; the CLI does **not** replace it with EWS “for more data” (different APIs are not interchangeable). |
| **`ews`** | **EWS only** (legacy / debugging). |

Helpers in [`exchange-backend.ts`](../src/lib/exchange-backend.ts): `shouldTryGraphFirst()`, `isAutoMode()`, `isGraphOnlyMode()`, `isEwsExclusiveMode()`, `mayUseEws()`.

### Delegated / shared mailbox (`--mailbox`) on Graph

EWS delegate access does **not** imply the same Microsoft Graph token scopes. Calling `GET /users/{other}/...` requires **delegated** Graph permissions including **`Mail.Read.Shared`**, **`Mail.ReadWrite.Shared`**, **`Calendars.Read.Shared`**, and **`Calendars.ReadWrite.Shared`** on the Entra app, then **`m365-agent-cli login`** again. Without them, Graph may return **Access is denied** for another user’s mailbox while the primary mailbox still works. Full scope list and CLI features: **[`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md)**; portal setup: [`ENTRA_SETUP.md`](./ENTRA_SETUP.md).

### Legend

| Marker | Meaning |
| --- | --- |
| 🟢 **GREEN** | Graph implementation exists on the **primary** command for typical use; EWS not required for that slice when using default `graph`. |
| 🟡 **YELLOW** | **Partial** — Graph covers some flags/subcommands; others still need EWS, `outlook-graph` / `graph-calendar`, or more code. |
| 🔴 **RED** | **No suitable Graph equivalent** for the *current* EWS UX (1:1), or Microsoft exposes a different product model — needs redesign, different command, or external tool. |

---

## Command matrix (`src/commands`)

| Command / area | Marker | Notes |
| --- | --- | --- |
| `whoami` | 🟢 | Graph: `GET /me` when `graph` or when `auto` succeeds on Graph. **`ews`:** EWS user info only. **`auto`:** tries Graph first, then EWS if Graph auth or `/me` fails. |
| `mail` | 🟢 | Graph: **primary path** in `mail-graph.ts` for supported flag combinations — list (**`--unread`**, **`--flagged`**, **`--search`**), **`--read`**, **`--download`**, **`--move`**, **`--mark-read` / `--mark-unread`**, **flag / unflag / complete**, **`--sensitivity`**, **`--set-categories` / `--clear-categories`**, **reply / reply-all / forward** (incl. `--draft`, `--attach`, `--attach-link`, `--with-category`, `--markdown`). **`graph`:** unsupported combinations error (use EWS or `outlook-graph`). **`auto`:** falls back to EWS when Graph does not handle the request or fails. |
| `send` | 🟢 | Graph: `sendMail` + file + **`--attach-link`** (`graph-send-mail.ts`); `auto` tries Graph then EWS on failure. Entra must include delegated **`Mail.Send`** (plus `Mail.ReadWrite`); tokens from before that addition need **`login`** again. |
| `folders` | 🟢 | Graph mail folders when `graph` / `auto`. |
| `drafts` | 🟢 | Graph: **list**, **read**, **`--create`** / **`--edit`** / **`--send`** / **`--delete`** (`createDraftMessage`, `patchMailMessage`, attachments, `sendMailMessage`, `deleteMailMessage` in `drafts-graph.ts`). `auto` falls back to EWS if Graph fails. |
| `calendar` (list range) | 🟢 | Graph `calendarView` when `graph` / `auto`. Rolling ranges (`--days`, `--business-days` / `--next-business-days`, …) and **`--now`** (clip window start to “now”) apply before the view; EWS uses the same resolved window. |
| `calendar` `--list-attachments` / `--download-attachments` | 🟢 | Graph: `listEventAttachments`, `downloadEventAttachmentBytes` (`graph-calendar-client`). EWS fallback if Graph auth fails in `auto`. |
| `findtime` | 🟢 | **Graph is the primary path** when `graph` / `auto`: **`findMeetingTimes`**, then **`getSchedule`** + merged `availabilityView` (`findtime-graph.ts`). **EWS** (`getScheduleViaOutlook`) only when `M365_EXCHANGE_BACKEND=ews`, or in **`auto`** after both Graph strategies fail. |
| `create-event` | 🟢 | Graph: **Places** for **`--list-rooms`**, **`--find-room`**, **`--room` by name**; attachments + `POST /me/events`. `auto` may fall back to EWS for rooms. |
| `update-event` | 🟡 | **Graph-first** for typical updates (PATCH + attachments + **attendees** + **`--room` by name** + **`--occurrence` / `--instance`**). **🟡** = mixed Graph/EWS ID story: with **`graph`**, there is **no EWS fallback** after Graph-backed data is used (see command errors). |
| `delete-event` | 🟡 | **Graph-first** cancel/delete + occurrence/instance matching via **`seriesMasterId`**. **Gap:** **`--scope future`** has **no Graph implementation** yet (trim series via PATCH on the **series master**); **EWS** implements it via SOAP `deleteEvent`. With **`auto`**, use EWS only when the listing path is EWS (e.g. Graph list failed). |
| `respond` | 🟢 | **Graph is the primary path** when `graph` / **`auto`**: list via **`calendarView`** + pending filter; **`accept` / `decline` / `tentative`** via Graph invitation APIs. **EWS** only when **`auto`** and Graph auth or **`getEvent`** fails (then list/respond use EWS). |
| `forward-event` / `counter` | 🟢 | Graph-only (`graph-event`). |
| `auto-reply` | 🔴 | EWS **Inbox Rules**–based templates (this command’s SOAP model). **Graph** offers **`oof`** (automatic replies) and **`rules`** (inbox rules), but **not** this CLI’s template UX — **no 1:1 replacement**. Prefer **`oof`** for OOF-style mail; use **`rules`** for Graph mail rules. |
| `oof` | 🟢 | Graph mailboxSettings. |
| `delegates` **list** | 🟢 | Graph **`calendarPermissions`** when `graph`. **`auto`:** Graph first; an **empty** Graph result is final (same message as `graph`). EWS **`GetDelegates`** only if the Graph **request fails** (auth/call error), not to “supplement” Graph. **`ews`:** EWS only. |
| `delegates` **add / update / remove** | 🔴 | EWS **delegate matrix** (folder permissions + deliver options) has **no 1:1 Graph API** — Graph uses **calendar sharing / permissions** with a different model. Use Outlook or a future redesigned CLI. |
| `login` / `auth` | 🟡 | Unified `token-cache-{identity}.json`; dual refresh slots (**EWS** + **Graph** scopes) for mixed-backend and migration. |
| `outlook-graph` | 🟢 | Graph REST mail (parallel surface). |
| `graph-calendar` | 🟢 | Graph calendar helpers (parallel surface). |
| `rules` | 🟢 | Graph inbox rules. |
| `todo` (core) | 🟢 | Graph To Do. `create --link` uses Graph **get message**. |
| `contacts` | 🟢 | Graph **`/me/contacts`** / **`/me/contactFolders`** (`Contacts.ReadWrite`). |
| `meeting` | 🟢 | Graph **`/me/onlineMeetings`** (`OnlineMeetings.ReadWrite`). Distinct from **`create-event --teams`** (calendar + Teams on the event). |
| `onenote` | 🟢 | Graph OneNote (`Notes.ReadWrite.All`). |
| `planner`, `files`, `sharepoint`, `find`, `rooms`, `subscribe`, … | 🟢 | Graph (no EWS in path). |

---

## Library / SOAP surface

| Module | Marker | Notes |
| --- | --- | --- |
| `ews-client.ts` | 🟡 | Large SOAP surface; shrinking as commands migrate. |
| `delegate-client.ts` | 🔴 | EWS-only; see delegates row. |

---

## How to use this doc

1. **Prioritize 🟡** — largest user-facing gaps; implement Graph in the **primary** command or document `outlook-graph` / `graph-calendar` workarounds.
2. **🔴** — decide product direction (drop feature, new Graph-native UX, or document “use Outlook”).
3. After each migration, update this file and [`GRAPH_V2_STATUS.md`](./GRAPH_V2_STATUS.md).

*Last updated: 2026-04-03 — **`Contacts.ReadWrite`**, **`OnlineMeetings.ReadWrite`**, **`Notes.ReadWrite.All`** in `graph-oauth-scopes.ts` + Entra scripts; **`contacts`**, **`meeting`**, **`onenote`** commands.*

---

## Graph response fixtures (tests)

See [`src/test/fixtures/graph/README.md`](../src/test/fixtures/graph/README.md). Canonical schemas remain Microsoft’s OpenAPI/docs; our JSON examples anchor unit tests for **error envelopes** and **sendMail** bodies.
