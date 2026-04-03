# EWS → Microsoft Graph migration tracking

**Purpose:** Single place to see **🟢 migrated**, **🟡 partial**, and **🔴 no Graph path** (or no 1:1 parity) for Exchange-related CLI behavior when `M365_EXCHANGE_BACKEND=graph` (default on `dev_v2`).

**Related:** [`GRAPH_V2_STATUS.md`](./GRAPH_V2_STATUS.md) (branch status log), [`EWS_TO_GRAPH_MIGRATION_EPIC.md`](./EWS_TO_GRAPH_MIGRATION_EPIC.md) (epic).

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
| `whoami` | 🟢 | Graph: `GET /me` when `graph` / `auto` (default). EWS path only when `M365_EXCHANGE_BACKEND=ews`. |
| `mail` | 🟢 | Graph: **full primary path** in `mail-graph.ts` — list (**`--unread`**, **`--flagged`**, **`--search`**), **`--read`**, **`--download`**, **`--move`**, **`--mark-read` / `--mark-unread`**, **flag / unflag / complete**, **`--sensitivity`**, **`--set-categories` / `--clear-categories`**, **reply / reply-all / forward** (incl. `--draft`, `--attach`, `--attach-link`, `--with-category`, `--markdown`). `M365_EXCHANGE_BACKEND=auto` falls back to EWS if Graph fails. |
| `send` | 🟢 | Graph: `sendMail` + file + **`--attach-link`** (`graph-send-mail.ts`); `auto` tries Graph then EWS on failure. |
| `folders` | 🟢 | Graph mail folders when `graph` / `auto`. |
| `drafts` | 🟢 | Graph: **list**, **read**, **`--create`** / **`--edit`** / **`--send`** / **`--delete`** (`createDraftMessage`, `patchMailMessage`, attachments, `sendMailMessage`, `deleteMailMessage` in `drafts-graph.ts`). `auto` falls back to EWS if Graph fails. |
| `calendar` (list range) | 🟢 | Graph `calendarView` when `graph` / `auto`. |
| `calendar` `--list-attachments` / `--download-attachments` | 🟢 | Graph: `listEventAttachments`, `downloadEventAttachmentBytes` (`graph-calendar-client`). EWS fallback if Graph auth fails in `auto`. |
| `findtime` | 🟡 | Graph: **`findMeetingTimes`**, then **`calendar/getSchedule`** merged availability (no EWS). EWS: `getScheduleViaOutlook` only if `M365_EXCHANGE_BACKEND=ews` or **`auto`** after both Graph strategies fail. |
| `create-event` | 🟢 | Graph: **Places** for **`--list-rooms`**, **`--find-room`**, **`--room` by name**; attachments + `POST /me/events`. `auto` may fall back to EWS for rooms. |
| `update-event` | 🟡 | Graph: PATCH + attachments + **attendees** + **`--room` by name** + **`--occurrence` / `--instance`** (`seriesMasterId` / instance id). |
| `delete-event` | 🟡 | Graph cancel/delete for many paths. **`--scope future`** and some series cases still limited vs EWS. |
| `respond` | 🟡 | Graph accept/decline/tentative + list. EWS fallback in `auto` on failure. |
| `forward-event` / `counter` | 🟢 | Graph-only (`graph-event`). |
| `auto-reply` | 🔴 | EWS **Inbox Rules** template model — **no 1:1 Graph clone**. Use **`oof`** (mailboxSettings automatic replies) instead. |
| `oof` | 🟢 | Graph mailboxSettings. |
| `delegates` **list** | 🟢 | Graph `calendar/calendarPermissions` when `graph`. |
| `delegates` **add / update / remove** | 🔴 | EWS **delegate matrix** (folder permissions + deliver options) has **no 1:1 Graph API** — Graph uses **calendar sharing / permissions** with a different model. Use Outlook or a future redesigned CLI. |
| `login` / `auth` | 🟡 | Unified `token-cache-{identity}.json`; dual refresh slots (**EWS** + **Graph** scopes) for mixed-backend and migration. |
| `outlook-graph` | 🟢 | Graph REST mail (parallel surface). |
| `graph-calendar` | 🟢 | Graph calendar helpers (parallel surface). |
| `rules` | 🟢 | Graph inbox rules. |
| `todo` (core) | 🟢 | Graph To Do. `create --link` uses Graph **get message**. |
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

*Last updated: 2026-04-02 — **`update-event` / `delete-event` recurring occurrence selection** on Graph (`seriesMasterId`); **`delete-event --scope future`** still unsupported.*
