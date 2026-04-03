# Graph-first (`dev_v2`) — status

**Branch:** `dev_v2`  
**Epic:** [#204 — EWS → Microsoft Graph migration](https://github.com/markus-lassfolk/m365-agent-cli/issues/204)  
**Goal:** Move toward **Microsoft Graph as the default** for Exchange-related flows, with **`M365_EXCHANGE_BACKEND`** to opt into EWS or `auto` during migration.

This file is the working log for `dev_v2`. Update it when slices land or decisions change.

**🟢 / 🟡 / 🔴 matrix (command-by-command):** [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md).

---

## Configuration

| Env | Values | Default on `dev_v2` |
| --- | --- | --- |
| `M365_EXCHANGE_BACKEND` | `graph` · `ews` · `auto` | **`graph`** (Graph-only for commands that honor the router) |
| Refresh token | **`M365_REFRESH_TOKEN`** (preferred), or `GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN` | Same value after `login`; unified cache **`token-cache-{identity}.json`** |

- **`graph`** — Graph APIs only (`resolveGraphAuth` + Graph REST).  
- **`ews`** — Legacy EWS only (`resolveAuth` + SOAP) where implemented.  
- **`auto`** — Try Graph first, then EWS if Graph auth or the call fails (per command).

Implementation: `src/lib/exchange-backend.ts`.

---

## Done (this branch)

| Item | Notes |
| --- | --- |
| Phase 0 stub | `getExchangeBackend()`, `DEFAULT_EXCHANGE_BACKEND='graph'`, helpers for tests |
| Auth / cache | **`m365-token-cache.ts`**: one **`token-cache-{identity}.json`** with EWS + Graph access slots; **`getUnifiedRefreshTokenFromEnv()`** |
| `whoami` | Uses **`GET /me`** on Graph when `graph` or `auto`; EWS path when `ews` only |
| `folders` | Graph: `listAllMailFoldersRecursive`, create/rename/delete via Graph mail folders; `ews` / `auto` unchanged semantics |
| `send` | Graph: `sendMail` + `buildGraphSendMailPayload` (file + **`referenceAttachment`** link URLs); **`auto`** tries Graph then EWS on failure |
| `mail` | Graph **`mail-graph.ts`**: list (**`--unread`**, **`--flagged`**, **`--search`**), **`--read`**, **`--download`**, **`--move`**, read state, **categories**, **follow-up flag** (incl. start/due), **`--sensitivity`**, **reply / reply-all / forward** (draft + attach + link + category + markdown); `auto` → EWS on Graph failure |
| `drafts` | Graph **`drafts-graph.ts`**: list, read, **create** / **edit** / **send** / **delete** (`createDraftMessage`, PATCH, attachment POSTs, `sendMailMessage`, `deleteMailMessage`); **`auto`** → EWS on Graph failure |
| `calendar` | Graph: **`listCalendarView`** for the resolved range when `graph` / `auto`; **`--list-attachments` / `--download-attachments`** use Graph **`/events/{id}/attachments`** when `graph` / `auto` (EWS fallback in `auto` if Graph auth fails); `auto` falls back to EWS on list-view failure |
| `findtime` | Graph: **`findMeetingTimes`**, then **`getSchedule`** + merged `availabilityView`; EWS `GetUserAvailability` only when `ews` or `auto` after both Graph paths fail |
| `create-event` | Graph: **`POST /me/events`** + attachments; **Places** (`/places/microsoft.graph.room`) for **`--list-rooms`**, **`--find-room`**, **`--room` by name**; calendar free/busy via room `calendarView`; `auto` falls back to EWS for room flows if Graph fails |
| `delete-event` | Graph: **`listCalendarView`** + organizer filter + `--search`; **`cancel`** vs **`delete`**; occurrence/instance id matching via **`seriesMasterId`**; `--scope future` still **unsupported**; `auto` falls back to EWS |
| `respond` | Graph: **`list`** via calendarView + pending filter; **`accept` / `decline` / `tentative`** via **`POST …/accept|decline|tentativelyAccept`**; organizer guard; `auto` falls back to EWS on Graph failure |
| `todo create --link` | Graph: **`GET …/messages/{id}`** (`getMessage`); shared mailbox via **`--user`** / **`--mailbox`** |
| `delegates list` | Graph: **`calendar/calendarPermissions`** when **`graph`**; **`auto`:** Graph then EWS; **`add`/`update`/`remove`:** EWS (blocked when **`graph`**) |
| `update-event` | Graph: list + **`PATCH …/events/{id}`** + attachments + **`PATCH attendees`** + **`--room` by name** + **`--occurrence` / `--instance`** (`seriesMasterId`); **`auto`** on Graph failure **does not** fall back to EWS if IDs came from Graph |
| `outlook-graph-client` | `listAllMailFoldersRecursive`; `MessagesQueryOptions.skip`; `OutlookMessage` body / lastModified |
| Unit tests | `src/lib/exchange-backend.test.ts`; **`graph-calendar-client`**: PATCH/DELETE/cancel; **`graph-event`**: accept/decline/tentativelyAccept |
| Integration | **`cli.integration.test.ts`**: `Graph backend` describe sets `M365_EXCHANGE_BACKEND=graph` — whoami, `update-event` / `delete-event` list, `update-event` PATCH, `respond list`; global `fetch` mock extended for **`/me`**, **`calendarView`**, event PATCH/DELETE/cancel/respond; **`runM365AgentCli`** clears leaked Commander flags (`--json`, `--id`, `--day`, …) on shared commands before each parse; **EWS** `--json` list-mode tests for `update-event` / `delete-event` re-enabled |

---

## Next (priority order — aligns with epic phases)

1. **Calendar parity** — Graph **`delete-event --scope future`** (or documented workaround).  
2. **Phase 6** — Remove EWS client usage when parity is verified — epic Phase 6 (optional: drop duplicate `GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN` env names after deprecation window).

---

## Open decisions

| Topic | Status |
| --- | --- |
| Default on `main` after merge: keep `graph` vs switch to `auto` | TBD before merge |
| Per-area env (`M365_MAIL_BACKEND`, …) vs single `M365_EXCHANGE_BACKEND` | Single var for now; split if needed |

---

*Last updated: 2026-04-02 — [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md): **recurring occurrence/instance** on Graph for **`update-event`** / **`delete-event`**; **`delete-event --scope future`** still in “Next”.*
