# Graph vs EWS parity and verification matrix

**Purpose:** Single reference for **what differs** between Microsoft Graph and EWS in this CLI, how **`M365_EXCHANGE_BACKEND`** (`graph` | `auto` | `ews`) affects routing, and how to verify behavior.

**Related:** [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md) (migration status), [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md) (Graph vs CLI coverage), [`exchange-backend.ts`](../src/lib/exchange-backend.ts) (mode semantics), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md) (OAuth scopes).

---

## 1. Backend modes (summary)

| Mode | Tries Graph first? | Uses EWS when? |
| --- | --- | --- |
| **`auto`** (default) | Yes | Graph auth fails, Graph API call fails for the operation, or the command has **no Graph implementation** for that path. **Successful Graph result (including empty list) is final** — EWS is **not** used to “fill in” different data. |
| **`graph`** | Yes (only path) | Never |
| **`ews`** | No | Always (Exchange Web Services only) |

---

## 2. Commands that honor `M365_EXCHANGE_BACKEND`

These commands **branch** on `getExchangeBackend()` in code (see `src/commands/*.ts`):

| Command | `graph` | `auto` | `ews` | Graph vs EWS output / behavior differences |
| --- | --- | --- | --- | --- |
| `whoami` | `/me` | Graph first → EWS if Graph auth/`/me` fails | EWS `GetUserOwaUserInfo` (SOAP) | **Different APIs:** Graph returns `displayName`, `mail`, `userPrincipalName`; EWS returns display name + primary SMTP from mailbox context. Rarely differs for same user. |
| `mail` | `mail-graph.ts` when flags eligible | Graph first if eligible + Graph auth OK; else EWS | EWS `GetItem` / folder APIs | **IDs:** Graph message and folder IDs differ from EWS `ItemId` / `ChangeKey`. **Not interchangeable.** Some flag combinations **only** on EWS or `outlook-graph` (Graph errors with hint). |
| `send` | `sendMail` + attachments | Graph then EWS on Graph failure | EWS `CreateItem` send | **Pipeline:** Graph uses JSON `sendMail`; EWS uses MIME/SOAP. Tenant policy may allow one and deny the other. |
| `folders` | Graph `mailFolders` | Graph first; EWS fallback on failure in `auto` | EWS folder list | **Folder IDs** and naming differ; custom folder resolution differs slightly. |
| `drafts` | `drafts-graph.ts` | Graph first; EWS fallback in `auto` on failure | EWS drafts | **Draft IDs** differ between APIs. |
| `calendar` | `calendarView` / Graph helpers | Graph first; EWS in `auto` if Graph fails | EWS `FindItem` calendar | **Event IDs:** Graph `id` vs EWS `ItemId` — not interchangeable. Recurrence/instance handling differs (`seriesMasterId` vs EWS calendar patterns). |
| `create-event` | `POST /me/events`, Places for rooms | Graph first; EWS fallback for some room paths in `auto` | EWS `CreateItem` | **Teams:** Graph `isOnlineMeeting` + `onlineMeetingProvider`; EWS `IsOnlineMeeting`. **Attachments** metadata may differ. |
| `update-event` | Graph PATCH + attachments | Same; **no EWS fallback** after Graph-backed listing when `graph` | EWS `UpdateItem` | **🟡 Mixed IDs:** If you loaded events via Graph, you must use Graph IDs — CLI errors if you try EWS fallback with Graph-only data. |
| `delete-event` | Graph cancel/delete | Graph first; EWS in `auto` when appropriate | EWS `DeleteItem` | **`--scope future`:** Graph **PATCH** on series master after **`…/instances`** (truncate recurrence); EWS uses SOAP `deleteEvent`. |
| `respond` | `calendarView` + invitation APIs | Graph first; EWS in `auto` if Graph fails | EWS respond | Same **event ID** caveats as calendar. |
| `findtime` | `findMeetingTimes` / `getSchedule` (Graph) | Graph strategies first; EWS `GetSchedule` in `auto` if both fail | EWS-only path | **Availability strings** may differ in formatting; merged free/busy logic is aligned in code but underlying data source differs. |
| `delegates` **list** | `calendarPermissions` (Graph) | Graph first; EWS `GetDelegates` **only if Graph request fails** (not if empty) | EWS | **Semantics:** Graph “calendar permissions” ≠ EWS “delegates” matrix — **different products**; empty Graph list is **not** supplemented from EWS. |
| `delegates` **add/update/remove** (EWS matrix) | **Error** — set `M365_EXCHANGE_BACKEND=ews` or **`auto`** | **EWS** (`requireEwsForDelegateMutations` blocks **only** when mode is `graph`) | EWS | Classic EWS delegate matrix — not Graph. |
| `delegates` **calendar-share** **add/update/remove** | **Graph** `calendarPermissions` | **Graph** (same) | **Graph** | **Calendar sharing** model — distinct from EWS delegates; see [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md). |

Commands that **do not** read `M365_EXCHANGE_BACKEND` for mail/calendar (always Microsoft Graph or other APIs):  
`contacts`, `meeting`, `onenote`, `todo`, `planner`, `files`, `sharepoint`, `find`, `graph-search`, `rooms`, `rules`, `oof`, `outlook-graph`, `graph-calendar`, `forward-event`, `counter`, `subscribe`, …

**OneNote:** Implemented exclusively via **Microsoft Graph** OneNote APIs (notebooks including **GetNotebookFromWebUrl**, section groups, sections, pages, HTML create/read, content PATCH, global page list, preview, page **copyToSection**, section **copyToNotebook** and **copyToSectionGroup**, async **onenoteOperation** poll, optional `/groups` and `/sites` roots). **Embedded resource binaries:** `onenote resource-download` (`GET …/resources/{id}/content`); page HTML with optional **`includeIDs`**: `onenote content` / `export --include-ids`. **EWS does not expose OneNote** — there is no EWS parity row or fallback. Advanced ink / multipart upload scenarios may still need Graph directly — see [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md).

---

## 3. Known “different output” scenarios (not bugs by default)

| Scenario | Graph | EWS / notes |
| --- | --- | --- |
| **Item/event IDs** | OData IDs, often long base64-like strings | EWS `ItemId` + `ChangeKey` — **never mix** across backends for the same operation sequence. |
| **Empty lists** | `[]` from Graph is authoritative in `auto` | EWS is **not** queried to “double-check” emptiness. |
| **Delegates vs calendar sharing** | Graph calendar permissions | EWS delegates — different permission models. |
| **`delete-event --scope future`** | `GET /events/{seriesMasterId}/instances` + `PATCH` recurrence `endDate` | EWS `deleteEvent` with scope — same CLI flag. |
| **`auto-reply` command** | Not this CLI’s Graph UX | EWS inbox-rule template model — use **`oof`** / **`rules`** for Graph-native behavior. |
| **Token / cache** | Graph access token `scp` must match app + consent | Stale `token-cache-*.json` can show wrong scopes until refresh — see [`ENTRA_SETUP.md`](./ENTRA_SETUP.md). |

---

## 4. Automated tests in this repo

| Layer | Location | What it proves |
| --- | --- | --- |
| Mode parsing | [`src/lib/exchange-backend.test.ts`](../src/lib/exchange-backend.test.ts) | `graph` / `auto` / `ews` / invalid env |
| Graph auth cache | [`src/test/graph-auth.test.ts`](../src/test/graph-auth.test.ts) | App id + critical `scp` invalidate stale cache; refresh fallback |
| CLI integration (mocked HTTP) | [`src/test/cli.integration.test.ts`](../src/test/cli.integration.test.ts) | Default `ews` for most tests; **`graph`** block: `whoami`, `update-event`, `delete-event` (incl. **`--scope future`** truncation path), `respond` use Graph mocks |
| Auto routing | [`src/test/cli.integration.test.ts`](../src/test/cli.integration.test.ts) (`describe('Auto backend (M365_EXCHANGE_BACKEND=auto)')`) | `whoami` / `whoami --json`: Graph success vs `/me` 401 → EWS; `delegates list`: empty Graph permissions → **no** EWS substitute; `calendar` / `delete-event --json` with graph token |
| Graph-only contrast | Same file (`describe('Graph backend (M365_EXCHANGE_BACKEND=graph)')`) | `whoami` on `/me` **401** → **throws** (no EWS fallback) |
| `getOwaUserInfo` + `EWS_USERNAME` | [`src/test/ews-client.test.ts`](../src/test/ews-client.test.ts) | `ResolveNames` SOAP embeds **current** `process.env.EWS_USERNAME` per call |
| CLI + `EWS_USERNAME` routing | [`src/test/cli.integration.test.ts`](../src/test/cli.integration.test.ts) (`whoami` › non-empty `EWS_USERNAME`) | Non-empty unresolved entry hits **people-search** mock (distinct from empty-whoami mock) |

**Not fully automated:** Live tenant policy, conditional access, admin consent gaps, delegated mailbox permissions, and every mail/calendar flag combination. Use **manual** runs (see §5).

---

## 5. Manual verification checklist (release / tenant change)

1. `m365-agent-cli verify-token` — **App ID** = **EWS_CLIENT_ID**; **`scp`** includes expected delegated scopes ([`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md)).
2. **`M365_EXCHANGE_BACKEND=graph`** — smoke: `whoami`, `mail inbox`, `calendar today`, `send` (test recipient), `folders`.
3. **`M365_EXCHANGE_BACKEND=ews`** — same smoke; confirm EWS-only paths still work (legacy).
4. **`M365_EXCHANGE_BACKEND=auto`** — repeat smoke; confirm Graph path used when Graph auth succeeds (see CLI output / “Graph” in messages where applicable).
5. **`--mailbox`** (shared/delegated) — confirm `Mail.Read*.Shared` / `Calendars.*.Shared` consent and no spurious EWS fallback on empty Graph results.

Optional: set `M365_AGENT_ENV_FILE` to `.env.beta` for a second app registration ([`ENTRA_SETUP.md`](./ENTRA_SETUP.md)).

---

## 6. When to treat a difference as a bug

- **Graph fails** with 403 but **EWS works** for the same user action → often **scopes**, **tenant policy**, or **delegation** — fix Entra / consent first.
- **Same backend, wrong CLI behavior** → file an issue with command, flags, and redacted error JSON.
- **`auto` calls EWS** when Graph returned **200 + empty array** for a list operation → likely **bug** vs documented policy — verify against §1.

*Last updated: 2026-04-03 — OneNote copy APIs + integration test note for `delete-event --scope future`.*
