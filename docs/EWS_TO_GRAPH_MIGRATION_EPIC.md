# Epic: Migrate Exchange Web Services (EWS) to Microsoft Graph

**Status:** In progress — **`dev_v2`** uses **`M365_EXCHANGE_BACKEND`** (`graph` \| `ews` \| `auto`, default **`graph`**) and Graph-first mail/calendar flows per **[`docs/GRAPH_V2_STATUS.md`](./GRAPH_V2_STATUS.md)** (EWS remains for gaps and `ews` / `auto` fallback). **🟢 / 🟡 / 🔴 command matrix:** **[`docs/MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md)**.  
**GitHub Epic:** [#204 — EWS → Microsoft Graph migration](https://github.com/markus-lassfolk/m365-agent-cli/issues/204) (sub-issues under the epic)  
**Driver:** [Exchange Online retirement of EWS](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-overview) (phased; confirm dates in Microsoft docs and Message Center).  
**Strategy:** Phased migration with **Microsoft Graph as the primary implementation** and **EWS as fallback** until each slice is verified; then remove EWS for that slice.

---

## Code review snapshot (2026-04-02)

**`M365_EXCHANGE_BACKEND`** (`graph` \| `ews` \| `auto`, default **`graph`** on `dev_v2`) is implemented in **`src/lib/exchange-backend.ts`**. See **[`GRAPH_V2_STATUS.md`](./GRAPH_V2_STATUS.md)** for the live matrix.

**Graph-first when backend is `graph` or `auto` (non-exhaustive):** `whoami` (`/me`), `calendar` (calendarView), `findtime` (findMeetingTimes), `mail` (list + read), `send`, `folders`, `drafts` (list), `create-event` / `update-event` / `delete-event`, `respond`, `todo create --link` (GET message), `delegates list` (calendarPermissions). **`auto-reply`** remains EWS; **`oof`** is the Graph path for automatic replies (help text on `auto-reply` points to Graph).

**EWS still used** where noted in `GRAPH_V2_STATUS.md` (e.g. some `mail` / `drafts` mutations, `delegates` add/update/remove, `auto-reply`, full `calendar` attachment download paths, optional `auto` fallback).

---

## How to track this in GitHub

1. Use Epic **[#204](https://github.com/markus-lassfolk/m365-agent-cli/issues/204)** and its **sub-issues** (see inventory table below).
2. New slices: **New issue** → template **EWS → Graph migration task** (`.github/ISSUE_TEMPLATE/ews-graph-migration.yml`), then link as a sub-issue of #204 if needed (`gh api graphql` `addSubIssue`, or GitHub UI).
3. **One-off script** that created #204–#221: `scripts/create-ews-graph-issues.ps1` — **do not re-run** (it would create duplicate issues). Use it only as a reference for `gh` + GraphQL `addSubIssue`.

Labels used: `epic`, `migration`, `ews`, `graph`.

---

## Fallback model (implementation pattern)

Until a slice is marked **EWS removed**, implementations should follow a consistent pattern:

1. **Single entry point per domain** (e.g. `calendar-read`, `mail-send`) that chooses backend:
   - `graph` — use Graph only (for tenants already cut off from EWS or for tests).
   - `ews` — use EWS only (legacy / emergency).
   - `auto` *(default during migration)* — try Graph first; on **definitive** failure (e.g. known unsupported case, or opt-in retry policy), fall back to EWS once.
2. **Configuration** (names are proposals; implement in code when starting Phase 1):
   - Env: `M365_MAIL_BACKEND`, `M365_CALENDAR_BACKEND`, etc., with values `graph` | `ews` | `auto`, **or** one `M365_EXCHANGE_BACKEND=auto`.
   - Document in README when introduced.
3. **Observability:** Log which backend served each request (debug/`--verbose` only) so support can see Graph vs EWS.
4. **Tests:** For each migrated command, add tests for Graph path; keep EWS mocks until EWS deletion phase.

**Definition of “slice complete”:** Graph path is default in `auto`, feature parity documented, EWS path behind flag for that slice only until hard cutover.

---

## Reference

- [Migrate EWS apps to Microsoft Graph (overview)](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-overview)
- [EWS to Graph API mapping](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-api-mapping)
- Repo architecture note: `docs/ARCHITECTURE.md` (EWS vs Graph priority)

---

## Inventory: EWS touchpoints in this repo

| Area | Commands / modules | Graph direction | Notes | Issue | Status |
| --- | --- | --- | --- | --- | --- |
| Phase 0 foundation | Router, env, Azure scopes inventory | — | **`M365_EXCHANGE_BACKEND`** + `exchange-backend.ts`; scopes table under Phase 0 below | [#205](https://github.com/markus-lassfolk/m365-agent-cli/issues/205) | 🟡 |
| Calendar read | `calendar` | `GET calendarView` / shared calendars | Graph when `graph`/`auto`; EWS fallback / attachment paths per status doc | [#206](https://github.com/markus-lassfolk/m365-agent-cli/issues/206) | 🟡 |
| Free-busy / findtime | `findtime`, parts of schedule | `findMeetingTimes` / `getSchedule` | **`findtime`** uses Graph `findMeetingTimes` when `graph`/`auto`; EWS `getScheduleViaOutlook` when `ews` | [#207](https://github.com/markus-lassfolk/m365-agent-cli/issues/207) | 🟡 |
| Whoami | `whoami` | `/me` (+ optional mailboxSettings) | Graph `/me` when `graph`/`auto`; EWS when `ews` | [#208](https://github.com/markus-lassfolk/m365-agent-cli/issues/208) | 🟡 |
| Mail CRUD + actions | `mail` | Messages, move, patch, send | Graph **list + read** when `graph`/`auto`; other subcommands EWS or error on `graph` | [#209](https://github.com/markus-lassfolk/m365-agent-cli/issues/209) | 🟡 |
| Send | `send` | `sendMail` / draft send | Graph `sendMail` (+ file attach); `--attach-link` not on Graph | [#210](https://github.com/markus-lassfolk/m365-agent-cli/issues/210) | 🟡 |
| Drafts | `drafts` | Graph draft messages | Graph **list**; other flows EWS | [#211](https://github.com/markus-lassfolk/m365-agent-cli/issues/211) | 🟡 |
| Folders | `folders` | mailFolders | Graph CRUD + recursive list when `graph`/`auto` | [#212](https://github.com/markus-lassfolk/m365-agent-cli/issues/212) | 🟡 |
| Todo link | `todo --link` | Graph get message | **`getMessage`** (`outlook-graph-client`) for `todo create --link` | [#213](https://github.com/markus-lassfolk/m365-agent-cli/issues/213) | 🟡 |
| Calendar write | `create-event`, `update-event`, `delete-event` | Events API + online meetings | Graph paths when `graph`/`auto` (see status doc for unsupported flags) | [#214](https://github.com/markus-lassfolk/m365-agent-cli/issues/214) | 🟡 |
| Meeting response | `respond` | Accept/decline/tentative via Graph | Graph when `graph`/`auto` | [#215](https://github.com/markus-lassfolk/m365-agent-cli/issues/215) | 🟡 |
| Forward / counter | `forward-event`, `counter` | Event forward / propose times | **Graph** (`graph-event`) | [#216](https://github.com/markus-lassfolk/m365-agent-cli/issues/216) | ✅ |
| Auto-reply (EWS) | `auto-reply` | Deprecate in favor of Graph `oof` / mailboxSettings | **`oof`** Graph; **`auto-reply`** EWS (help text prefers `oof`) | [#217](https://github.com/markus-lassfolk/m365-agent-cli/issues/217) | 🟡 |
| Delegates | `delegates`, `delegate-client.ts` | Calendar permission / share APIs | **`list`:** Graph `calendarPermissions` when `graph`; **`list`:** Graph then EWS when `auto`; **mutations:** EWS (`graph` mode blocks) | [#218](https://github.com/markus-lassfolk/m365-agent-cli/issues/218) | 🟡 |
| Auth | `auth.ts`, `graph-auth.ts`, `m365-token-cache.ts` | Unified cache + `M365_REFRESH_TOKEN` | One `token-cache-{identity}.json` (EWS + Graph slots); legacy env names supported | [#219](https://github.com/markus-lassfolk/m365-agent-cli/issues/219) | ✅ |
| Tests / mocks | `src/test/mocks`, integration tests | Graph-shaped mocks | Mixed; `cli.integration.test` Graph backend cases | [#220](https://github.com/markus-lassfolk/m365-agent-cli/issues/220) | 🟡 |
| Docs | README, ENTRA_SETUP, SKILL | Remove EWS setup when cut over | [`ENTRA_SETUP.md`](./ENTRA_SETUP.md) lists Graph + EWS; Phase 6 cutover | [#221](https://github.com/markus-lassfolk/m365-agent-cli/issues/221) | 🟡 |

**Other Graph-only domains (not in original rows):** `planner`, `todo` (core), `files`, `sharepoint`, `site-pages`, `find`, `rooms`, `subscribe` / `subscriptions` — **no EWS** in those paths.

Legend: ⬜ not started / EWS-only · 🟡 in progress / partial Graph · ✅ done for stated slice (EWS fallback may still exist elsewhere)

---

## Phased roadmap

### Phase 0 — Foundation

- [x] Create GitHub Epic + child issues from inventory table ([#204](https://github.com/markus-lassfolk/m365-agent-cli/issues/204), [#205](https://github.com/markus-lassfolk/m365-agent-cli/issues/205)–[#221](https://github.com/markus-lassfolk/m365-agent-cli/issues/221))  
- [x] Env var **`M365_EXCHANGE_BACKEND`** + module **`exchange-backend.ts`** (default **`graph`** on `dev_v2`) — see [`GRAPH_V2_STATUS.md`](./GRAPH_V2_STATUS.md)  
- [ ] Agree default for **`main`** after merge (`graph` vs `auto`)  
- [x] Inventory Azure AD **delegated** permissions (manual setup — see [`ENTRA_SETUP.md`](./ENTRA_SETUP.md); scripts may add subsets):

| API | Delegated permissions (typical full CLI) |
| --- | --- |
| Microsoft Graph | `User.Read`, `Calendars.ReadWrite`, `Mail.ReadWrite`, `MailboxSettings.ReadWrite`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Tasks.ReadWrite`, `Group.ReadWrite.All`, `offline_access` |
| Office 365 Exchange Online | `EWS.AccessAsUser.All` |

`login` / device-code flows may request a combined scope string; `verify-token` validates granted scopes. `graph-auth` refresh uses Graph resource scopes (see `src/lib/graph-auth.ts`).

**Exit:** Epic linked; Phase 1 issue open.

### Phase 1 — Read-only paths

- [x] `whoami` → Graph (`/me`) when backend is `graph` or `auto`  
- [x] `calendar` calendarView on Graph when `graph`/`auto`  
- [x] `schedule` / `suggest` / `findtime` (Graph paths) — `findtime` uses EWS only when `M365_EXCHANGE_BACKEND=ews`  
- [x] Read paths: EWS fallback via `auto` where implemented  

**Exit:** Default `auto` uses Graph for reads; EWS fallback tested.

### Phase 2 — Mail stack

- [x] Graph mail REST (`outlook-graph`) — [x] **`dev_v2`:** `folders` / `send` / `mail` (list+read) / `drafts` (list) honor `M365_EXCHANGE_BACKEND`  
- [ ] Full Graph parity on `mail` / `drafts` (reply, search, draft CRUD, …)  

**Exit:** Mail commands use Graph in `auto`; EWS optional per env.

### Phase 3 — Calendar writes + meeting actions

- [x] `create-event`, `update-event`, `delete-event` (Graph when `graph`/`auto`; see `GRAPH_V2_STATUS.md` for gaps)  
- [x] `respond` (Graph when `graph`/`auto`)  
- [x] `forward-event`, `counter` (Graph)  

**Exit:** Calendar lifecycle on Graph in `auto` (with documented EWS-only flags).

### Phase 4 — Rules / OOF consolidation

- [x] Inbox rules Graph-only (`rules` today)  
- [x] `oof` = Graph mailbox settings; `auto-reply` documented as legacy (EWS rules)  

**Exit:** Users directed to Graph for automatic replies; `auto-reply` optional until EWS removal.

### Phase 5 — Delegates (redesign)

- [x] Spike: `delegates list` uses Graph **`calendar/calendarPermissions`** when `M365_EXCHANGE_BACKEND=graph`  
- [ ] Full Graph parity for add/update/remove (or document permanent EWS-only mutations)  
- [x] EWS mutations when `M365_EXCHANGE_BACKEND=ews` or `auto`; blocked when `graph`  

**Exit:** Documented parity or known limitations.

### Phase 6 — EWS removal

- [ ] Remove `callEws`, `ews-client` usage, SOAP mocks  
- [ ] Remove legacy **`GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN`** env aliases (optional; `M365_REFRESH_TOKEN` only)  
- [ ] Update Entra scripts, README, skills  

**Exit:** No EWS in repo; CI green.

---

## Child issue checklist (copy into each issue)

- [ ] Scope: one row (or one small group) from the inventory table  
- [ ] Graph implementation + tests  
- [ ] `auto` fallback to EWS (until slice signed off)  
- [ ] README / `--help` if user-visible flags added  
- [ ] This doc updated: Issue #, Status ✅ for that row  

---

## Open decisions (record answers here)

| Question | Decision |
| --- | --- |
| One env var vs per-area (`MAIL`, `CALENDAR`, …)? | *TBD* |
| Default during migration: `auto` everywhere? | *TBD* (recommended: yes) |
| Breaking CLI changes for `delegates`? | *TBD* |

---

*Last updated: 2026-04-02 — `dev_v2`: `M365_EXCHANGE_BACKEND`, Graph-first commands; see [`GRAPH_V2_STATUS.md`](./GRAPH_V2_STATUS.md).*
