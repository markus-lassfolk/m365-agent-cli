---
name: m365-agent-cli
version: 2026.7.3
description: Microsoft 365 CLI (EWS + Graph) for calendar, mail, OneDrive, Planner, SharePoint, To Do, Teams, Bookings, Excel-on-drive, presence, inbox rules, delegates, subscriptions, Graph Search, Microsoft Viva / employee experience (`viva` — Graph beta: tenant `/employeeExperience`, user work time/insights/roles/learning, admin+org itemInsights, workHoursAndLocations, meeting Engage Q&A), Microsoft 365 Copilot APIs (`copilot` — retrieval, search, chat, reports, packages, meeting insights, interaction export, notify-help), and raw `graph invoke`/`graph batch`. Use when the user needs Outlook/Exchange, Graph, or M365 automation from the terminal.
# `version` matches the npm CLI release; run `npm run sync-skill` after bumping package.json.
metadata:
  openclaw:
    requires:
      bins:
        - m365-agent-cli
---

# m365-agent-cli

CLI for Microsoft 365: **Exchange Web Services (EWS)** and **Microsoft Graph**. Prefer `m365-agent-cli <command> --help` for exact flags on each command.

## Authentication and profiles

- Config directory: `~/.config/m365-agent-cli/` (`.env`, token caches).
- **Unified OAuth cache:** `token-cache-{identity}.json` — holds both EWS and Graph access-token slots (default identity: `default`). Env: **`M365_REFRESH_TOKEN`** preferred, or `GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN` (legacy aliases).
- **`--identity <name>`** — use a named cache profile (Graph- and EWS-backed commands that expose the flag). Default is `default`.
- **`--token <token>`** — override cached access token for that request (advanced).
- Interactive login: `m365-agent-cli login` (device code); tokens land in `.env` / caches. **Delegated Graph scopes** (Entra app + CLI features): **`docs/GRAPH_SCOPES.md`**. For **another user’s** mail/calendar (`--mailbox`), include **Mail/Calendars \*.Shared**; for **`find`**, **Places**/**`rooms`**, add scopes listed there; re-`login` after changing the app registration.
- Check session: `m365-agent-cli whoami`, `m365-agent-cli verify-token [--identity <name>]`, `verify-token --capabilities` for a read/write feature matrix from token scopes.

## Quick command choice (agents)

| User goal | Start here |
| --- | --- |
| Today’s calendar / shared mailbox (EWS-style) | `calendar`, `mail` — **`--mailbox`** where supported |
| Graph REST mail (OData, folders, move) | `outlook-graph` |
| Graph REST calendar (view, REST ids, respond) | `graph-calendar` |
| OneDrive / SharePoint files, sharing, delta | `files` |
| Excel workbook on drive (ranges, tables, pivots, charts, sessions, **comments beta**) | `excel` |
| Word / PowerPoint preview, upload, share, versions, … | `word`, `powerpoint` — per-item ops mirror **`files`** (list/delta/search stay **`files`**); workflows: **`docs/AGENT_WORKFLOWS.md`** §7, **`docs/WORD_POWERPOINT_EDITING.md`** |
| Teams channel/chat message + **@mentions** | `teams channel-message-send` / `chat-message-send` — **`--at userId:DisplayName`** + **`--text`** with `@DisplayName` |
| Microsoft Search (stable hits JSON) | `graph-search … --json-hits` |
| Anything else on Graph | `graph invoke` / `graph batch` |

Longer workflows (deltas, read-only, drive flags, Teams + file links): **`docs/AGENT_WORKFLOWS.md`**. **`--json`** / read-only per module: **`docs/CLI_SCRIPTING_APPENDIX.md`** + generated **`docs/CLI_SCRIPTING_INVENTORY.md`**. Optional **Cursor MCP** adapter: **`packages/m365-agent-cli-mcp`** in this repo.

## Weekly planning and approval-gated calendar writes (OpenClaw / PA)

Personal-assistant stacks (e.g. **openclaw-personal-assistant**) should use the verbs below. Listing and creating support **two equivalent spellings**: the default form (**`calendar <range>`** / **`create-event`**) and explicit subcommands **`calendar list …`** / **`calendar create …`** (same flags as **`calendar list --help`** / **`create-event --help`**).

| Goal | Commands |
| --- | --- |
| List events for a day or range | **`calendar <range>`** or **`calendar list <range>`** — keywords (`today`, `week`, `nextweek`, …), optional **`[end]`**, **`--json`**, span modes **`--days`**, **`--business-days`**, **`--now`**, etc. ([`CLI_REFERENCE.md`](../../docs/CLI_REFERENCE.md) § Calendar.) |
| Create one event | Top-level **`create-event "Subject" <start> <end>`** or **`calendar create "Subject" <start> <end>`** — same options (**`--day`**, **`--timezone`**, **`--category`**, **`--mailbox`**, **`--calendar`**, **`--teams`**, …). |
| Many events after one user approval | **`graph batch`** — up to **20** requests per file; each item can **`POST /me/events`**. Alternatively run **`create-event`** / **`calendar create`** once per block. Patterns: [`AGENT_WORKFLOWS.md`](../../docs/AGENT_WORKFLOWS.md) §5a and §5b. |
| Free/busy slots, overload signals | **`findtime`**, **`schedule`**, **`suggest`**. |
| Graph-native **`calendarView`**, REST event ids, incremental sync | **`graph-calendar list-view`**, **`get-event`**, **`graph-calendar events-delta --state-file …`**. |

**Approval gate:** the CLI performs writes immediately when invoked; the **agent** must only call **`create-event`** / **`calendar create`** / **`graph batch`** **after** explicit user confirmation (never auto-commit a proposed week).

**Behavioral tracking** (daily focus hours, mood, CSV/JSON in the user workspace) is **out of scope** for this CLI—keep it in the PA repo; weekly planning can **read** that file before proposing calendar steps.

## Delegation and shared access

- **`delegates` vs `delegates calendar-share`:** **Classic delegates** (`delegates add|update|remove`) are **EWS-only** (per-folder matrix). **`delegates calendar-share`** uses **Microsoft Graph** `calendarPermission` on the default calendar — prefer this for Graph-first workflows. **`M365_EXCHANGE_BACKEND=graph`** blocks classic delegate mutations; use **`calendar-share`** or switch backend. Full stance: repo **`docs/GRAPH_EWS_PARITY_MATRIX.md`** §2a.
- **EWS shared mailbox:** `--mailbox <email>` on calendar, mail, send, folders, drafts, respond, findtime, delegates, auto-reply (and similar) to act as that mailbox where supported.
- **Graph drive roots (`files`, `excel`, `word`, `powerpoint`):** use **at most one** of **`--user <upn-or-id>`** (delegated **`/users/{id}/drive`**), **`--drive-id`**, **`--site-id`**, or **`--site-id`** + **`--library-drive-id`**. Default is **`/me/drive`**. These flags are **not** the same as EWS **`--mailbox`**.
- **Graph delegation (other areas):** **`--user`** on supported commands (e.g. inbox **rules**, **oof**, **`todo`**, **`outlook-categories`**, **`contacts`**, **`schedule`** / meeting-time helpers, **`graph-calendar`**, **`outlook-graph`**, **`subscribe`**, **`rooms`/places**, **`teams list`**, **`org manager`**, **`org direct-reports`**) switches paths to **`/users/{id}/...`** where Graph exposes that shape. Requires scopes per **`docs/GRAPH_SCOPES.md`** (**`Contacts.Read*.Shared`** for delegated contacts; **`Team.ReadBasic.All`** for **`teams list --user`**; **`User.Read.All`** often needed for **`org … --user`**). **`teams chats`** remains **`GET /me/chats`** only—there is no delegated list of another user’s chats in Graph. **`bookings`**, most other **`teams`** subcommands, **`presence`**, and **`graph invoke`/`batch`** may still use **`/me/...`** or fixed paths—use **`graph invoke`** when you need a **`/users/{id}/...`** URL the CLI does not wrap.

## Safety

- **`--read-only`** (root) or **`READ_ONLY_MODE=true`** in env / `.env` runs `checkReadOnly()` before specific mutating actions (exits before the request). The **authoritative list** is the **Read-Only Mode** table in this repo’s [`docs/CLI_REFERENCE.md`](../../docs/CLI_REFERENCE.md) (includes **`contacts`**, **`onenote`**, **`meeting`** mutating subcommands; kept in sync with `grep checkReadOnly src` in source).
- Read/query commands (e.g. `calendar`, `schedule`, `suggest`, **`outlook-graph list-mail` / `list-messages` / `get-message` / attachment list-get-download** / folder list-get, `subscriptions list`, `rules list`, **`outlook-categories list`**) are **not** gated unless they call `checkReadOnly`—see README. **`outlook-categories create` / `update` / `delete`** are mutating and **are** gated.
- **`m365-agent-cli --help`** only lists root flags (e.g. `--read-only`). Per-command flags are on each subcommand’s help.

## Categories, labels, and colors (Outlook vs To Do vs Planner)

- **Mail and calendar (EWS):** Items use **category name strings** (same names as in Outlook). **Colors** are not stored per message/event in the CLI; they come from the mailbox **master category list**. **List** names + preset colors: **`outlook-categories list`**. **Create / rename / delete** master entries (and colors `preset0`..`preset24`): **`outlook-categories create`**, **`update`**, **`delete`** (Graph; needs **`MailboxSettings.ReadWrite`**). Set categories on existing messages: **`mail --set-categories <id> --category <name>`**, **`mail --clear-categories <id>`**. On **outgoing** reply/forward: **`mail --with-category <name>`** (repeatable) with **`--reply` / `--reply-all` / `--forward`**. Drafts: **`drafts --create` / `--edit`** with **`--category`**; **`--clear-categories`** on edit. Calendar: **`create-event` / `update-event`** with **`--category`**; verbose **`calendar`** shows categories when present.
- **Outlook Graph REST** (distinct from EWS **`mail`** / **`folders`):** **`outlook-graph`** — mail folder **`list-folders`**, **`child-folders`**, **`get-folder`**, **`create-folder`**, **`update-folder`**, **`delete-folder`**; messages **`list-messages`** (**`--folder`**, **`--top`**, **`--all`**, **`--filter`**, **`--orderby`**, **`--select`**), **`list-mail`** (mailbox-wide **`GET /messages`**, **`--search`**, **`--skip`**, same OData flags), **`get-message`**; **`send-mail`** (**`--json-file`** `{ message, saveToSentItems }`); **`patch-message`** (**`--json-file`**); **`delete-message`** (**`--confirm`**); **`move-message`** / **`copy-message`** (**`--destination`**); **`list-message-attachments`**, **`get-message-attachment`**, **`download-message-attachment`** (**`-o`**); **`create-reply`**, **`create-reply-all`**, **`create-forward`** (**`--to`** comma-separated, **`--comment`**); **`send-message`** (draft id). Contacts **`list-contacts`**, **`get-contact`**, **`create-contact`** / **`update-contact`** (**`--json-file`**), **`delete-contact`**. Use **`--user`** for delegation where supported.
- **Graph calendar REST** (distinct from EWS **`calendar`** / **`respond`):** **`graph-calendar`** — **`list-calendars`**, **`get-calendar`**, **`list-view`** (**`--start`**, **`--end`**, optional **`--calendar`**), **`get-event`** (**`--select`**); **`accept`**, **`decline`**, **`tentative`** (**`--comment`**, **`--no-notify`**). Use **`--user`** for delegation where supported.
- **Microsoft To Do (Graph):** Tasks use **`categories[]`** as **string labels** (independent of Outlook master categories). **`todo create --category`**, **`todo update --category` / `--clear-categories`**. **`todo get`** supports OData **`--filter`**, **`--orderby`**, **`--select`**, **`--top`**, **`--skip`**, **`--expand`**, **`--count`**, or **`--status` / `--importance`**; **`--task`** (single task) passes **`--select`** to Graph. Create/update: **`--start`**, per-field time zones (**`--timezone`**, **`--due-tz`**, **`--start-tz`**, **`--reminder-tz`**), **`todo update --clear-start`**. Lists: **`todo create-list`**, **`todo update-list`**, **`todo delete-list`**. Checklist: **`todo add-checklist`**, **`todo update-checklist`**, **`todo delete-checklist`**, **`todo list-checklist-items`**, **`todo get-checklist-item`**. Recurrence: **`--recurrence-json`** / **`--clear-recurrence`** on create/update. Attachments: **`todo list-attachments`**, **`todo get-attachment`**, **`todo download-attachment`** (**`--output`**; Graph **`.../attachments/{id}/$value`** for file bytes), **`todo add-attachment`**, **`todo add-reference-attachment`**, **`todo upload-attachment-large`**, **`todo delete-attachment`**, **`todo attachment-session`** (list/get/patch/delete + **content** get/put/delete; v1 has no POST on the collection). Todo container: **`todo root`** get/patch/delete (**`--confirm`** on delete). Linked resources: Graph **`linkedResource`** (**`todo linked-resource`** `list`/`create`/`get`/`update`/`delete`), or merge on **`todo add-linked-resource`** / **`remove-linked-resource`** (**`displayName`** / **`--display-name`**). Delta sync: **`todo delta`**. Task extensions: **`todo extension list`**, **`get`**, **`set`**, **`update`**, **`delete`**. List extensions: **`todo list-extension list`**, **`get`**, **`set`**, **`update`**, **`delete`**.
- **Planner (Graph):** Tasks use **six fixed slots** `category1`..`category6` (**`appliedCategories`**). Display names are in **plan details**: **`planner get-plan-details`**, **`planner update-plan-details`** (**`--names-json`**, **`--shared-with-json`**), rare destructive **`planner delete-plan-details`** / **`planner delete-task-details`** (**`--confirm`**). **Tasks assigned to you:** **`planner list-my-tasks`** (alias **`planner tasks`**). Tasks: **`planner create-task`** / **`update-task`** (**`--assign`**, **`--priority`** 0–10, **`--preview-type`**, **`--conversation-thread`**, **`--order-hint`**, **`--assignee-priority`**, **`--due` / `--start`**, **`--clear-due` / `--clear-start`**, label flags). **`planner get-task --with-details`**, **`planner delete-task`**. Per-user lists (Graph **`/users/{id}/planner/...`**; may **403**): **`planner list-user-tasks --user`**, **`planner list-user-plans --user`**. Plans/buckets: **`create-plan --group`**, beta **`create-plan --roster`**, or beta **`create-plan --me`** (personal / user container; **`POST /me/planner/plans`**), **`update-plan`**, **`delete-plan`**, **`create-bucket`**, **`update-bucket`** (**`--order-hint`**), **`delete-bucket`**. Task details: **`get-task-details`**, **`update-task-details`**, **`add-checklist-item`**, **`update-checklist-item`**, **`remove-checklist-item`**, **`add-reference`**, **`remove-reference`**. Task board ordering (Graph format resources): **`get-task-board --view`** `assignedTo` \| `bucket` \| `progress`, **`update-task-board`** (**`--json-file`**). Beta roster container APIs: **`planner roster`** `create` \| `get` \| `list-members` \| `add-member` \| `remove-member`. Other beta: **`list-my-day-tasks`**, **`list-recent-plans`**, **`list-favorite-plans`**, **`list-roster-plans`**, **`get-me`**, **`add-favorite`**, **`remove-favorite`** (optional **`--user`** for delegated **`/users/{id}/...`** where supported); **`update-me`** (**`--etag`** from **`get-me`**, **`--json-file`** merge body, optional **`--user`**); **`delta`** (use **`--url`** from **`nextLink`** or **`deltaLink`**).

## Attachments (EWS)

File and link attachments in the **EWS** flows below apply to **messages** and **calendar items**. **Microsoft To Do** file and link attachments use **Graph** via **`todo`** (see the To Do bullet above), not these EWS commands.

| Flow | Command / flags |
| ------ | ----------------- |
| Send email with files or links | **`send --attach <paths>`** (comma-separated), **`send --attach-link <spec>`** (repeatable; spec is **`Title&#124;https://url`** or a bare **`https://`** URL) |
| Drafts | **`drafts --create` / `--edit`** with **`--attach`** and **`--attach-link`** (same pattern as `send`) |
| Download from a message | **`mail -d <id>`** (or **`--download`**), **`--output <dir>`** for save location |
| Reply / forward with attachments or outgoing categories | **`mail --reply` / `--reply-all` / `--forward`** with **`--attach`**, **`--attach-link`**, **`--with-category`** (uses draft + send; **`--draft`** to save only). Use **message id** from list/read, not the numeric index, for non-interactive scripts. |
| Add CC / BCC recipients when replying or forwarding | **`mail --reply` / `--reply-all` / `--forward`** with **`--cc <emails>`** and/or **`--bcc <emails>`** (comma-separated). These **add** recipients on top of the To/Cc a reply-all already includes. |
| Calendar event files / links | **`create-event`** / **`update-event`**: **`--attach`**, **`--attach-link`** (adds after the item exists; paths relative to cwd where documented) |
| List or download event attachments | **`calendar --list-attachments <eventId>`**, **`calendar --download-attachments <eventId>`** with **`--output`**, **`--force`** to overwrite |
| Inbox rules | **`rules`**: conditions like **`--hasAttachments`**, actions like **`--forwardAsAttachmentTo`** (see **`rules --help`**) |

Paths are validated (no unsafe traversal); large files may hit Exchange limits—see command help and README.

## Calendar: date ranges and weekday (“business day”) windows

The **`calendar`** command accepts a start day (and optional end day) plus **mutually exclusive** span modes:

- **`--days <n>`** — **N consecutive calendar days** forward from the start day (includes the start day). No end-date argument.
- **`--previous-days <n>`** — **N consecutive calendar days** ending on the start day.
- **`--business-days <n>`**, **`--next-business-days <n>`** (alias), or **`--busness-days <n>`** (typo alias) — **N weekdays (Mon–Fri)** forward from the start anchor; if the anchor falls on a weekend, the first counted weekday is the next Monday (see implementation in `calendar-range.ts`). Only one of these three may be set; conflicting values error.
- **`--previous-business-days <n>`** — **N weekdays** backward ending on the last weekday on or before the start anchor.
- **`--now`** — After the range is computed, **raise the API query start to the current instant** so the list omits meetings that **already ended** and only shows what is **ongoing or upcoming** within the window. Combine with e.g. **`calendar today --business-days 5 --now`**. Not valid with **`--previous-days`** or **`--previous-business-days`**.

Do **not** combine the span modes (`--days` / `--previous-days` / `--business-days` / …) with each other or with an explicit **`[end]`** date argument. Do **not** combine them with week keywords **`week` / `thisweek` / `lastweek` / `nextweek`** — use a single day (e.g. `today`) as the start when using `--business-days` / `--days` / etc.

## Command map (high level)

| Area | Commands / notes |
| ------ | ------------------ |
| Calendar | `calendar` / **`calendar list`** (default list subcommand), **`calendar create`** (alias for `create-event`), **`--list-attachments`**, **`--download-attachments`**, `create-event` / `update-event` (**`--attach`**, **`--attach-link`**), `delete-event` (**`--scope all` \| `this` \| `future`**; recurring: **`--occurrence`**, **`--instance`**), `respond`, `findtime`, `forward-event`, `counter`, `schedule`, `suggest`, **`graph-calendar`** (Graph list/view/get + invitation responses) |
| Mail | `mail` (**`-d`**, reply/forward **`--attach`**, **`--attach-link`**, **`--with-category`**), `send`, `drafts`, `folders`; **`outlook-graph`** (Graph mail: **`list-mail`**, **`send-mail`**, **`patch-message`**, move/copy, attachments, reply/forward drafts + **`send-message`**) |
| Outlook categories (Graph) | `outlook-categories` **`list`**, **`create`**, **`update`**, **`delete`** — master list **names + colors** |
| Contacts (Graph only) | **`contacts`** — **`folders`** / **`folder`** (list/get/create/update/delete/children), **`list`** (**`--filter`**), **`show`**, **`create`** / **`update`** / **`delete`** (**`--json-file`**), **`search`**, **`delta`**, **`photo`** get/set/delete, **`attachments`** list/add file/**`add-link`**/show/download/delete, **`extension`** list/get/set/update/delete (**`-f`/`--folder`**, **`--child-folder`** for nested contact-folder extension paths), **`merge-suggestions`** get/set/delete (Graph **beta** userSettings); **`--user`** for delegated mailbox |
| OneNote (Graph only) | **`onenote`** — **`notebooks`** / **`notebook`** (list/get/create/update/delete/**`from-web-url`**), **`section-group`**, **`section`** (list/get/create/update/delete/**`copy-to-notebook`**/**`copy-to-section-group`**), **`pages`** / **`list-pages`**, **`page`**, **`page-preview`**, **`content`**, **`export`**, **`create-page`**, **`delete-page`**, **`patch-page-content`**, **`copy-page`**, **`operation`** (poll async copy); **`--group`** / **`--site`** OneNote roots |
| Online meetings (Graph) | **`meeting`** — create (simple or **`--json-file`**), get, update, delete (`/me/onlineMeetings`). Calendar+Teams invites: **`create-event … --teams`**. |
| Files | `files` (list, search, **delta**, **shared-with-me**, upload, **copy**, **move**, download, share, **invite**, **permissions**, **permission-remove**, **permission-update**, versions, checkout/checkin, …) — drive flags **`--user`** / **`--drive-id`** / **`--site-id`** / **`--library-drive-id`** |
| Planner | `planner` (tasks, plans, buckets; **labels** on tasks; **`delete-plan-details`** / **`delete-task-details`** for rare details-facet deletes) |
| SharePoint | **`sharepoint`** / **`sp`** — **`resolve-site`**, **`get-site`**, **`drives`**, lists, **`get-list`**, **`columns`**, **`items`** (**`--filter`** / **`--orderby`** / **`--top`** / **`--url`** / **`--all-pages`**), **`get-item`**, **`create-item`** / **`update-item`** (**`--fields`** or **`--json-file`**), **`delete-item`**, **`items-delta`**, followed sites; **`pages`** (site pages) |
| Directory / rooms | `find`, `rooms` |
| Teams (Graph) | **`teams`** — **list** (**`--user`** for another user’s joined teams), **get**, **`channel-files-folder`** (Files tab → **`files --drive-id`**), **channels** / **all-channels** (**`--filter`**) / **incoming-channels** / **primary-channel** / **channel-get**, **channel-members**, **messages** / **channel-message-get** / **channel-message-send** / **message-replies** / **channel-message-reply**, **`channel-message-react`** / **`chat-message-react`** (`--unset`), **tabs**, **members**, **apps**, **chats** / **chat-get** / **chat-messages** / **chat-message-get** / **chat-message-replies** / **chat-message-send** / **chat-message-reply** / **chat-members** / **chat-pinned** |
| Bookings (Graph) | **`bookings`** — full **`/solutions/bookingBusinesses`** surface (incl. **business-create** / **delete** / **publish** / **unpublish**, **currency-get**, **custom-question**); **`staff-availability`** — **application-only** via **`--token`**, body **`--json-file`** |
| Excel on drive (Graph) | **`excel`** — same drive flags as **`files`**; worksheets; **range** / **range-patch** / **range-clear**; **used-range**; **tables** (CRUD, rows, columns); **pivot-tables** + **pivot-table-***; **names** + **name-get** + **worksheet-names** / **worksheet-name-get**; **charts**; **workbook-get**; **application-calculate**; **session-create** / **session-refresh** / **session-close**; optional **`--session-id`** on mutating calls; **`comments-*`** (Graph **beta**) |
| Presence (Graph) | **`presence`** — **me**, **user**, **bulk**, **set-me** / **set-user** (prints **`sessionId`**), **clear-me** / **clear-user** (**`--session-id`**); webhook flows → **`subscribe`** + **`serve`** |
| Word / PowerPoint (drive item) | **`word`** / **`powerpoint`** — full **`files`** per-item mirror (**`list-item`**, **`follow`**, **`sensitivity-assign`/`extract`**, **`retention-label`**, **`permanent-delete`**, upload/share/versions/checkout/…) + **preview**/**thumbnails**; **`files`** for **list**/**delta**/**search**; no in-file **comment** / slide OM APIs like **`excel comments-*`** — **`graph invoke`** if documented |
| Raw Graph | **`graph`** — **invoke**, **batch** (JSON `$batch`); pair with scopes for the target API |
| Graph Search | **`graph-search`** — `POST /search/query` (presets / `--types`, `--merge-json-file`, `--body-file`, plus `--fields` / `--content-sources` / `--region` / … — see command help) |
| Viva / employee experience (Graph beta) | **`viva`** — user **`employee-experience-*`**, work time, insights, **`engage-assigned-role*`** (+ **`engage-assigned-role-member-user-*`** nested user/mailbox/errors), **`learning-*`**; tenant **`tenant-*`** (`/employeeExperience`, incl. **`tenant-community-owner-*`** / **`tenant-engagement-role-member-user-*`**); **`admin-item-insights-*`**, **`org-item-insights-*`**; **`work-hours-*`**; **`meeting-engage-*`** (conversations, messages, replies, reactions) — scopes in **`docs/GRAPH_SCOPES.md`** |
| Microsoft 365 Copilot (Graph) | **`copilot`** — `retrieval`, `search`, `search-next`, `conversation-create`, `chat`, `chat-stream`, `interactions-export`, `meeting-insights-list`, `meeting-insight-get`, `reports` (summary/trend/usage-user-detail), `packages` (list/get/update/block/unblock/reassign), `notify-help`; **`subscribe copilot-interactions`** for [AI interaction change notifications](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/change-notifications/aiinteraction-changenotifications) (see [Copilot APIs overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/copilot-apis-overview)) |
| Graph mail extras | `rules` (inbox message rules), `oof` (automatic replies), `todo` (Microsoft To Do), **`approvals`** (list/get/steps/respond/**`cancel`** — beta `/me/approvals`) |
| EWS admin-style | `delegates`, `auto-reply` |
| Push | `subscribe`, `subscriptions` (incl. **`subscriptions renew-all`**) |
| Other | `login`, `whoami`, `verify-token`, `serve` |

## EWS writes (mail/calendar)

- Mutating EWS calls (reply, forward, move, flag/read state, drafts send, calendar respond/cancel/delete/update, **message categories**, etc.) are implemented in **`ews-client.ts`** to resolve **ItemId + ChangeKey** via **`GetItem`** / **`getCalendarEvent`** before **`CreateItem`**, **`UpdateItem`**, **`MoveItem`**, **`SendItem`**, and similar—especially important for **delegated/shared mailbox** use (`--mailbox`), where Exchange may return **`ErrorChangeKeyRequiredForWriteOperations`** if ChangeKey is omitted.
- Callers pass **message or event IDs** from list/read output as today; they do **not** supply ChangeKey manually.

## Graph maintainers

- **Out of scope:** **Teams Phone / PSTN** (telephony admin, direct routing, carrier scenarios) — not wrapped or documented here; use Teams admin center or provider tooling (**[`docs/GRAPH_INVOKE_BOUNDARIES.md`](../../docs/GRAPH_INVOKE_BOUNDARIES.md)**).
- Repo docs: **[`docs/GRAPH_SCOPES.md`](../../docs/GRAPH_SCOPES.md)** (scopes), **[`docs/GRAPH_TROUBLESHOOTING.md`](../../docs/GRAPH_TROUBLESHOOTING.md)** (OData headers, `$search`, people edge cases).
- Live Graph behavior and **`$filter`** checks: use the **msgraph** Cursor skill — [graph.pm introduction](https://graph.pm/getting-started/introduction/). Install with `npx skills add merill/msgraph` if you use Cursor skills.

## Agent tips

### For agents

End users can describe intent in **natural language** (e.g. “read mail in the shared mailbox”). The **agent** maps that to the right flags: use **`--mailbox`** for **EWS** commands and **`--user`** for **Microsoft Graph** commands, according to the command’s API (see the next bullet). The end user does **not** need to know whether a call is EWS or Graph.

- Start with **list/read** commands, then use IDs from output for updates.
- If auth fails, suggest `verify-token` and re-`login`; wrong **identity** profile means wrong cache file—check `--identity`.
- **Graph vs EWS “acting as another mailbox”:** use **`--user`** (or **`--drive-id`** / **`--site-id`**) on Graph commands per **`--help`**. For **OneDrive/SharePoint drive items**, **`--user`** selects **`/users/{id}/drive`**, not the mailbox SOAP model. Use **`--mailbox <email>`** on **EWS** commands (`calendar`, `mail`, `send`, `drafts`, `respond`, `findtime`, `delegates`, …) for **shared mailboxes** via Exchange SOAP. They are **not** interchangeable. Confirm flags on that subcommand’s **`--help`**.
- **External agents / orchestration:** Cursor or other agents should map natural language to **`files`** / **`excel`** / **`word`** / **`powerpoint`** ( **`word`**/**`powerpoint`** cover the same per-item lifecycle as **`files`** for `.docx`/`.pptx`); for Word/PowerPoint **comments** or undocumented Graph paths, prefer **`graph invoke`** with a JSON body file and scopes from **`docs/GRAPH_SCOPES.md`**. For MCP-based hosts, see **`packages/m365-agent-cli-mcp`** (stdio tools wrapping a subset of CLI calls).
- If a user still sees EWS change-key or conflict errors after an update, suggest **re-fetching the item ID** (another process may have modified the message/event) and retrying.
- For **Outlook-colored categories** on mail/calendar items, use **names** that exist in **`outlook-categories list`** (or Outlook) so colors match; **Planner** and **To Do** use different label models.
- For **attachments**, use **`mail -d`**, **`send`/`drafts`/`create-event`/`update-event`** flags above; **OneDrive/SharePoint files** use **`files`**, not EWS attachments.
