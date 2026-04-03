---
name: m365-agent-cli
version: 2.0.0-beta.0
description: Microsoft 365 CLI (EWS + Graph) for calendar, mail, OneDrive, Planner, SharePoint, To Do, inbox rules, delegates, and subscriptions. Use when the user needs Outlook/Exchange, Graph, or M365 automation from the terminal.
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
- Interactive login: `m365-agent-cli login` (device code); tokens land in `.env` / caches.
- Check session: `m365-agent-cli whoami`, `m365-agent-cli verify-token [--identity <name>]`.

## Delegation and shared access

- **EWS shared mailbox:** `--mailbox <email>` on calendar, mail, send, folders, drafts, respond, findtime, delegates, auto-reply (and similar) to act as that mailbox where supported.
- **Graph delegation:** **`--user <upn-or-id>`** on supported commands (e.g. inbox **rules**, **oof**, **todo**, **outlook-categories**, **schedule** / meeting-time helpers, **`graph-calendar`**, **`outlook-graph`**, **subscribe**, **rooms**/places, **files** where implemented) — calls Graph as `/users/{id}/...` instead of `/me/...`. Requires app permissions + admin consent where applicable.

## Safety

- **`--read-only`** (root) or **`READ_ONLY_MODE=true`** in env / `.env` runs `checkReadOnly()` before specific mutating actions (exits before the request). The **authoritative list** is the **Read-Only Mode** table in this repo’s `README.md` (kept in sync with `grep checkReadOnly src` in source).
- Read/query commands (e.g. `calendar`, `schedule`, `suggest`, **`outlook-graph list-mail` / `list-messages` / `get-message` / attachment list-get-download** / folder list-get, `subscriptions list`, `rules list`, **`outlook-categories list`**) are **not** gated unless they call `checkReadOnly`—see README. **`outlook-categories create` / `update` / `delete`** are mutating and **are** gated.
- **`m365-agent-cli --help`** only lists root flags (e.g. `--read-only`). Per-command flags are on each subcommand’s help.

## Categories, labels, and colors (Outlook vs To Do vs Planner)

- **Mail and calendar (EWS):** Items use **category name strings** (same names as in Outlook). **Colors** are not stored per message/event in the CLI; they come from the mailbox **master category list**. **List** names + preset colors: **`outlook-categories list`**. **Create / rename / delete** master entries (and colors `preset0`..`preset24`): **`outlook-categories create`**, **`update`**, **`delete`** (Graph; needs **`MailboxSettings.ReadWrite`**). Set categories on existing messages: **`mail --set-categories <id> --category <name>`**, **`mail --clear-categories <id>`**. On **outgoing** reply/forward: **`mail --with-category <name>`** (repeatable) with **`--reply` / `--reply-all` / `--forward`**. Drafts: **`drafts --create` / `--edit`** with **`--category`**; **`--clear-categories`** on edit. Calendar: **`create-event` / `update-event`** with **`--category`**; verbose **`calendar`** shows categories when present.
- **Outlook Graph REST** (distinct from EWS **`mail`** / **`folders`):** **`outlook-graph`** — mail folder **`list-folders`**, **`child-folders`**, **`get-folder`**, **`create-folder`**, **`update-folder`**, **`delete-folder`**; messages **`list-messages`** (**`--folder`**, **`--top`**, **`--all`**, **`--filter`**, **`--orderby`**, **`--select`**), **`list-mail`** (mailbox-wide **`GET /messages`**, **`--search`**, **`--skip`**, same OData flags), **`get-message`**; **`send-mail`** (**`--json-file`** `{ message, saveToSentItems }`); **`patch-message`** (**`--json-file`**); **`delete-message`** (**`--confirm`**); **`move-message`** / **`copy-message`** (**`--destination`**); **`list-message-attachments`**, **`get-message-attachment`**, **`download-message-attachment`** (**`-o`**); **`create-reply`**, **`create-reply-all`**, **`create-forward`** (**`--to`** comma-separated, **`--comment`**); **`send-message`** (draft id). Contacts **`list-contacts`**, **`get-contact`**, **`create-contact`** / **`update-contact`** (**`--json-file`**), **`delete-contact`**. Use **`--user`** for delegation where supported.
- **Graph calendar REST** (distinct from EWS **`calendar`** / **`respond`):** **`graph-calendar`** — **`list-calendars`**, **`get-calendar`**, **`list-view`** (**`--start`**, **`--end`**, optional **`--calendar`**), **`get-event`** (**`--select`**); **`accept`**, **`decline`**, **`tentative`** (**`--comment`**, **`--no-notify`**). Use **`--user`** for delegation where supported.
- **Microsoft To Do (Graph):** Tasks use **`categories[]`** as **string labels** (independent of Outlook master categories). **`todo create --category`**, **`todo update --category` / `--clear-categories`**. **`todo get`** supports OData **`--filter`**, **`--orderby`**, **`--select`**, **`--top`**, **`--skip`**, **`--expand`**, **`--count`**, or **`--status` / `--importance`**; **`--task`** (single task) passes **`--select`** to Graph. Create/update: **`--start`**, per-field time zones (**`--timezone`**, **`--due-tz`**, **`--start-tz`**, **`--reminder-tz`**), **`todo update --clear-start`**. Lists: **`todo create-list`**, **`todo update-list`**, **`todo delete-list`**. Checklist: **`todo add-checklist`**, **`todo update-checklist`**, **`todo delete-checklist`**, **`todo list-checklist-items`**, **`todo get-checklist-item`**. Recurrence: **`--recurrence-json`** / **`--clear-recurrence`** on create/update. Attachments: **`todo list-attachments`**, **`todo get-attachment`**, **`todo download-attachment`** (**`--output`**; Graph **`.../attachments/{id}/$value`** for file bytes), **`todo add-attachment`**, **`todo add-reference-attachment`**, **`todo upload-attachment-large`**, **`todo delete-attachment`**. Linked resources: Graph **`linkedResource`** (**`todo linked-resource`** `list`/`create`/`get`/`update`/`delete`), or merge on **`todo add-linked-resource`** / **`remove-linked-resource`** (**`displayName`** / **`--display-name`**). Delta sync: **`todo delta`**. Task extensions: **`todo extension list`**, **`get`**, **`set`**, **`update`**, **`delete`**. List extensions: **`todo list-extension list`**, **`get`**, **`set`**, **`update`**, **`delete`**.
- **Planner (Graph):** Tasks use **six fixed slots** `category1`..`category6` (**`appliedCategories`**). Display names are in **plan details**: **`planner get-plan-details`**, **`planner update-plan-details`** (**`--names-json`**, **`--shared-with-json`**). Tasks: **`planner create-task`** / **`update-task`** (**`--assign`**, **`--priority`** 0–10, **`--preview-type`**, **`--conversation-thread`**, **`--order-hint`**, **`--assignee-priority`**, **`--due` / `--start`**, **`--clear-due` / `--clear-start`**, label flags). **`planner get-task --with-details`**, **`planner delete-task`**. Per-user lists (Graph **`/users/{id}/planner/...`**; may **403**): **`planner list-user-tasks --user`**, **`planner list-user-plans --user`**. Plans/buckets: **`create-plan --group`** or beta **`create-plan --roster`**, **`update-plan`**, **`delete-plan`**, **`create-bucket`**, **`update-bucket`** (**`--order-hint`**), **`delete-bucket`**. Task details: **`get-task-details`**, **`update-task-details`**, **`add-checklist-item`**, **`update-checklist-item`**, **`remove-checklist-item`**, **`add-reference`**, **`remove-reference`**. Task board ordering (Graph format resources): **`get-task-board --view`** `assignedTo` \| `bucket` \| `progress`, **`update-task-board`** (**`--json-file`**). Beta roster container APIs: **`planner roster`** `create` \| `get` \| `list-members` \| `add-member` \| `remove-member`. Other beta: **`list-favorite-plans`**, **`list-roster-plans`**, **`get-me`**, **`add-favorite`**, **`remove-favorite`**, **`delta`** (use **`--url`** from **`nextLink`** or **`deltaLink`**).

## Attachments (EWS)

File and link attachments in the **EWS** flows below apply to **messages** and **calendar items**. **Microsoft To Do** file and link attachments use **Graph** via **`todo`** (see the To Do bullet above), not these EWS commands.

| Flow | Command / flags |
| ------ | ----------------- |
| Send email with files or links | **`send --attach <paths>`** (comma-separated), **`send --attach-link <spec>`** (repeatable; spec is **`Title&#124;https://url`** or a bare **`https://`** URL) |
| Drafts | **`drafts --create` / `--edit`** with **`--attach`** and **`--attach-link`** (same pattern as `send`) |
| Download from a message | **`mail -d <id>`** (or **`--download`**), **`--output <dir>`** for save location |
| Reply / forward with attachments or outgoing categories | **`mail --reply` / `--reply-all` / `--forward`** with **`--attach`**, **`--attach-link`**, **`--with-category`** (uses draft + send; **`--draft`** to save only). Use **message id** from list/read, not the numeric index, for non-interactive scripts. |
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
| Calendar | `calendar` (**`--list-attachments`**, **`--download-attachments`**), `create-event` / `update-event` (**`--attach`**, **`--attach-link`**), `delete-event`, `respond`, `findtime`, `forward-event`, `counter`, `schedule`, `suggest`, **`graph-calendar`** (Graph list/view/get + invitation responses) |
| Mail | `mail` (**`-d`**, reply/forward **`--attach`**, **`--attach-link`**, **`--with-category`**), `send`, `drafts`, `folders`; **`outlook-graph`** (Graph mail: **`list-mail`**, **`send-mail`**, **`patch-message`**, move/copy, attachments, reply/forward drafts + **`send-message`**) |
| Outlook categories (Graph) | `outlook-categories` **`list`**, **`create`**, **`update`**, **`delete`** — master list **names + colors** |
| Files | `files` (list, search, upload, download, share, versions, …) |
| Planner | `planner` (tasks, plans, buckets; **labels** on tasks) |
| SharePoint | `sharepoint` / `sp`, `pages` (site pages) |
| Directory / rooms | `find`, `rooms` |
| Graph mail extras | `rules` (inbox message rules), `oof` (automatic replies), `todo` (Microsoft To Do) |
| EWS admin-style | `delegates`, `auto-reply` |
| Push | `subscribe`, `subscriptions` |
| Other | `login`, `whoami`, `verify-token`, `serve` |

## EWS writes (mail/calendar)

- Mutating EWS calls (reply, forward, move, flag/read state, drafts send, calendar respond/cancel/delete/update, **message categories**, etc.) are implemented in **`ews-client.ts`** to resolve **ItemId + ChangeKey** via **`GetItem`** / **`getCalendarEvent`** before **`CreateItem`**, **`UpdateItem`**, **`MoveItem`**, **`SendItem`**, and similar—especially important for **delegated/shared mailbox** use (`--mailbox`), where Exchange may return **`ErrorChangeKeyRequiredForWriteOperations`** if ChangeKey is omitted.
- Callers pass **message or event IDs** from list/read output as today; they do **not** supply ChangeKey manually.

## Agent tips

### For agents

End users can describe intent in **natural language** (e.g. “read mail in the shared mailbox”). The **agent** maps that to the right flags: use **`--mailbox`** for **EWS** commands and **`--user`** for **Microsoft Graph** commands, according to the command’s API (see the next bullet). The end user does **not** need to know whether a call is EWS or Graph.

- Start with **list/read** commands, then use IDs from output for updates.
- If auth fails, suggest `verify-token` and re-`login`; wrong **identity** profile means wrong cache file—check `--identity`.
- **Graph vs EWS “acting as another mailbox”:** use **`--user <upn-or-id>`** only on commands that call **Microsoft Graph** (`outlook-graph`, `graph-calendar`, `todo`, `rules`, `oof`, `schedule`, etc.) — it switches the API path to `/users/{id}/...`. Use **`--mailbox <email>`** on **EWS** commands (`calendar`, `mail`, `send`, `drafts`, `respond`, `findtime`, `delegates`, …) for **shared mailboxes** via Exchange SOAP. They are **not** interchangeable: a Graph subcommand will not accept `--mailbox`, and typical EWS mail/calendar flows do not use `--user`. Confirm the flag exists on that subcommand’s **`--help`** before assuming delegation works.
- If a user still sees EWS change-key or conflict errors after an update, suggest **re-fetching the item ID** (another process may have modified the message/event) and retrying.
- For **Outlook-colored categories** on mail/calendar items, use **names** that exist in **`outlook-categories list`** (or Outlook) so colors match; **Planner** and **To Do** use different label models.
- For **attachments**, use **`mail -d`**, **`send`/`drafts`/`create-event`/`update-event`** flags above; **OneDrive/SharePoint files** use **`files`**, not EWS attachments.
