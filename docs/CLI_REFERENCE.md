# CLI reference

Complete command-line reference for **m365-agent-cli**: global flags, read-only mode, calendar, email, OneDrive, Planner, SharePoint, Graph helpers, and examples.

- [README (overview)](../README.md) — install, scenarios, and doc index
- [Agent workflows](./AGENT_WORKFLOWS.md) — AI/script patterns (deltas, Teams + files, Word/PPT loop)
- [CLI scripting appendix](./CLI_SCRIPTING_APPENDIX.md) — `--json` / read-only inventory ([generated table](./CLI_SCRIPTING_INVENTORY.md))
- [Authentication](./AUTHENTICATION.md) — tokens, shared mailboxes, Graph vs EWS
- [Entra app setup](./ENTRA_SETUP.md) — register an app and permissions

---

## Options: root vs per-command

`m365-agent-cli --help` shows only options registered on the **root** program, currently:

```bash
--read-only         # Run in read-only mode, blocking mutating operations
--dry-run           # Preview the resolved request for a mutating command without sending it
--cache <duration>  # Cache idempotent Graph GET responses on disk (e.g. 30s, 5m, 1h). Off by default.
--version, -V       # CLI version (semver from package)
```

Many **subcommands** accept their own flags. Common patterns (not every command has every flag—use `m365-agent-cli <command> --help`):

```bash
--json                    # Machine-readable output (where supported)
--token <token>           # Use a specific access token (overrides cache)
--identity <name>         # Token cache profile (default: default). Selects EWS and Graph cache files for that name.
--user <email>            # Graph delegation: target another user/mailbox (supported commands only; needs permissions)
```

**EWS shared mailbox:** use `--mailbox <email>` on calendar, mail, send, folders, drafts, respond, findtime, delegates, auto-reply, and related flows.

### Structured `--json` errors

When a command fails under `--json`, the error is a structured object (mirroring Microsoft
Graph's own `{ "error": { "code", "message" } }` shape) instead of a bare string:

```json
{
  "error": {
    "message": "The specified object was not found in the store.",
    "code": "ErrorItemNotFound",
    "status": 404,
    "retriable": false,
    "requestId": "a1b2c3d4-..."
  }
}
```

Only `message` is guaranteed present — `code`/`status`/`retriable`/`requestId` are included when
the underlying Graph/EWS error carried them. `retriable: true` marks throttling/service-unavailable
errors (HTTP 429/502/503/504, or a `tooManyRequests`/`serviceNotAvailable` error code) — a signal
scripts/agents can use to decide whether to back off and retry.

### Read-Only Mode

When read-only mode is on (`READ_ONLY_MODE=true` in env / `~/.config/m365-agent-cli/.env`, or `--read-only` on the **root** command), the CLI calls `checkReadOnly()` before the listed actions and exits with an error **before** the mutating request is sent.

The table below matches **`checkReadOnly` in the source** (search the repo for `checkReadOnly(` to verify after changes). Anything **not** listed here is either read-only or not wired to read-only yet.

| Command | Blocked actions (read-only on) |
| --- | --- |
| `create-event` | Entire command |
| `calendar` | `create` subcommand only (alias for `create-event`; **`list`** and the default no-subcommand path are read-only) |
| `update-event` | Entire command |
| `delete-event` | Entire command |
| `forward-event` | Entire command |
| `counter` | Entire command |
| `respond` | `accept`, `decline`, `tentative` (not `respond list`) |
| `send` | Entire command |
| `mail` | Mutating flags only: `--flag`, `--unflag`, `--mark-read`, `--mark-unread`, `--complete`, `--sensitivity`, `--move`, `--reply`, `--reply-all`, `--forward`, `--set-categories`, `--clear-categories` (listing, `--read`, `--download` stay allowed) |
| `drafts` | `--create`, `--edit`, `--send`, `--delete` (plain list/read allowed) |
| `folders` | `--create`, `--rename` (with `--to`), `--delete` (listing folders allowed) |
| `files` | `upload`, `upload-large`, `delete`, `share` (including `--collab`), `invite`, `permission-remove`, `permission-update`, `copy`, `move`, `restore`, `checkout`, `checkin` (read/query: **`thumbnails`**, **`delta`**, **`shared-with-me`**, …) |
| `planner` | `create-task`, `update-task`, `delete-task`, **`bulk-complete-task`**, `create-plan` (`--group` or beta `--roster`), `update-plan`, `delete-plan`, `delete-plan-details`, `delete-task-details`, `plan-archive`, `plan-unarchive` (beta), `create-bucket`, `update-bucket`, `delete-bucket`, `list-user-tasks`, `list-user-plans`, `update-task-details`, `update-plan-details`, `add-checklist-item`, `update-checklist-item`, `remove-checklist-item`, `add-reference`, `remove-reference`, `update-task-board`, `add-favorite`, `remove-favorite`, `roster` (beta: `create`, `get`, `list-members`, `add-member`, `remove-member`) |
| `sharepoint` / `sp` | `create-item`, `update-item`, `delete-item`, `follow`, `unfollow`, `site-permission-update`, `site-permission-create`, `site-permission-delete` |
| `pages` | `update`, `publish` |
| `rules` | `create`, `update`, `delete` |
| `todo` | `create`, `update`, `complete`, `delete`, **`bulk-complete`**, **`bulk-delete`**, `add-checklist`, `update-checklist`, `delete-checklist`, `get-checklist-item`, `create-list`, `update-list`, `delete-list`, `add-attachment`, `get-attachment`, `download-attachment`, `delete-attachment`, `add-reference-attachment`, `add-linked-resource`, `remove-linked-resource`, `upload-attachment-large`, `attachment-session` (patch/delete/content-put/content-delete), `root` (patch/delete), `linked-resource` (`create`, `update`, `delete`), `extension` (`set`, `update`, `delete`), `list-extension` (`set`, `update`, `delete`) |
| `subscribe` | Creating a subscription; `subscribe cancel <id>` |
| `delegates` | `add`, `update`, `remove` |
| `oof` | Write path only (when `--status`, `--internal-message`, `--external-message`, `--start`, or `--end` is used to change settings) |
| `auto-reply` | Entire command (EWS auto-reply rules) |
| `outlook-categories` | `create`, `update`, `delete` (not `list`) |
| `outlook-graph` | `create-folder`, `update-folder`, `delete-folder`, `send-mail`, `patch-message`, `delete-message`, `move-message`, `copy-message`, `create-reply`, `create-reply-all`, `create-forward`, `send-message`, `create-contact`, `update-contact`, `delete-contact` |
| `graph-calendar` | `accept`, `decline`, `tentative`, `create-calendar-group`, `delete-calendar-group`, `create-calendar`, `update-calendar`, `delete-calendar` |
| `mailbox-settings` | `set` (root command is read-only GET) |
| `contacts` | Mutating: `folder` (create/update/delete), `create` / `update` / `delete`, `photo` (set/delete), `attachments` (add file, **add-link**, delete), `extension` (set/update/delete), `merge-suggestions` (`set`, `delete`) |
| `onenote` | Mutating: notebook/section-group/section create/update/delete, `create-page`, `delete-page`, `patch-page-content`, `copy-page`, `section copy-to-notebook`, `section copy-to-section-group` (read helpers such as `notebook from-web-url`, list/get/page-preview are not gated) |
| `meeting` | `create`, `update`, `delete` (read-only paths: `recordings`, `recording-download`, `recordings-all`, `transcripts`, `transcript-download`, `transcripts-all`) |
| `excel` | `worksheet-add`, `worksheet-update`, `worksheet-delete`, `range-patch`, `range-clear`, `table-add`, `table-patch`, `table-delete`, `table-rows-add`, `table-row-patch`, `table-row-delete`, `table-column-patch`, `pivot-table-create`, `pivot-table-patch`, `pivot-table-delete`, `pivot-table-refresh`, `pivot-tables-refresh-all`, `chart-create`, `chart-patch`, `chart-delete`, `application-calculate`, `session-create`, `session-refresh`, `session-close`, `comments-create`, `comments-reply`, `comments-patch` |
| `bookings` | `business-create`, `business-delete`, `business-publish`, `business-unpublish`, `business-update`, `appointment-create`, `appointment-update`, `appointment-delete`, `appointment-cancel`, `customer-create`, `customer-update`, `customer-delete`, `service-create`, `service-update`, `service-delete`, `staff-create`, `staff-update`, `staff-delete`, `custom-question-create`, `custom-question-update`, `custom-question-delete` |
| `teams` | `activity-notify`, `channel-message-send`, `channel-message-reply`, `channel-message-patch`, `channel-message-delete`, `chat-message-send`, `chat-message-reply`, `chat-message-patch`, `chat-message-reply-patch`, `chat-message-delete`, `chat-create`, `chat-member-add`, `team-member-add`, `channel-member-add`, `tab-create`, `tab-update`, `tab-delete`, `app-add`, `app-patch`, `app-upgrade`, `app-delete`, `chat-app-add`, `chat-app-patch`, `chat-app-upgrade`, `chat-app-delete`, `user-app-add`, `user-app-delete` |
| `copilot` | `conversation-create`, `chat`, `chat-stream`, `packages update`, `packages block`, `packages unblock`, `packages reassign` |
| `presence` | `set-me`, `set-user`, `clear-me`, `clear-user`, `status-message-set`, `preferred-set`, `preferred-clear`, `clear-location` |
| `groups` | `post-reply` |
| `approvals` | `respond`, `cancel` |
| `viva` | Every **`viva`** subcommand that issues Graph **POST**, **PATCH**, or **DELETE** on **beta** (user-scoped, **`tenant-*`**, **`admin-item-insights-*`**, **`org-item-insights-*`**, **`work-hours-*`**, **`meeting-engage-*`**). Pure **GET** / list helpers are not gated; see **`m365-agent-cli viva --help`** for names. |

**Intentionally not gated** (no `checkReadOnly` today): read/query helpers such as `schedule`, `suggest`, `findtime`, `calendar`, `graph-calendar list-calendars` / `get-calendar` / `list-calendar-groups` / `list-view` / `get-event` / `events-delta`, `mailbox-settings` (root GET), `outlook-graph list-mail` / `list-messages` / `list-message-attachments` / `get-message-attachment` / `download-message-attachment` / `get-message` / `list-folders` / `list-contacts` / `get-contact` / `get-folder`, `subscriptions list`, `rules list` / `rules get`, `todo` list-only usage, **`outlook-categories list`** (mutating `outlook-categories create|update|delete` **are** gated), `files` list/search/delta/meta/download/convert/analytics/versions/`shared-with-me`/`permissions`, etc. Those calls do not use the guard in code; if a new subcommand adds writes, it should call `checkReadOnly` and this table should be updated.

You can enable Read-Only mode in two ways:

1. **Global flag**: Pass **`--read-only` immediately after** `m365-agent-cli` (before the subcommand). Commander treats this as a root option; placing it after the subcommand will not enable read-only mode.

   ```bash
   m365-agent-cli --read-only create-event "Test" 09:00 10:00
   # Error: Command blocked. The CLI is running in read-only mode.
   ```

2. **Environment variable**: Set `READ_ONLY_MODE=true` in your environment or `~/.config/m365-agent-cli/.env` file.

   ```bash
   export READ_ONLY_MODE=true
   m365-agent-cli planner update-task <taskId> --title "New"
   # Error: Command blocked. The CLI is running in read-only mode.
   ```

### Dry-Run Mode

`--dry-run` (root flag — works before or after the subcommand; or `M365_DRY_RUN=1` in the
environment) previews the exact request a mutating command would send, without sending it.
Unlike `--read-only` (which blocks before any client call), `--dry-run` lets read-only lookups
the command needs (e.g. resolving a message's current `ChangeKey`) go through as normal, and only
intercepts the actual mutating write, right at the transport layer:

- **Graph**: prints `{ dryRun: true, backend: "graph", method, url, headers, body }` — the exact
  resolved HTTP method, URL, and JSON body — instead of calling `fetch`.
- **EWS**: prints `{ dryRun: true, backend: "ews", operation, endpoint, mailbox, envelope }` — the
  resolved SOAP operation name and full envelope XML — instead of POSTing it.

```bash
m365-agent-cli mail --reply msg-123 --message "Thanks!" --cc alice@contoso.com --dry-run
# { "dryRun": true, "backend": "graph", "method": "POST", "url": "https://graph.microsoft.com/v1.0/me/messages/msg-123/createReply", "body": { "comment": "Thanks!" } }
```

For a multi-step command (e.g. create a draft, add an attachment, then send), only the **first**
mutating request is shown — the CLI exits immediately after printing it, exactly like a real run
would if that first call failed, so nothing downstream can execute against state that was never
actually written. Re-run without `--dry-run` to send it, or run again after fixing the previewed
request. GET-only commands are unaffected by `--dry-run` (there's nothing to preview).

### Read Cache

`--cache <duration>` (root flag — works before or after the subcommand; or `M365_CACHE_TTL=<duration>`
in the environment) opt-in caches successful Microsoft Graph **GET** responses on disk for
`<duration>`, so repeated agent calls to the same read (folder lists, room lists, user lookups,
etc.) within that window skip the network round trip. Off by default — without `--cache`, every
run hits Graph fresh, same as today.

`<duration>` accepts a bare number (seconds), or a number with a unit suffix: `30s`, `5m`, `2h`, `1d`.

```bash
m365-agent-cli folders --cache 5m
# first call hits Graph and populates the cache; a second identical call within 5 minutes
# returns the cached response without a network request
```

Only `GET` requests are cached — mutating requests (POST/PATCH/PUT/DELETE) always go straight to
Graph. The cache lives at `~/.config/m365-agent-cli/graph-cache/` (or under `XDG_CONFIG_HOME`),
keyed by a hash of the bearer token, HTTP method, and full URL, so two different signed-in
identities on the same machine never share cache entries, and different query parameters (e.g.
`$top`, `$filter`) get separate entries. Entries past their TTL are treated as a miss and pruned
opportunistically. EWS requests are not cached (EWS operations are POST/SOAP, not idempotent GETs
in the way Graph's REST API is).

### Output Shaping (`--fields` / `--ndjson`)

On commands that list many rows, `--json` combined with `--fields <dot-paths>` and/or `--ndjson`
lets an agent shape and stream output instead of buffering a large pretty-printed array:

- **`--fields "id,subject,from.emailAddress.address"`** — projects each row down to only the
  listed dot-paths (nested paths keep their nesting; e.g. `from.emailAddress.address` produces
  `{ "from": { "emailAddress": { "address": "..." } } }`). A path missing on a given row is
  silently omitted from that row rather than erroring — Graph and EWS payloads aren't uniform.
- **`--ndjson`** — prints one compact JSON object per row, one per line, instead of a single
  pretty-printed `{ ... : [ ... ] }` array, so a large list can be parsed line-by-line as it's
  produced rather than requiring the whole response to be buffered and parsed at once.

```bash
m365-agent-cli mail inbox --json --fields "id,subject,from.emailAddress.address" --ndjson --limit 200
# {"id":"...","subject":"...","from":{"emailAddress":{"address":"..."}}}
# {"id":"...","subject":"...","from":{"emailAddress":{"address":"..."}}}
# ...
```

This is a **per-command opt-in**, not a global flag — unlike `--dry-run`/`--cache` (which hook
into the single Graph/EWS transport layer and so apply uniformly to every command), output
shaping happens after each command's own data-fetching and formatting logic, so each command
wires it in explicitly. Today `mail` (list view, both the EWS and Graph backends) supports
`--fields`/`--ndjson`; other high-volume list commands are natural candidates for the same
`src/lib/output-shape.ts` helper as a follow-up.

---

## Calendar Commands

### View Calendar

**Subcommands:** **`calendar list`** is an explicit alias for the default listing behavior (`calendar` with no subcommand still lists). **`calendar create`** is an alias for top-level **`create-event`** (same arguments and options). Use **`calendar list --help`** and **`calendar create --help`** for flags.

```bash
# Today's events
m365-agent-cli calendar
# Equivalent explicit form:
m365-agent-cli calendar list

# Specific day
m365-agent-cli calendar tomorrow
m365-agent-cli calendar monday
m365-agent-cli calendar 2024-02-15

# Date ranges
m365-agent-cli calendar monday friday
m365-agent-cli calendar 2024-02-15 2024-02-20

# Week views
m365-agent-cli calendar week          # This week (Mon-Sun)
m365-agent-cli calendar lastweek
m365-agent-cli calendar nextweek

# Include details (attendees, body preview, categories)
m365-agent-cli calendar -v
m365-agent-cli calendar week --verbose

# Shared mailbox calendar
m365-agent-cli calendar --mailbox shared@company.com
m365-agent-cli calendar nextweek --mailbox shared@company.com

# Non-default calendar (Graph calendar id from `graph-calendar list-calendars`; Graph path only)
m365-agent-cli calendar week --calendar <calendarId>
m365-agent-cli calendar today --calendar <calendarId> --mailbox shared@company.com
```

### Calendar: rolling ranges and business (weekday) windows

Besides a single day or `start end` date range, **`calendar`** supports **one** of these span modes (not combinable with each other or with an `[end]` argument; not with `week` / `lastweek` / `nextweek`):

```bash
# Next 5 calendar days starting today (includes today)
m365-agent-cli calendar today --days 5

# Previous 3 calendar days ending on today
m365-agent-cli calendar today --previous-days 3

# Next 10 weekdays (Mon–Fri) starting from the anchor day
# (if anchor is Sat/Sun, counting starts from the next Monday)
m365-agent-cli calendar today --business-days 10

# Same as --business-days (readable alias)
m365-agent-cli calendar today --next-business-days 5

# Typo-tolerant alias for --business-days
m365-agent-cli calendar today --busness-days 5

# 5 weekdays backward ending on or before the anchor
m365-agent-cli calendar today --previous-business-days 5

# From the current time onward within the range (hide meetings that already ended today)
m365-agent-cli calendar today --now
m365-agent-cli calendar today --business-days 5 --now
```

### Create Events

```bash
# Basic event
m365-agent-cli create-event "Team Standup" 09:00 09:30

# With options
m365-agent-cli create-event "Project Review" 14:00 15:00 \
  --day tomorrow \
  --description "Q1 review meeting" \
  --attendees "alice@company.com,bob@company.com" \
  --teams \
  --room "Conference Room A"

# Specify a timezone explicitly
m365-agent-cli create-event "Global Sync" 09:00 10:00 --timezone "Pacific Standard Time"

# All-day event with category and sensitivity
m365-agent-cli create-event "Holiday" --all-day --category "Personal" --sensitivity private

# Find an available room automatically
m365-agent-cli create-event "Workshop" 10:00 12:00 --find-room

# List available rooms
m365-agent-cli create-event "x" 10:00 11:00 --list-rooms

# Create in shared mailbox calendar
m365-agent-cli create-event "Team Standup" 09:00 09:30 --mailbox shared@company.com

# Create on a secondary calendar (Graph only; calendar id from `graph-calendar list-calendars`)
m365-agent-cli create-event "Side project sync" 15:00 15:30 --calendar <calendarId>
```

### Recurring Events

```bash
# Daily standup
m365-agent-cli create-event "Daily Standup" 09:00 09:15 --repeat daily

# Weekly on specific days
m365-agent-cli create-event "Team Sync" 14:00 15:00 \
  --repeat weekly \
  --days mon,wed,fri

# Monthly, 10 occurrences
m365-agent-cli create-event "Monthly Review" 10:00 11:00 \
  --repeat monthly \
  --count 10

# Every 2 weeks until a date
m365-agent-cli create-event "Sprint Planning" 09:00 11:00 \
  --repeat weekly \
  --every 2 \
  --until 2024-12-31
```

### Update Events

```bash
# List today's events
m365-agent-cli update-event

# Update by event ID
m365-agent-cli update-event --id <eventId> --title "New Title"
m365-agent-cli update-event --id <eventId> --start 10:00 --end 11:00
m365-agent-cli update-event --id <eventId> --add-attendee "new@company.com"
m365-agent-cli update-event --id <eventId> --room "Room B"
m365-agent-cli update-event --id <eventId> --location "Off-site"
m365-agent-cli update-event --id <eventId> --teams        # Add Teams meeting
m365-agent-cli update-event --id <eventId> --no-teams      # Remove Teams meeting
m365-agent-cli update-event --id <eventId> --all-day       # Make all-day
m365-agent-cli update-event --id <eventId> --sensitivity private
m365-agent-cli update-event --id <eventId> --category "Project A" --category Review
m365-agent-cli update-event --id <eventId> --clear-categories

# Show events from a specific day
m365-agent-cli update-event --day tomorrow

# Update event in shared mailbox calendar
m365-agent-cli update-event --id <eventId> --title "Updated Title" --mailbox shared@company.com
```

### Delete/Cancel Events

**Scopes:** `--scope all` (default: entire series or single meeting), **`--scope this`** (one occurrence of a recurring series), **`--scope future`** (this occurrence and all later — **Graph:** truncates recurrence on the series master via `…/instances` + PATCH; **EWS:** SOAP `deleteEvent`). Use **`--occurrence N`** or **`--instance YYYY-MM-DD`** with a recurring master/occurrence id when you need a specific occurrence.

```bash
# List your events
m365-agent-cli delete-event

# Delete event by ID
m365-agent-cli delete-event --id <eventId>

# Recurring: only this occurrence, or this and all future
m365-agent-cli delete-event --id <eventId> --scope this
m365-agent-cli delete-event --id <eventId> --scope future

# With cancellation message
m365-agent-cli delete-event --id <eventId> --message "Sorry, need to reschedule"

# Force delete without sending cancellation
m365-agent-cli delete-event --id <eventId> --force-delete

# Search for events by title
m365-agent-cli delete-event --search "standup"

# Delete event in shared mailbox calendar
m365-agent-cli delete-event --id <eventId> --mailbox shared@company.com
```

### Respond to Invitations

```bash
# List events needing response
m365-agent-cli respond

# Accept/decline/tentative by event ID
m365-agent-cli respond accept --id <eventId>
m365-agent-cli respond decline --id <eventId> --comment "Conflict with another meeting"
m365-agent-cli respond tentative --id <eventId>

# Don't send response to organizer
m365-agent-cli respond accept --id <eventId> --no-notify

# Only show required invitations (exclude optional)
m365-agent-cli respond list --only-required

# Respond to invitation in shared mailbox calendar
m365-agent-cli respond accept --id <eventId> --mailbox shared@company.com
```

### Find Meeting Times

```bash
# Find free slots next week for yourself and others
m365-agent-cli findtime nextweek alice@company.com bob@company.com

# Specific date range (keywords or YYYY-MM-DD)
m365-agent-cli findtime monday friday alice@company.com
m365-agent-cli findtime 2026-04-01 2026-04-03 alice@company.com

# Custom duration and working hours
m365-agent-cli findtime nextweek alice@company.com --duration 60 --start 10 --end 16

# Only check specified people (exclude yourself from availability check)
m365-agent-cli findtime nextweek alice@company.com --solo
```

---

## Email Commands

### List & Read Email

```bash
# Inbox (default)
m365-agent-cli mail

# Other folders
m365-agent-cli mail sent
m365-agent-cli mail drafts
m365-agent-cli mail deleted
m365-agent-cli mail archive

# Pagination
m365-agent-cli mail -n 20           # Show 20 emails
m365-agent-cli mail -p 2            # Page 2

# Filters
m365-agent-cli mail --unread        # Only unread
m365-agent-cli mail --flagged       # Only flagged
m365-agent-cli mail -s "invoice"    # Search

# Read an email
m365-agent-cli mail -r 1            # Read email #1

# Download attachments
m365-agent-cli mail -d 3            # Download from email #3
m365-agent-cli mail -d 3 -o ~/Downloads

# Shared mailbox inbox
m365-agent-cli mail --mailbox shared@company.com
```

### Categories on mail (Outlook)

Messages use **category name strings** (same as Outlook). **Colors** come from the mailbox **master category list**, not from a separate field per message. List master categories (names + preset color) via Graph:

```bash
m365-agent-cli outlook-categories list
m365-agent-cli outlook-categories list --user colleague@company.com   # delegation, if permitted
```

Manage the master list (names + **preset** colors `preset0`..`preset24`; requires Graph **`MailboxSettings.ReadWrite`**):

```bash
m365-agent-cli outlook-categories create --name "Project A" --color preset9
m365-agent-cli outlook-categories update --id <categoryGuid> --name "Project A (new)" --color preset12
m365-agent-cli outlook-categories delete --id <categoryGuid>
```

Set or clear categories on a message **by ID** from list/read output:

```bash
m365-agent-cli mail --set-categories <messageId> --category Work --category "Follow up"
m365-agent-cli mail --clear-categories <messageId>
```

Category names appear in **`mail`** list (text and JSON) and when reading a message (`-r`).

### Send Email

```bash
# Simple email (--body is optional)
m365-agent-cli send \
  --to "recipient@example.com" \
  --subject "Hello"

# With body
m365-agent-cli send \
  --to "recipient@example.com" \
  --subject "Hello" \
  --body "This is the message body"

# Multiple recipients, CC, BCC
m365-agent-cli send \
  --to "alice@example.com,bob@example.com" \
  --cc "manager@example.com" \
  --bcc "archive@example.com" \
  --subject "Team Update" \
  --body "..."

# With markdown formatting
m365-agent-cli send \
  --to "user@example.com" \
  --subject "Update" \
  --body "**Bold text** and a [link](https://example.com)" \
  --markdown

# With attachments
m365-agent-cli send \
  --to "user@example.com" \
  --subject "Report" \
  --body "Please find attached." \
  --attach "report.pdf,data.xlsx"

# Send from shared mailbox
m365-agent-cli send \
  --to "recipient@example.com" \
  --subject "From shared mailbox" \
  --body "..." \
  --mailbox shared@company.com
```

### Reply & Forward

```bash
# Reply to an email
m365-agent-cli mail --reply 1 --message "Thanks for your email!"

# Reply all
m365-agent-cli mail --reply-all 1 --message "Thanks everyone!"

# Reply with markdown
m365-agent-cli mail --reply 1 --message "**Got it!** Will do." --markdown

# Reply (or reply-all) with additional CC / BCC recipients (comma-separated)
m365-agent-cli mail --reply 1 --message "Looping in" --cc "manager@example.com"
m365-agent-cli mail --reply-all 1 --message "For the record" --bcc "archive@example.com"
m365-agent-cli mail --reply 1 --message "..." --cc "a@example.com,b@example.com" --bcc "audit@example.com"

# Save reply as draft instead of sending
m365-agent-cli mail --reply 1 --message "Draft reply" --draft

# Forward an email (uses --to-addr, not --to)
m365-agent-cli mail --forward 1 --to-addr "colleague@example.com"
m365-agent-cli mail --forward 1 --to-addr "a@example.com,b@example.com" --message "FYI"

# Forward with extra CC / BCC recipients
m365-agent-cli mail --forward 1 --to-addr "colleague@example.com" --cc "team@example.com" --bcc "archive@example.com"

# Reply or forward with file/link attachments and/or Outlook categories (draft workflow)
m365-agent-cli mail --reply <messageId> --message "See attached" --attach "report.pdf"
m365-agent-cli mail --forward <messageId> --to-addr "boss@company.com" --message "FYI" --attach-link "https://contoso.com/doc"
m365-agent-cli mail --reply <messageId> --message "Tagged" --with-category Work --with-category "Follow up"

# Forward or reply as draft (optionally with --attach / --attach-link / --with-category)
m365-agent-cli mail --forward <messageId> --to-addr "a@b.com" --draft
m365-agent-cli mail --reply <messageId> --message "Will edit later" --draft

# Reply/forward from shared mailbox
m365-agent-cli mail --reply 1 --message "..." --mailbox shared@company.com
m365-agent-cli mail --reply-all 1 --message "..." --mailbox shared@company.com
m365-agent-cli mail --forward 1 --to-addr "colleague@example.com" --mailbox shared@company.com
```

### Email Actions

```bash
# Mark as read/unread
m365-agent-cli mail --mark-read 1
m365-agent-cli mail --mark-unread 2

# Flag emails
m365-agent-cli mail --flag 1
m365-agent-cli mail --unflag 2
m365-agent-cli mail --complete 3    # Mark flag as complete
m365-agent-cli mail --flag 1 --start-date 2026-05-01 --due 2026-05-05

# Set sensitivity
m365-agent-cli mail --sensitivity <emailId> --level confidential

# Move to folder (--to here is for folder destination, not email recipient)
m365-agent-cli mail --move 1 --to archive
m365-agent-cli mail --move 2 --to deleted
m365-agent-cli mail --move 3 --to "My Custom Folder"
```

See **Categories on mail** above for `--set-categories` / `--clear-categories`.

### Manage Drafts

```bash
# List drafts
m365-agent-cli drafts

# Read a draft
m365-agent-cli drafts -r 1

# Create a draft
m365-agent-cli drafts --create \
  --to "recipient@example.com" \
  --subject "Draft Email" \
  --body "Work in progress..."

# Categories on drafts (same name strings as Outlook; see outlook-categories list)
m365-agent-cli drafts --create --to "a@b.com" --subject "Hi" --body "..." --category Work
m365-agent-cli drafts --edit <draftId> --category Review --category "Follow up"
m365-agent-cli drafts --edit <draftId> --clear-categories

# CC/BCC on drafts (create or edit; comma-separated)
m365-agent-cli drafts --create --to "a@b.com" --cc "cc@b.com" --bcc "hidden@b.com" --subject "Hi" --body "..."
m365-agent-cli drafts --edit <draftId> --bcc "hidden@b.com"

# Create with attachment
m365-agent-cli drafts --create \
  --to "user@example.com" \
  --subject "Report" \
  --body "See attached" \
  --attach "report.pdf"

# Edit a draft
m365-agent-cli drafts --edit 1 --body "Updated content"
m365-agent-cli drafts --edit 1 --subject "New Subject"

# Send a draft
m365-agent-cli drafts --send 1

# Delete a draft
m365-agent-cli drafts --delete 1
```

### Manage Folders

```bash
# List all folders
m365-agent-cli folders

# Create a folder
m365-agent-cli folders --create "Projects"

# Rename a folder
m365-agent-cli folders --rename "Projects" --to "Active Projects"

# Delete a folder
m365-agent-cli folders --delete "Old Folder"
```

---

## OneDrive / Office Online Commands

### List, Search, and Inspect Files

```bash
# List root files
m365-agent-cli files list

# List a folder by item ID
m365-agent-cli files list --folder <folderId>

# Search OneDrive
m365-agent-cli files search "budget 2026"

# Inspect metadata
m365-agent-cli files meta <fileId>

# Get file analytics
m365-agent-cli files analytics <fileId>

# File versions
m365-agent-cli files versions <fileId>
m365-agent-cli files restore <fileId> <versionId>
```

### Upload, Download, Delete, and Share

```bash
# Upload a normal file (<=250MB)
m365-agent-cli files upload ./report.docx

# Upload to a specific folder
m365-agent-cli files upload ./report.docx --folder <folderId>

# Upload a large file (>250MB, up to 4GB via chunked upload)
m365-agent-cli files upload-large ./video.mp4
m365-agent-cli files upload-large ./backup.zip --folder <folderId>

# Download a file
m365-agent-cli files download <fileId>
m365-agent-cli files download <fileId> --out ./local-copy.docx

# Convert and download (e.g., to PDF)
m365-agent-cli files convert <fileId> --format pdf --out ./converted.pdf

# Delete a file
m365-agent-cli files delete <fileId>

# Create a share link
m365-agent-cli files share <fileId> --type view --scope org
m365-agent-cli files share <fileId> --type edit --scope anonymous

# Share link with an expiration, a password, and without retaining prior inherited permissions
m365-agent-cli files share <fileId> --type view --scope anonymous --expiration 2026-01-01T00:00:00Z --password "hunter2" --no-retain-inherited-permissions
```

### Drive root, named invites, permissions, Excel workbook comments

All **`files`** subcommands (and **`excel`**, **`word`**, **`powerpoint`** where they touch drive items) accept **one** drive root selector: default **`/me/drive`**, or **`--user <upn|id>`**, **`--drive-id <id>`**, **`--site-id <id>`** (tenant default library), or **`--site-id`** + **`--library-drive-id`** for another library. Mixing multiple selectors errors.

```bash
# Another user's OneDrive (delegated Graph)
m365-agent-cli files list --user user@contoso.com
m365-agent-cli files search "Q1" --user user@contoso.com

# Explicit drive or SharePoint site library
m365-agent-cli files meta <itemId> --drive-id b!xxx
m365-agent-cli files download <itemId> --site-id contoso.sharepoint.com,abc-123-def

# Invite people (POST body per Microsoft Graph driveItem invite)
m365-agent-cli files invite <fileId> --body ./invite.json

# List, get one, or remove sharing entries on an item
m365-agent-cli files permissions <fileId>
m365-agent-cli files permission-get <fileId> <permissionId>
m365-agent-cli files permission-remove <fileId> <permissionId>

# Excel threaded comments on the workbook (Microsoft Graph beta)
m365-agent-cli excel comments-list <fileId>
m365-agent-cli excel comments-get <fileId> <commentId>
# create / reply / patch use --json-file per --help

# SharePoint list columns / listItem for a file (often 404 on personal OneDrive)
m365-agent-cli files list-item <fileId> --json

# Follow a file (OneDrive for Business)
m365-agent-cli files follow <fileId>
m365-agent-cli files unfollow <fileId>

# Microsoft Purview / MIP (JSON body per Graph assignSensitivityLabel)
m365-agent-cli files sensitivity-assign <fileId> --json-file ./mip-assign.json
m365-agent-cli files sensitivity-extract <fileId> --json

# Retention label on item
m365-agent-cli files retention-label <fileId> --json
m365-agent-cli files retention-label-remove <fileId> --if-match "<etag>"

# Permanent delete (irreversible where allowed)
m365-agent-cli files permanent-delete <fileId>
```

The same **`list-item`**, **`follow`**, **`sensitivity-*`**, **`retention-label*`**, and **`permanent-delete`** subcommands exist on **`word`** and **`powerpoint`** with identical Graph behavior.

### Delta sync, “shared with me”, copy/move, permission PATCH

```bash
# Drive item delta (root or folder); optional --state-file + --url for paging (kind: driveDelta)
m365-agent-cli files delta --state-file ./drive.delta.json
m365-agent-cli files delta --folder <folderItemId> --url "<nextLink from previous page>"

# Items shared with you (GET /me/drive/sharedWithMe only — no --user/--site-id)
m365-agent-cli files shared-with-me

# Copy / move (use --wait on copy to poll the async monitor URL)
m365-agent-cli files copy <itemId> --parent-id <folderId> --wait
m365-agent-cli files move <itemId> --parent-id <folderId>

# PATCH permission roles (body per Graph driveItem permission)
m365-agent-cli files permission-update <fileId> <permissionId> --json-file ./perm-patch.json
```

### SharePoint lists and site resolution

```bash
m365-agent-cli sharepoint resolve-site 'contoso.sharepoint.com:/sites/YourTeam'
m365-agent-cli sharepoint get-site <siteGraphId>
m365-agent-cli sharepoint drives --site-id <id>
m365-agent-cli sharepoint get-list --site-id <id> --list-id <id>
m365-agent-cli sharepoint columns --site-id <id> --list-id <id>
m365-agent-cli sharepoint items --site-id <id> --list-id <id> --top 50
m365-agent-cli sharepoint items --site-id <id> --list-id <id> --filter "fields/Title eq 'Q1'" --all-pages
m365-agent-cli sharepoint get-item --site-id <id> --list-id <id> --item-id <id>
m365-agent-cli sharepoint delete-item --site-id <id> --list-id <id> --item-id <id>
m365-agent-cli sharepoint create-item --site-id <id> --list-id <id> --json-file ./fields.json
m365-agent-cli sharepoint items-delta --site-id <id> --list-id <id> --state-file ./list.delta.json
```

### Teams channel Files folder

```bash
m365-agent-cli teams channel-files-folder <teamId> <channelId>
# Then: m365-agent-cli files list --drive-id "<driveId>" --folder "<folderItemId>"
```

### Word (.docx) / PowerPoint (.pptx) on Graph: `files` vs `word` / `powerpoint`

**`word`** and **`powerpoint`** expose **preview**, **meta**, **download**, **thumbnails**, and **mirrored** per-item verbs that call the same Graph endpoints as **`files`**: **upload**, **upload-large**, **delete**, **share**, **invite**, **permissions**, **permission-remove**, **permission-update**, **copy**, **move**, **versions**, **restore**, **checkout**, **checkin**, **convert**, **analytics**, **activities** (same **`--user`** / **`--drive-id`** / **`--site-id`** / **`--library-drive-id`**). Use **`files`** for **list**, **search**, **delta**, **shared-with-me**, and other drive-root workflows. See **`docs/GRAPH_API_GAPS.md`** (Word + PowerPoint matrices) and **`docs/WORD_POWERPOINT_EDITING.md`** (checkout, convert, OOXML round-trip).

| Need | Command |
| --- | --- |
| Thumbnails (small/medium/large URLs) | **`files thumbnails`** or **`word thumbnails`** / **`powerpoint thumbnails`** |
| List / search / sync changes | **`files list`**, **`files search`**, **`files delta`** (not on **`word`** / **`powerpoint`**) |
| Upload / delete / copy / move | **`files …`** or **`word …`** / **`powerpoint …`** (e.g. **`word upload`**, **`powerpoint copy`**) |
| Share / invite / permissions | **`files …`** or **`word …`** / **`powerpoint …`** |
| PDF or other format | **`files convert`** or **`word convert`** / **`powerpoint convert`** |
| Checkout / versions | **`word checkout`** / **`powerpoint checkout`**, **`checkin`**, **`versions`**, **`restore`** — or **`files`** equivalents |

```bash
m365-agent-cli files thumbnails <docItemId>
m365-agent-cli word thumbnails <docItemId> --json   # same Graph call
m365-agent-cli powerpoint share <slideDeckId> --collab
```

### Excel long tail and Word/PowerPoint comments (`graph invoke`)

**`excel`** wraps worksheets, ranges (read/patch/clear), used-range, **tables** (CRUD, rows, columns), **pivot tables** (list/get/create/patch/delete/refresh), **names** (workbook + worksheet scope), charts, **workbook-get**, **application-calculate**, sessions (**create** / **refresh** / **close**, optional **`--session-id`** on mutating calls), and **`excel comments-*`** (Graph **beta**). For **Excel** features still not in the CLI (e.g. **workbook images**, **shapes**, deep **`range()`** method chains), use **`graph invoke`** — confirm path and schema in Microsoft Graph docs for your API version.

**Word / PowerPoint:** use the same drive flags on **`word`** / **`powerpoint`** for preview/meta/download/thumbnails **and** for mirrored lifecycle subcommands (see table above). **Folder-level** operations stay on **`files`**. There is **no** first-class CLI for Word/PowerPoint **in-file comments** on drive items the way **`excel comments-*`** wraps **`…/workbook/comments`**; use **`graph invoke`** against current Graph docs for your scenario if available.

### Collaborative Editing via Office Online

Microsoft Graph cannot join or control a live Office Online editing session. What m365-agent-cli can do is prepare the handoff properly:

1. Find the document in OneDrive
2. Create an organization-scoped edit link (or anonymous)
3. Return the Office Online URL (`webUrl`) for the user to open
4. Optionally checkout the file first for exclusive editing workflows

```bash
# Search for the document first
m365-agent-cli files search "budget 2026.xlsx"

# Create a collaboration handoff for Word/Excel/PowerPoint Online
m365-agent-cli files share <fileId> --collab

# Same, but checkout the file first (exclusive edit lock)
m365-agent-cli files share <fileId> --collab --lock

# Explicit checkout without creating a collaboration link (pair with checkin)
m365-agent-cli files checkout <fileId>

# When the exclusive-edit workflow is done, check the file back in
m365-agent-cli files checkin <fileId> --comment "Updated Q1 numbers"
```

**Supported collaboration file types:** `.docx`, `.xlsx`, `.pptx`

Legacy Office formats such as `.doc`, `.xls`, and `.ppt` must be converted first.

**Important clarification:**

- m365-agent-cli does **not** participate in the real-time editing session
- Office Online handles the actual co-authoring once the user opens the returned URL
- m365-agent-cli handles the file lifecycle around that workflow

---

## Microsoft Planner Commands

Manage tasks and plans in Microsoft Planner. Planner uses **six label slots** per task (`category1`..`category6`); **display names** for those slots are defined in **plan details**. The CLI accepts slots as **`1`..`6`** or **`category1`..`category6`**.

```bash
# List tasks assigned to you (label names shown when plan details are available)
m365-agent-cli planner list-my-tasks

# List your plans
m365-agent-cli planner list-plans
m365-agent-cli planner list-plans -g <groupId>

# List another user's Planner tasks/plans (Graph may return 403 depending on tenant/token)
m365-agent-cli planner list-user-tasks --user <azureAdObjectId>
m365-agent-cli planner list-user-plans --user <azureAdObjectId>

# Beta: roster container (create roster → add members → create-plan --roster <rosterId>)
m365-agent-cli planner roster create
m365-agent-cli planner roster add-member -r <rosterId> --user <userId>
m365-agent-cli planner create-plan --roster <rosterId> -t "Roster plan"

# View plan structure
m365-agent-cli planner list-buckets --plan <planId>
m365-agent-cli planner list-tasks --plan <planId>

# Create and update tasks
m365-agent-cli planner create-task --plan <planId> --title "New Task" -b <bucketId>
m365-agent-cli planner create-task --plan <planId> --title "Labeled" --label 1 --label category3

m365-agent-cli planner update-task -i <taskId> --title "Updated Task" --percent 50 --assign <userId>
m365-agent-cli planner update-task -i <taskId> --label 2 --unlabel 1
m365-agent-cli planner update-task -i <taskId> --clear-labels

# Mark many tasks complete in one call (batched: GETs to fetch each @odata.etag, then PATCHes)
m365-agent-cli planner bulk-complete-task --ids <taskId1>,<taskId2>,<taskId3>

# Beta: archive / unarchive a plan (Graph requires a justification string)
m365-agent-cli planner plan-archive -p <planId> -j "Project closed"
m365-agent-cli planner plan-unarchive -p <planId> -j "Reopened for Q2"
```

---

## Microsoft To Do

To Do tasks support **string categories** (independent of Outlook mailbox master categories).

```bash
# Create with categories
m365-agent-cli todo create -t "Buy milk" --category Shopping --category Errands

# Update fields including categories (see: m365-agent-cli todo update --help)
m365-agent-cli todo update -l Tasks -t <taskId> --category Work --category Urgent
m365-agent-cli todo update -l Tasks -t <taskId> --clear-categories

# One checklist row (Graph GET checklistItems/{id}); download file attachment bytes ($value)
m365-agent-cli todo get-checklist-item -l Tasks -t <taskId> -c <checklistItemId>
m365-agent-cli todo download-attachment -l Tasks -t <taskId> -a <attachmentId> -o ./file.bin

# Incremental sync: tasks in a list (`todo delta`) vs task **lists** themselves (`todo lists-delta` + `--state-file`)
m365-agent-cli todo lists-delta --state-file ./todo-lists-sync.json

# Bulk complete / delete many tasks in one call (Graph JSON $batch, instead of one call per task)
m365-agent-cli todo bulk-complete -l Tasks --ids <taskId1>,<taskId2>,<taskId3>
m365-agent-cli todo bulk-delete -l Tasks --ids <taskId1>,<taskId2> --confirm
```

## Outlook Graph REST (`outlook-graph`)

Microsoft Graph endpoints for **mail folders**, **messages** (folder or mailbox-wide list, **sendMail**, PATCH, move, copy, attachments, reply/reply-all/forward drafts + send), and **personal contacts** (complements EWS **`mail`** / **`folders`**). Requires appropriate **Mail.** and **Contacts.** scopes.

```bash
m365-agent-cli outlook-graph list-folders
m365-agent-cli outlook-graph list-messages --folder inbox --top 25
m365-agent-cli outlook-graph list-mail --top 25
m365-agent-cli outlook-graph list-mail --search "quarterly report" --all
m365-agent-cli outlook-graph get-message -i <messageId>
m365-agent-cli outlook-graph send-mail --json-file mail.json
m365-agent-cli outlook-graph patch-message <id> --json-file patch.json
m365-agent-cli outlook-graph list-message-attachments -i <messageId>
m365-agent-cli outlook-graph download-message-attachment -i <id> -a <attId> -o ./file.bin
m365-agent-cli outlook-graph create-reply <messageId>
m365-agent-cli outlook-graph send-message <draftId>
m365-agent-cli outlook-graph list-contacts
```

## Graph calendar REST (`graph-calendar`)

Microsoft Graph endpoints for **calendars**, **calendarView** (time-range queries), **single events**, and **invitation responses** (`accept` / `decline` / `tentative`). Complements EWS **`calendar`** and **`respond`** when you need Graph IDs or REST-only flows. Requires **`Calendars.Read`** (read) or **`Calendars.ReadWrite`** (writes / responses).

```bash
m365-agent-cli graph-calendar list-calendars
m365-agent-cli graph-calendar list-calendar-groups
m365-agent-cli graph-calendar create-calendar --name "Team outings" --color preset9
m365-agent-cli graph-calendar create-calendar --name "Client A" --group-id <calendarGroupId>
m365-agent-cli graph-calendar update-calendar <calendarId> --name "Team outings (Q2)"
m365-agent-cli graph-calendar delete-calendar <calendarId>
m365-agent-cli graph-calendar list-view --start 2026-04-01T00:00:00Z --end 2026-04-08T00:00:00Z
m365-agent-cli graph-calendar list-view --start ... --end ... --calendar <calendarId>
m365-agent-cli graph-calendar get-event <eventId>
m365-agent-cli graph-calendar accept <eventId> --comment "Will attend"
```

## Microsoft Graph Search (`graph-search`)

Cross-workload search via **`POST /search/query`** (messages, events, drive items, list items, people, etc.). Uses **entity-specific** Graph delegated permissions (e.g. mail, files, calendars) — see [Microsoft Graph Search](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview) and [GRAPH_SCOPES.md](./GRAPH_SCOPES.md). This is distinct from directory **`find`** (people/groups).

Advanced **`searchRequest`** properties: **`--merge-json-file`** (object merged into the built request), **`--fields`**, **`--content-sources`**, **`--region`**, **`--aggregation-filters`**, **`--sort-json-file`** (JSON array of `sortProperty` objects), **`--enable-top-results`**, **`--trim-duplicates`**. For multi-request bodies or uncommon shapes, **`--body-file`** accepts the full JSON payload (`{ "requests": [ … ] }`) and is exclusive of query-template flags.

```bash
m365-agent-cli graph-search "project alpha"
m365-agent-cli graph-search "subject:invoice" --types message,event --size 50
m365-agent-cli graph-search "contoso" --json
m365-agent-cli graph-search "report" --fields id,lastModifiedDateTime --enable-top-results
m365-agent-cli graph-search --body-file ./search-requests.json --json-hits
```

## SharePoint Commands

Manage SharePoint lists and Site Pages.

### SharePoint Lists (`m365-agent-cli sharepoint` or `m365-agent-cli sp`)

```bash
# Site metadata and libraries (pick --library-drive-id for non-default document libraries)
m365-agent-cli sp get-site <siteGraphId>
m365-agent-cli sp drives --site-id <siteId>

# List all SharePoint lists in a site
m365-agent-cli sp lists --site-id <siteId>

# List schema / rows (items default = all pages; use --top / --filter / --url for paging)
m365-agent-cli sp get-list --site-id <siteId> --list-id <listId>
m365-agent-cli sp columns --site-id <siteId> --list-id <listId>
m365-agent-cli sp items --site-id <siteId> --list-id <listId>
m365-agent-cli sp items --site-id <siteId> --list-id <listId> --top 100 --filter "startswith(fields/Title,'Proj')"

# Create and update items
m365-agent-cli sp create-item --site-id <siteId> --list-id <listId> --fields '{"Title": "New Item"}'
m365-agent-cli sp update-item --site-id <siteId> --list-id <listId> --item-id <itemId> --fields '{"Title": "Updated Item"}'

# Site sharing permissions (owner/admin scenarios; create/delete are app-permission-only per Graph)
m365-agent-cli sp site-permissions --site-id <siteId>
m365-agent-cli sp site-permission-get --site-id <siteId> --permission-id <permissionId>
m365-agent-cli sp site-permission-update --site-id <siteId> --permission-id <permissionId> --json-file ./roles.json
m365-agent-cli sp site-permission-create --site-id <siteId> --json-file ./new-permission.json
m365-agent-cli sp site-permission-delete --site-id <siteId> --permission-id <permissionId>
```

### SharePoint Site Pages (`m365-agent-cli pages`)

```bash
# List site pages
m365-agent-cli pages list <siteId>

# Get a site page
m365-agent-cli pages get <siteId> <pageId>

# Update a site page
m365-agent-cli pages update <siteId> <pageId> --title "New Title" --name "new-name.aspx"

# Publish a site page
m365-agent-cli pages publish <siteId> <pageId>
```

---

## People & Room Search

```bash
# Search for people (relevant people + directory users) and groups
m365-agent-cli find "john"

# Search for meeting rooms by name, email, building, or tags (Places API)
m365-agent-cli find "conference" --rooms

# List relevant people (GET /me/people); optional --search, --top, --user
m365-agent-cli people list
m365-agent-cli people list --search "contoso" --top 25
m365-agent-cli people get <person-id>

# Room lists, rooms in a list, filtered search, single place
m365-agent-cli rooms lists
m365-agent-cli rooms rooms <roomListSmtp>
m365-agent-cli rooms find --query "board"
m365-agent-cli rooms get <place-id>

# Org profile and full report subtree
m365-agent-cli org user
m365-agent-cli org user someone@contoso.com
m365-agent-cli org transitive-reports --user someone@contoso.com
```

---

## Additional commands

These commands are not expanded step-by-step above; use **`m365-agent-cli <command> --help`** for flags and examples.

| Command | What it does |
| --- | --- |
| **`describe`** | Machine-readable **JSON manifest** of every command/subcommand, option, and argument — for agents/tools discovering the CLI surface programmatically instead of parsing `--help` text. **`--list`** for a fast top-level overview, **`--command "rules create"`** to scope to one (sub)command. |
| **`mcp`** | Starts a native **MCP (Model Context Protocol) stdio server**: reflects the `describe` manifest into one MCP tool per leaf command (e.g. `rules create` → tool `rules_create`), with a JSON schema built from that command's own arguments/options. A tool call runs this same CLI as a subprocess with the equivalent argv (`--json` auto-appended when the command supports it), so behavior — read-only mode, `--dry-run`, structured `--json` errors — is identical to running the command directly. `mcp`, `serve`, and `login` are not exposed as tools (self-referential / interactive / long-running). Point an MCP client at `{ "command": "m365-agent-cli", "args": ["mcp"] }`. |
| **`contacts`** | **Graph-only** Outlook contacts: folders (CRUD), list/search/delta, photo, attachments (file + **link** via `attachments add-link`), **`--user`** for delegated mailboxes ([GRAPH_SCOPES.md](./GRAPH_SCOPES.md)). |
| **`onenote`** | **Graph-only** OneNote: notebooks (incl. **resolve by web URL** — `notebook from-web-url`), section groups, sections (**copy-to-notebook**, **copy-to-section-group**), pages, HTML export/create, **patch-page-content**, **copy-page**, async **operation** poll; **`--group`** / **`--site`** roots. |
| **`meeting`** | **Graph** standalone Teams meetings (`/me/onlineMeetings`): create (simple or **`--json-file`**), get, update, delete. Calendar invitations with Teams: use **`create-event … --teams`**. |
| **`forward-event`** (`forward`) | Forward a calendar invitation to more recipients (Graph). |
| **`graph-calendar`** | Graph **calendars** (list/get/create/update/delete), **calendar groups** (list/create/delete), **calendarView**, **events-delta**, **get-event**, **accept** / **decline** / **tentative** (vs EWS `calendar` / `respond`). |
| **`mailbox-settings`** | Graph **mailboxSettings** read + **`set`** (**`--timezone`**, **`--work-days`** / **`--work-start`** / **`--work-end`**, **`--json-file`** for advanced PATCH). |
| **`graph-search`** | Microsoft Graph **Search** (`POST /search/query`) — **`--preset`** `default` \| `extended` \| `connectors` or **`--types`**; **`--merge-json-file`**, **`--body-file`**, and flags for **`fields`**, **`contentSources`**, **`region`**, **`aggregationFilters`**, **`sortProperties`** (via **`--sort-json-file`**), **`enableTopResults`**, **`trimDuplicates`** (entity-specific scopes per Graph docs). |
| **`teams`** | **Graph-only** Microsoft Teams: **list** joined teams (optional **`list --user`**), team get, **channels** / **all-channels** / **incoming-channels** / **primary-channel** / **channel-get** / **channel-files-folder**, **channel-members**, **messages** / **channel-message-get** / **channel-message-send** / **channel-message-patch** / **channel-message-delete** / **message-replies** / **channel-message-reply**, **tabs** / **tab-get** / **tab-create** / **tab-update** / **tab-delete**, **members**, **team-member-add**, **channel-member-add**, **app-catalog** / **app-catalog-get**, **apps** / **app-get** / **app-add** / **app-patch** / **app-upgrade** / **app-delete**, **chat-apps** / **chat-app-***, **user-apps** / **user-app-*** (personal scope; optional **`--user`**), **activity-notify**, **chats** / **chat-create** / **chat-member-add** / **chat-get** / **chat-messages** / **chat-message-get** / **chat-message-patch** / **chat-message-reply-patch** / **chat-message-delete** / **chat-message-replies** / **chat-message-send** / **chat-message-reply** / **chat-members** / **chat-pinned** — **chats** list is **`/me/chats`** only ([GRAPH_SCOPES.md](./GRAPH_SCOPES.md)). |
| **`org`** | **Graph-only** **user** (GET /me or GET /users/{id}), **manager**, **direct-reports**, **transitive-reports**; **`--user`** for another user’s hierarchy ([GRAPH_SCOPES.md](./GRAPH_SCOPES.md), [PERSONAL_ASSISTANT_DELEGATION.md](./PERSONAL_ASSISTANT_DELEGATION.md)). |
| **`people`** | **Graph-only** **list** / **get** on **`/me/people`** or **`/users/{id}/people`** (**`--user`**, **`--search`**, **`--top`**, **`--json`**) — [GRAPH_SCOPES.md](./GRAPH_SCOPES.md). |
| **`bookings`** | **Graph-only** Microsoft Bookings: **businesses**, **business-get** / **business-create** / **business-update** / **business-delete** / **business-publish** / **business-unpublish**, **currencies** + **currency-get**, appointments (**list**, **appointment**, **appointment-create** / **update** / **delete** / **cancel**), customers (**list**, **customer**, CRUD), **custom-questions** + **custom-question** (get) + CRUD, services + **service-get** + CRUD, staff + **staff-get** + CRUD, **calendar-view**, **staff-availability** (app-only **`--token`**). |
| **`excel`** | **Graph-only** Excel on a drive item: **worksheets** + **worksheet-get** / **add** / **update** / **delete**; **range** / **range-patch** / **range-clear**; **used-range**; **tables** / **table-get** / **table-add** / **table-patch** / **table-delete** / **table-rows** / **table-rows-add** / **table-row-patch** / **table-row-delete** / **table-columns** / **table-column-get** / **table-column-patch**; **pivot-tables** / **pivot-table-get** / **pivot-table-create** / **pivot-table-patch** / **pivot-table-delete** / **pivot-table-refresh** / **pivot-tables-refresh-all**; **names** / **name-get** / **worksheet-names** / **worksheet-name-get**; **charts** + **chart-create** / **chart-patch** / **chart-delete**; **workbook-get**; **application-calculate**; **session-create** / **session-refresh** / **session-close**; **comments-*** (beta). Same drive location flags as **`files`**. |
| **`word`** / **`powerpoint`** | Full **per-item** parity with **`files`** (incl. **list-item**, **follow**/**unfollow**, **sensitivity-assign**/**extract**, **retention-label**/**remove**, **permanent-delete**) plus **preview**/**meta**/**download**/**thumbnails**; same drive flags as **`files`**. |
| **`graph`** | **Graph-only** escape hatch: **`graph invoke`** (any JSON path/method; repeatable **`-H` / `--header "Name: value"`** for OData headers such as **`ConsistencyLevel: eventual`**) and **`graph batch`** (JSON **`$batch`** file — any number of requests; auto-chunked into **≤20**-request POSTs, sent sequentially, with `responses` merged back into one array in request order; a `dependsOn` chain must stay within one 20-request chunk); respects **`--read-only`** for non-GET. |
| **`presence`** | **Graph-only** presence: **`me`**, **`user`**, **`bulk`**, **`set-me`** / **`set-user`** (session; output includes `sessionId`), **`clear-me`** / **`clear-user`**, **`status-message-set`**, **`preferred-set`** / **`preferred-clear`**, **`clear-location`** ([GRAPH_SCOPES.md](./GRAPH_SCOPES.md)). |
| **`counter`** (`propose-new-time`) | Propose a new time for an existing event (Graph). |
| **`schedule`** | Merged free/busy for one or more people over a time window (`getSchedule`). |
| **`suggest`** | Meeting-time suggestions via Graph (`findMeetingTimes`). |
| **`rooms`** | **Graph-only** Places: **lists**, **rooms** (per room list SMTP), **find** (**`--query`**, **`--building`**, **`--capacity`**, **`--equipment`**, **`--start`/`--end`** availability), **get** (place id). |
| **`oof`** | Mailbox **automatic replies** slice of **mailboxSettings** (Graph). |
| **`mailbox-settings`** | Remaining **mailboxSettings** (time zone, working hours, formats) — see command **`--help`**. |
| **`auto-reply`** | EWS **inbox-rule** based auto-reply templates (distinct from `oof`). |
| **`rules`** | Inbox message rules (Graph). |
| **`delegates`** | Calendar/mailbox delegate permissions (EWS). |
| **`subscribe`** / **`subscriptions`** | Graph **change notifications** (create or list/cancel subscriptions). |
| **`serve`** | Local **webhook receiver** for subscription callbacks (pair with `subscribe`). |

---

## Microsoft Graph HTTP tuning (environment)

These variables apply to Graph calls routed through **`src/lib/graph-client.ts`** (most Graph-backed commands, including **`graph invoke`**).

| Variable | Purpose |
| --- | --- |
| **`GRAPH_TIMEOUT_MS`** | Per-attempt HTTP timeout in milliseconds (default **60000** if unset or invalid). |
| **`GRAPH_PAGE_DELAY_MS`** | Optional delay between **`@odata.nextLink`** page fetches in **`fetchAllPages`** (default **0**). |
| **`GRAPH_MAX_RETRIES`** | Max attempts for throttling / transient GET|HEAD network retries (default **4**, clamped to **1–8**). |
| **`GRAPH_RETRY_MAX_WAIT_MS`** | Cap for **`Retry-After`**-driven waits in milliseconds (default **60000**). |

---

## Examples

### Morning Routine Script

```bash
#!/bin/bash
echo "=== Today's Calendar ==="
m365-agent-cli calendar

# Only ongoing and upcoming (hide items that already ended today):
# m365-agent-cli calendar today --now

echo -e "=== Unread Emails ==="
m365-agent-cli mail --unread -n 5

echo -e "=== Pending Invitations ==="
m365-agent-cli respond
```

### Quick Meeting Setup

```bash
# Find a time when everyone is free and create the meeting
m365-agent-cli create-event "Project Kickoff" 14:00 15:00 \
  --day tomorrow \
  --attendees "team@company.com" \
  --teams \
  --find-room \
  --description "Initial project planning session"
```

### Email Report with Attachment

```bash
m365-agent-cli send \
  --to "manager@company.com" \
  --subject "Weekly Report - $(date +%Y-%m-%d)" \
  --body "Please find this week's report attached." \
  --attach "weekly-report.pdf"
```

### Shared Mailbox Operations

```bash
# Check shared mailbox calendar
m365-agent-cli calendar --mailbox shared@company.com

# Send email from shared mailbox
m365-agent-cli send \
  --to "team@company.com" \
  --subject "Team Update" \
  --body "..." \
  --mailbox shared@company.com

# Read shared mailbox inbox
m365-agent-cli mail --mailbox shared@company.com

# Reply from shared mailbox
m365-agent-cli mail --reply 1 --message "Done!" --mailbox shared@company.com
```

---

## Security practices

- **Graph search queries** and **markdown links** are validated/escaped where applicable
- **Date and email inputs** are validated before API calls
- **Token cache files** under `~/.config/m365-agent-cli/` are written with owner-only permissions where supported
- **String pattern replacement** avoids regex injection from malformed `$pattern` values

---

## Requirements

- [Bun](https://bun.sh) runtime
- Microsoft 365 account
- Azure AD app registration with EWS permissions (`EWS.AccessAsUser.All`)
