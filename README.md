# m365-agent-cli

> **Credits:** This repository is heavily extended from the original project by [foeken/clippy](https://github.com/foeken/clippy).

A powerful command-line interface for Microsoft 365 using Exchange Web Services (EWS) and Microsoft Graph. Manage your calendar, email, OneDrive files, Microsoft Planner tasks, and SharePoint Sites directly from the terminal.

## The Ultimate AI Personal Assistant (PA)

The Personal Assistant (PA) playbook and skills have been moved to their own dedicated repository.

To use this tool as a highly effective Personal Assistant within OpenClaw, please visit the **[openclaw-personal-assistant](https://github.com/markus-lassfolk/openclaw-personal-assistant)** repository for the Master Guide, ecosystem requirements, and skill installation instructions.

---

## Installation

```bash
# Clone the repository
git clone https://github.com/markus-lassfolk/m365-agent-cli.git
cd m365-agent-cli

# Install dependencies
bun install
# Install OpenClaw Skills (optional, gives your AI agent superpowers)
mkdir -p ~/.openclaw/workspace/skills
cp -r skills/* ~/.openclaw/workspace/skills/


# Run directly
bun run src/cli.ts <command>

# Or link globally
bun link
m365-agent-cli <command>
```

### Published package (npm)

```bash
npm install -g m365-agent-cli
m365-agent-cli update
```

`update` checks the npm registry and runs `npm install -g m365-agent-cli@latest` (or `bun install -g` when you run the CLI with Bun). Maintainers: versioning, tags, and npm publish are documented in [docs/RELEASE.md](docs/RELEASE.md).

## Authentication

**Need help setting up the Azure AD App?** Follow our [Automated Entra ID App Setup Guide](docs/ENTRA_SETUP.md) for bash and PowerShell scripts that configure the exact permissions you need in seconds.

**EWS retirement:** Microsoft is phasing out EWS for Exchange Online in favor of Microsoft Graph. Track migration work in [docs/EWS_TO_GRAPH_MIGRATION_EPIC.md](docs/EWS_TO_GRAPH_MIGRATION_EPIC.md) (phased plan, inventory, Graph-primary + EWS-fallback strategy).

**Optional error reporting:** To receive CLI crashes and unhandled errors in [GlitchTip](https://glitchtip.com/) (Sentry-compatible), set **`GLITCHTIP_DSN`** (or **`SENTRY_DSN`**) in your environment. See [docs/GLITCHTIP.md](docs/GLITCHTIP.md).

m365-agent-cli uses OAuth2 with a refresh token to authenticate against Microsoft 365. You need an Azure AD app registration.

### Setup

If you used the setup scripts from `docs/ENTRA_SETUP.md`, your `EWS_CLIENT_ID` is already appended to your `~/.config/m365-agent-cli/.env` file.

The easiest way to obtain your refresh tokens is to run the interactive login command:

```bash
m365-agent-cli login
```

This will initiate the Microsoft Device Code flow and automatically save **`M365_REFRESH_TOKEN`** (preferred single name) plus legacy `EWS_REFRESH_TOKEN` and `GRAPH_REFRESH_TOKEN` (same value) into your `~/.config/m365-agent-cli/.env` file upon successful authentication.

Alternatively, you can manually create a `~/.config/m365-agent-cli/.env` file (or set environment variables):

```bash
EWS_CLIENT_ID=your-azure-app-client-id
M365_REFRESH_TOKEN=your-refresh-token
EWS_USERNAME=your@email.com
EWS_ENDPOINT=https://outlook.office365.com/EWS/Exchange.asmx
EWS_TENANT_ID=common  # or your tenant ID
```

### Shared Mailbox Access

To send from or access a shared mailbox, set the default in your env:

```bash
EWS_TARGET_MAILBOX=shared@company.com
```

Or pass `--mailbox` per-command (see examples below).

### How It Works

1. m365-agent-cli uses the refresh token to obtain a short-lived access token via Microsoft's OAuth2 endpoint
2. Access tokens are cached under `~/.config/m365-agent-cli/`:
   - **Unified OAuth cache:** `token-cache-{identity}.json` (default identity: `default`) holds separate **EWS** and **Microsoft Graph** access tokens and refresh token metadata. Legacy `graph-token-cache-{identity}.json` is merged on load and removed after save.
   Tokens are refreshed automatically when expired.
3. Microsoft may rotate the refresh token on each use — the latest one is cached automatically in the same directory

### Verify Authentication

```bash
# Check who you're logged in as
m365-agent-cli whoami

# Verify your Graph API token scopes
m365-agent-cli verify-token
```

---

## Options: root vs per-command

`m365-agent-cli --help` shows only options registered on the **root** program, currently:

```bash
--read-only         # Run in read-only mode, blocking mutating operations
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

### Read-Only Mode

When read-only mode is on (`READ_ONLY_MODE=true` in env / `~/.config/m365-agent-cli/.env`, or `--read-only` on the **root** command), the CLI calls `checkReadOnly()` before the listed actions and exits with an error **before** the mutating request is sent.

The table below matches **`checkReadOnly` in the source** (search the repo for `checkReadOnly(` to verify after changes). Anything **not** listed here is either read-only or not wired to read-only yet.

| Command | Blocked actions (read-only on) |
| --- | --- |
| `create-event` | Entire command |
| `update-event` | Entire command |
| `delete-event` | Entire command |
| `forward-event` | Entire command |
| `counter` | Entire command |
| `respond` | `accept`, `decline`, `tentative` (not `respond list`) |
| `send` | Entire command |
| `mail` | Mutating flags only: `--flag`, `--unflag`, `--mark-read`, `--mark-unread`, `--complete`, `--sensitivity`, `--move`, `--reply`, `--reply-all`, `--forward`, `--set-categories`, `--clear-categories` (listing, `--read`, `--download` stay allowed) |
| `drafts` | `--create`, `--edit`, `--send`, `--delete` (plain list/read allowed) |
| `folders` | `--create`, `--rename` (with `--to`), `--delete` (listing folders allowed) |
| `files` | `upload`, `upload-large`, `delete`, `share`, `restore`, `checkin` |
| `planner` | `create-task`, `update-task`, `delete-task`, `create-plan` (`--group` or beta `--roster`), `update-plan`, `delete-plan`, `create-bucket`, `update-bucket`, `delete-bucket`, `list-user-tasks`, `list-user-plans`, `update-task-details`, `update-plan-details`, `add-checklist-item`, `update-checklist-item`, `remove-checklist-item`, `add-reference`, `remove-reference`, `update-task-board`, `add-favorite`, `remove-favorite`, `roster` (beta: `create`, `get`, `list-members`, `add-member`, `remove-member`) |
| `sharepoint` / `sp` | `create-item`, `update-item` |
| `pages` | `update`, `publish` |
| `rules` | `create`, `update`, `delete` |
| `todo` | `create`, `update`, `complete`, `delete`, `add-checklist`, `update-checklist`, `delete-checklist`, `get-checklist-item`, `create-list`, `update-list`, `delete-list`, `add-attachment`, `get-attachment`, `download-attachment`, `delete-attachment`, `add-reference-attachment`, `add-linked-resource`, `remove-linked-resource`, `upload-attachment-large`, `linked-resource` (`create`, `update`, `delete`), `extension` (`set`, `update`, `delete`), `list-extension` (`set`, `update`, `delete`) |
| `subscribe` | Creating a subscription; `subscribe cancel <id>` |
| `delegates` | `add`, `update`, `remove` |
| `oof` | Write path only (when `--status`, `--internal-message`, `--external-message`, `--start`, or `--end` is used to change settings) |
| `auto-reply` | Entire command (EWS auto-reply rules) |
| `outlook-categories` | `create`, `update`, `delete` (not `list`) |
| `outlook-graph` | `create-folder`, `update-folder`, `delete-folder`, `send-mail`, `patch-message`, `delete-message`, `move-message`, `copy-message`, `create-reply`, `create-reply-all`, `create-forward`, `send-message`, `create-contact`, `update-contact`, `delete-contact` |
| `graph-calendar` | `accept`, `decline`, `tentative` |

**Intentionally not gated** (no `checkReadOnly` today): read/query helpers such as `schedule`, `suggest`, `findtime`, `calendar`, `graph-calendar list-calendars` / `get-calendar` / `list-view` / `get-event`, `outlook-graph list-mail` / `list-messages` / `list-message-attachments` / `get-message-attachment` / `download-message-attachment` / `get-message` / `list-folders` / `list-contacts` / `get-contact` / `get-folder`, `subscriptions list`, `rules list` / `rules get`, `todo` list-only usage, **`outlook-categories list`** (mutating `outlook-categories create|update|delete` **are** gated), `files` list/search/meta/download/convert/analytics/versions, etc. Those calls do not use the guard in code; if a new subcommand adds writes, it should call `checkReadOnly` and this table should be updated.

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

---

## Calendar Commands

### View Calendar

```bash
# Today's events
m365-agent-cli calendar

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

```bash
# List your events
m365-agent-cli delete-event

# Delete event by ID
m365-agent-cli delete-event --id <eventId>

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

# Save reply as draft instead of sending
m365-agent-cli mail --reply 1 --message "Draft reply" --draft

# Forward an email (uses --to-addr, not --to)
m365-agent-cli mail --forward 1 --to-addr "colleague@example.com"
m365-agent-cli mail --forward 1 --to-addr "a@example.com,b@example.com" --message "FYI"

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
```

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
m365-agent-cli graph-calendar list-view --start 2026-04-01T00:00:00Z --end 2026-04-08T00:00:00Z
m365-agent-cli graph-calendar list-view --start ... --end ... --calendar <calendarId>
m365-agent-cli graph-calendar get-event <eventId>
m365-agent-cli graph-calendar accept <eventId> --comment "Will attend"
```

## SharePoint Commands

Manage SharePoint lists and Site Pages.

### SharePoint Lists (`m365-agent-cli sharepoint` or `m365-agent-cli sp`)

```bash
# List all SharePoint lists in a site
m365-agent-cli sp lists --site-id <siteId>

# Get items from a list
m365-agent-cli sp items --site-id <siteId> --list-id <listId>

# Create and update items
m365-agent-cli sp create-item --site-id <siteId> --list-id <listId> --fields '{"Title": "New Item"}'
m365-agent-cli sp update-item --site-id <siteId> --list-id <listId> --item-id <itemId> --fields '{"Title": "Updated Item"}'
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
# Search for people
m365-agent-cli find "john"

# Search for rooms
m365-agent-cli find "conference" --rooms

# Only people (exclude rooms)
m365-agent-cli find "smith" --people
```

---

## Additional commands

These commands are not expanded step-by-step above; use **`m365-agent-cli <command> --help`** for flags and examples.

| Command | What it does |
| --- | --- |
| **`forward-event`** (`forward`) | Forward a calendar invitation to more recipients (Graph). |
| **`graph-calendar`** | Graph **calendars**, **calendarView**, **get-event**, **accept** / **decline** / **tentative** (vs EWS `calendar` / `respond`). |
| **`counter`** (`propose-new-time`) | Propose a new time for an existing event (Graph). |
| **`schedule`** | Merged free/busy for one or more people over a time window (`getSchedule`). |
| **`suggest`** | Meeting-time suggestions via Graph (`findMeetingTimes`). |
| **`rooms`** | Search and filter room resources (capacity, equipment, floor, etc.). |
| **`oof`** | Mailbox out-of-office **settings** (Graph mailbox settings API). |
| **`auto-reply`** | EWS **inbox-rule** based auto-reply templates (distinct from `oof`). |
| **`rules`** | Inbox message rules (Graph). |
| **`delegates`** | Calendar/mailbox delegate permissions (EWS). |
| **`subscribe`** / **`subscriptions`** | Graph **change notifications** (create or list/cancel subscriptions). |
| **`serve`** | Local **webhook receiver** for subscription callbacks (pair with `subscribe`). |

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

## License

MIT
