---
name: clippy
description: Microsoft 365 / Outlook CLI using EWS SOAP API + OAuth2. Manage calendar (view, create, update, delete events, find meeting times, respond to invitations), send/read/search email, shared mailbox support, and OneDrive file operations via Microsoft Graph.
metadata: {"clawdbot":{"requires":{"bins":["clippy"]}}}
---

# Clippy - Microsoft 365 CLI

Source: https://github.com/markus-lassfolk/clippy

Uses EWS SOAP API with OAuth2 refresh token auth. Runs on Bun.

## Install

```bash
git clone https://github.com/markus-lassfolk/clippy.git
cd clippy && bun install
```

## Auth Setup

Set these env vars (e.g. in `.env`):

```bash
EWS_CLIENT_ID=<Azure AD app client ID>
EWS_REFRESH_TOKEN=<OAuth2 refresh token>
EWS_USERNAME=<your email>
EWS_ENDPOINT=https://outlook.office365.com/EWS/Exchange.asmx
EWS_TENANT_ID=common  # or your tenant ID
```

For shared mailbox access, either set a default:
```bash
EWS_TARGET_MAILBOX=<shared@mailbox.com>
```

Or pass `--mailbox` per-command on any mail or calendar command.

**Global options** (all commands):
- `--json` — output as JSON
- `--token <token>` — use a specific access token (overrides cached token)

Check auth: `clippy whoami`

## Commands (13 total)

---

### Calendar

#### `clippy calendar [start] [end]`
View calendar events.

```
clippy calendar                        # today's events
clippy calendar tomorrow
clippy calendar monday friday
clippy calendar 2026-04-01 2026-04-03
clippy calendar week
clippy calendar nextweek
clippy calendar --verbose              # -v: show attendees + details
clippy calendar --json
clippy calendar --mailbox shared@co.com
```

#### `clippy create-event <title> <start> <end>`
Create a new calendar event.

```
clippy create-event "Meeting" 14:00 15:00 --day tomorrow --description "Notes"
clippy create-event "Workshop" 10:00 12:00 --find-room
clippy create-event "x" 10:00 11:00 --list-rooms          # list available rooms
clippy create-event "Sync" 14:00 15:00 --attendees "a@co.com,b@co.com" --teams
clippy create-event "Daily" 09:00 09:15 --repeat daily
clippy create-event "Sync" 14:00 15:00 --repeat weekly --days mon,wed,fri
clippy create-event "Monthly" 10:00 11:00 --repeat monthly --count 10
clippy create-event "Sprint" 09:00 11:00 --repeat weekly --every 2 --until 2026-12-31
clippy create-event "Team Standup" 09:00 09:30 --mailbox shared@co.com
```

Options: `--day`, `--description`, `--attendees`, `--room`, `--teams`, `--list-rooms`, `--find-room`, `--repeat` (daily|weekly|monthly|yearly), `--every`, `--days`, `--until`, `--count`, `--json`, `--token`, `--mailbox`

#### `clippy update-event [eventIndex]`
Update a calendar event by index or stable ID.

```
clippy update-event --id <eventId> --title "New Title"
clippy update-event --id <eventId> --start 10:00 --end 11:00
clippy update-event --id <eventId> --add-attendee "new@co.com"
clippy update-event --id <eventId> --room "Room B"
clippy update-event --id <eventId> --location "Off-site"
clippy update-event --id <eventId> --teams       # add Teams meeting
clippy update-event --id <eventId> --no-teams    # remove Teams meeting
clippy update-event --day tomorrow               # list events to pick from
clippy update-event --id <eventId> --mailbox shared@co.com
```

Options: `--id`, `--day`, `--title`, `--description`, `--start`, `--end`, `--add-attendee`, `--room`, `--location`, `--teams`, `--no-teams`, `--json`, `--token`, `--mailbox`

#### `clippy delete-event [eventIndex]`
Delete/cancel a calendar event.

```
clippy delete-event --id <eventId>
clippy delete-event --id <eventId> --message "Sorry, need to reschedule"
clippy delete-event --id <eventId> --force-delete   # no cancellation sent
clippy delete-event --search "standup"
clippy delete-event --day tomorrow
clippy delete-event --id <eventId> --mailbox shared@co.com
```

Options: `--id`, `--day`, `--search`, `--message`, `--force-delete`, `--json`, `--token`, `--mailbox`

#### `clippy findtime [start] [endOrEmails...]`
Find available meeting times.

```
clippy findtime nextweek alice@co.com bob@co.com
clippy findtime monday friday alice@co.com
clippy findtime 2026-04-01 2026-04-03 alice@co.com
clippy findtime nextweek alice@co.com --duration 60 --start 10 --end 16
clippy findtime nextweek alice@co.com --solo    # exclude yourself
```

Options: `--duration`, `--start`, `--end`, `--solo`, `--json`, `--token`

#### `clippy respond [action] [eventIndex]`
Respond to calendar invitations.

```
clippy respond                           # list pending invitations
clippy respond list --only-required     # only required (not optional)
clippy respond accept --id <eventId>
clippy respond decline --id <eventId> --comment "Conflict"
clippy respond tentative --id <eventId>
clippy respond accept --id <eventId> --no-notify   # don't tell organizer
clippy respond accept --id <eventId> --mailbox shared@co.com
```

Options: `--id`, `--comment`, `--no-notify`, `--include-optional`, `--only-required`, `--json`, `--token`, `--mailbox`

---

### Email

#### `clippy mail [folder]`
List and read emails.

```
clippy mail                              # inbox
clippy mail sent / drafts / deleted / archive / junk
clippy mail -n 20                        # 20 emails per page
clippy mail -p 2                         # page 2
clippy mail --unread
clippy mail --flagged
clippy mail -s "invoice"                 # search
clippy mail -r 1                         # read email #1
clippy mail -d 3 -o ~/Downloads         # download attachments from #3
clippy mail --mark-read 1
clippy mail --mark-unread 2
clippy mail --flag 1
clippy mail --unflag 2
clippy mail --complete 3                 # mark flag as complete
clippy mail --move 1 --to archive        # move to folder
clippy mail --mailbox shared@co.com      # shared mailbox inbox
```

**Reply/Forward:**
```
clippy mail --reply 1 --message "Thanks!"           # reply
clippy mail --reply-all 1 --message "Thanks all!"  # reply all
clippy mail --reply 1 --message "..." --markdown   # markdown
clippy mail --reply 1 --message "Draft" --draft    # save as draft
clippy mail --forward 1 --to-addr "colleague@co.com"  # forward (--to-addr!)
clippy mail --forward 1 --to-addr "a@co.com,b@co.com" --message "FYI"
clippy mail --reply 1 --message "..." --mailbox shared@co.com
clippy mail --reply-all 1 --message "..." --mailbox shared@co.com
clippy mail --forward 1 --to-addr "colleague@co.com" --mailbox shared@co.com
```

Options: `-n/--limit`, `-p/--page`, `--unread`, `--flagged`, `-s/--search`, `-r/--read`, `-d/--download`, `-o/--output`, `--mark-read`, `--mark-unread`, `--flag`, `--unflag`, `--complete`, `--move`, `--to`, `--reply`, `--reply-all`, `--draft`, `--forward`, `--to-addr`, `--message`, `--markdown`, `--json`, `--token`, `--mailbox`

#### `clippy send`
Send an email. **`--body` is optional** — allows sending a subject-only or even empty email.

```
clippy send --to "a@co.com" --subject "Hello"
clippy send --to "a@co.com" --subject "Hello" --body "Body text"
clippy send --to "a@co.com,b@co.com" --cc "c@co.com" --bcc "d@co.com" \
  --subject "Update" --body "..."
clippy send --to "user@co.com" --subject "Report" --body "Attached" \
  --attach "report.pdf,data.xlsx"
clippy send --to "user@co.com" --subject "With markdown" \
  --body "**Bold** and [link](https://example.com)" --markdown
clippy send --to "user@co.com" --subject "HTML" --body "<b>Bold</b>" --html
clippy send --to "user@co.com" --subject "From shared" --body "..." \
  --mailbox shared@co.com
```

Options: `--to`, `--subject`, `--body` (default: ""), `--cc`, `--bcc`, `--attach`, `--html`, `--markdown`, `--json`, `--token`, `--mailbox`

#### `clippy drafts`
Manage email drafts.

```
clippy drafts -n 10                    # list drafts
clippy drafts -r 1                      # read draft #1
clippy drafts --create --to "a@co.com" --subject "Draft" --body "WIP..."
clippy drafts --create --to "a@co.com" --subject "Report" \
  --body "See attached" --attach "report.pdf"
clippy drafts --edit 1 --body "Updated" --subject "New Subject"
clippy drafts --send 1                 # send draft
clippy drafts --delete 1              # delete draft
```

Options: `-n/--limit`, `-r/--read`, `--create`, `--edit`, `--send`, `--delete`, `--to`, `--cc`, `--subject`, `--body`, `--attach`, `--markdown`, `--html`, `--json`, `--token`

#### `clippy folders`
Manage mail folders.

```
clippy folders                        # list all folders
clippy folders --create "Projects"   # create folder
clippy folders --rename "Projects" --to "Active Projects"
clippy folders --delete "Old Folder"
```

Options: `--create`, `--rename`, `--delete`, `--to`, `--json`, `--token`

---

### Files (OneDrive via Microsoft Graph)

#### `clippy files list`
```
clippy files list                     # root files
clippy files list --folder <folderId>
```

#### `clippy files search <query>`
```
clippy files search "budget 2026"
```

#### `clippy files meta <fileId>`
```
clippy files meta <fileId>
```

#### `clippy files upload <path>`
```
clippy files upload ./report.docx
clippy files upload ./report.docx --folder <folderId>
```

#### `clippy files upload-large <path>`
```
clippy files upload-large ./large-video.mp4
```

#### `clippy files download <fileId>`
```
clippy files download <fileId>
clippy files download <fileId> --out ./local-copy.docx
```

#### `clippy files delete <fileId>`
```
clippy files delete <fileId>
```

#### `clippy files share <fileId>`
Create a sharing link or Office Online collaboration handoff.

```
clippy files share <fileId> --type view --scope org
clippy files share <fileId> --type edit --scope anonymous

# Office Online collaboration handoff (Word/Excel/PowerPoint .docx/.xlsx/.pptx)
clippy files share <fileId> --collab                    # org edit + webUrl
clippy files share <fileId> --collab --lock            # checkout first (exclusive edit)
```

Supported `--collab` types: `.docx`, `.xlsx`, `.pptx`. Legacy formats (`.doc`, `.xls`, `.ppt`) must be converted first.
Clippy does **not** join the live Office Online session — it only returns the `webUrl` for the user to open. File lifecycle (lock/checkin) is managed around that workflow.

Options: `--type` (view|edit), `--scope` (org|anonymous), `--collab`, `--lock`, `--json`, `--token`

#### `clippy files checkin <fileId>`
```
clippy files checkin <fileId> --comment "Updated Q1 numbers"
```

---

### People

#### `clippy find <query>`
```
clippy find "john"
clippy find "conference" --rooms
clippy find "smith" --people
```

Options: `--rooms`, `--people`, `--json`, `--token`

---

### Utility

#### `clippy whoami`
Show authenticated user info.
```
clippy whoami
clippy whoami --json
```

Options: `--json`, `--token`

---

## Recent Security Hardening

- Graph search queries and markdown links are properly escaped/validated
- Date and email input validation tightened
- Token cache file permissions secured (owner-only read)
- String `$pattern` replacement bug fixed (prevents injection via malformed patterns)
