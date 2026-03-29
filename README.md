# Clippy

A command-line interface for Microsoft 365 using Exchange Web Services (EWS) and Microsoft Graph. Manage your calendar, email, OneDrive files, Microsoft Planner tasks, and SharePoint Sites directly from the terminal.

## Installation

```bash
# Clone the repository
git clone https://github.com/markus-lassfolk/clippy.git
cd clippy

# Install dependencies
bun install

# Run directly
bun run src/cli.ts <command>

# Or link globally
bun link
clippy <command>
```

## Authentication

Clippy uses OAuth2 with a refresh token to authenticate against Microsoft 365. You need an Azure AD app registration.

### Setup

Create a `.env` file in the project root (or set environment variables):

```bash
EWS_CLIENT_ID=your-azure-app-client-id
EWS_REFRESH_TOKEN=your-refresh-token
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

1. Clippy uses the refresh token to obtain a short-lived access token via Microsoft's OAuth2 endpoint
2. Access tokens are cached in `~/.config/clippy/token-cache.json` and refreshed automatically when expired
3. Microsoft may rotate the refresh token on each use — the latest one is cached automatically

### Verify Authentication

```bash
# Check who you're logged in as
clippy whoami

# Verify your Graph API token scopes
clippy verify-token
```

---

## Global Options

All commands support these global options:

```bash
--json              # Output as JSON (for scripting)
--token <token>     # Use a specific access token (overrides cached token)
```

---

## Calendar Commands

### View Calendar

```bash
# Today's events
clippy calendar

# Specific day
clippy calendar tomorrow
clippy calendar monday
clippy calendar 2024-02-15

# Date ranges
clippy calendar monday friday
clippy calendar 2024-02-15 2024-02-20

# Week views
clippy calendar week          # This week (Mon-Sun)
clippy calendar lastweek
clippy calendar nextweek

# Include details (attendees, body preview, categories)
clippy calendar -v
clippy calendar week --verbose

# Shared mailbox calendar
clippy calendar --mailbox shared@company.com
clippy calendar nextweek --mailbox shared@company.com
```

### Create Events

```bash
# Basic event
clippy create-event "Team Standup" 09:00 09:30

# With options
clippy create-event "Project Review" 14:00 15:00 \
  --day tomorrow \
  --description "Q1 review meeting" \
  --attendees "alice@company.com,bob@company.com" \
  --teams \
  --room "Conference Room A"

# Specify a timezone explicitly
clippy create-event "Global Sync" 09:00 10:00 --timezone "Pacific Standard Time"

# All-day event with category and sensitivity
clippy create-event "Holiday" --all-day --category "Personal" --sensitivity private

# Find an available room automatically
clippy create-event "Workshop" 10:00 12:00 --find-room

# List available rooms
clippy create-event "x" 10:00 11:00 --list-rooms

# Create in shared mailbox calendar
clippy create-event "Team Standup" 09:00 09:30 --mailbox shared@company.com
```

### Recurring Events

```bash
# Daily standup
clippy create-event "Daily Standup" 09:00 09:15 --repeat daily

# Weekly on specific days
clippy create-event "Team Sync" 14:00 15:00 \
  --repeat weekly \
  --days mon,wed,fri

# Monthly, 10 occurrences
clippy create-event "Monthly Review" 10:00 11:00 \
  --repeat monthly \
  --count 10

# Every 2 weeks until a date
clippy create-event "Sprint Planning" 09:00 11:00 \
  --repeat weekly \
  --every 2 \
  --until 2024-12-31
```

### Update Events

```bash
# List today's events
clippy update-event

# Update by event ID
clippy update-event --id <eventId> --title "New Title"
clippy update-event --id <eventId> --start 10:00 --end 11:00
clippy update-event --id <eventId> --add-attendee "new@company.com"
clippy update-event --id <eventId> --room "Room B"
clippy update-event --id <eventId> --location "Off-site"
clippy update-event --id <eventId> --teams        # Add Teams meeting
clippy update-event --id <eventId> --no-teams      # Remove Teams meeting
clippy update-event --id <eventId> --all-day       # Make all-day
clippy update-event --id <eventId> --sensitivity private

# Show events from a specific day
clippy update-event --day tomorrow

# Update event in shared mailbox calendar
clippy update-event --id <eventId> --title "Updated Title" --mailbox shared@company.com
```

### Delete/Cancel Events

```bash
# List your events
clippy delete-event

# Delete event by ID
clippy delete-event --id <eventId>

# With cancellation message
clippy delete-event --id <eventId> --message "Sorry, need to reschedule"

# Force delete without sending cancellation
clippy delete-event --id <eventId> --force-delete

# Search for events by title
clippy delete-event --search "standup"

# Delete event in shared mailbox calendar
clippy delete-event --id <eventId> --mailbox shared@company.com
```

### Respond to Invitations

```bash
# List events needing response
clippy respond

# Accept/decline/tentative by event ID
clippy respond accept --id <eventId>
clippy respond decline --id <eventId> --comment "Conflict with another meeting"
clippy respond tentative --id <eventId>

# Don't send response to organizer
clippy respond accept --id <eventId> --no-notify

# Only show required invitations (exclude optional)
clippy respond list --only-required

# Respond to invitation in shared mailbox calendar
clippy respond accept --id <eventId> --mailbox shared@company.com
```

### Find Meeting Times

```bash
# Find free slots next week for yourself and others
clippy findtime nextweek alice@company.com bob@company.com

# Specific date range (keywords or YYYY-MM-DD)
clippy findtime monday friday alice@company.com
clippy findtime 2026-04-01 2026-04-03 alice@company.com

# Custom duration and working hours
clippy findtime nextweek alice@company.com --duration 60 --start 10 --end 16

# Only check specified people (exclude yourself from availability check)
clippy findtime nextweek alice@company.com --solo
```

---

## Email Commands

### List & Read Email

```bash
# Inbox (default)
clippy mail

# Other folders
clippy mail sent
clippy mail drafts
clippy mail deleted
clippy mail archive

# Pagination
clippy mail -n 20           # Show 20 emails
clippy mail -p 2            # Page 2

# Filters
clippy mail --unread        # Only unread
clippy mail --flagged       # Only flagged
clippy mail -s "invoice"    # Search

# Read an email
clippy mail -r 1            # Read email #1

# Download attachments
clippy mail -d 3            # Download from email #3
clippy mail -d 3 -o ~/Downloads

# Shared mailbox inbox
clippy mail --mailbox shared@company.com
```

### Send Email

```bash
# Simple email (--body is optional)
clippy send \
  --to "recipient@example.com" \
  --subject "Hello"

# With body
clippy send \
  --to "recipient@example.com" \
  --subject "Hello" \
  --body "This is the message body"

# Multiple recipients, CC, BCC
clippy send \
  --to "alice@example.com,bob@example.com" \
  --cc "manager@example.com" \
  --bcc "archive@example.com" \
  --subject "Team Update" \
  --body "..."

# With markdown formatting
clippy send \
  --to "user@example.com" \
  --subject "Update" \
  --body "**Bold text** and a [link](https://example.com)" \
  --markdown

# With attachments
clippy send \
  --to "user@example.com" \
  --subject "Report" \
  --body "Please find attached." \
  --attach "report.pdf,data.xlsx"

# Send from shared mailbox
clippy send \
  --to "recipient@example.com" \
  --subject "From shared mailbox" \
  --body "..." \
  --mailbox shared@company.com
```

### Reply & Forward

```bash
# Reply to an email
clippy mail --reply 1 --message "Thanks for your email!"

# Reply all
clippy mail --reply-all 1 --message "Thanks everyone!"

# Reply with markdown
clippy mail --reply 1 --message "**Got it!** Will do." --markdown

# Save reply as draft instead of sending
clippy mail --reply 1 --message "Draft reply" --draft

# Forward an email (uses --to-addr, not --to)
clippy mail --forward 1 --to-addr "colleague@example.com"
clippy mail --forward 1 --to-addr "a@example.com,b@example.com" --message "FYI"

# Reply/forward from shared mailbox
clippy mail --reply 1 --message "..." --mailbox shared@company.com
clippy mail --reply-all 1 --message "..." --mailbox shared@company.com
clippy mail --forward 1 --to-addr "colleague@example.com" --mailbox shared@company.com
```

### Email Actions

```bash
# Mark as read/unread
clippy mail --mark-read 1
clippy mail --mark-unread 2

# Flag emails
clippy mail --flag 1
clippy mail --unflag 2
clippy mail --complete 3    # Mark flag as complete
clippy mail --flag 1 --start-date 2026-05-01 --due 2026-05-05

# Set sensitivity
clippy mail --sensitivity <emailId> --level confidential

# Move to folder (--to here is for folder destination, not email recipient)
clippy mail --move 1 --to archive
clippy mail --move 2 --to deleted
clippy mail --move 3 --to "My Custom Folder"
```

### Manage Drafts

```bash
# List drafts
clippy drafts

# Read a draft
clippy drafts -r 1

# Create a draft
clippy drafts --create \
  --to "recipient@example.com" \
  --subject "Draft Email" \
  --body "Work in progress..."

# Create with attachment
clippy drafts --create \
  --to "user@example.com" \
  --subject "Report" \
  --body "See attached" \
  --attach "report.pdf"

# Edit a draft
clippy drafts --edit 1 --body "Updated content"
clippy drafts --edit 1 --subject "New Subject"

# Send a draft
clippy drafts --send 1

# Delete a draft
clippy drafts --delete 1
```

### Manage Folders

```bash
# List all folders
clippy folders

# Create a folder
clippy folders --create "Projects"

# Rename a folder
clippy folders --rename "Projects" --to "Active Projects"

# Delete a folder
clippy folders --delete "Old Folder"
```

---

## OneDrive / Office Online Commands

### List, Search, and Inspect Files

```bash
# List root files
clippy files list

# List a folder by item ID
clippy files list --folder <folderId>

# Search OneDrive
clippy files search "budget 2026"

# Inspect metadata
clippy files meta <fileId>

# Get file analytics
clippy files analytics <fileId>

# File versions
clippy files versions <fileId>
clippy files restore <fileId> <versionId>
```

### Upload, Download, Delete, and Share

```bash
# Upload a normal file (<=250MB)
clippy files upload ./report.docx

# Upload to a specific folder
clippy files upload ./report.docx --folder <folderId>

# Upload a large file (>250MB, up to 4GB via chunked upload)
clippy files upload-large ./video.mp4
clippy files upload-large ./backup.zip --folder <folderId>

# Download a file
clippy files download <fileId>
clippy files download <fileId> --out ./local-copy.docx

# Convert and download (e.g., to PDF)
clippy files convert <fileId> --format pdf --out ./converted.pdf

# Delete a file
clippy files delete <fileId>

# Create a share link
clippy files share <fileId> --type view --scope org
clippy files share <fileId> --type edit --scope anonymous
```

### Collaborative Editing via Office Online

Microsoft Graph cannot join or control a live Office Online editing session. What Clippy can do is prepare the handoff properly:

1. Find the document in OneDrive
2. Create an organization-scoped edit link (or anonymous)
3. Return the Office Online URL (`webUrl`) for the user to open
4. Optionally checkout the file first for exclusive editing workflows

```bash
# Search for the document first
clippy files search "budget 2026.xlsx"

# Create a collaboration handoff for Word/Excel/PowerPoint Online
clippy files share <fileId> --collab

# Same, but checkout the file first (exclusive edit lock)
clippy files share <fileId> --collab --lock

# When the exclusive-edit workflow is done, check the file back in
clippy files checkin <fileId> --comment "Updated Q1 numbers"
```

**Supported collaboration file types:** `.docx`, `.xlsx`, `.pptx`

Legacy Office formats such as `.doc`, `.xls`, and `.ppt` must be converted first.

**Important clarification:**
- Clippy does **not** participate in the real-time editing session
- Office Online handles the actual co-authoring once the user opens the returned URL
- Clippy handles the file lifecycle around that workflow

---

---

## Microsoft Planner Commands

Manage tasks and plans in Microsoft Planner.

```bash
# List tasks assigned to you
clippy planner list-my-tasks

# List your plans
clippy planner list-plans
clippy planner list-plans -g <groupId>

# View plan structure
clippy planner list-buckets --plan <planId>
clippy planner list-tasks --plan <planId>

# Create and update tasks
clippy planner create-task --plan <planId> --title "New Task" -b <bucketId>
clippy planner update-task <taskId> --title "Updated Task" --percent 50 --assign <userId>
```

---

## SharePoint Commands

Manage SharePoint lists and Site Pages.

### SharePoint Lists (`clippy sharepoint` or `clippy sp`)

```bash
# List all SharePoint lists in a site
clippy sp lists --site-id <siteId>

# Get items from a list
clippy sp items --site-id <siteId> --list-id <listId>

# Create and update items
clippy sp create-item --site-id <siteId> --list-id <listId> --fields '{"Title": "New Item"}'
clippy sp update-item --site-id <siteId> --list-id <listId> --item-id <itemId> --fields '{"Title": "Updated Item"}'
```

### SharePoint Site Pages (`clippy pages`)

```bash
# List site pages
clippy pages list <siteId>

# Get a site page
clippy pages get <siteId> <pageId>

# Update a site page
clippy pages update <siteId> <pageId> --title "New Title" --name "new-name.aspx"

# Publish a site page
clippy pages publish <siteId> <pageId>
```

---
## People & Room Search

```bash
# Search for people
clippy find "john"

# Search for rooms
clippy find "conference" --rooms

# Only people (exclude rooms)
clippy find "smith" --people
```

---

## Examples

### Morning Routine Script

```bash
#!/bin/bash
echo "=== Today's Calendar ==="
clippy calendar

echo -e "\n=== Unread Emails ==="
clippy mail --unread -n 5

echo -e "\n=== Pending Invitations ==="
clippy respond
```

### Quick Meeting Setup

```bash
# Find a time when everyone is free and create the meeting
clippy create-event "Project Kickoff" 14:00 15:00 \
  --day tomorrow \
  --attendees "team@company.com" \
  --teams \
  --find-room \
  --description "Initial project planning session"
```

### Email Report with Attachment

```bash
clippy send \
  --to "manager@company.com" \
  --subject "Weekly Report - $(date +%Y-%m-%d)" \
  --body "Please find this week's report attached." \
  --attach "weekly-report.pdf"
```

### Shared Mailbox Operations

```bash
# Check shared mailbox calendar
clippy calendar --mailbox shared@company.com

# Send email from shared mailbox
clippy send \
  --to "team@company.com" \
  --subject "Team Update" \
  --body "..." \
  --mailbox shared@company.com

# Read shared mailbox inbox
clippy mail --mailbox shared@company.com

# Reply from shared mailbox
clippy mail --reply 1 --message "Done!" --mailbox shared@company.com
```

---

## Recent Security Hardening

Recent commits have strengthened input validation and API security:

- **Graph search queries** and **markdown links** are now properly escaped/validated
- **Date and email input validation** has been tightened
- **Token cache file permissions** are secured (readable only by owner)
- **String pattern replacement** bug fixed (prevents regex injection via malformed `$pattern`)

---

## Requirements

- [Bun](https://bun.sh) runtime
- Microsoft 365 account
- Azure AD app registration with EWS permissions (`EWS.AccessAsUser.All`)

## License

MIT
