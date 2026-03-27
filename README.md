# Clippy

A command-line interface for Microsoft 365 using Exchange Web Services (EWS) and Microsoft Graph. Manage your calendar, email, and OneDrive files directly from the terminal.

## Installation

```bash
# Clone the repository
git clone https://github.com/foeken/clippy.git
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

Clippy uses OAuth2 with a refresh token to authenticate against EWS. You need an Azure AD app registration with EWS permissions.

### Setup

Create a `.env` file in the project root (or set environment variables):

```bash
EWS_CLIENT_ID=your-azure-app-client-id
EWS_REFRESH_TOKEN=your-refresh-token
```

### How It Works

1. Clippy uses the refresh token to obtain a short-lived access token via Microsoft's OAuth2 endpoint
2. Access tokens are cached in `~/.config/clippy/token-cache.json` and refreshed automatically when expired
3. Microsoft may rotate the refresh token on each use — the latest one is cached automatically

### Verify Authentication

```bash
# Check who you're logged in as
clippy whoami
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

# Find an available room automatically
clippy create-event "Workshop" 10:00 12:00 --find-room

# List available rooms
clippy create-event "x" 10:00 11:00 --list-rooms
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
clippy update-event --id <eventId> --no-teams     # Remove Teams meeting

# Show events from a specific day
clippy update-event --day tomorrow
```

### Delete/Cancel Events

```bash
# List your events
clippy delete-event

# Delete event by ID
clippy delete-event --id <eventId>

# With cancellation message
clippy delete-event --id <eventId> --message "Sorry, need to reschedule"

# Force delete without notification
clippy delete-event --id <eventId> --force-delete

# Search for events by title
clippy delete-event --search "standup"
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

# Only show required invitations
clippy respond list --only-required
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

# Only check specified people (exclude yourself)
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
```

### Send Email

```bash
# Simple email
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
```

### Upload, Download, Delete, and Share

```bash
# Upload a normal file (<=250MB)
clippy files upload ./report.docx

# Create a large upload session (>250MB, <=4GB)
clippy files upload-large ./large-video.mp4

# Download a file
clippy files download <fileId>
clippy files download <fileId> --out ./local-copy.docx

# Delete a file
clippy files delete <fileId>

# Create a normal share link
clippy files share <fileId> --type view --scope org
clippy files share <fileId> --type edit --scope anonymous
```

### Collaborative Editing via Office Online

Microsoft Graph cannot join or control a live Office Online editing session. What Clippy can do is prepare the handoff properly:

1. Find the document in OneDrive
2. Create an organization-scoped edit link
3. Return the Office Online URL (`webUrl`) for the user to open
4. Optionally checkout the file first for exclusive editing workflows

```bash
# Search for the document first
clippy files search "budget 2026.xlsx"

# Create a collaboration handoff for Word/Excel/PowerPoint Online
clippy files share <fileId> --collab

# Same, but checkout the file first
clippy files share <fileId> --collab --lock

# When the exclusive-edit workflow is done, check the file back in
clippy files checkin <fileId> --comment "Updated Q1 numbers"
```

**Supported collaboration file types:**
- `.docx`
- `.xlsx`
- `.pptx`

Legacy Office formats such as `.doc`, `.xls`, and `.ppt` must be converted first.

**Important clarification:**
- Clippy does **not** participate in the real-time editing session
- Office Online handles the actual co-authoring once the user opens the returned URL
- Clippy handles the file lifecycle around that workflow

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

# Forward an email
clippy mail --forward 1 --to-addr "colleague@example.com"
clippy mail --forward 1 --to-addr "a@example.com,b@example.com" --message "FYI"
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

# Move to folder
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

## Global Options

All commands support:

```bash
--json              # Output as JSON (for scripting)
--token <token>     # Use a specific access token
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

---

## Requirements

- [Bun](https://bun.sh) runtime
- Microsoft 365 account
- Azure AD app registration with EWS permissions (`EWS.AccessAsUser.All`)

## License

MIT
