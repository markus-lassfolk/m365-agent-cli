---
name: m365-agent-cli
description: Microsoft 365 integration CLI for managing calendars, emails, OneDrive files, Planner, SharePoint, and FindTime scheduling. Use when the user requests actions involving Outlook, OneDrive, Teams tasks, or SharePoint.
metadata: {"clawdbot":{"requires":{"bins":["m365-agent-cli"]}}}
---

# m365-agent-cli Microsoft 365 CLI

A command-line tool for interacting with Microsoft 365 services.

## Core Commands

### Email (Outlook)
* `m365-agent-cli mail` - List recent emails.
* `m365-agent-cli mail [folder]` - List emails in a specific folder (e.g., `inbox`, `sent`).
* `m365-agent-cli mail --unread` - List unread emails.
* `m365-agent-cli mail --read <id>` - Read a specific email.
* `m365-agent-cli mail --flag <id>` - Flag an email.
* `m365-agent-cli drafts` - List and manage mail drafts.
* `m365-agent-cli mail --reply <id> --draft` - Create a draft reply to a specific email.

### Calendar (Outlook)
* `m365-agent-cli calendar` - View upcoming events.
* `m365-agent-cli create-event --title "Meeting" --start "YYYY-MM-DDTHH:MM:SS" --end "YYYY-MM-DDTHH:MM:SS"` - Schedule a new event.
* `m365-agent-cli findtime` - Propose meeting times using FindTime.
* `m365-agent-cli counter` - Counter-propose a meeting time.

### Tasks (Planner/To Do)
* `m365-agent-cli planner` - List Planner/To Do tasks.

### Files (OneDrive/SharePoint)
* `m365-agent-cli files` - List files in OneDrive.
* `m365-agent-cli files upload --file <local_path> --dest <remote_path>` - Upload or replace a file in-place.
* `m365-agent-cli sharepoint` - Interact with SharePoint sites and document libraries.
* `m365-agent-cli pages` - Manage SharePoint pages.

### Authentication
* `m365-agent-cli verify-token` - Check or refresh your M365 authentication token.

## Notes
* **Always verify tokens** if commands start failing with unauthorized errors.
* **Progressive Disclosure:** Start by listing items (`m365-agent-cli mail`, `m365-agent-cli files`), then drill down using specific IDs or paths.
