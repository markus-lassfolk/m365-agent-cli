---
name: clippy
description: Microsoft 365 integration CLI for managing calendars, emails, OneDrive files, Planner, SharePoint, and FindTime scheduling. Use when the user requests actions involving Outlook, OneDrive, Teams tasks, or SharePoint.
---

# Clippy Microsoft 365 CLI

A command-line tool for interacting with Microsoft 365 services.

## Core Commands

### Email (Outlook)
* `clippy mail` - List recent emails.
* `clippy mail [folder]` - List emails in a specific folder (e.g., `inbox`, `sent`).
* `clippy mail --unread` - List unread emails.
* `clippy mail --read <id>` - Read a specific email.
* `clippy mail --flag <id>` - Flag an email.
* `clippy drafts` - List and manage mail drafts.
* `clippy mail --reply <id> --draft` - Create a draft reply to a specific email.

### Calendar (Outlook)
* `clippy calendar` - View upcoming events.
* `clippy create-event --title "Meeting" --start "YYYY-MM-DDTHH:MM:SS" --end "YYYY-MM-DDTHH:MM:SS"` - Schedule a new event.
* `clippy findtime` - Propose meeting times using FindTime.
* `clippy counter` - Counter-propose a meeting time.

### Tasks (Planner/To Do)
* `clippy planner` - List Planner/To Do tasks.

### Files (OneDrive/SharePoint)
* `clippy files` - List files in OneDrive.
* `clippy files upload --file <local_path> --dest <remote_path>` - Upload or replace a file in-place.
* `clippy sharepoint` - Interact with SharePoint sites and document libraries.
* `clippy pages` - Manage SharePoint pages.

### Authentication
* `clippy verify-token` - Check or refresh your M365 authentication token.

## Notes
* **Always verify tokens** if commands start failing with unauthorized errors.
* **Progressive Disclosure:** Start by listing items (`clippy mail`, `clippy files`), then drill down using specific IDs or paths.
