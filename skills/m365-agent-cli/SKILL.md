---
name: m365-agent-cli
description: Microsoft 365 CLI (EWS + Graph) for calendar, mail, OneDrive, Planner, SharePoint, To Do, inbox rules, delegates, and subscriptions. Use when the user needs Outlook/Exchange, Graph, or M365 automation from the terminal.
metadata: {"clawdbot":{"requires":{"bins":["m365-agent-cli"]}}}
---

# m365-agent-cli

CLI for Microsoft 365: **Exchange Web Services (EWS)** and **Microsoft Graph**. Prefer `m365-agent-cli <command> --help` for exact flags on each command.

## Authentication and profiles

- Config directory: `~/.config/m365-agent-cli/` (`.env`, token caches).
- **EWS** cache: `token-cache-{identity}.json` — default identity name: `default`.
- **Graph** cache: `graph-token-cache-{identity}.json` — same identity string as EWS for that “profile”.
- **`--identity <name>`** — use a named cache profile (Graph- and EWS-backed commands that expose the flag). Default is `default`.
- **`--token <token>`** — override cached access token for that request (advanced).
- Interactive login: `m365-agent-cli login` (device code); tokens land in `.env` / caches.
- Check session: `m365-agent-cli whoami`, `m365-agent-cli verify-token [--identity <name>]`.

## Delegation and shared access

- **EWS shared mailbox:** `--mailbox <email>` on calendar, mail, send, folders, drafts, respond, findtime, delegates, auto-reply (and similar) to act as that mailbox where supported.
- **Graph delegation:** **`--user <upn-or-id>`** on supported commands (e.g. inbox **rules**, **oof**, **todo**, **schedule** / meeting-time helpers, **subscribe**, **rooms**/places, **files** where implemented) — calls Graph as `/users/{id}/...` instead of `/me/...`. Requires app permissions + admin consent where applicable.

## Safety

- **`--read-only`** (root) or **`READ_ONLY_MODE=true`** in env / `.env` runs `checkReadOnly()` before specific mutating actions (exits before the request). The **authoritative list** is the **Read-Only Mode** table in this repo’s `README.md` (kept in sync with `grep checkReadOnly src` in source).
- Read/query commands (e.g. `calendar`, `schedule`, `suggest`, `subscriptions list`, `rules list`) are **not** gated unless they call `checkReadOnly`—see README.
- **`m365-agent-cli --help`** only lists root flags (e.g. `--read-only`). Per-command flags are on each subcommand’s help.

## Command map (high level)

| Area | Commands / notes |
|------|------------------|
| Calendar | `calendar`, `create-event`, `update-event`, `delete-event`, `respond`, `findtime`, `forward-event`, `counter`, `schedule`, `suggest` |
| Mail | `mail`, `send`, `drafts`, `folders` |
| Files | `files` (list, search, upload, download, share, versions, …) |
| Planner | `planner` |
| SharePoint | `sharepoint` / `sp`, `pages` (site pages) |
| Directory / rooms | `find`, `rooms` |
| Graph mail extras | `rules` (inbox message rules), `oof` (automatic replies), `todo` (Microsoft To Do) |
| EWS admin-style | `delegates`, `auto-reply` |
| Push | `subscribe`, `subscriptions` |
| Other | `login`, `whoami`, `verify-token`, `serve` |

## Agent tips

- Start with **list/read** commands, then use IDs from output for updates.
- If auth fails, suggest `verify-token` and re-`login`; wrong **identity** profile means wrong cache file—check `--identity`.
- For “on behalf of user X” Graph work, confirm **`--user`** is available on that subcommand via `--help` before assuming it works everywhere.
