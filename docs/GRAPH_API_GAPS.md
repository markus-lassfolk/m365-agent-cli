# Microsoft Graph vs this CLI — capabilities and gaps

**Purpose:** Track **Graph API areas** that are **implemented** in `m365-agent-cli`, **partially** covered, or **not** exposed, so we can prioritize work and set expectations.

**Related:** [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md), [`GRAPH_EWS_PARITY_MATRIX.md`](./GRAPH_EWS_PARITY_MATRIX.md), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md).

---

## Legend

| Status | Meaning |
| --- | --- |
| **Implemented** | Primary or parallel command covers typical use. |
| **Partial** | Some APIs or flags; not exhaustive vs Graph. |
| **Gap** | Graph supports it; this CLI does not wrap it (use Graph directly, another tool, or contribute). |

---

## Exchange / Outlook (mail & calendar)

| Graph area | CLI | Notes |
| --- | --- | --- |
| Messages CRUD, send, attachments | **Implemented** | `mail`, `send`, `drafts`, `folders`, `outlook-graph` |
| Message search / list filters | **Implemented** | `mail` Graph path (`mail-graph.ts`) — not every flag combo |
| **messages/delta** (sync) | **Implemented** | **`outlook-graph messages-delta`** — first page or `--next` for `@odata.nextLink` |
| Calendar view, events CRUD | **Implemented** | `calendar`, `create-event`, `update-event`, `delete-event`, `graph-calendar` |
| **events/delta** | **Implemented** | **`graph-calendar events-delta`** — optional `--calendar`; `--next` for paging |
| Attachments on events | **Implemented** | `calendar --list-attachments` / `--download-attachments` |
| Calendar sharing (calendarPermission) | **Implemented** | `delegates list`; **`delegates calendar-share add, update, remove`** (Graph model) |
| Classic EWS delegates (folder matrix) | **Implemented (EWS)** | `delegates add, update, remove` — **not** Graph 1:1 |
| Inbox rules | **Implemented** | `rules` |
| Automatic replies (mailboxSettings) | **Implemented** | `oof` |
| EWS-style auto-reply templates | **Partial** | `auto-reply` (EWS); **`oof`** / **`rules`** for Graph-native |

---

## OneNote

| Graph area | CLI | Notes |
| --- | --- | --- |
| Notebooks, sections, pages, HTML, PATCH | **Implemented** | `onenote` |
| Copy page/section, operations poll | **Implemented** | `onenote copy-page`, `section copy-to-*`, `onenote operation` |
| **GET …/pages/{id}/content** | **Implemented** | `onenote content`, `export`; **`--include-ids`** for `includeIDs=true` |
| **GET …/resources/{id}/content** (binary) | **Implemented** | **`onenote resource-download`** — resource ids from page HTML |
| Page **resources** upload / multipart | **Implemented** | **`onenote create-page-multipart`**, **`onenote patch-page-content-multipart`** |

---

## Files, Teams, To Do, Contacts, Planner, etc

| Graph area | CLI | Notes |
| --- | --- | --- |
| Drive / SharePoint (subset) | **Implemented** | `files`, `sharepoint` |
| Excel workbook (worksheets, range read) | **Partial** | **`excel`** — worksheets, **tables**, **range**, **used-range** (`valuesOnly`); no session/chart mutations |
| Teams (joined teams, channels, messages) | **Partial** | **`teams`** — **all-channels** (`$filter`), **incoming-channels**, **channel-get**, **channel-members**, primary-channel, **tabs** (`$expand=teamsApp`), chat-get / chat-pinned / chat-messages / chat-members, team **members**, **apps**, list **channels** / **messages** / **message-replies**; not full meetings / RSC |
| Bookings | **Partial** | **`bookings`** — businesses + **business-get**, **`currencies`**, appointments + appointment get, customers + **customer** get, custom-questions, services + **service-get**, staff + **staff-get**, calendar-view |
| Bookings **getStaffAvailability** | **Gap (delegated)** | Microsoft documents **no delegated** support — application-only; use **`graph invoke`** with app-only token if applicable |
| Presence | **Partial** | **`presence`** — `/me/presence` and `/users/.../presence` |
| Raw REST + JSON `$batch` | **Partial** | **`graph invoke`**, **`graph batch`** — escape hatch for any JSON Graph API |
| Online meetings | **Implemented** | `meeting` |
| To Do | **Implemented** | `todo` |
| Contacts + delta / photo | **Implemented** | `contacts` |
| Planner | **Implemented** | `planner` |
| Search (query) | **Partial** | `graph-search` + `find` — not every Search vertical |
| Cloud communications (calls, PSTN, etc.) | **Gap** | Use **`graph invoke`** or dedicated apps; not wrapped |

---

## Subscriptions & change notifications

| Graph area | CLI | Notes |
| --- | --- | --- |
| Create/delete subscription | **Implemented** | `subscribe` |
| Webhook validation | **Implemented** | `webhook-server` helper |
| List / renew subscriptions | **Implemented** | **`subscribe list`**, **`subscribe renew`** (plus create/cancel) — no built-in “renew all” daemon |

---

## What “new” Graph features usually mean here

1. **Delta + long-running sync** — mail/events/contacts/todo/planner expose **paged** delta (`--next` / `--url`); a **persisted state file** for full sync loops is still a product choice, not a single command.
2. **Microsoft Search** — **`graph-search`** covers typical `entityTypes` + KQL-style queries; exotic search verticals / connectors may still need raw Graph.
3. **OneNote** — advanced ink scenarios beyond multipart HTML + binary parts may need Graph directly.

---

*Last updated: 2026-04-03 — **`teams incoming-channels`**; **`bookings` business-get / service-get / staff-get**.*
