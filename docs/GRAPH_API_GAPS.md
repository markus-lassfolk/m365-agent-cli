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
| **messages/delta** (sync) | **Gap** | No delta query / `@odata.deltaLink` loop in CLI |
| Calendar view, events CRUD | **Implemented** | `calendar`, `create-event`, `update-event`, `delete-event`, `graph-calendar` |
| **events/delta** | **Gap** | Same as mail delta |
| Attachments on events | **Implemented** | `calendar --list-attachments` / `--download-attachments` |
| Calendar sharing (calendarPermission) | **Implemented** | `delegates list`; **`delegates calendar-share add|update|remove`** (Graph model) |
| Classic EWS delegates (folder matrix) | **Implemented (EWS)** | `delegates add|update|remove` — **not** Graph 1:1 |
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
| Page **resources** upload / multipart ink | **Gap** | Graph supports richer edit APIs; CLI focuses on HTML + download |

---

## Files, Teams, To Do, Contacts, Planner, etc.

| Graph area | CLI | Notes |
| --- | --- | --- |
| Drive / SharePoint (subset) | **Implemented** | `files`, `sharepoint` |
| Online meetings | **Implemented** | `meeting` |
| To Do | **Implemented** | `todo` |
| Contacts + delta / photo | **Implemented** | `contacts` |
| Planner | **Implemented** | `planner` |
| Search (query) | **Partial** | `find` — not full Microsoft Search |

---

## Subscriptions & change notifications

| Graph area | CLI | Notes |
| --- | --- | --- |
| Create/delete subscription | **Implemented** | `subscribe` |
| Webhook validation | **Implemented** | `webhook-server` helper |
| Rich subscription lifecycle / renewal CLI | **Partial** | Manual renewal; no dedicated “renew all” daemon in CLI |

---

## What “new” Graph features usually mean here

1. **Delta + long-running sync** — high value for automation; requires design (state file, paging, scopes).
2. **Full Microsoft Search** — broad API; overlaps with `find` and product scope.
3. **OneNote ink / multi-part resource upload** — niche; Graph docs cover advanced page update.

---

*Last updated: 2026-04-03 — OneNote `resource-download` + `content --include-ids`; `delegates calendar-share`.*
