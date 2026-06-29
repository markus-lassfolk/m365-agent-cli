# Graph permissions ↔ CLI features (Entra admin matrix)

**Purpose:** Help Entra administrators choose **delegated** Microsoft Graph permissions based on which **m365-agent-cli** capabilities should work, and to see **which features each permission unlocks**.

**Sources of truth**

- **Feature ↔ scope logic (read/write evaluation):** [`src/lib/graph-capability-matrix.ts`](../src/lib/graph-capability-matrix.ts) (`GRAPH_CAPABILITY_MATRIX`) — also drives **`m365-agent-cli verify-token --capabilities`**. This file is **generated** from that matrix via **`npm run docs:graph-permission-matrix`**.
- **Scopes requested on `login` / refresh:** [`src/lib/graph-oauth-scopes.ts`](../src/lib/graph-oauth-scopes.ts).
- **Narrative scope guide:** [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md).
- **Exchange Web Services:** not represented in Graph `scp`; add **`EWS.AccessAsUser.All`** (Exchange Online) when using EWS-backed mail/calendar — see [`ENTRA_SETUP.md`](./ENTRA_SETUP.md).

**How to verify after consent:** `m365-agent-cli verify-token` (inspect `scp`) and **`verify-token --capabilities`** (checklist). Add **`--json`** for automation.

**Legend**

- **Read:** user can use read-only flows for that area if the token includes **any** listed scope (a **Write** scope also satisfies **Read** for that row, unless the row marks read as not applicable).
- **Write:** user can mutate data if the token includes **any** listed scope. **—** means the row is read-only, send-only, or otherwise has no separate “write” column meaning.
- **Least privilege:** prefer narrower permissions where Graph allows; this table lists what the CLI’s capability checker understands, not every possible Graph alternative.

---

## 1) Feature / capability → Graph permissions (pick scopes by product need)

| Feature area | CLI context (summary) | Read — grant one or more | Write — grant one or more |
| --- | --- | --- | --- |
| Profile / sign-in | `whoami`, basic `/me` | `User.Read`, `User.ReadWrite` | `User.ReadWrite` |
| Directory users | `find` user/group search; `org user` — often admin consent for `/users` reads | `User.Read.All`, `Directory.Read.All`, `Directory.ReadWrite.All` | `Directory.ReadWrite.All` |
| Org hierarchy & profile | `org manager`, `direct-reports`, `user`, `transitive-reports` — self with `User.Read`; others typically `User.Read.All` | `User.Read`, `User.Read.All`, `Directory.Read.All`, `Directory.ReadWrite.All` | — |
| Calendar (your mailbox) | `calendar`, `create-event`, `respond`, … | `Calendars.Read`, `Calendars.ReadWrite` | `Calendars.ReadWrite` |
| Calendar (shared / delegated) | `calendar --mailbox`, delegated calendars | `Calendars.Read.Shared`, `Calendars.ReadWrite.Shared` | `Calendars.ReadWrite.Shared` |
| Mail (your mailbox) | `mail`, `folders`, `drafts` (Graph path) | `Mail.Read`, `Mail.ReadWrite` | `Mail.ReadWrite` |
| Mail (shared / delegated) | `mail --mailbox`, shared folders | `Mail.Read.Shared`, `Mail.ReadWrite.Shared` | `Mail.ReadWrite.Shared` |
| Send mail (Graph) | `send` — `Mail.Send` alone can send; read mail needs `Mail.ReadWrite` | — | `Mail.Send`, `Mail.ReadWrite` |
| Mailbox settings | `oof`, categories, mailbox settings | `MailboxSettings.Read`, `MailboxSettings.ReadWrite` | `MailboxSettings.ReadWrite` |
| Rooms & places | `rooms` (lists, rooms, find, get), `find --rooms`, Places in `create-event` | `Place.Read.All` | — |
| People / relevance | `people` list / get, `find` (People + directory users) | `People.Read`, `People.Read.All` | — |
| OneDrive / files | `files` (delta, shared-with-me, thumbnails, list-item, follow, sensitivity-assign/extract, retention-label, permanent-delete, copy/move, invite, permissions), `excel` workbooks (tables/pivots/ranges/sessions/application) + beta comments, `word`/`powerpoint` full per-item mirror + preview/meta/download/thumbnails | `Files.Read`, `Files.Read.All`, `Files.ReadWrite`, `Files.ReadWrite.All` | `Files.ReadWrite`, `Files.ReadWrite.All` |
| SharePoint sites | `sharepoint` resolve-site/get-site/drives/lists/get-list/columns/items (OData paging)/create/update/get/delete, `items-delta`, followed sites; `site-pages` | `Sites.Read.All`, `Sites.ReadWrite.All`, `Sites.Manage.All` | `Sites.ReadWrite.All`, `Sites.Manage.All` |
| SharePoint followed sites | `sharepoint followed-sites`, `follow`, `unfollow` | `Sites.Read.All`, `Sites.ReadWrite.All` | `Sites.ReadWrite.All` |
| Discovery / Insights | `insights` trending / used / shared, `files recent`, `files activities`, `files preview` | `Sites.Read.All`, `Sites.ReadWrite.All`, `Files.Read`, `Files.Read.All`, `Files.ReadWrite`, `Files.ReadWrite.All` | — |
| Viva / employee experience (Graph beta) | `viva` — user + tenant `/employeeExperience` (communities, goals, learning, roles), work time + insights, admin/org itemInsights, workHoursAndLocations, meeting Engage Q&A | `User.Read`, `LearningAssignedCourse.Read`, `EngagementRole.Read`, `EngagementRole.Read.All`, `MailboxSettings.Read`, `MailboxSettings.ReadWrite` — tenant learning / communities / goals need product-specific admin-consented scopes per Microsoft; work-time mutations may be app-only (`Schedule-WorkingTime.*`) in some tenants | `EngagementRole.ReadWrite.All`, `MailboxSettings.ReadWrite`, `User.ReadWrite`, `User.ReadWrite.All` |
| Approvals | `approvals` list / get / steps / respond / cancel (DELETE) — Teams Approvals + Power Automate (beta `/me/approvals`); `ApprovalSolution.ReadWrite` (canonical) or narrower `ApprovalSolutionResponse.ReadWrite` | `ApprovalSolution.Read.All`, `ApprovalSolution.ReadWrite`, `ApprovalSolution.ReadWrite.All`, `ApprovalSolutionResponse.ReadWrite` | `ApprovalSolution.ReadWrite`, `ApprovalSolutionResponse.ReadWrite` |
| Microsoft To Do | `todo` (incl. `attachment-session` list/get/patch/delete/content-*; `root` get/patch/delete for …/todo — destructive delete requires `--confirm`) | `Tasks.Read`, `Tasks.ReadWrite` | `Tasks.ReadWrite` |
| Planner & group-backed Teams | `planner` (incl. `delete-plan-details`, `delete-task-details`), `teams` members/channels/apps/tabs — broad group scope | `Group.Read.All`, `Group.ReadWrite.All` | `Group.ReadWrite.All` |
| Outlook Groups (Microsoft 365 groups) | `groups list`, `conversations`, `thread`, `posts`, `post-reply` | `Group.Read.All`, `Group.ReadWrite.All` | `Group.ReadWrite.All` |
| Contacts (your mailbox) | `contacts` (extensions: `-f/--folder`, `--child-folder` for nested contact folder paths) | `Contacts.Read`, `Contacts.ReadWrite` | `Contacts.ReadWrite` |
| Contacts (shared mailbox) | `contacts --user` (extensions: `-f/--folder`, `--child-folder`) | `Contacts.Read.Shared`, `Contacts.ReadWrite.Shared` | `Contacts.ReadWrite.Shared` |
| Contact merge suggestions (Graph beta) | `contacts merge-suggestions` get/set/delete — `/settings/contactMergeSuggestions` (see Microsoft Graph beta docs for exact permission names) | `User.Read`, `User.ReadWrite`, `User.Read.All` | `User.ReadWrite`, `User.ReadWrite.All` |
| Online meetings | `meeting`, Teams links in `create-event` | `OnlineMeetings.Read`, `OnlineMeetings.ReadWrite` | `OnlineMeetings.ReadWrite` |
| Meeting recordings | `meeting recordings`, `recording-download`, `recordings-all` (+ `--delta`) — tenant Stream/Teams policy applies | `OnlineMeetingRecording.Read.All` | — |
| Meeting transcripts | `meeting transcripts`, `transcript-download`, `transcripts-all` (+ `--delta`) | `OnlineMeetingTranscript.Read.All` | — |
| OneNote | `onenote` | `Notes.Read`, `Notes.ReadWrite`, `Notes.ReadWrite.All` | `Notes.ReadWrite`, `Notes.ReadWrite.All` |
| Teams (teams & channels) | `teams` list (incl. `list --user`), channels, `channel-files-folder`, metadata | `Team.ReadBasic.All`, `Channel.ReadBasic.All` | — |
| Teams channel messages | `teams messages`, read channel posts — often admin consent | `ChannelMessage.Read.All` | — |
| Teams channel send | `teams channel-message-send`, replies, `channel-message-patch`, `channel-message-delete` (soft/hard), `tab-create` / `tab-update` / `tab-delete` | — | `ChannelMessage.Send` |
| Teams chats (1:1 / group) | `teams chats`, messages, `chat-create`, `chat-member-add`, `chat-message-patch`, `chat-message-delete`, `chat-apps` / `chat-app-*` | `Chat.Read`, `Chat.ReadWrite` | `Chat.ReadWrite` |
| Teams activity feed | `teams activity-notify` — POST /me/teamwork/sendActivityNotification, /chats/{id}/sendActivityNotification, or /users/{id}/teamwork/sendActivityNotification (--user-id; typically app token) | — | `TeamsActivity.Send` |
| Teams membership (provision) | `teams team-member-add`, `teams channel-member-add` | — | `TeamMember.ReadWrite.All`, `TeamMember.ReadWriteNonGuestRole.All`, `ChannelMember.ReadWrite.All`, `Group.ReadWrite.All` |
| Teams app catalog | `teams app-catalog`, `teams app-catalog-get` | `AppCatalog.Read.All`, `AppCatalog.Submit` | — |
| Teams apps on a team | `teams apps`, `app-get`, `app-add`, `app-patch`, `app-upgrade`, `app-delete` | `Group.ReadWrite.All` | `Group.ReadWrite.All` |
| Teams apps (personal user scope) | `teams user-apps`, `user-app-get`, `user-app-add`, `user-app-delete` (`--user` may need broader Teams app permissions) | `TeamsAppInstallation.ReadWriteSelfForUser` | `TeamsAppInstallation.ReadWriteSelfForUser` |
| Presence (read) | `presence me`, `presence user`, bulk | `Presence.Read.All` | — |
| Presence (set/clear) | `presence set-*`, `presence clear-*`, `status-message-set`, `preferred-set`, `preferred-clear`, `clear-location` | — | `Presence.ReadWrite` |
| Bookings | `bookings` | `Bookings.Read.All`, `Bookings.ReadWrite.All` | `Bookings.ReadWrite.All` |
| Graph Search | `graph-search` — presets/types, `--merge-json-file` / `--body-file`, entity-specific mail/files/site scopes | `Mail.Read`, `Mail.ReadWrite`, `Files.Read.All`, `Files.ReadWrite.All`, `Sites.Read.All`, `Sites.ReadWrite.All` | — |
| Copilot Retrieval API | `copilot retrieval` — SharePoint/OneDrive need Files+Sites read; connectors need `ExternalItem.Read.All` | `Files.Read.All`, `Files.ReadWrite.All`, `Sites.Read.All`, `Sites.ReadWrite.All`, `ExternalItem.Read.All` | — |
| Copilot Search API (preview) | `copilot search` — OneDrive; `Files.Read.All`+`Sites.Read.All` (or ReadWrite equivalents) | `Files.Read.All`, `Files.ReadWrite.All`, `Sites.Read.All`, `Sites.ReadWrite.All` | — |
| Copilot Chat API (preview) | `copilot conversation-create`, `conversations-list`, `conversation-*`, `messages-*`, `chat`, `chat-stream` (OData actions) — Microsoft requires the full permission bundle documented for Chat API | `Sites.Read.All`, `Sites.ReadWrite.All`, `Mail.Read`, `Mail.ReadWrite`, `Mail.ReadWrite.Shared`, `People.Read.All`, `OnlineMeetingTranscript.Read.All`, `Chat.Read`, `Chat.ReadWrite`, `ChannelMessage.Read.All`, `ExternalItem.Read.All` | — |
| Copilot interaction export | `copilot interactions-export`, `interactions-export-tenant` — application `AiEnterpriseInteraction.Read.All` (delegated not supported) | `AiEnterpriseInteraction.Read.All` (application permission typical) | — |
| Copilot interaction change notifications | `subscribe copilot-interactions` — delegated `AiEnterpriseInteraction.Read` (per-user); tenant subscription is app-only | `AiEnterpriseInteraction.Read`, `AiEnterpriseInteraction.Read.All` | — |
| Copilot meeting insights | `copilot meeting-insights-list`, `meeting-insight-get`, `meeting-insights-count`, `meeting-insight-create\|patch\|delete` | `OnlineMeetingAiInsight.Read.All` | — |
| Copilot usage reports | `copilot reports user-count-*`, `usage-user-detail`, `nav-get\|nav-patch\|nav-delete` — `Reports.Read.All` + admin reader role per Microsoft | `Reports.Read.All` | — |
| Copilot package catalog (admin) | `copilot packages` list\|get\|create\|delete\|zip-download\|zip-upload\|update\|block\|unblock\|reassign — `CopilotPackages.Read*.` | `CopilotPackages.Read.All`, `CopilotPackages.ReadWrite.All` | `CopilotPackages.ReadWrite.All` |
| Copilot agents (Graph) | `copilot agents-list`, `agent-get` — permissions per Microsoft Graph `/copilot/agents` docs | — | — |
| Copilot user settings | `copilot settings-*` (+ `settings-delete`, `settings-people-delete`, `settings-enhanced-personalization-delete`) | — | — |
| Copilot admin settings | `copilot admin-settings-*`, `admin-limited-mode-*`, `admin-settings-delete`, `admin-limited-mode-delete`, `admin-nav-*`, `admin-catalog-*` (beta) | — | — |
| Copilot realtime activity feed | `copilot activity-feed …` incl. `patch-root`, `delete-root`, `*-count` — `/copilot/communications/realtimeActivityFeed/...` | — | — |
| Copilot aiUser (/copilot/users) | `copilot ai-user …` — list/count/CRUD, interactionHistory, onlineMeetings | — | — |
| Graph invoke / batch | `graph invoke`, `graph batch` — depends on path you call | *Depends on URL and method* | *Depends on URL and method* |
| Exchange Web Services (EWS) | EWS mail/calendar when configured — not in Graph `scp`; add `EWS.AccessAsUser.All` (Exchange Online) on the same Entra app | *Not in Graph `scp`* | *Not in Graph `scp`* — add `EWS.AccessAsUser.All` |

---

## 2) Graph permission → features (pick features enabled by each consent)

Alphabetical **delegated** (and noted **application**) permissions referenced by the capability matrix. “Features” names match §1 **Feature area** column.

| Permission | Enables (feature areas) | Notes |
| --- | --- | --- |
| `AiEnterpriseInteraction.Read` | Copilot interaction change notifications | Delegated per-user subscriptions |
| `AiEnterpriseInteraction.Read.All` | Copilot interaction change notifications; Copilot interaction export | Export often **application**; `.All` also listed for notifications in matrix |
| `AppCatalog.Read.All` | Teams app catalog |  |
| `AppCatalog.Submit` | Teams app catalog |  |
| `ApprovalSolution.Read.All` | Approvals | Read/list steps |
| `ApprovalSolution.ReadWrite` | Approvals | Create/respond; canonical name in [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md) |
| `ApprovalSolution.ReadWrite.All` | Approvals | Read side in matrix |
| `ApprovalSolutionResponse.ReadWrite` | Approvals | Narrower respond-only variant |
| `Bookings.Read.All` | Bookings | Read |
| `Bookings.ReadWrite.All` | Bookings | Read + write |
| `Calendars.Read` | Calendar (your mailbox) | Read |
| `Calendars.Read.Shared` | Calendar (shared / delegated) | Read |
| `Calendars.ReadWrite` | Calendar (your mailbox) | Read + write |
| `Calendars.ReadWrite.Shared` | Calendar (shared / delegated) | Read + write |
| `Channel.ReadBasic.All` | Teams (teams & channels) | Channels metadata |
| `ChannelMember.ReadWrite.All` | Teams membership (provision) | Add channel members |
| `ChannelMessage.Read.All` | Copilot Chat API (preview); Teams channel messages | Often admin consent |
| `ChannelMessage.Send` | Teams channel send | Post/reply/patch/delete channel messages & tab CRUD per matrix |
| `Chat.Read` | Copilot Chat API (preview); Teams chats (1:1 / group) | Read |
| `Chat.ReadWrite` | Copilot Chat API (preview); Teams chats (1:1 / group) | Read + write chats |
| `Contacts.Read` | Contacts (your mailbox) | Read |
| `Contacts.Read.Shared` | Contacts (shared mailbox) | Read |
| `Contacts.ReadWrite` | Contacts (your mailbox) | Read + write |
| `Contacts.ReadWrite.Shared` | Contacts (shared mailbox) | Read + write |
| `CopilotPackages.Read.All` | Copilot package catalog (admin) | Read |
| `CopilotPackages.ReadWrite.All` | Copilot package catalog (admin) | Mutations |
| `Directory.Read.All` | Directory users; Org hierarchy & profile | With `User.Read.All`-style directory reads |
| `Directory.ReadWrite.All` | Directory users; Org hierarchy & profile | Write |
| `EngagementRole.Read` | Viva / employee experience (Graph beta) |  |
| `EngagementRole.Read.All` | Viva / employee experience (Graph beta) |  |
| `EngagementRole.ReadWrite.All` | Viva / employee experience (Graph beta) |  |
| `ExternalItem.Read.All` | Copilot Chat API (preview); Copilot Retrieval API | Connectors / external items |
| `Files.Read` | Discovery / Insights; OneDrive / files | Narrower file read |
| `Files.Read.All` | Copilot Retrieval API; Copilot Search API (preview); Discovery / Insights; Graph Search; OneDrive / files | All drives user can reach |
| `Files.ReadWrite` | Discovery / Insights; OneDrive / files | File write |
| `Files.ReadWrite.All` | Copilot Retrieval API; Copilot Search API (preview); Discovery / Insights; Graph Search; OneDrive / files | Broad file read/write |
| `Group.Read.All` | Outlook Groups (Microsoft 365 groups); Planner & group-backed Teams | Read |
| `Group.ReadWrite.All` | Outlook Groups (Microsoft 365 groups); Planner & group-backed Teams; Teams apps on a team; Teams membership (provision) | Includes member add paths per matrix |
| `LearningAssignedCourse.Read` | Viva / employee experience (Graph beta) |  |
| `Mail.Read` | Copilot Chat API (preview); Graph Search; Mail (your mailbox) | Read |
| `Mail.Read.Shared` | Mail (shared / delegated) | Read |
| `Mail.ReadWrite` | Copilot Chat API (preview); Graph Search; Mail (your mailbox); Send mail (Graph) | Read + write mailbox |
| `Mail.ReadWrite.Shared` | Copilot Chat API (preview); Mail (shared / delegated) | Shared mailbox write |
| `Mail.Send` | Send mail (Graph) | Sending without full mailbox read |
| `MailboxSettings.Read` | Mailbox settings; Viva / employee experience (Graph beta) | Read |
| `MailboxSettings.ReadWrite` | Mailbox settings; Viva / employee experience (Graph beta) | Read + write |
| `Notes.Read` | OneNote | Read |
| `Notes.ReadWrite` | OneNote | Read + write |
| `Notes.ReadWrite.All` | OneNote | All notebooks |
| `OnlineMeetingAiInsight.Read.All` | Copilot meeting insights |  |
| `OnlineMeetingRecording.Read.All` | Meeting recordings | Tenant Stream/Teams policy may still block |
| `OnlineMeetings.Read` | Online meetings | Read |
| `OnlineMeetings.ReadWrite` | Online meetings | Create/update/delete |
| `OnlineMeetingTranscript.Read.All` | Copilot Chat API (preview); Meeting transcripts |  |
| `People.Read` | People / relevance | `/me/people`, etc. |
| `People.Read.All` | Copilot Chat API (preview); People / relevance | Directory-style people reads |
| `Place.Read.All` | Rooms & places | Often admin consent |
| `Presence.Read.All` | Presence (read) |  |
| `Presence.ReadWrite` | Presence (set/clear) |  |
| `Reports.Read.All` | Copilot usage reports | Plus Microsoft’s reports reader role |
| `Sites.Manage.All` | SharePoint sites | Elevated site manage |
| `Sites.Read.All` | Copilot Chat API (preview); Copilot Retrieval API; Copilot Search API (preview); Discovery / Insights; Graph Search; SharePoint followed sites; SharePoint sites |  |
| `Sites.ReadWrite.All` | Copilot Chat API (preview); Copilot Retrieval API; Copilot Search API (preview); Discovery / Insights; Graph Search; SharePoint followed sites; SharePoint sites |  |
| `Tasks.Read` | Microsoft To Do |  |
| `Tasks.ReadWrite` | Microsoft To Do |  |
| `Team.ReadBasic.All` | Teams (teams & channels) | Joined teams |
| `TeamMember.ReadWrite.All` | Teams membership (provision) | Add team members |
| `TeamMember.ReadWriteNonGuestRole.All` | Teams membership (provision) | Non-guest variant |
| `TeamsActivity.Send` | Teams activity feed | Activity notifications |
| `TeamsAppInstallation.ReadWriteSelfForUser` | Teams apps (personal user scope) |  |
| `User.Read` | Contact merge suggestions (Graph beta); Org hierarchy & profile; Profile / sign-in; Viva / employee experience (Graph beta) |  |
| `User.Read.All` | Contact merge suggestions (Graph beta); Directory users; Org hierarchy & profile | Often admin consent |
| `User.ReadWrite` | Contact merge suggestions (Graph beta); Profile / sign-in; Viva / employee experience (Graph beta) |  |
| `User.ReadWrite.All` | Contact merge suggestions (Graph beta); Viva / employee experience (Graph beta) |  |

---

## Related

- [`GRAPH_PRODUCT_PARITY_MATRIX.md`](./GRAPH_PRODUCT_PARITY_MATRIX.md) — workloads vs CLI commands (coverage, not permissions).
- [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md) — raw **`graph invoke`** surfaces; consent must match each API.
- [`GRAPH_TROUBLESHOOTING.md`](./GRAPH_TROUBLESHOOTING.md)

*Auto-generated from `GRAPH_CAPABILITY_MATRIX` — run `npm run docs:graph-permission-matrix` after editing the matrix. Generated: 2026-06-29.*
