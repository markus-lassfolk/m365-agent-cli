# Goals: m365-agent-cli (Ralph Wiggum Pattern — Personal Assistant Edition)

> Target: v1.0 PA-ready release
> Focus: Complete PA workflow coverage, single-token auth, zero-friction Office 365 integration, real-time collaboration

---

## Part 1: Product & Experience Goals

### 1. The OpenClaw PA — M365 Agent CLI as the Executive Assistant

M365 Agent CLI exists to make the OpenClaw agent the most capable Personal Assistant possible. Every feature should serve this mission: the PA reads email, manages calendars, finds people, searches the directory, tracks tasks, handles files, and communicates — all from the terminal, all driven by an agent, all without the human lifting a finger.

The benchmark is not "can we do this in Outlook" — it is "can the agent do this autonomously without prompting the user." A task that requires the user to open a browser, log in, and click through a UI is not yet solved.

### 2. Single Authentication, Zero Re-authentication

The user authenticates once. M365 Agent CLI requests all necessary scopes via Microsoft OAuth2 incremental consent. All APIs (EWS SOAP, Microsoft Graph) reuse the same token. No second Azure AD app, no separate credentials, no re-authentication flows.

This is not a nice-to-have — it is the foundation. If a new feature requires a separate auth mechanism, it is not compatible with M365 Agent CLI unless that is explicitly resolved first.

**Current gap:** EWS and Graph each have their own token cache. This must be consolidated into a single cache file.

### 3. Zero Hardcoded Assumptions

M365 Agent CLI must never assume:
- A specific timezone (CET, UTC, etc.) — read from `/me/mailboxSettings`
- A specific locale or language — read from `/me/mailboxSettings`
- A specific date/time format — derive from locale
- A specific user's working hours — read from `workingHours`
- A specific country or regional setting

Every output to the user must reflect their actual Microsoft 365 profile settings.

### 4. Full PA Workflow Coverage

A capable PA needs to handle all of these areas. Each is tracked as a separate issue or epic:

| Area | Status | Key Issues |
|------|--------|------------|
| Email (send/read/search/reply/forward/move) | ✅ Implemented | #40, #54 security fixes |
| Calendar (create/update/delete/list/respond) | ✅ Implemented | #73–#78 recurrence gaps |
| Recurring events | ⚠️ Partial | RelativeMonthly/RelativeYearly broken (#74) |
| People / GAL search | ⚠️ Basic | Full directory + DL expansion needed (#85) |
| Delegate management | ❌ Missing | #79 — AddDelegate/UpdateDelegate/RemoveDelegate |
| Shared mailbox access | ⚠️ Partial | Read-only via --mailbox; send-as not complete |
| Inbox rules | ❌ Missing | #81 — full rules CRUD |
| Room discovery | ❌ Missing | #80 — Places API integration |
| Free/busy + meeting suggestions | ⚠️ Basic | findtime exists; getSchedule/findMeetingTimes missing (#86) |
| Out-of-Office | ❌ Missing | #83 — automatic replies via mailboxSettings |
| To-Do / task tracking | ❌ Missing | #82 — use To-Do API, NOT Outlook Tasks (deprecated) |
| Push notifications | ❌ Missing | #84 — Graph subscriptions |
| Contacts / distribution lists | ❌ Missing | #85 — GAL lookup, DL expansion |
| Office 365 Files (SharePoint + OneDrive) | ❌ Missing | See Part 2 — this is critical |
| Delegation (act on behalf of) | ❌ Missing | #79 — shared mailbox, delegate permissions |

---

## Part 2: The Collaboration Imperative (SharePoint + OneDrive)

> This is the highest-priority gap. Real PA productivity requires the PA and user to work on the **same file at the same time**, not pass files back and forth.

### Vision

The PA (agent) and the user both work in SharePoint and OneDrive. When the PA creates a document, it creates it in a shared location the user can immediately open. When the user creates a document, the PA can reference, update, and comment on it — without downloading a copy, without creating duplicates, without version conflicts.

### What Must Work

**File lifecycle (agent ↔ user same file):**
- Agent creates a file in a shared SharePoint document library or OneDrive — user opens it in Office Online immediately
- Agent finds and opens a file the user is working on — reads the current content, adds comments, makes edits
- Edits are applied in-place (Office Online co-authoring), not by uploading a new version
- Comments are added via the Comments API — visible in the document without a new save
- Document check-in/check-out is supported for exclusive-edit workflows

**Reference, don't copy:**
- Emails reference a OneDrive/SharePoint URL — the PA and user both open the same file
- Meeting invitations link to the shared document, not an attachment
- M365 Agent CLI never uploads a file as an email attachment when a sharing link would be better

**Scope:** This applies to:
- Word (.docx), Excel (.xlsx), PowerPoint (.pptx) — full Office Online integration
- PDF — annotation and commenting via Graph
- Text files — direct Graph API read/write
- NOT: legacy Office formats (.doc, .xls, .ppt) — must be converted first

**The PLATEAUS stack for collaboration:**
1. **P**ut the file where both can reach it (SharePoint/OneDrive)
2. **L**ink instead of attach (sharing links, not email attachments)
3. **A**nnotate in-place (Comments API)
4. **T**rack without duplication (check-in/check-out)
5. **E**dit co-authoringly (Office Online)
6. **U**pdate without upload (in-place edits via Graph)
7. **S**ync state (notification when shared docs change)

---

## Part 3: Engineering & Architecture Goals

### 1. Single Token Cache (Critical Path)

**Issue:** Currently EWS and Graph each have their own token cache (`token-cache-${identity}.json` and `graph-token-cache.json`). This violates the single-auth principle and causes sync issues.

**Target:** One file: `~/.config/m365-agent-cli/token-cache.json`. All APIs reuse the same refresh token. Incremental scope consent adds new scopes to the token on next refresh.

**Steps:**
1. Audit current scopes in `auth.ts` and `graph-auth.ts`
2. Merge into single cache with scope metadata
3. Update token refresh to request all known scopes in one call
4. Remove `graph-token-cache.json` migration path

### 2. Security Hardening (Non-Negotiable)

All security issues from the Consultant review are prerequisite to any feature work:

| Priority | Issue | Risk |
|----------|-------|------|
| P0 | #40 — download URL validation | Token exfiltration |
| P0 | #69 — EWS_ENDPOINT/GRAPH_BASE_URL redirect | Token exfiltration |
| P0 | #46 — createLargeUploadSession is a no-op — ✅ FIXED (implemented chunked upload via `uploadLargeFile` + `files upload-large`) | UX fraud |
| P1 | #54 — drafts path traversal bypass | Path traversal |
| P1 | #62 — getFreeBusy returns wrong data | Wrong decisions |
| P1 | #66 — streamWebToFile partial file leak | Disk pollution |

All P0/P1 issues must be resolved before release.

### 3. Full Command Test Coverage

**Issue #36:** Currently 0 command-level integration tests exist.

All 13 CLI commands must have tests covering:
- `--help` output
- `--json` output validity
- Error cases (invalid IDs, missing args, network errors)
- Authentication failure handling
- Edge cases (empty results, unicode, special characters)

Tests must use mocked network calls — no live API calls in CI.

### 4. Dynamic Settings Infrastructure

Before any date/time/locale feature work, establish a shared settings reader:

```typescript
// src/lib/settings.ts
interface UserSettings {
  timezone: string;       // e.g., "W. Europe Standard Time"
  locale: string;         // e.g., "sv-SE"
  dateFormat: string;     // e.g., "yyyy-MM-dd"
  timeFormat: string;     // e.g., "HH:mm"
  workingHours: { start: string; end: string; days: string[]; };
  displayName: string;
  email: string;
}

async function getUserSettings(token: string): Promise<UserSettings>
// Reads from GET /me/mailboxSettings + GET /me
// Caches for 5 minutes to avoid repeated calls
// Falls back to Intl.DateTimeFormat().resolvedOptions() on failure
```

All date/time formatting must use this — never `new Date().toLocaleString()` without locale context.

### 5. No PowerShell Dependencies

Any feature that requires PowerShell remoting (WinRM, remote Exchange management) is explicitly out of scope unless discussed and approved as a design exception. This includes:
- Server-side auto-reply template rules (`ServerReplyWithMessage` — not available in Graph)
- Exchange Admin operations (eDiscovery, litigation hold, permission granting)
- Any operation requiring `*-Mailbox`, `*-Recipient`, `*-ExchangeServer` PowerShell cmdlets

If an operation requires PowerShell, the correct answer is: "that requires admin access and is not an M365 Agent CLI feature."

---

## Part 4: Release Roadmap

### v0.x — Security & Stability (current)

Focus: fix the Consultant issues, establish test coverage, consolidate token cache.

### v0.x+1 — PA Foundation

Focus: delegate management, inbox rules, room discovery, dynamic settings.

### v0.x+2 — Collaboration First

Focus: SharePoint/OneDrive full integration, PLATEAUS workflow, commenting, check-in/check-out.

### v0.x+3 — Intelligence

Focus: To-Do integration, OOF management, push notifications, free/busy AI suggestions, GAL deep search.

---

## Part 5: What "Done" Looks Like

A new user should be able to:

1. Register an Azure AD app, set 4 environment variables
2. Run `m365-agent-cli whoami` — authenticated, shows their profile
3. Run `m365-agent-cli mail` — their inbox, correct locale formatting
4. Run `m365-agent-cli calendar` — their calendar, their timezone
5. Run `m365-agent-cli files list` — their OneDrive
6. Have the OpenClaw agent do all of the above autonomously, every morning, without prompting

That is the benchmark.

---

## Active Tech Debt Targets

- **Token consolidation** — merge EWS and Graph token caches into one
- **Security P0/P1 fixes** — issues #40, #69, #46, #54, #62, #66
- **Command test coverage** — #36: 0 → full coverage for all 13 commands
- **Recurrence fixes** — #74 (RelativeMonthly/RelativeYearly), #73 (view recurrence info)
- **Dynamic settings reader** — shared `/me/mailboxSettings` infrastructure
- **Auth scope audit** — document all scopes currently requested, remove unnecessary ones

## References

- Architecture: `docs/ARCHITECTURE.md`
- Open issues: `github.com/markus-lassfolk/m365-agent-cli/issues`
- Health metrics: CI pass rate, test coverage %, security issue count
