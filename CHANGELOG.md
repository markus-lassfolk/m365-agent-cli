# Changelog

All notable changes to **m365-agent-cli** are documented here. Release **2026.4.50** was the first stable line after **1.2.4** with the Graph-first stack and unified auth. **2026.5.51** is a follow-up patch focused on **discoverability** (grouped `--help` for the full tree and key workloads), **clearer Microsoft Graph errors** when a **v1.0** call returns **404** (hint to try **beta** or `GRAPH_BASE_URL`), and small **reliability / analysis** fixes (trailing-slash handling on Graph base URLs, refreshed path inventory). **2026.5.50** remains the larger feature drop: agent packaging (OpenClaw skill + MCP), Copilot and Viva surfaces, deeper Teams / Excel / OneDrive / SharePoint coverage, meeting recordings and transcripts, and expanded docs and CI—see that section below.

For install and tagging, see [docs/RELEASE.md](docs/RELEASE.md).

---

## [Unreleased]

### Mail

- **`mail --reply` / `--reply-all` / `--forward` now accept `--cc` and `--bcc`** (comma-separated) to add CC/BCC recipients to the outgoing reply or forward. Previously these commands had no way to set CC/BCC, so agents could only reply to the original participants. Recipients are **added** on top of the To/Cc a reply-all already carries (deduped case-insensitively), rather than replacing them. Works on both the Microsoft Graph path (patches the reply/forward draft before send) and the EWS path (adds `CcRecipients`/`BccRecipients` to the response object).

### Fixed (QA sweep)

Correctness and robustness fixes surfaced by a full-repo review:

- **Graph large-attachment threshold was 2 MB, below Graph's 3 MB minimum.** Files sized 2–3 MB were routed to an upload session and rejected with `ErrorAttachmentSizeShouldNotBeLessThanMinimumSize`. The crossover is now the correct 3 MB (single POST below, upload session at/above).
- **Large (upload-session) attachments now return the real attachment id** parsed from the final PUT's `Location` header instead of the placeholder string `"uploaded"`.
- **EWS multi-attachment sends used a stale ChangeKey.** `CreateAttachment` returns the parent's new change key as attributes on `<t:AttachmentId>` (`RootItemId` / `RootItemChangeKey`), not as a `<RootItemId>` element — the previous extraction always missed it, so a second attachment or the final `SendItem` could fail with `ErrorIrresolvableConflict`.
- **EWS `sendEmail` with attachments dropped all BCC recipients** — the attachment path built the draft via `createDraft`, which had no `bcc` parameter. BCC is now threaded through.
- **`drafts --create` / `--edit` (EWS) reported success when a file attachment failed.** The create path logged the error but didn't exit; the edit path discarded the result entirely. Both now exit non-zero on attachment failure.
- **`planner create-task --priority` / `--preview-type` were parsed (and `--priority` range-checked) but never sent.** Both are now applied to the created task.
- **Planner `reference add`/`remove` sent raw URLs as Open Type property names.** OData forbids `.`, `:`, `%`, `@`, `#` in those names; they are now percent-encoded, so adding/removing a normal `https://` reference works.
- **`copilot packages zip-download`** now refuses to overwrite an existing output file unless `--force` is passed.
- **`mail --force`** (used by the attachment-download overwrite logic) was referenced but never registered as an option; it is now a real flag.
- **`update-event` (auto backend)** no longer crashes with no message when a Graph attachments-only attempt fails without an EWS fallback context — it now exits with a clear, actionable error.
- **Date parsing** (`--day`) no longer silently rolls over out-of-range values (`2026-02-30` → Mar 2); invalid calendar dates are rejected.
- **Markdown → HTML** no longer corrupts link URLs/labels containing `_` or `*` (emphasis was applied after links were restored, rewriting `href` internals); links are now restored after all other transforms.
- **Input validation:** empty comma-split entries are dropped for `mail`/`create-event`/`suggest` attendee/recipient lists and `presence bulk --json-file`; `excel table-rows --top` rejects non-numeric values; `excel table-rows-add --json` always emits JSON; `todo create` validates `--importance`/`--status`; `onenote` and `presence` file reads report a clear error instead of an empty non-zero exit.
- **Webhook receiver hardening (`serve`):** caps request body size, sets request/header timeouts (slow-loris), handles bind errors (`EADDRINUSE`/`EACCES`) gracefully, redacts `clientState` from logs, compares `clientState` in constant time, warns when verification is disabled, and logs the actual bind interface.
- **Auth:** EWS token refresh now preserves the `graphNarrowScopeAccepted` flag, avoiding a redundant Graph refresh on the next call.
- **Mime types:** `.tgz` now maps to `application/gzip`.

---

## [2026.7.3] — 2026-07-02

Patch release: **`interaction_required` surfacing on EWS token refresh**, cached-token app id / tenant mismatch checks, and secret-directory permission hardening. Upgrade with `npm install -g m365-agent-cli@2026.7.3` (or `@latest` once published).

### Auth

- **Surfaces AADSTS `interaction_required` (500133) on EWS refresh.** `resolveAuth` now detects `error: 'interaction_required'` or error code `500133` from the token endpoint and appends a re-authentication hint (`Run \`m365-agent-cli login\` again.`) to the failure message instead of a generic refresh error.
- **Rejects cached EWS access tokens that don't match the configured client or pinned tenant.** Cached tokens are now checked against `EWS_CLIENT_ID` (`appid`/`azp` claim) and, when the operator pinned a concrete tenant GUID, the token's `tid` claim; a mismatch on either forces a refresh instead of silently returning a stale token for the wrong app or tenant.
- **Hardens `sanitizeRefreshError`** to strip all Unicode control characters (`\p{Cc}`), not just `\r\n\t\0`, before surfacing OAuth error text.
- **`login` device-code flow uses `isValidJwtStructure`** instead of a bare `parts.length === 3` check before decoding the access token payload.

### Security

- **Re-tightens secret-bearing config directories to `0700`** (and cache files to `0600`) after `mkdir`/`rename`, since `mkdir`'s `mode` only applies to directories it creates and `rename` preserves the source file's prior mode — both `atomic-write.ts` and the legacy token-cache migration paths in `m365-token-cache.ts` now re-`chmod` explicitly.

### Tests

- New tests cover `interaction_required` surfacing, cached-token app id / tenant mismatch, `isValidJwtStructure`/`getJwtPayloadTenantId`/`isPinnedTenantGuid`, and the control-character sanitizer in `src/test/auth.test.ts`, `src/test/graph-auth.test.ts`, `src/lib/jwt-utils.test.ts`, and `src/lib/m365-token-cache.test.ts`.

---

## [2026.7.2] — 2026-07-02

Patch release: **auth and token handling hardening** for active env-file refresh persistence, clearer Microsoft Entra refresh failures, stricter JWT validation, scope parity tests, and tenant-id precedence. Upgrade with `npm install -g m365-agent-cli@2026.7.2` (or `@latest` once published).

### Auth and token handling hardening

- **Fixes rotated refresh token persistence to the active env file (H-1 / H-2).** `m365-agent-cli resolveAuth` and `resolveGraphAuth` now accept an `envPath` option and persist refreshed `M365_REFRESH_TOKEN` (plus legacy `EWS_REFRESH_TOKEN` / `GRAPH_REFRESH_TOKEN`) to that same file the CLI loaded from — including `login --env-file` and `verify-token --env-file` (and `M365_AGENT_ENV_FILE`). Previously, `login --env-file` followed by any token refresh wrote the rotated token to the default global `~/.config/m365-agent-cli/.env`, silently losing the file the user pointed at. A new `getActiveEnvFilePath(explicit?)` helper in `src/lib/active-env.ts` centralizes the precedence: explicit caller path > `M365_AGENT_ENV_FILE` > default global.
- **Surfaces AADSTS / `interaction_required` / refresh errors in `AuthResult` and `GraphAuthResult` (M-1).** Failed token refreshes previously only emitted a `console.warn`; callers saw a generic `Token refresh failed.` message. The refreshed `AuthResult` and `GraphAuthResult` now carry a sanitized `lastRefreshError` field, and the human-facing `error` string includes the last `error: error_description` (or HTTP status) with a re-authentication hint when the response carries AADSTS code `500133` / `interaction_required`. Refresh tokens and access tokens are stripped before surfacing.
- **M-3 regression test asserts Viva / Engage scope parity.** `EngagementRole.Read.All`, `EngagementRole.ReadWrite.All`, and `LearningAssignedCourse.Read` are checked to appear in BOTH `GRAPH_DEVICE_CODE_LOGIN_SCOPES` and `GRAPH_REFRESH_SCOPE_CANDIDATES`, with an additional allowlist-driven parity check for primary refresh scope URLs.
- **Hardens JWT structure validation (M-5).** `isValidJwtStructure` now requires three non-empty base64url segments and JSON-decodable object header and payload. The previous implementation called `Buffer.from(parts[1], 'base64url').toString()` (which never throws and yields an empty string for malformed input), accepting strings like `"a.."` and `"...."` as valid tokens. New tests cover empty input, wrong part count, empty segments, non-JSON payload, JSON-array payload, non-object header, and undecodable base64url.
- **Tenant ID precedence (M-2).** `getMicrosoftTenantPathSegment` now reads `M365_TENANT_ID` > `MICROSOFT_TENANT_ID` > `EWS_TENANT_ID` (legacy) > `common`. `EWS_TENANT_ID` remains supported for backwards compatibility. The error message lists all three names. Documented in `docs/AUTHENTICATION.md`.

### Tests

- New `src/lib/active-env.test.ts` covers `getActiveEnvFilePath` precedence and tilde expansion.
- New `isValidJwtStructure` and tenant precedence test blocks in `src/lib/jwt-utils.test.ts`.
- New AADSTS surfacing test in `src/test/graph-auth.test.ts` and `src/test/auth.test.ts`.
- New envPath-threaded refresh tests in `src/test/auth.test.ts` and `src/test/graph-auth.test.ts`.
- New M-3 scope parity test in `src/lib/graph-oauth-scopes.test.ts`.

---

## [2026.6.30] — 2026-06-30

Patch release: **shared Microsoft To Do delegated scopes** and clearer cross-user To Do guidance. Upgrade with `npm install -g m365-agent-cli@2026.6.30` (or `@latest` once published).

### OAuth scopes (Microsoft To Do)

- Adds **`Tasks.Read.Shared`** and **`Tasks.ReadWrite.Shared`** to Graph login and refresh scope sets so shared/delegated Microsoft To Do scenarios request the same consent the Entra setup now documents.
- Treats **`Tasks.ReadWrite.Shared`** as a critical delegated scope; stale or narrow cached Graph tokens refresh instead of silently continuing without the shared To Do write permission.
- Updates the Graph capability matrix, Entra setup, authentication, and personal-assistant delegation docs to include the shared To Do scopes.

### Microsoft To Do cross-user semantics

- Clarifies `todo --user` help text and docs: Microsoft To Do `/users/{id}/todo/...` is **not mailbox delegation** and can still depend on target-user To Do provisioning/sharing even after the shared scopes are present.
- Documents the practical troubleshooting path: verify `/me/todo/lists` and `/users/{self}/todo/lists` first, then investigate target-user To Do service/sharing state if another user returns Graph `Invalid request`.

---

## [2026.6.29] — 2026-06-29

Patch release: **refresh-token `.env` sync** ([#230](https://github.com/markus-lassfolk/m365-agent-cli/issues/230)) and **Viva Learning OAuth scope name** fix. Upgrade with `npm install -g m365-agent-cli@2026.6.29` (or `@latest` once published).

### Auth / token persistence

- After any successful **Graph** or **EWS** OAuth refresh, the CLI now **upserts** `M365_REFRESH_TOKEN` (and legacy `GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN`) in the global `.env` — same keys as `login`. The JSON cache and `.env` stay aligned when Entra rotates refresh tokens.
- Skips the write when the refresh token string is unchanged, when `NODE_ENV=test`, or when `M365_AGENT_SKIP_GLOBAL_ENV=1`.
- **`login`** uses the same helper (behavior unchanged for operators).

### OAuth scopes (Viva Learning)

- **`LearningAssignedCourse.Read.All`** → **`LearningAssignedCourse.Read`** in `login` / refresh scope lists and Entra setup scripts. Microsoft Graph exposes only the `.Read` delegated permission; the old name never existed on the service principal (stale GUID in setup scripts corrected to `ac08cdae-e845-41db-adf9-5899a0ec9ef6`).
- Re-run **`m365-agent-cli login`** after upgrading if Viva Learning calls previously failed consent for the wrong scope name.

---

## [2026.5.51] — 2026-05-05

Patch release after **2026.5.50**. Upgrade with `npm install -g m365-agent-cli@2026.5.51` (or `@latest` once published). No new Entra permissions are required for this patch; behavior changes are limited to help output, error messaging, and internal URL normalization.

### Highlights (what changed at a glance)

| Area | What you get |
|------|----------------|
| **Help and discovery** | **`m365-agent-cli --help`** now lists commands in **logical groups** (calendar, mail, files, Teams, Graph tools, etc.) instead of a flat wall of names. Heavy commands (**`teams`**, **`files`**, **`calendar`**, **`mail`**, **`create-event`**) show **grouped subcommand help** with short summaries so you can scan options faster. |
| **Graph 404 guidance** | When Microsoft Graph returns **404** on a request that still looks like **v1.0** (not **beta**), the CLI appends a **one-line tip**: retry with **`--beta`** on commands that support it, or point **`GRAPH_BASE_URL`** at the beta root—see [docs/CLI_REFERENCE.md](docs/CLI_REFERENCE.md). This does not change routing; it only clarifies a common “wrong API version” situation. |
| **Documentation** | [README.md](README.md) **Supported workloads** expanded with clearer command-to-scenario mapping; [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) notes how grouped help is wired from the registry. |
| **Reliability and maintainers** | **Trailing slashes** on the configured Graph base URL are stripped with a **CodeQL-safe** implementation (no risky regex on user-controlled strings). Generated **[docs/GRAPH_PATH_INVENTORY.json](docs/GRAPH_PATH_INVENTORY.json)** refreshed; tests and formatting aligned with **Biome** and stricter **`fetch`** mock typing. |

### User-facing details

- **Root help (`m365-agent-cli --help`)** — New **`m365-help`** and **`root-command-groups`** plumbing reads the live command registry and prints **sections** (for example sign-in, calendar, mail, files, Teams, tasks, Graph utilities). Same binaries and flags; only the **layout and copy** of help changed.
- **Subcommand help** — **`teams`**, **`files`**, **`calendar`**, **`mail`**, and **`create-event`** include **grouped subcommands** via **`subcommand-help-groups`**, plus **`addHelpText`** / summary lines where it improves scanability. Run **`m365-agent-cli <command> --help`** to see the new layout.
- **Beta-only APIs** — If you call a path that exists only under **Microsoft Graph beta** while the client is still on **v1.0**, a **404** response now includes the **beta hint** above (when the response URL looks like v1, not beta). Prefer explicit **`--beta`** or **`GRAPH_BASE_URL`** for beta-only flows as documented.
- **README** — The **Supported workloads** table lists more top-level commands in context (sign-in, calendar, mail, files, Teams, etc.) so newcomers map **tasks → commands** without opening every `--help` first.

### Upgrading from 2026.5.50

1. **Install:** `npm install -g m365-agent-cli@2026.5.51` or `@latest`.
2. **Scripts / automation:** Output of **`--help`** is **different** (grouped). If you parse help text with brittle string matching, prefer **`--json`** where the command supports it, or match on **command names** rather than exact help layout.
3. **No auth changes** for this patch.

### Compare on GitHub

**`v2026.5.50...v2026.5.51`** — [compare on GitHub](https://github.com/markus-lassfolk/m365-agent-cli/compare/v2026.5.50...v2026.5.51) after the tag exists.

---

## [2026.5.50] — 2026-05-05

This release ships everything from **PR #228** (“Packaging, Graph ergonomics, docs, and expanded CLI surface”): a much wider Microsoft Graph footprint, first-class support for **AI agents and automation** (OpenClaw skill in the npm package, optional MCP server, scripting inventories), and stronger tests and CI. If you are upgrading from **2026.4.50**, you usually only need `npm install -g m365-agent-cli@latest` (or your usual install path); use **`m365-agent-cli login`** again if you enable new Entra permissions for Copilot, Viva, recordings, or approvals.

### Highlights (what changed at a glance)

| Area | What you get |
|------|----------------|
| **Agents & skills** | Bundled **OpenClaw / ClawHub** skill metadata in the npm tarball; optional **postinstall** copy into `OPENCLAW_SKILLS_DIR`; **`TOOLS.md` patcher** for agent tool lists. |
| **MCP** | New **`packages/m365-agent-cli-mcp`**: stdio MCP server that shells to the CLI for **`whoami`**, **`graph-search`** (with **`--json-hits`**), and **read-only** **`graph invoke` GET**. |
| **Copilot & Viva** | Large **`copilot`** command tree (Graph-backed Copilot APIs) and **`viva`** (employee experience, tenant/user surfaces, meeting Engage, learning, insights—many paths are **beta**; see **`--help`** and [docs/GRAPH_SCOPES.md](docs/GRAPH_SCOPES.md)). |
| **Meetings & recordings** | **`meeting`** commands for online meetings, **tenant-style recordings/transcripts** (`getAllRecordings` / `getAllTranscripts`), downloads, and **delta sync** with **`--state-file`** (including **initial delta without a date window** when only the organizer is required, per Graph delta docs). |
| **Files & SharePoint** | Richer **OneDrive / SharePoint** flows (including large-file **upload sessions**, **drive batch** helpers, **async copy** polling improvements, and more **`files` / `sharepoint` / `site-pages`** options). |
| **Excel & Office** | Deeper **Excel** on drive items (tables, charts, names, sessions, comments) and **Word / PowerPoint** drive-backed editing via shared **office-docs** plumbing. |
| **Teams & Planner** | Many new or expanded **Teams** subcommands (messages, reactions, notifications, compose helpers with **`@` mentions**), **Planner**, **To Do**, **presence**, **Bookings**, **groups**, **org**, **people**, **insights**, **mailbox settings**, **approvals**. |
| **Contacts & mail** | Expanded **contacts** (folders, delta, merge suggestions, extensions) and ongoing Graph alignment for mail/calendar paths. |
| **Docs & compliance** | New guides (**agent workflows**, **Graph invoke boundaries**, **troubleshooting**, **Word/PowerPoint editing**, **delegation**), **Graph path inventory**, **permission feature matrix** generation, and **OpenAPI compliance** scripting for maintainers. |

---

### Packaging and OpenClaw (for users who use skills / agents)

- The **npm package** now ships **`skills/m365-agent-cli/SKILL.md`**, **`skills/README.md`**, **`packaging/tools-md-snippet.md`**, and the installer scripts **`scripts/install-tools-md.mjs`** and **`scripts/install-openclaw-skill.mjs`**, so a normal **`npm install -g m365-agent-cli`** includes the same skill metadata the repo uses.
- **`npm run install-tools-md -- <path-to-TOOLS.md>`** (or **`node scripts/install-tools-md.mjs …`**) updates one **HTML-comment–delimited** block inside your **`TOOLS.md`** so you do not get duplicate snippets on re-run.
- **Postinstall is opt-in:** if **`OPENCLAW_SKILLS_DIR`** is set, **`npm install`** copies the bundled skill into that directory; if it is unset, postinstall does nothing (safe for CI and minimal installs).

---

### MCP server (Model Context Protocol)

- New workspace package **`packages/m365-agent-cli-mcp`** (see its **README**): run a small **Node** stdio server that exposes tools such as **`m365_whoami`**, **`m365_graph_search`**, and a read-only **`m365_graph_invoke_get`** wrapper around **`m365-agent-cli graph invoke`**.
- Configure your MCP client with the **`M365_AGENT_CLI_BIN`** / config paths described in **`packages/m365-agent-cli-mcp/README.md`** so the server invokes **your** CLI binary and config directory.

---

### Microsoft Graph: new or substantially expanded commands

Use **`m365-agent-cli <command> --help`** for exact flags; permissions are summarized in **[docs/GRAPH_SCOPES.md](docs/GRAPH_SCOPES.md)**.

- **`approvals`** — Approval workflows (beta Graph), including steps and responses where exposed by the CLI.
- **`copilot`** — Broad surface for **Microsoft 365 Copilot**-related Graph APIs (retrieval, search, chat, packages, meeting insights, interaction export, etc.); many operations need **beta** endpoints and explicit consent.
- **`viva`** — **Microsoft Viva / employee experience**: user and tenant-oriented subcommands (learning, insights, roles, work hours/locations, meeting Engage, tenant admin paths). Subcommands are split across **`viva`**, **`viva-extra-subcommands`**, and **`viva-tenant-subcommands`** in the codebase; the root CLI presents them under **`viva`**.
- **`meeting`** — Online meetings, **recordings** and **transcripts** (per-meeting and **getAllRecordings** / **getAllTranscripts** rollups), **content download**, and **delta** sync with **`--state-file`**.
- **`groups`** — Group-oriented Graph operations exposed by the CLI.
- **`insights`** — Item insights / followed sites–style operations (as implemented in this release).
- **`org`** — Organization directory lookups (as implemented in this release).
- **`people`** — People / profile-related Graph helpers.
- **`mailbox-settings`** — Mailbox settings (automatic replies, regional options, etc., per implemented flags).
- **`word`** / **`powerpoint`** — Entry points that register **drive-backed** document flows (see **[docs/WORD_POWERPOINT_EDITING.md](docs/WORD_POWERPOINT_EDITING.md)**).
- **`excel`** — Expanded workbook/worksheet/table/chart/range/session/comments coverage on drive items.
- **`teams`** — Expanded channel and chat messaging, replies, reactions, activity notifications, and **mention** ergonomics (**`--at userId:displayName`**) for **`--text`** bodies.
- **`files`** — Upload sessions, batch operations, copy/move with monitoring, and related Graph client improvements.
- **`sharepoint`** / **`site-pages`** — Deeper site, list, and page helpers.
- **`planner`** / **`todo`** — More plans, buckets, tasks, and lists (including delta and open extensions where supported).
- **`contacts`** — Folders, delta, photos, attachments, open extensions, merge suggestions.
- **`graph-search`** — **`--json-hits`** for **flattened** search hits suitable for scripts and MCP.
- **`graph`** / **`graph-calendar`** — Invoke, batch, and calendar helpers aligned with new **[docs/GRAPH_INVOKE_BOUNDARIES.md](docs/GRAPH_INVOKE_BOUNDARIES.md)** guidance (paths relative to **`GRAPH_BASE_URL`**, use **`-X` / `--method`** for HTTP method).

Under the hood, large additions landed in **`src/lib/graph-client.ts`** (including async job polling and drive batch patterns), **`graph-teams-client.ts`**, **`graph-excel-client.ts`**, **`copilot-graph-client.ts`**, **`graph-viva-*.ts`**, **`graph-meeting-recordings-client.ts`**, **`graph-delta-state-file.ts`**, and many more libraries—see **[docs/GRAPH_PATH_INVENTORY.json](docs/GRAPH_PATH_INVENTORY.json)** for a generated index of Graph call sites.

---

### Agent-friendly scripting and safety

- **[docs/AGENT_WORKFLOWS.md](docs/AGENT_WORKFLOWS.md)** — End-to-end patterns: auth, **`--read-only`**, default drive roots, **delta** with **`--state-file`**, Teams and files, Word/PowerPoint round-trips, search → drive item.
- **[docs/CLI_SCRIPTING_APPENDIX.md](docs/CLI_SCRIPTING_APPENDIX.md)** and generated **[docs/CLI_SCRIPTING_INVENTORY.md](docs/CLI_SCRIPTING_INVENTORY.md)** — Refresh with **`npm run inventory:scripting`**: maps commands to **`--json`** support and **`checkReadOnly`** behavior so agents know what is safe to run.
- **`counter --json`** — Stable JSON success payload for trivial automation checks.
- **Teams mentions** — For **`channel-message-send`**, **`channel-message-reply`**, **`chat-message-send`**, and **`chat-message-reply`**, repeatable **`--at userId:displayName`** pairs with **`@displayName`** in **`--text`** produce compatible **`mentions`** payloads.

---

### Authentication, cache, and scopes

- **Delegated scopes** in **`src/lib/graph-oauth-scopes.ts`** were extended for new areas (Copilot packages, meeting recordings/transcripts, approvals, app catalog, learning, engagement roles, etc.); duplicates were cleaned up where noted in review.
- **Token cache** behavior and **`M365_AGENT_CLI_CONFIG_DIR`** handling were tightened (prefer predictable, absolute config dirs when overriding).
- If you use **only a subset** of features, you can still **`login`** once; for **Copilot**, **Viva beta**, **recordings**, or **approvals**, expect to **add permissions in Entra** and **re-consent** as needed.

---

### Reliability, tests, and CI

- **Tests** run with **`--isolate`** in **`npm run test:coverage`** to reduce **`mock.module`** and **`fetch`** leakage between files; many **`fetch`** mocks use **`as unknown as typeof fetch`** for TypeScript compatibility with DOM **`fetch`** typing.
- **Graph auth tests** restore real module implementations after mocks; disk-cache cases remain in **`src/test/auth.test.ts`** without a duplicate **`zzz-*`** file.
- **Lib-only coverage gate** (**`scripts/check-coverage-lib.mjs`**) measures **`src/lib/**`** (excluding **`ews-client.ts`**) with correct **LF / DA** merging for duplicate **`SF:`** records; **`scripts/report-lib-coverage-gaps.mjs`** lists the worst-covered files. CI **`COVERAGE_MIN_LINES_LIB`** is **50%** (see **`.github/workflows/ci.yml`** and **`.github/workflows/release.yml`**—they stay aligned).
- Maintainer scripts: **`graph-call-inventory.mjs`**, **`graph-openapi-compliance.mjs`**, **`generate-graph-permission-feature-matrix.ts`**, **`run-bun-test.mjs`** improvements.

---

### Documentation added or expanded

- **[docs/GRAPH_INVOKE_BOUNDARIES.md](docs/GRAPH_INVOKE_BOUNDARIES.md)** — What **`graph invoke`** is for, path rules, and beta caveats.
- **[docs/GRAPH_TROUBLESHOOTING.md](docs/GRAPH_TROUBLESHOOTING.md)** — Common Graph errors and fixes.
- **[docs/GRAPH_PRODUCT_PARITY_MATRIX.md](docs/GRAPH_PRODUCT_PARITY_MATRIX.md)**, **[docs/GRAPH_WRAPPER_GAP_AUDIT.md](docs/GRAPH_WRAPPER_GAP_AUDIT.md)** — Product and wrapper coverage views.
- **[docs/GRAPH_PERMISSION_FEATURE_MATRIX.md](docs/GRAPH_PERMISSION_FEATURE_MATRIX.md)** — Generated permission × feature matrix (see **`npm run docs:graph-permission-matrix`**).
- **[docs/PERSONAL_ASSISTANT_DELEGATION.md](docs/PERSONAL_ASSISTANT_DELEGATION.md)**, **[docs/PHASE6_EWS_REMOVAL.md](docs/PHASE6_EWS_REMOVAL.md)** — Roadmap and delegation notes.
- **[docs/CLI_REFERENCE.md](docs/CLI_REFERENCE.md)**, **[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)**, **[README.md](README.md)** — Updated for the larger surface.

---

### Upgrading from 2026.4.50

1. **Install:** `npm install -g m365-agent-cli@latest` (or `m365-agent-cli update` if you use the self-update path).
2. **Scopes:** open **[docs/GRAPH_SCOPES.md](docs/GRAPH_SCOPES.md)** and your Entra app; add permissions for any new areas you need (Copilot, Viva beta, **OnlineMeetingRecording.Read.All**, **OnlineMeetingTranscript.Read.All**, approvals, etc.), then run **`m365-agent-cli login`** again.
3. **Agents:** optionally set **`OPENCLAW_SKILLS_DIR`** so postinstall installs the bundled skill; configure **MCP** via **`packages/m365-agent-cli-mcp/README.md`** if you use Claude Desktop / Cursor / other MCP hosts.
4. **Graph invoke in scripts:** use **`graph invoke -X GET "/me/..."`** (path without duplicating **`/v1.0`**); see **[docs/GRAPH_INVOKE_BOUNDARIES.md](docs/GRAPH_INVOKE_BOUNDARIES.md)**.

### Compare on GitHub

**`v2026.4.50...v2026.5.50`** — [compare on GitHub](https://github.com/markus-lassfolk/m365-agent-cli/compare/v2026.4.50...v2026.5.50) after the tag exists.

---

## [2026.4.50] — 2026-04-04

### Highlights

- **Microsoft Graph first, EWS when needed.** Set **`M365_EXCHANGE_BACKEND`** to `graph` (Graph only), `ews` (EWS only), or **`auto`** (try Graph, fall back to EWS). Default is **`auto`**, aligned with Exchange Online’s move away from EWS over time.
- **One sign-in, one refresh token, one cache file.** Prefer **`M365_REFRESH_TOKEN`** in your environment; legacy `GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN` still work. Access tokens for EWS and Graph live in **`token-cache-{identity}.json`** (default identity `default`), with migration from older `graph-token-cache-*.json` files.
- **Many more Graph-backed commands** — Teams, Bookings, Excel on OneDrive, presence, Microsoft Search, raw **`graph invoke`** / **`graph batch`**, contacts, OneNote, online meetings, and more — with documentation in-repo ([docs/GRAPH_SCOPES.md](docs/GRAPH_SCOPES.md), [docs/MIGRATION_TRACKING.md](docs/MIGRATION_TRACKING.md)).

### Authentication and Entra app

- Canonical **delegated Graph scopes** live in **`src/lib/graph-oauth-scopes.ts`** and are documented in [docs/GRAPH_SCOPES.md](docs/GRAPH_SCOPES.md) (including **\*.Shared** scopes for delegated mail/calendar, **Place.Read.All**, **People.Read**, **User.Read.All**, Teams, Bookings, presence, OneNote, etc.).
- **`m365-agent-cli login`** uses those scopes; **`verify-token`** can show raw `scp` or **`--capabilities`** for a feature matrix. Entra setup scripts (Bash / PowerShell) and [docs/ENTRA_SETUP.md](docs/ENTRA_SETUP.md) cover a full permission list, beta app / **`.env.beta`** workflows, and PowerShell 7.4 LTS notes.
- **JWT / cache safety:** refresh prefers critical scopes; cache can be invalidated when the token’s app id does not match **`EWS_CLIENT_ID`** or when delegated scopes are too narrow (e.g. after moving between machines).

### Calendar and meetings

- Graph-backed **`calendar`**, **`create-event`**, **`update-event`**, **`delete-event`** (including recurring **`--scope this`** / **`future`**, Teams links, room / Places resolution, attachments).
- **`calendar`**: **`--now`** (hide meetings that already ended today), **`--next-business-days`** (alias for business-day windows), typo-tolerant **`--busness-days`**.
- **`findtime`** / schedule helpers: Graph **`findMeetingTimes`**, **`getSchedule`**, merged availability, work-hours and timezone fixes.
- **`delegates`**: Graph calendar permissions where applicable; EWS remains for some delegate operations.

### Mail, drafts, send, folders

- Graph-first listing, read, send, and folder operations under **`auto`** / **`graph`**, with clear errors and **EWS fallback** in **`auto`** when Graph cannot satisfy the request.
- **Shared / delegated mailboxes:** use **`--mailbox`** plus the correct **\*.Shared** Graph scopes (see [docs/GRAPH_SCOPES.md](docs/GRAPH_SCOPES.md)).

### New or expanded command areas

- **Contacts** (Graph), **OneNote** (Graph-only), **online meetings** (`meeting`), **Teams** (channels, chats, messages), **Bookings**, **Excel** (worksheets on drive items), **presence**, **`graph-search`** (Microsoft Search), **`graph`** / **`graph-calendar`** (invoke, batch, calendar helpers, delta, etc.).
- **`todo`**, **`planner`**, **`files`**, **`sharepoint`**, **`subscribe`**, and others gained fixes and Graph alignment where noted in migration docs.

### Security and reliability

- Safer attachment and **`.url`** handling (path sanitization, HTTP(S) rules, CodeQL-oriented patterns).
- **Graph URL validation** for absolute URLs (e.g. paging / `nextLink`) to avoid sending tokens to untrusted hosts.
- **GlitchTip / Sentry:** centralized **`beforeSend`** policy to drop noisy network and OAuth failures; release builds embed git SHA for support correlation.

### Developer experience

- Run from source with **Bun** (CI default) or **`tsx`** for the TypeScript entry; **`npm run sync-skill`** keeps **`skills/m365-agent-cli/SKILL.md`** `version` in sync with **`package.json`**.
- CI: typecheck, Biome, tests with coverage floor, Knip, Gitleaks (with documented allowlists where needed).

### Documentation

- [docs/AUTHENTICATION.md](docs/AUTHENTICATION.md), [docs/CLI_REFERENCE.md](docs/CLI_REFERENCE.md), migration and parity docs (**[docs/GRAPH_V2_STATUS.md](docs/GRAPH_V2_STATUS.md)**, **[docs/GRAPH_EWS_PARITY_MATRIX.md](docs/GRAPH_EWS_PARITY_MATRIX.md)**, **[docs/GRAPH_API_GAPS.md](docs/GRAPH_API_GAPS.md)**), [docs/GLITCHTIP.md](docs/GLITCHTIP.md), streamlined [README.md](README.md).

### Upgrading from 1.2.4

1. Upgrade the global package: `npm install -g m365-agent-cli@latest` (or use your usual install path).
2. Prefer **`M365_REFRESH_TOKEN`** in **`~/.config/m365-agent-cli/.env`**; run **`m365-agent-cli login`** again if you add scopes in Entra.
3. Set **`M365_EXCHANGE_BACKEND`** if you need **`graph`** or **`ews`** only; default **`auto`** matches the new Graph-first behavior.
4. Re-read [docs/GRAPH_SCOPES.md](docs/GRAPH_SCOPES.md) if you use **delegated** or **shared** mailboxes.

### Full commit list (since v1.2.4)

See GitHub compare: **`v1.2.4...v2026.4.50`** (after the release tag exists), or browse history on `main` / `dev_v2` for individual commits.

---

## [1.2.4] and earlier

See git tags and [releases](https://github.com/markus-lassfolk/m365-agent-cli/releases) for prior versions.
