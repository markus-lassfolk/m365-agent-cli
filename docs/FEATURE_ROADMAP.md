# Feature Roadmap

Tracking document for the post-2026.7.4 feature push. Generated from a review of the
codebase's agent-usability gaps and coverage holes. Status is updated as each item lands;
each landed item references the commit(s) that implemented it.

Legend: ⬜ Not started · 🟡 In progress · ✅ Done · ⏭️ Deferred (needs live Exchange / external
verification / a product decision)

## Top picks (agent-usability infrastructure)

| # | Feature | Status | Notes |
|---|---------|--------|-------|
| 1 | `describe` — machine-readable command/option manifest | ✅ | `m365 describe` (full JSON manifest), `--list` (fast top-level overview), `--command "rules create"` (scoped lookup). `src/lib/command-manifest.ts` + `src/commands/describe.ts`. |
| 2 | Native MCP server mode | ⬜ | Expose each CLI command as an MCP tool (built on #1's manifest for schema generation). Extends `serve.ts`. |
| 3 | `--dry-run` for mutations | ✅ | Transport-level: `callGraphAt`/`callEws` halt and print the resolved request instead of sending it, gated by `M365_DRY_RUN` (synced from the root `--dry-run` flag via a Commander `preAction` hook). Works uniformly for all ~60 commands with no per-command wiring. For multi-step flows, only the first mutating request is shown. `src/lib/dry-run.ts`. |
| 4 | Output shaping: `--select`/`--fields` projection + `--ndjson` streaming | ⬜ | Client-side field projection where Graph `$select` doesn't apply; NDJSON mode for large lists so agents can stream-process instead of buffering. |

## Reliability & safety

| # | Feature | Status | Notes |
|---|---------|--------|-------|
| 5 | Structured error envelope in `--json` mode | ✅ | `{ error: { message, code?, status?, retriable?, requestId? } }` (mirrors Graph's own error shape) instead of a bare string. `toJsonError()` in `src/lib/json-error.ts` normalizes any input (GraphError/OwaError object, string, Error, undefined); applied at all ~168 `--json` error print sites across 21 command files. Sites that already had the full error object (not just its pre-extracted `.message`) get full `code`/`status`/`requestId` fidelity; sites that only ever held a message string still get the consistent shape but without those extra fields — a follow-up could thread the full error object through more call sites for richer fidelity everywhere. |
| 6 | Auto-chunking Graph `$batch` helper | ⬜ | Accept N requests, transparently split into ≤20-request batches, merge results. |
| 7 | Opt-in read cache with TTL | ⬜ | `--cache <dur>` / `M365_CACHE_TTL` for idempotent GETs (folder lists, room lists, user lookups). |
| 8 | `whoami --capabilities` — decode token scopes → usable command groups | ✅ (investigation reversed the premise) | Traced `verify-token --capabilities`: it already fully implements token-scope → per-area read/write decoding via `graph-capability-matrix.ts`. Duplicating that in `whoami` would fork the logic. Instead `whoami` (both Graph and EWS paths, `--json` and text) now carries a `hint`/`Tip:` cross-reference pointing agents at `verify-token --capabilities`. |

## New workload / API coverage

| # | Feature | Status | Notes |
|---|---------|--------|-------|
| 9 | OneDrive/SharePoint sharing & permissions (create/list/revoke links, manage item permissions) | ⬜ | |
| 10 | Bulk mutation commands (ID-list / `--filter` driven) for mail/todo/planner | ⬜ | Avoid one-call-per-item loops. |
| 11 | `findtime` multi-attendee scheduling depth (working hours, timezone, free/busy in one call) | ⬜ | |
| 12 | Reusable mail/draft templates with variable substitution | ⬜ | |

## Quick wins / gaps found during QA review

| # | Feature | Status | Notes |
|---|---------|--------|-------|
| 13 | EWS `createEvent` attendees+attachments ordering fix | ⏭️ | Documented as a known limitation in `ews-client.ts` (2026.7.4). Needs a live Exchange to verify the create→attach→resend-invite flow before shipping. |
| 14 | `teams chat-message-react --reply` (chat variant lacks what `channel-message-react` has) | ✅ | Added `--reply <replyId>` to `chat-message-react` (both set/unset), targeting `/chats/{id}/messages/{id}/replies/{id}/(un)setReaction` — verified against Microsoft's Graph SDK endpoint reference before implementing, matching the existing `channel-message-react --reply` pattern. |
| 15 | EWS draft BCC parity (`drafts --edit` has `--cc` but no `--bcc`; `parseEmailMessage` doesn't read `BccRecipients`) | ✅ | Added `--bcc` to `drafts` (both `--create` and `--edit`, EWS backend); `updateDraft` gained a `bcc` param + `message:BccRecipients` `SetItemField`; `getEmail` now requests and parses `BccRecipients` (visible in `drafts --read --json`). |
| 16 | `excel comments-*` `--beta` flag | ✅ (investigation reversed the premise) | Traced `graph-excel-comments-client.ts`: all 5 `comments-*` functions already call `getGraphBetaUrl()` unconditionally — `workbookComment` has no v1.0 equivalent, so they're *always* beta regardless of any flag. A `--beta` flag would have been a **no-op that misleads users** into thinking it's needed/optional. Fixed the actual bug instead: reworded all 5 command descriptions from the ambiguous "requires beta or GRAPH_BETA_URL" to state plainly that they always use the beta root and only `GRAPH_BETA_URL` customizes *which* beta host. Added regression tests locking in the always-beta behavior and the `GRAPH_BETA_URL` override. |

## Working notes

- Branch: `claude/feature-roadmap-implementation`
- Each item lands as its own commit (or small group of related commits), gated by the full
  local suite (typecheck, biome, knip, tests, coverage, graph inventory, permission matrix)
  before commit.
- Items are implemented roughly in table order, but reordered opportunistically when one
  unblocks another (e.g. #1 before #2).
