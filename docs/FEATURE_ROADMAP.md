# Feature Roadmap

Tracking document for the post-2026.7.4 feature push. Generated from a review of the
codebase's agent-usability gaps and coverage holes. Status is updated as each item lands;
each landed item references the commit(s) that implemented it.

Legend: ⬜ Not started · 🟡 In progress · ✅ Done · ⏭️ Deferred (needs live Exchange / external
verification / a product decision)

## Top picks (agent-usability infrastructure)

| # | Feature | Status | Notes |
|---|---------|--------|-------|
| 1 | `describe` — machine-readable command/option manifest | ⬜ | Reflects the Commander tree to JSON: every command, option (name/type/required/description), and `--json` output shape. Lets an agent discover the surface without parsing `--help` prose. |
| 2 | Native MCP server mode | ⬜ | Expose each CLI command as an MCP tool (built on #1's manifest for schema generation). Extends `serve.ts`. |
| 3 | `--dry-run` for mutations | ⬜ | Print the resolved method/URL/body (Graph) or SOAP envelope (EWS) without sending. Composes with `checkReadOnly`. |
| 4 | Output shaping: `--select`/`--fields` projection + `--ndjson` streaming | ⬜ | Client-side field projection where Graph `$select` doesn't apply; NDJSON mode for large lists so agents can stream-process instead of buffering. |

## Reliability & safety

| # | Feature | Status | Notes |
|---|---------|--------|-------|
| 5 | Structured error envelope in `--json` mode | ⬜ | `{ error: { code, message, retriable, graphCode, requestId } }` instead of a bare string, surfacing what `graph-client.ts` already captures internally. |
| 6 | Auto-chunking Graph `$batch` helper | ⬜ | Accept N requests, transparently split into ≤20-request batches, merge results. |
| 7 | Opt-in read cache with TTL | ⬜ | `--cache <dur>` / `M365_CACHE_TTL` for idempotent GETs (folder lists, room lists, user lookups). |
| 8 | `whoami --capabilities` — decode token scopes → usable command groups | ⬜ | Presentation layer over the scope-decoding `verify-token` already does. |

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
| 14 | `teams chat-message-react --reply` (chat variant lacks what `channel-message-react` has) | ⬜ | |
| 15 | EWS draft BCC parity (`drafts --edit` has `--cc` but no `--bcc`; `parseEmailMessage` doesn't read `BccRecipients`) | ⬜ | |
| 16 | `excel comments-*` `--beta` flag (help text says "requires beta" but no flag exists; only `GRAPH_BETA_URL` works) | ⬜ | |

## Working notes

- Branch: `claude/feature-roadmap-implementation`
- Each item lands as its own commit (or small group of related commits), gated by the full
  local suite (typecheck, biome, knip, tests, coverage, graph inventory, permission matrix)
  before commit.
- Items are implemented roughly in table order, but reordered opportunistically when one
  unblocks another (e.g. #1 before #2).
