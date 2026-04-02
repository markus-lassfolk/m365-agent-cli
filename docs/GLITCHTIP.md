# GlitchTip (optional error reporting)

[m365-agent-cli](https://github.com/markus-lassfolk/m365-agent-cli) can send **crashes and unhandled errors** to a self-hosted or cloud [GlitchTip](https://glitchtip.com/) project using the **Sentry-compatible** Node SDK ([GlitchTip Node docs](https://glitchtip.com/sdkdocs/node)).

## Enable

1. Create a project in GlitchTip and copy the **DSN** (same format as Sentry).
2. Set in the environment or `~/.config/m365-agent-cli/.env`:

| Variable | Meaning |
|----------|---------|
| **`GLITCHTIP_DSN`** or **`SENTRY_DSN`** | Project DSN. If unset, **no** reporting runs. |
| `GLITCHTIP_ENABLED` | Set to `0` or `false` to disable even when DSN is set. |
| `GLITCHTIP_ENVIRONMENT` | e.g. `production`, `ci` (defaults to `NODE_ENV` or `production`). |
| `GLITCHTIP_RELEASE` | Optional **override** for the release tag in GlitchTip. If unset, the CLI sets **`m365-agent-cli@` + version** from the **installed package** (`package.json` next to the running binary — same source as `m365-agent-cli --version`). Use a git SHA here **not** recommended; prefer the default so release tracks the published semver. |

### Example `.env` for agents / production

Copy **[`env.glitchtip.example`](./env.glitchtip.example)** to `~/.config/m365-agent-cli/.env`. You do **not** need to set **`GLITCHTIP_RELEASE`** unless you want a custom label; the running CLI version is applied automatically.

## When reporting runs (release gating)

To avoid flooding GlitchTip from **dev installs**, **pre-release builds**, or **outdated** npm installs, the CLI only initializes Sentry when **all** of the following hold:

1. **`package.json` version** equals the **latest version on npm** for this package (checked against the registry; result cached for about an hour).
2. The **embedded git commit** in the build matches the commit that GitHub’s tag **`v{version}`** points at (the release tag for that npm version).

The **Release** GitHub Action runs **`npm run embed-sha`** before `npm publish`, so published builds embed the tag commit. For manual publishes, run **`npm run embed-sha`** yourself. Tag **`vX.Y.Z`** on GitHub at the same commit you publish to npm (see [RELEASE.md](./RELEASE.md)).

| Variable | Meaning |
|----------|---------|
| **`GLITCHTIP_SKIP_VERSION_CHECK`** | Set to `1` or `true` to **skip** the checks above and report whenever DSN is set (e.g. internal testing). |
| **`GLITCHTIP_ALLOW_UNVERIFIED_COMMIT`** | If the embedded commit is `unknown` (e.g. local dev without `embed-sha`), set to `1` to still allow reporting when the npm version check passes. |
| **`GLITCHTIP_DEBUG_ELIGIBILITY`** | Set to `1` or `true` to print why reporting was disabled to **stderr**. |

## What gets reported

- **Uncaught exceptions** and **unhandled promise rejections** (via `@sentry/node`).
- **Commander parse failures** (e.g. invalid CLI usage that throws).

Tracing is disabled (`tracesSampleRate: 0`) to keep events small and GlitchTip-friendly.

## What is filtered out (by default)

To reduce noise from environment and auth (not “our code bug”):

- Common **network errno** values: `ECONNREFUSED`, `ETIMEDOUT`, `ENOTFOUND`, etc.
- Messages that look like **OAuth refresh / AAD token** failures (`invalid_grant`, `AADSTS…`).

Set **`GLITCHTIP_REPORT_ALL=1`** to **disable** those filters (send everything that Sentry would still accept). **PII scrubbing still applies** (see below); this flag does not re-enable raw argv or user-identifiable fields.

## Privacy and PII

The DSN is a **write-only** project key; it does not grant read access to your GlitchTip org. Do not commit real DSNs to public repos; use env or a private `.env`.

Before an event is sent, the CLI **strips or redacts** data that is not needed to fix code bugs:

- **`sendDefaultPii` is off** — no default IP/user identity from the SDK.
- **No command-line text** — only **`cli.argc`** (argument count) and, when safe, **`cli.command`** (first token if it looks like a subcommand name, e.g. `mail`). Full argv is never included.
- **No user, request, server name, or breadcrumbs** on the outbound payload; breadcrumbs are disabled at capture time.
- **Stack traces** keep file paths with **home directory segments redacted**; **locals (`vars`)** on frames are removed; **exception messages** are run through the same redaction as strings.
- **`extra` / tags** — keys that often carry mail or secrets (e.g. `body`, `subject`, `token`, `email`, …) are dropped; remaining string values are redacted for **emails**, **Bearer/JWT-like tokens**, and **user home paths**.
- **Contexts** are reduced to **OS name/version**, **runtime**, and minimal **app** timing/memory — no device names or free-form app fields.

What remains is mainly **error type**, **redacted message**, **stack frames** (with safe paths), and **environment/release** — enough to locate bugs without usernames, mail bodies, or similar content.

## Verify

After setting `GLITCHTIP_DSN`, run a command that throws inside the CLI (or temporarily break a local branch). You should see the event in GlitchTip within a few seconds.

### DSN / host connectivity

The ingest API lives at `{scheme}://{host}/api/{project_id}/store/` (Sentry-compatible). A quick check that the host and project path respond (expect **405** on `GET` — only `POST` is valid for ingest):

```bash
curl -s -o /dev/null -w "%{http_code}\n" "http://glitchtip.lassfolk.cc/api/6/store/"
```

You should see **`405`**. A connection error or **long timeout** means DNS/firewall or the server is down. (**`403`** on the site root URL alone is normal if the UI is restricted.)

To send a **real test event** (same SDK as the CLI), from the repo root:

```bash
node scripts/test-glitchtip-send.mjs
```

Uses **`GLITCHTIP_DSN`** / **`SENTRY_DSN`** if set; otherwise reads **`GLITCHTIP_DSN`** from **`docs/env.glitchtip.example`**. On success you should see the message in GlitchTip within seconds (search for `connectivity test`).

If nothing appears, set **`GLITCHTIP_DEBUG_ELIGIBILITY=1`** — common reasons are: not on the latest npm version, commit not matching tag `v{version}`, or missing network to npm/GitHub.

For local testing without matching a release, use **`GLITCHTIP_SKIP_VERSION_CHECK=1`** (or **`GLITCHTIP_ALLOW_UNVERIFIED_COMMIT=1`** when commit is `unknown`).
