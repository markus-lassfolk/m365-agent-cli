# Changelog

All notable changes to **m365-agent-cli** are documented here. Release **2026.4.50** is the first stable line after **1.2.4** that ships the Graph-first stack, unified auth, and the expanded command surface described below.

For install and tagging, see [docs/RELEASE.md](docs/RELEASE.md).

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
