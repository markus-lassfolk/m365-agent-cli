# Release notes — m365-agent-cli **2026.5.51**

**Release date:** 2026-05-05  
**Type:** Patch (follow-up to **2026.5.50**; see [CHANGELOG.md](../CHANGELOG.md))

This document is a **user-oriented** summary of **2026.5.51**. The authoritative per-file change list remains **[CHANGELOG.md](../CHANGELOG.md)**.

---

## Summary in one paragraph

Version **2026.5.51** makes the CLI **easier to navigate** through **grouped `--help`** at the root and for several large command trees, adds a **practical hint when Microsoft Graph returns 404** on what looks like a **v1.0** URL (often a sign you need **beta** or a different `GRAPH_BASE_URL`), improves the **README workload map**, and includes small **hardening and housekeeping** changes (safe trailing-slash trimming on Graph base URLs, refreshed Graph path inventory, test and formatter alignment). It does **not** introduce new OAuth scopes or new top-level product areas compared to **2026.5.50**.

---

## Who should upgrade

- **Everyone on 2026.5.x** — Safe, low-risk patch; recommended for the help and Graph error clarity alone.
- **Interactive terminal users** — You will notice **clearer `--help`** immediately.
- **Automation authors** — If any script **parses `--help` text** with fixed strings or line positions, re-check those scripts; prefer **`--json`** or stable **subcommand names** where available.

---

## What you will notice

### 1. Grouped help at the top level

Running:

```bash
m365-agent-cli --help
```

now shows commands arranged in **named groups** (for example sign-in and meta commands, calendar and meetings, mail, files and SharePoint, Teams, tasks, Graph utilities). The goal is to answer “**where is the command for X?**” without scrolling an alphabetical list of every subcommand in the product.

### 2. Richer help on big command trees

For **`teams`**, **`files`**, **`calendar`**, **`mail`**, and **`create-event`**, **`--help`** is organized so related subcommands sit together with **short descriptions**. This mirrors how people think about work (**channels**, **chats**, **uploads**, **calendars**) rather than how the implementation happens to register commands.

### 3. Clearer message when Graph says “not found”

Some Microsoft Graph capabilities exist only under the **beta** endpoint. If the CLI issues a **v1.0** request and Graph responds with **404**, the error output may now include a **single-line tip**: try **`--beta`** on commands that support it, or set **`GRAPH_BASE_URL`** to the beta root, with a pointer to **CLI_REFERENCE**. The CLI does not automatically switch API versions for you; it only **explains a frequent cause** of confusing 404s.

### 4. README “Supported workloads”

The main **[README.md](../README.md)** maps **common goals** to **commands** more explicitly, so new users can find the right entry point before diving into nested help.

---

## Operational and security notes

- **Graph base URL** — Trailing slashes on the configured base URL are normalized using logic that satisfies **static analysis (CodeQL)** for safe string handling. Behavior for well-formed URLs is unchanged; edge cases with duplicated slashes at the end are tidied consistently.
- **Permissions** — No new **Entra ID** delegated or application permissions are required for this release compared to **2026.5.50**.
- **OpenClaw skill version** — The bundled skill’s frontmatter **`version:`** matches **`package.json`** (**2026.5.51**). If you rely on **`npm run sync-skill`** in your fork, run it whenever you bump the package version.

---

## Install and verify

```bash
npm install -g m365-agent-cli@2026.5.51
m365-agent-cli --version
m365-agent-cli --help
```

If you use **Git tags** for releases, create **`v2026.5.51`** on the release commit per **[docs/RELEASE.md](RELEASE.md)**.

---

## References

| Resource | Purpose |
|----------|---------|
| [CHANGELOG.md](../CHANGELOG.md) | Full changelog including **2026.5.51** and **2026.5.50** |
| [docs/CLI_REFERENCE.md](CLI_REFERENCE.md) | Flags, env vars, and Graph invoke conventions |
| [docs/GRAPH_SCOPES.md](GRAPH_SCOPES.md) | OAuth scopes by feature area |
| Compare **`v2026.5.50...v2026.5.51`** on GitHub | Exact code diff after tags exist |
