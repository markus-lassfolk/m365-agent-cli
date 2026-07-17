# Releases and version control

Releases are aligned with **npm**, **Git tags**, and **GlitchTip** eligibility (see [GLITCHTIP.md](./GLITCHTIP.md)): the published tarball embeds the git SHA, and GitHub tag `v{version}` must point at that commit.

**Why isn’t there a new npm version after every merge?** Publishing runs only when someone pushes a **version tag** (`v…` matching `package.json`). Merging to `main` updates the repo but does **not** trigger npm — that is intentional so releases stay deliberate and tagged. See **Cutting a release** below.

User-facing changes for each version are summarized in the root [`CHANGELOG.md`](../CHANGELOG.md).

## One-time setup (maintainer): npm Trusted Publishing

This repo publishes from [`.github/workflows/release.yml`](../.github/workflows/release.yml) with **`npm publish --access public --provenance`** and `permissions: id-token: write`.

1. Sign in to [npmjs.com](https://www.npmjs.com/) and open the **`m365-agent-cli`** package (or create it if the name is first published from your account).
2. Go to **Package → Access** (or **Publishing access**), then **Trusted publishers** / **Configure trusted publisher**.
3. Add **GitHub Actions** and specify:
   - **Organization or user:** `markus-lassfolk` (or your fork’s owner if you publish from a fork)
   - **Repository:** `m365-agent-cli`
   - **Workflow filename:** `release.yml`
   - **Environment:** leave empty unless you use a [GitHub Environment](https://docs.github.com/en/actions/deployment/targeting-different-environments/using-environments-for-deployment) with protection rules (then match the name here and in the workflow).
4. Click **Save changes** on npm and confirm the trusted publisher still appears after refresh.

Official npm docs: [Trusted publishers](https://docs.npmjs.com/trusted-publishers).

### If publish still returns `E404` / `Not found` after saving npm

The workflow **does not** use `actions/setup-node`’s `registry-url` input (that can inject an empty `_authToken` into `.npmrc` and break OIDC).

**Optional unblock (temporary):** add a repository secret **`NPM_TOKEN`** — a [granular access token](https://docs.npmjs.com/creating-and-viewing-access-tokens) with **Read and write** for `m365-agent-cli`. For **CI**, the token must have **Bypass two-factor authentication (2FA)** enabled; otherwise automated `npm publish` fails (often with misleading `E404`/`E403`). npm warns about bypass; that is expected until you move to Trusted Publishing.

When set, the workflow publishes with **`npm publish --access public`** (no `--provenance`). **When Trusted Publishing is verified:** revoke the token on npm, delete the **`NPM_TOKEN`** secret — the next tag run uses **OIDC only** (`npm publish` with provenance). The Release workflow uses **Node 24** and **npm ≥11.5.1**, which npm requires for Trusted Publishing.

```bash
gh secret set NPM_TOKEN --repo markus-lassfolk/m365-agent-cli
# paste token when prompted
```

### First publish (name not yet on the registry)

Trusted Publishing only applies **after** the package exists under your npm account. For the **initial** `npm publish`:

1. **Create an access token** on npm: [Access tokens](https://docs.npmjs.com/creating-and-viewing-access-tokens) — use a **Granular Access Token** with **Read and write** for the package (or **Automation** classic token if you prefer).
2. **Locally** (or in a one-off CI job), from this repo at the release commit:
   - `npm run embed-sha`
   - `npm pack` (optional sanity check — expect `dist/cli.js`, `dist/index.js`, `skills/m365-agent-cli/SKILL.md`, `packaging/tools-md-snippet.md`, and `scripts/install-*.mjs` in the file list)
   - `npm publish --access public`  
     With a token: `npm config set //registry.npmjs.org/:_authToken=YOUR_TOKEN` (or `NPM_TOKEN` env with `npm publish` per npm docs). Do **not** commit the token.
3. On [npmjs.com](https://www.npmjs.com/) open **`m365-agent-cli` → Package → Access → Trusted publishers** and add **GitHub Actions** with repository `markus-lassfolk/m365-agent-cli` and workflow file **`release.yml`** (see section above).
4. From then on, push tags to trigger [`.github/workflows/release.yml`](../.github/workflows/release.yml); CI publishes via **Trusted Publishing (OIDC)**. Optionally add repository secret **`NPM_TOKEN`** if you need a token-based publish path (see above).

If the package name is already taken on npm, you must rename the package in `package.json` or obtain access from the owner before publishing.

## Cutting a release

Releases are cut automatically — see **Automatic releases** below. You only need to do this manually if automation is disabled or you're releasing from a branch other than `main`.

1. On `main` (or your release branch), set **`version`** in `package.json` to the new version (semver or calendar-style, e.g. `2026.4.50`). Update **`CHANGELOG.md`** for the release. Run **`npm run sync-skill`** so `skills/m365-agent-cli/SKILL.md` frontmatter `version:` matches the package (OpenClaw/ClawHub skill metadata). Commit, for example: `chore(release): 2026.4.50`.
2. Tag that commit: `git tag v2026.4.50` (the tag **must** be `v` + the exact version string from `package.json`).
3. Push the tag: `git push origin v2026.4.50` (and push the commit to `main` if needed).

Pushing the tag runs **Release** in GitHub Actions (see [`.github/workflows/release.yml`](../.github/workflows/release.yml)): it checks that the tag matches `package.json`, runs checks, runs **`npm run embed-sha`**, publishes to npm with **`npm publish --access public`** (Trusted Publishing OIDC when **`NPM_TOKEN`** is unset), and creates a GitHub Release (with auto-generated notes; you can paste from **`CHANGELOG.md`** if you want richer text).

## Automatic releases

[`.github/workflows/auto-release.yml`](../.github/workflows/auto-release.yml) watches every **CI** run on `main`. When CI finishes green and `package.json`'s `version` has no matching `vX.Y.Z` tag yet, it creates that tag on the green commit and dispatches **Release** (via `workflow_dispatch`, since the default `GITHUB_TOKEN` used to push the tag doesn't itself fire the `push: tags` trigger — see the comments in both workflow files). Every other green CI run on `main` (no version bump) is a no-op.

In practice: merge a `chore(release): X` PR (step 1 above, without steps 2–3) to `main`, and once CI passes the release ships on its own — no local tag push needed. Steps 2–3 above remain available as a manual fallback (e.g. from an environment without push access to trigger Actions, or to release a commit that isn't `main`'s tip).

## After publish

Users can upgrade with:

```bash
m365-agent-cli update
```

or:

```bash
npm install -g m365-agent-cli@latest
```

## Build

`npm pack` / `npm publish` compile `src/` to a Node-runnable `dist/` via the `prepack` script
(`npm run build`, i.e. `tsc -p tsconfig.build.json` plus a `#!/usr/bin/env node` shebang on
`dist/cli.js`) — this is what `bin`/`main` in `package.json` point at, so `npm install -g
m365-agent-cli` produces a working executable without a separately installed Bun runtime (see
issue #239). `dist/` is git-ignored and generated fresh on every pack/publish; there is nothing to
commit for it.

## Local dry run (optional)

From a clean checkout at the release commit:

```bash
npm run embed-sha
npm pack
```

`npm pack` triggers the build above automatically. Inspect the tarball; `package.json` and
`src/lib/git-commit.ts` (compiled into `dist/lib/git-commit.js`) should reflect the release.
