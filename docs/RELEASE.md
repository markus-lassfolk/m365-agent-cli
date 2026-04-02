# Releases and version control

Releases are aligned with **npm**, **Git tags**, and **GlitchTip** eligibility (see [GLITCHTIP.md](./GLITCHTIP.md)): the published tarball embeds the git SHA, and GitHub tag `v{version}` must point at that commit.

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

**Optional unblock:** add a repository secret **`NPM_TOKEN`** (a [granular access token](https://docs.npmjs.com/creating-and-viewing-access-tokens) with **Read and write** for `m365-agent-cli`). When set, the workflow publishes with **`npm publish --access public`** (no `--provenance`, which expects OIDC-only auth). Remove the secret later to use **Trusted Publishing + provenance** only.

```bash
gh secret set NPM_TOKEN --repo markus-lassfolk/m365-agent-cli
# paste token when prompted
```

### First publish (name not yet on the registry)

Trusted Publishing only applies **after** the package exists under your npm account. For the **initial** `npm publish`:

1. **Create an access token** on npm: [Access tokens](https://docs.npmjs.com/creating-and-viewing-access-tokens) — use a **Granular Access Token** with **Read and write** for the package (or **Automation** classic token if you prefer).
2. **Locally** (or in a one-off CI job), from this repo at the release commit:
   - `npm run embed-sha`
   - `npm pack` (optional sanity check)
   - `npm publish --access public`  
     With a token: `npm config set //registry.npmjs.org/:_authToken=YOUR_TOKEN` (or `NPM_TOKEN` env with `npm publish` per npm docs). Do **not** commit the token.
3. On [npmjs.com](https://www.npmjs.com/) open **`m365-agent-cli` → Package → Access → Trusted publishers** and add **GitHub Actions** with repository `markus-lassfolk/m365-agent-cli` and workflow file **`release.yml`** (see section above).
4. From then on, push tags to trigger [`.github/workflows/release.yml`](../.github/workflows/release.yml); CI publishes via **Trusted Publishing (OIDC)**. Optionally add repository secret **`NPM_TOKEN`** if you need a token-based publish path (see above).

If the package name is already taken on npm, you must rename the package in `package.json` or obtain access from the owner before publishing.

## Cutting a release

1. On `main`, set **`version`** in `package.json` to the new semver (e.g. `1.2.3`). Run **`npm run sync-skill`** so `skills/m365-agent-cli/SKILL.md` frontmatter `version:` matches the package (OpenClaw/ClawHub skill metadata). Commit, for example: `chore: release 1.2.3`.
2. Tag that commit: `git tag v1.2.3` (the tag **must** be `v` + the exact version string from `package.json`).
3. Push the tag: `git push origin v1.2.3` (and push the commit to `main` if needed).

Pushing the tag runs **Release** in GitHub Actions: it checks that the tag matches `package.json`, runs checks, runs **`npm run embed-sha`**, publishes to npm, and creates a GitHub Release with generated notes.

## After publish

Users can upgrade with:

```bash
m365-agent-cli update
```

or:

```bash
npm install -g m365-agent-cli@latest
```

## Local dry run (optional)

From a clean checkout at the release commit:

```bash
npm run embed-sha
npm pack
```

Inspect the tarball; `package.json` and `src/lib/git-commit.ts` should reflect the release.
