# Releases and version control

Releases are aligned with **npm**, **Git tags**, and **GlitchTip** eligibility (see [GLITCHTIP.md](./GLITCHTIP.md)): the published tarball embeds the git SHA, and GitHub tag `v{version}` must point at that commit.

## One-time setup (maintainer): npm Trusted Publishing

This repo publishes with **Trusted Publishing** (OIDC from GitHub Actions). You **do not** add an automation token or secret named `NPM_TOKEN` for CI.

1. Sign in to [npmjs.com](https://www.npmjs.com/) and open the **`m365-agent-cli`** package (or create it if the name is first published from your account).
2. Go to **Package → Access** (or **Publishing access**), then **Trusted publishers** / **Configure trusted publisher**.
3. Add **GitHub Actions** and specify:
   - **Organization or user:** `markus-lassfolk` (or your fork’s owner if you publish from a fork)
   - **Repository:** `m365-agent-cli`
   - **Workflow filename:** `release.yml`
   - **Environment:** leave empty unless you use a [GitHub Environment](https://docs.github.com/en/actions/deployment/targeting-different-environments/using-environments-for-deployment) with protection rules (then match the name here and in the workflow).

Official npm docs: [Trusted publishers](https://docs.npmjs.com/trusted-publishers).

The workflow [`.github/workflows/release.yml`](../.github/workflows/release.yml) uses `permissions: id-token: write` and runs `npm publish --access public --provenance` so the package is published with [provenance](https://docs.npmjs.com/generating-provenance-statements) (no long-lived npm token in GitHub secrets).

If publish fails with an authentication error, the trusted publisher is usually misconfigured (wrong repo, workflow name, or package owner on npm).

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
