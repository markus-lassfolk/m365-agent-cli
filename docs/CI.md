# Continuous integration

Workflows live under [`.github/workflows/`](../.github/workflows/).

## What runs on every push / PR to `main`

| Workflow | Purpose |
|----------|---------|
| **CI** | TypeScript (`tsc --noEmit`), **Biome** (`biome check` = lint + format + assists), tests with **LCOV** coverage, **minimum line coverage** (40% by default, see `scripts/check-coverage.mjs`), **Knip** (unused deps/files/exports), TruffleHog + Gitleaks + Trivy. |
| **CodeQL** | Semantic analysis for TypeScript (`security-extended` query pack). |

PRs also get an **lcov** comment from `romeovs/lcov-reporter-action` when coverage is uploaded.

## Reproducing locally

```bash
bun install --frozen-lockfile
bun run typecheck
bun run biome:check
bun run test:coverage
bun run verify:coverage
bun run knip
```

## Monitoring

- [Actions tab](https://github.com/markus-lassfolk/m365-agent-cli/actions) — green = latest `main` / PR checks passed.
- **Release** workflow runs on tag `v*.*.*` (npm Trusted Publishing + GitHub Release).

## Notes

- **`bun install --frozen-lockfile`** in CI ensures installs match `bun.lock`.
- **Knip** is configured in [`knip.json`](../knip.json); some public API surfaces are ignored to avoid false positives.
- **TruffleHog** uses `continue-on-error: true` (verified hits only); **Gitleaks** remains the strict gate for secret patterns in the repo.
