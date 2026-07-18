# Unattended device-code login — reference example

Adaptable, secret-free scripts that complete an `m365-agent-cli` device-code sign-in **without a human**,
by driving a headless browser. **This is example code to copy and adapt — it is not installed or run by the
CLI.**

Read **[../../docs/UNATTENDED_LOGIN.md](../../docs/UNATTENDED_LOGIN.md)** first — especially the security
tradeoff (this needs a stored **password** *and* **TOTP seed**, which collapses MFA into stored secrets)
and the section on when *not* to use it.

## Files

- **`device-login.mjs`** — Playwright automation of the Microsoft sign-in pages (code → account → password
  → consent → TOTP → done). Reads everything from env vars; writes a screenshot only on failure, into a
  temp dir.
- **`refresh-token.sh`** — orchestration: fetch creds from *your* secret store, run `m365-agent-cli login
  --json`, parse the `device_code` event, run `device-login.mjs`, then verify. Caps retries.
- **`package.json`** — the two dependencies (`playwright`, `otplib`).

## Setup

```bash
cd examples/unattended-login
npm install                     # installs playwright + otplib
npx playwright install chromium # if you don't already have a Chromium build
```

## Adapt before using

1. Implement `fetch_secret()` in `refresh-token.sh` for your secret manager (1Password, Vault, Key Vault,
   AWS Secrets Manager, `pass`, …). **Never hardcode secrets.**
2. Make sure `EWS_CLIENT_ID` (+ tenant vars) are configured for `m365-agent-cli`, and the Entra app allows
   public client flows.
3. The Microsoft page selectors in `device-login.mjs` are a **starting point** — they change over time.

## Run

```bash
bash refresh-token.sh
# runs the flow and prints FINAL_RESULT=SUCCESS or FINAL_RESULT=FAILURE
```

The `login --json` events this consumes (`device_code`, `authenticated`, `complete`, `error`) are
documented in [../../docs/UNATTENDED_LOGIN.md](../../docs/UNATTENDED_LOGIN.md).
