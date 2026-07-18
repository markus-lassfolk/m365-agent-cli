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
- **`enroll-totp.mjs`** — one-time helper for a *fresh* account: signs in with the password, drives the
  Security info wizard, scrapes the base32 seed off the **Can't scan image?** screen, activates it, and
  prints `{"totp_secret":"…","account_name":"…"}` to **stdout** for you to store. See the "Automated
  first-time TOTP enrollment" section of the docs for when this is (and isn't) possible.
- **`enroll.sh`** — orchestration for enrollment: fetch email + password from *your* secret store, run
  `enroll-totp.mjs`, store the returned seed back, then verify end-to-end by running `refresh-token.sh`
  (which re-reads the stored seed — so it also proves the round-trip).
- **`refresh-token.sh`** — orchestration for steady-state login: fetch creds from *your* secret store, run
  `m365-agent-cli login --json`, parse the `device_code` event, run `device-login.mjs`, then verify. Caps
  retries.
- **`package.json`** — the two dependencies (`playwright`, `otplib`).

## One-time enrollment (optional)

If you're bootstrapping a brand-new account and want the automation to obtain the TOTP seed itself,
implement `fetch_secret()` + `store_secret()` in `enroll.sh` and run it once:

```bash
bash enroll.sh
# fetch email+password -> enroll-totp.mjs -> store the seed -> verify via refresh-token.sh
# prints FINAL_RESULT=SUCCESS or FINAL_RESULT=FAILURE
```

Or drive `enroll-totp.mjs` directly and capture the seed straight into your vault — never a log file:

```bash
M365_EMAIL=agent@contoso.com M365_PASSWORD="$(fetch_secret agent-password)" \
  node enroll-totp.mjs | your-vault put agent-totp-seed
```

This only works when the tenant allows self-service registration from this host (no "require MFA to
register security info" Conditional Access policy — that needs a Temporary Access Pass). It captures
**Microsoft's** generated seed; to inject a seed *you* generate you need the admin OATH-token API instead.
Both paths are covered in the docs.

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
