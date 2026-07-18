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
- **`enroll-totp.mjs`** — helper to register a software authenticator: signs in with a **password or a
  Temporary Access Pass**, confirms the sign-in worked, drives the Security info wizard, scrapes the base32
  seed off the **Can't scan image?** screen, activates it, and prints `{"totp_secret":"…","account_name":"…"}`
  to **stdout** for you to store. See the "Automated first-time TOTP enrollment" section of the docs.
- **`enroll.sh`** — orchestration for enrollment: fetch email + a credential (a **TAP** `m365-tap` if present,
  else a **password** `m365-password`) from *your* secret store, run `enroll-totp.mjs`, store the returned seed
  back, then verify end-to-end by running `refresh-token.sh` (which re-reads the stored seed — so it also proves
  the round-trip).
- **`refresh-token.sh`** — orchestration for steady-state login: fetch creds from *your* secret store, run
  `m365-agent-cli login --json`, parse the `device_code` event, run `device-login.mjs`, then verify. Caps
  retries.
- **`package.json`** — the two dependencies (`playwright`, `otplib`).

## Enrollment (optional)

To have the automation obtain the TOTP seed itself, implement `fetch_secret()` + `store_secret()` in
`enroll.sh` and run it once:

```bash
bash enroll.sh
# fetch credential (TAP m365-tap, else password m365-password) -> enroll-totp.mjs
#   -> store the seed -> verify via refresh-token.sh
# prints FINAL_RESULT=SUCCESS or FINAL_RESULT=FAILURE
```

Or drive `enroll-totp.mjs` directly with a password **or** a TAP, capturing the seed straight into your vault
(never a log file):

```bash
# password (permissive tenants):
M365_EMAIL=agent@contoso.com M365_PASSWORD="$(fetch_secret agent-password)" \
  node enroll-totp.mjs | your-vault put agent-totp-seed

# Temporary Access Pass (hardened tenants, or an account that already has MFA):
M365_EMAIL=agent@contoso.com M365_TAP="$(fetch_secret agent-tap)" \
  node enroll-totp.mjs | your-vault put agent-totp-seed
```

A **password** works when the tenant allows self-service registration (no "require MFA to register security
info" policy). A **TAP** clears that policy and also covers accounts that already have MFA. Either way it
captures **Microsoft's** generated seed; to inject a seed *you* generate you need the admin OATH-token API.
Note a TAP can't set a password — and TOTP is only a second factor — so a first-factor password must still
exist for device-code refreshes. All covered in the docs.

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
