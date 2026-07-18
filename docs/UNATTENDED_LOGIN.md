# Unattended (automated) device-code login

`m365-agent-cli` signs in with the **OAuth2 device code** flow, which normally needs a human to open a
browser and approve once. This guide shows how to **automate that human step** with a headless browser so
an agent can re-authenticate on its own, using the CLI's machine-readable **`login --json`** mode.

It is still the same device-code flow — nothing here adds a different grant type. It just drives the
interactive step programmatically.

> **Reference example:** [`examples/unattended-login/`](../examples/unattended-login/) contains an
> adaptable, secret-free Playwright script plus an orchestration wrapper. Read this page first, then copy
> and adapt those to your environment.

---

## Read this first — the security tradeoff

Automating the sign-in means a machine must be able to complete **every** factor without you. In practice
that means storing, and handing to a browser, both:

- the account **password**, and
- the **TOTP (authenticator) seed** used for MFA.

**This collapses MFA into secrets your automation holds.** Anyone who can read your secret store then has
full access to the account, second factor included. Only do this when:

- The account is one you **own and intend to fully automate** — ideally a **dedicated, least-privilege**
  account, not a person's primary identity.
- The secrets live in a **real secret manager** you fetch at runtime — never hardcoded in a script, a
  `.env` you commit, or anywhere in the repo.
- The account's second factor is a **TOTP authenticator app**, not a **Temporary Access Pass** (TAP is a
  one-time bootstrap; each TAP rotation **revokes all refresh tokens**) and not push / number-matching
  (those need a human tap).

If you can instead keep a human in the loop for the occasional device-code login, **prefer that** — it is
simpler and strictly safer. Automating Microsoft's sign-in UI is also inherently brittle (the markup
changes without notice) and repeated automated sign-ins can raise **Conditional Access** risk scoring, so
cap your attempts and be ready to fall back to a person.

---

## Prerequisites

- An **Entra app registration** with the device code flow enabled (**Allow public client flows = Yes**) —
  the same app you already use for `m365-agent-cli login`. See [ENTRA_SETUP.md](./ENTRA_SETUP.md).
- **`EWS_CLIENT_ID`** and any tenant variables (`M365_TENANT_ID` / …) set in the environment or `.env`.
  In `--json` mode the CLI does **not** prompt for a missing client id — it errors instead.
- **Node.js** and **Playwright** with a Chromium build on the host (`npm i playwright`).
- A **TOTP library** to generate codes from your seed — Node: `otplib`; Python: `pyotp`.
- The account **password** and **base32 TOTP seed** — see [Setting up software TOTP](#setting-up-software-totp-one-time) below for how to obtain the seed. Both are read from your secret store at runtime.

---

## Setting up software TOTP (one-time)

The automation needs the account's **TOTP shared secret** — a base32 string — so it can generate codes the
way a phone authenticator would. A normal phone setup hides that secret behind a QR code; for automation you
need the **raw value**. Two ways to get one:

### Option A — reveal the secret during self-service registration

Best when you (or the account owner) can sign in to the account's security-info page once.

1. Sign in at **<https://mysignins.microsoft.com/security-info>** (also **<https://aka.ms/mfasetup>**) as the
   account. A brand-new account with no method yet needs a **one-time bootstrap**: have an admin issue a
   **Temporary Access Pass** (Entra admin center → the user → **Authentication methods → Add → Temporary
   Access Pass**) and sign in with it. TAP is a bootstrap *only* — you switch to TOTP in the same sitting and
   never rely on TAP for the running automation (**each TAP rotation revokes all refresh tokens**).
2. **+ Add sign-in method → Authenticator app → Add**.
3. On "Start by getting the app", click **I want to use a different authenticator app → Next**.
4. On "Set up your account", click **Can't scan image?**. This reveals the **Secret key** (the base32 TOTP
   seed) plus an account name. **Copy the secret key straight into your secret manager — do not scan the QR
   into a phone.**
5. Generate a current 6-digit code from that secret (see [Generating codes](#generating-codes)) and enter it
   to finish registration.

### Option B — provision your own secret as an admin (OATH token)

Cleaner for a controlled automation account: *you* generate the secret, so there's no screen to scrape.
Requires an admin and that the tenant allows OATH tokens.

1. Generate a random base32 secret with your TOTP library — Node `otplib`: `authenticator.generateSecret()`;
   Python `pyotp`: `pyotp.random_base32()`.
2. In the **Entra admin center → Protection → Authentication methods → OATH tokens**, upload a CSV row for the
   account (UPN, serial number, secret key, … per the on-screen template), then **Activate** the token —
   activation asks for a current code, which you generate from that secret.
3. Store the same secret in your secret manager.

Either path leaves you with one thing: a **base32 seed** in your secret store.

### Tenant / admin prerequisites

- The tenant's **Authentication methods policy** (Entra admin center → **Protection → Authentication
  methods**) must permit the method you chose — allow non-Microsoft / "different" authenticator apps for
  Option A, or **Software OATH tokens** for Option B. An admin may need to enable it.
- **Conditional Access** must not block the automation host. A new location or an unmanaged device can trigger
  an extra challenge the script can't answer — see [When it stops working](#when-it-stops-working).

### Generating codes

Microsoft software TOTP uses the standard parameters: **SHA-1, 6 digits, 30-second period** — the default in
most libraries, so no extra configuration is needed.

```bash
# Node (otplib) — what the reference example uses
node -e "import('otplib').then(({authenticator}) => console.log(authenticator.generate(process.env.M365_TOTP_SECRET)))"

# Python (pyotp)
python3 -c "import os, pyotp; print(pyotp.TOTP(os.environ['M365_TOTP_SECRET']).now())"
```

**Verify once before relying on it:** generate a code and confirm it matches what a phone authenticator shows
for the same secret (or that it completes one manual sign-in). A wrong secret — or non-default parameters
(SHA-256, 8 digits) — produces codes Microsoft silently rejects.

---

## How it works (five phases)

1. **Fetch credentials** — pull the password and TOTP seed from your secret store into environment
   variables (never onto disk in plaintext).
2. **Start the CLI login and capture the code** — run **`m365-agent-cli login --json`** and read the
   `device_code` event from stdout to get `user_code` and `verification_uri`.
3. **Drive the browser** — a headless Playwright session opens the verification URL, enters the code,
   signs in with the password, completes MFA with a freshly generated TOTP code, and approves the app.
4. **Let the CLI finish** — the still-running `login` process polls, receives the token, persists it, and
   emits `authenticated` then `complete`.
5. **Verify** — confirm with `whoami` and `verify-token --capabilities`.

---

## The `login --json` contract

`--json` writes **newline-delimited JSON events to stdout**; all human-readable text goes to **stderr**.
Parse stdout one line at a time.

```console
$ m365-agent-cli login --json
{"event":"device_code","user_code":"ABCD1234","verification_uri":"https://microsoft.com/devicelogin","verification_uri_complete":null,"expires_in":900,"interval":5,"message":"To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code ABCD1234 to authenticate."}
# ... the process keeps polling while you complete the browser side ...
{"event":"authenticated","username":"agent@contoso.com"}
{"event":"complete","env_path":"/home/agent/.config/m365-agent-cli/.env"}
```

| Event | When | Key fields |
| --- | --- | --- |
| `device_code` | Immediately, once the code is issued | `user_code`, `verification_uri`, `verification_uri_complete`, `expires_in`, `interval`, `message` |
| `authenticated` | Sign-in succeeded, token received | `username` |
| `complete` | Refresh token persisted to the env file | `env_path` |
| `error` | Any failure (bad request, expired code, missing client id, auth failed) | `error`, `error_description` |

**Keep the process alive.** `login` is a **synchronous foreground poll** — it keeps calling the token
endpoint until sign-in completes or the code expires (`expires_in`, ~15 min). Background it if you like,
but do **not** kill it until it emits `complete` (or `error`), even if the browser already shows "signed
in" — killing it early discards the pending device-code session and the token is never saved.

---

## The Microsoft sign-in page sequence

Drive these in order (selectors change over time — treat them as a starting point, not a contract):

1. **Device-code entry** — field `input[name="otc"]` (fallback `input[type="text"]`) → **Next**.
2. **Account** — either a fresh email field `input[name="loginfmt"]`, or, if a persisted browser profile
   already has an account, a **"Pick an account"** screen — click the tile matching the email (match by
   the visible email text, not a brittle generated selector).
3. **Password** — `input[type="password"]` → **Sign in**.
4. **App consent** — "Are you trying to sign in to …?" with **Cancel** / **Continue**. This screen can
   **re-render 1–2 times**; click **Continue** in a loop until it goes away, don't assume one click.
5. **TOTP** — when prompted, fill the 6-digit code from your TOTP library → **Verify**. If it's rejected,
   wait for the next 30-second window and retry **once**.
6. **"Stay signed in?"** — click **Yes** if it appears.
7. **Done** — the final page confirms "You have signed in … you may now close this window."

---

## Pitfalls (learned the hard way — don't reintroduce them)

- **Don't kill the login process early.** Phases 2–4 all depend on the same `login` process staying alive
  until it emits `complete`. This is the single most common cause of a "browser said success but the token
  never refreshed" failure.
- **Wait for real DOM elements, never a blind timeout.** Use `waitForSelector` / `waitForFunction` for the
  element you expect next; a fixed `waitForTimeout` as your only wait will read a half-loaded page and act
  on the wrong field. A fixed delay is fine only as extra margin *after* an event-based wait.
- **Never treat a URL substring as success.** The flow's base URL contains fragments like `deviceauth` on
  *every* page, so a check like `url.includes('deviceauth')` is true from the first load and always
  "passes." Only trust the explicit final confirmation **text** (e.g. `you have signed in` /
  `you may now close this window`).
- **Log every step** — the current URL plus the first ~200 characters of the page body — so a failed run
  tells you exactly where it stopped instead of a bare pass/fail.

---

## Hygiene and safety (every run)

- **Never hardcode secrets** in scripts or files under the repo/workspace. Pass them via environment
  variables set at execution time; **never log their values — not even lengths**, which can trip secret
  scanners. Log a generic status line instead.
- **Keep the browser profile and screenshots in a temp dir** (`mktemp -d` / `os.tmpdir()`), never in your
  workspace or the repo, and delete them on success.
- **Cap automated attempts** (e.g. ≤ 5 per incident). Repeated automated sign-ins against the same account
  raise Conditional Access risk and can trigger a harder lockout than the one you're fixing — escalate to a
  human after that.
- **Always verify afterwards** with `m365-agent-cli whoami` **and** `m365-agent-cli verify-token
  --capabilities`. Don't trust the automation's own "success"; the CLI's `complete` event or a verified
  token is the real signal.
- **No manual lock cleanup.** Refresh-token exchange is serialized per identity via `.refresh-{identity}.lock`
  in the config dir, which **auto-heals** stale locks (dead holder PID, or older than 120 s). You never
  need to delete a lock file by hand.

---

## When it stops working

If Microsoft changes the sign-in markup enough that selectors stop matching, or Conditional Access starts
challenging the automated browser (new device/location risk), **fall back to a human** completing
`m365-agent-cli login` once — same flow, same URL, a person just clicks through. Don't hammer the automated
path; a few failed automated attempts is your signal to hand off.

---

## See also

- [AUTHENTICATION.md](./AUTHENTICATION.md) — device-code login, token cache, tenant precedence.
- [ENTRA_SETUP.md](./ENTRA_SETUP.md) — app registration (enable public client flows).
- [`examples/unattended-login/`](../examples/unattended-login/) — the adaptable reference scripts.
