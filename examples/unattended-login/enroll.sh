#!/usr/bin/env bash
# Reference orchestration: software-TOTP enrollment for an M365 account.
#
# Fetches a sign-in credential from YOUR secret store — a Temporary Access Pass (m365-tap) if present,
# else a password (m365-password) — runs enroll-totp.mjs to register a software authenticator and
# scrape its seed, stores that seed back, then verifies by driving a real device-code sign-in (via
# refresh-token.sh). Use a TAP when the account already has MFA or the tenant has a "require MFA to
# register security info" policy (a TAP satisfies that gate); a plain password works on permissive
# tenants. NOTE: a TAP can't set a password, and TOTP is only a SECOND factor — so a first-factor
# password must already live in the vault (m365-password) for future device-code refreshes.
#
# This is EXAMPLE code to copy and adapt — it is NOT shipped or executed by the CLI. Read
# docs/UNATTENDED_LOGIN.md ("Automated first-time TOTP enrollment") first. It captures Microsoft's
# generated seed rather than one you choose.
#
# Expects:
#   - node + this directory's deps installed (npm install; npx playwright install chromium)
#   - fetch_secret() and store_secret() implemented for YOUR secret store (placeholders below)
#   - for the verify step: m365-agent-cli on PATH with EWS_CLIENT_ID (+ tenant vars) configured,
#     and fetch_secret() in refresh-token.sh wired to the SAME store this script writes to
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# --- Implement these for your secret manager. NEVER hardcode secrets here. ---
# fetch_secret: print the named secret's value to stdout.
# store_secret: read a value from stdin and persist it under the given name (stdin, so the secret
#               never appears in argv / `ps`). It must write where refresh-token.sh's fetch_secret reads.
# Examples: `op` (1Password), `vault kv`, `az keyvault secret`, `aws secretsmanager`, `pass`, `gcloud secrets`.
fetch_secret() {
  : "${1:?secret name required}"
  echo "REPLACE_ME: read secret '$1' from your secret store" >&2
  return 1
}
store_secret() {
  : "${1:?secret name required}"
  echo "REPLACE_ME: write secret '$1' to your secret store (value on stdin)" >&2
  return 1
}

# Extract one field from enroll-totp.mjs's JSON output without a jq dependency (matches the
# node-inline style used in refresh-token.sh).
json_field() {
  node -e 'let s="";process.stdin.on("data",d=>s+=d).on("end",()=>{try{process.stdout.write(String(JSON.parse(s)[process.argv[1]]??""))}catch{process.stdout.write("")}})' "$1"
}

main() {
  local email tap password enroll_out totp_secret account_name

  email="$(fetch_secret m365-email)"
  if [ -z "$email" ]; then
    echo "FAIL: empty email (implement fetch_secret for your secret store)" >&2
    exit 1
  fi

  # 1) Enroll. Credential: prefer a Temporary Access Pass (works on hardened / existing-MFA accounts),
  #    else a password. enroll-totp.mjs prints ONE JSON line to stdout on success; diagnostics to
  #    stderr. Capture stdout only — the seed must never land in a log. Never log credential values.
  tap="$(fetch_secret m365-tap || true)"
  if [ -n "$tap" ]; then
    echo "signing in with a Temporary Access Pass to enroll a software authenticator" >&2
    if ! enroll_out="$(M365_EMAIL="$email" M365_TAP="$tap" node "$SCRIPT_DIR/enroll-totp.mjs")"; then
      echo "FINAL_RESULT=FAILURE — enrollment did not complete (see stderr; failure.png in a temp dir)" >&2
      exit 1
    fi
  else
    password="$(fetch_secret m365-password)"
    if [ -z "$password" ]; then
      echo "FAIL: no m365-tap and no m365-password available (implement fetch_secret)" >&2
      exit 1
    fi
    echo "signing in with a password to enroll a software authenticator" >&2
    if ! enroll_out="$(M365_EMAIL="$email" M365_PASSWORD="$password" node "$SCRIPT_DIR/enroll-totp.mjs")"; then
      echo "FINAL_RESULT=FAILURE — enrollment did not complete (see stderr; failure.png in a temp dir)" >&2
      exit 1
    fi
  fi

  # 2) Parse the seed and account name.
  totp_secret="$(printf '%s' "$enroll_out" | json_field totp_secret)"
  account_name="$(printf '%s' "$enroll_out" | json_field account_name)"
  if [ -z "$totp_secret" ]; then
    echo "FINAL_RESULT=FAILURE — enrollment produced no seed" >&2
    exit 1
  fi
  echo "authenticator registered for ${account_name:-the account}; storing seed" >&2

  # 3) Persist the seed. Pipe on stdin so it never appears in argv / `ps`.
  #    If this fails you must RE-ENROLL — the seed exists only in this shell and is otherwise lost.
  if ! printf '%s' "$totp_secret" | store_secret m365-totp-secret; then
    echo "FINAL_RESULT=FAILURE — enrolled but could NOT store the seed; re-enroll (seed is lost)" >&2
    exit 1
  fi

  # 4) Verify end-to-end (optional): a normal device-code login needs a FIRST-FACTOR password (TOTP is
  #    only a second factor) plus the seed you just stored. The password must already be in the vault —
  #    a TAP can't set one. Skipped if there's no password, no CLI, or ENROLL_VERIFY=0.
  [ -n "${password:-}" ] || password="$(fetch_secret m365-password || true)"
  if [ "${ENROLL_VERIFY:-1}" = "1" ] && command -v m365-agent-cli >/dev/null 2>&1 && [ -n "${password:-}" ]; then
    echo "seed stored; verifying with a real device-code sign-in" >&2
    if bash "$SCRIPT_DIR/refresh-token.sh"; then
      echo "FINAL_RESULT=SUCCESS — enrolled, stored, and verified"
      exit 0
    fi
    echo "FINAL_RESULT=FAILURE — seed stored but the verification login failed; check the account/seed" >&2
    exit 1
  fi

  if [ -z "${password:-}" ]; then
    echo "FINAL_RESULT=SUCCESS — TOTP enrolled and stored. NOTE: no m365-password in the vault; device-code refreshes need a first-factor password (TOTP is only a second factor), so store one before relying on unattended login."
  else
    echo "FINAL_RESULT=SUCCESS — enrolled and stored (verification skipped)"
  fi
  exit 0
}

main "$@"
