#!/usr/bin/env bash
# Reference orchestration: one-time software-TOTP enrollment for a FRESH M365 account.
#
# Fetches the account password from YOUR secret store, runs enroll-totp.mjs to register a software
# authenticator and scrape its seed, stores that seed back in your secret store, then verifies by
# driving a real device-code sign-in (via refresh-token.sh) with the freshly enrolled seed.
#
# This is EXAMPLE code to copy and adapt — it is NOT shipped or executed by the CLI. Read
# docs/UNATTENDED_LOGIN.md ("Automated first-time TOTP enrollment") first: it only works when the
# tenant allows self-service registration from this host (no "require MFA to register security info"
# Conditional Access policy), and it captures Microsoft's generated seed rather than one you choose.
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
  local email password enroll_out totp_secret account_name

  email="$(fetch_secret m365-email)"
  password="$(fetch_secret m365-password)"
  if [ -z "$email" ] || [ -z "$password" ]; then
    echo "FAIL: empty email/password (implement fetch_secret for your secret store)" >&2
    exit 1
  fi
  # Never log credential values (or their lengths).
  echo "credentials loaded; enrolling a software authenticator for the account" >&2

  # 1) Enroll. enroll-totp.mjs prints ONE JSON line to stdout on success; diagnostics go to stderr.
  #    Capture stdout only — the seed must never land in a log.
  if ! enroll_out="$(M365_EMAIL="$email" M365_PASSWORD="$password" node "$SCRIPT_DIR/enroll-totp.mjs")"; then
    echo "FINAL_RESULT=FAILURE — enrollment did not complete (see stderr; failure.png in a temp dir)" >&2
    exit 1
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

  # 4) Verify end-to-end (optional): a normal unattended login — which re-reads the seed you just
  #    stored — should now succeed, proving MFA works AND that the seed round-tripped your store.
  #    Skipped if the CLI isn't on this host or ENROLL_VERIFY=0.
  if [ "${ENROLL_VERIFY:-1}" = "1" ] && command -v m365-agent-cli >/dev/null 2>&1; then
    echo "seed stored; verifying with a real device-code sign-in" >&2
    if bash "$SCRIPT_DIR/refresh-token.sh"; then
      echo "FINAL_RESULT=SUCCESS — enrolled, stored, and verified"
      exit 0
    fi
    echo "FINAL_RESULT=FAILURE — seed stored but the verification login failed; check the account/seed" >&2
    exit 1
  fi

  echo "FINAL_RESULT=SUCCESS — enrolled and stored (verification skipped)"
  exit 0
}

main "$@"
