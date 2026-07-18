#!/usr/bin/env bash
# Reference orchestration for unattended m365-agent-cli device-code login.
#
# This is EXAMPLE code to copy and adapt — it is NOT shipped or executed by the CLI.
# Read docs/UNATTENDED_LOGIN.md first for the security tradeoffs and when NOT to use this.
#
# Expects:
#   - m365-agent-cli on PATH, with EWS_CLIENT_ID (+ tenant vars) already configured
#   - node + this directory's deps installed (npm install; npx playwright install chromium)
#   - a fetch_secret() you implement for YOUR secret store (the body below is a placeholder)
set -euo pipefail

MAX_ATTEMPTS="${MAX_ATTEMPTS:-3}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# --- Implement this for your secret manager. NEVER hardcode secrets here. ---
# Examples: `op read ...` (1Password), `vault kv get ...`, `az keyvault secret show ...`,
#           `aws secretsmanager get-secret-value ...`, `pass show ...`, `gcloud secrets ...`.
fetch_secret() {
  : "${1:?secret name required}"
  echo "REPLACE_ME: fetch secret '$1' from your secret store" >&2
  return 1
}

run_once() {
  local email="$1" password="$2" totp_secret="$3"
  local out line user_code verification_uri deadline login_pid ok
  out="$(mktemp)"

  # Start the CLI login in JSON mode in the background. It stays alive polling until it
  # emits `complete` (or `error`); do NOT kill it before the browser side finishes.
  m365-agent-cli login --json >"$out" 2>/dev/null &
  login_pid=$!

  # Wait for the device_code event, then extract fields with node (no jq dependency).
  deadline=$(( $(date +%s) + 30 ))
  while :; do
    if line="$(grep -m1 '"event":"device_code"' "$out" 2>/dev/null)"; then
      user_code="$(printf '%s' "$line" | node -e 'let s="";process.stdin.on("data",d=>s+=d).on("end",()=>process.stdout.write((JSON.parse(s).user_code)||""))')"
      verification_uri="$(printf '%s' "$line" | node -e 'let s="";process.stdin.on("data",d=>s+=d).on("end",()=>process.stdout.write((JSON.parse(s).verification_uri)||""))')"
      break
    fi
    # Fail fast if the CLI emitted an error (e.g. missing EWS_CLIENT_ID) instead of a device code,
    # rather than waiting out the full 30s for a device_code event that will never arrive.
    if err_line="$(grep -m1 '"event":"error"' "$out" 2>/dev/null)"; then
      echo "login reported an error before any device code: ${err_line}" >&2
      kill "$login_pid" 2>/dev/null || true
      wait "$login_pid" 2>/dev/null || true
      rm -f "$out"
      return 1
    fi
    if [ "$(date +%s)" -ge "$deadline" ]; then
      echo "no device_code event within 30s" >&2
      kill "$login_pid" 2>/dev/null || true
      rm -f "$out"
      return 1
    fi
    sleep 1
  done

  # Complete the browser side (the login process must stay alive during this).
  M365_EMAIL="$email" \
  M365_PASSWORD="$password" \
  M365_TOTP_SECRET="$totp_secret" \
  M365_USER_CODE="$user_code" \
  M365_VERIFICATION_URI="$verification_uri" \
    node "$SCRIPT_DIR/device-login.mjs" || true

  # Wait for the CLI to persist the token; it exits after emitting `complete`.
  wait "$login_pid" 2>/dev/null || true
  ok=1
  grep -q '"event":"complete"' "$out" 2>/dev/null && ok=0
  rm -f "$out"
  return "$ok"
}

main() {
  local email password totp_secret attempt
  email="$(fetch_secret m365-email)"
  password="$(fetch_secret m365-password)"
  totp_secret="$(fetch_secret m365-totp-secret)"
  if [ -z "$email" ] || [ -z "$password" ] || [ -z "$totp_secret" ]; then
    echo "FAIL: empty credentials (implement fetch_secret for your secret store)" >&2
    exit 1
  fi
  # Never log the credential values (or their lengths).
  echo "credentials loaded; starting sign-in" >&2

  for attempt in $(seq 1 "$MAX_ATTEMPTS"); do
    echo "=== attempt ${attempt}/${MAX_ATTEMPTS} ===" >&2
    if run_once "$email" "$password" "$totp_secret"; then
      # The token is already persisted (run_once saw the `complete` event). Keep these confirmation
      # commands best-effort so a transient failure here doesn't flip a real success to a failure
      # (they run under `set -e`, which would otherwise abort before FINAL_RESULT=SUCCESS).
      m365-agent-cli whoami || echo "warning: whoami failed after a successful login (transient?)" >&2
      m365-agent-cli verify-token --capabilities ||
        echo "warning: verify-token failed after a successful login (transient?)" >&2
      echo "FINAL_RESULT=SUCCESS"
      exit 0
    fi
    echo "attempt ${attempt} failed" >&2
  done

  echo "FINAL_RESULT=FAILURE — escalate to a human device-code login" >&2
  exit 1
}

main "$@"
