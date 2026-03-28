#!/bin/bash
# Ralph Wiggum loop for clippy — the repo that notices dead birds
# Runs every 15 min. Monitors, enriches, and IMPLEMENTS.
# Loops: after each Forge dispatch, re-checks time budget and picks the next
# item if >8 minutes remain.
set -euo pipefail

REPO="markus-lassfolk/clippy"
BRANCH="dev"
WORKSPACE="/home/markus/.openclaw/workspace/markus-lassfolk-clippy"
GOALS_FILE="$WORKSPACE/docs/GOALS.md"
LOG_DIR="$HOME/.openclaw/workspace/ralph/logs"
LOG_FILE="$LOG_DIR/clippy-$(date +%Y-%m-%d).md"
GUARD_FILE="$HOME/.openclaw/workspace/tmp/guards/ralph-clippy-wiggum.ts"
FORGE_GUARD="$HOME/.openclaw/workspace/tmp/guards/ralph-clippy-forge-running.txt"

mkdir -p "$LOG_DIR"
mkdir -p "$(dirname "$GUARD_FILE")"

# Guard: skip if ran within 10 min
if [[ -f "$GUARD_FILE" ]]; then
  LAST=$(cat "$GUARD_FILE")
  NOW=$(date +%s)
  AGE=$((NOW - LAST))
  if (( AGE < 600 )); then
    echo "Ralph clippy: guard active (${AGE}s ago)"
    exit 0
  fi
fi
echo "$(date +%s)" > "$GUARD_FILE"

cd "$WORKSPACE"

SESSION_START=$(date +%s)
MIN_BUDGET=480  # must have >8 min left to dispatch another Forge

log() {
  echo "$1" >> "$LOG_FILE"
}

remaining_budget() {
  local elapsed=$(($(date +%s) - SESSION_START))
  echo $((900 - elapsed))
}

log "Ralph clippy Wiggum run: $(date -Iminutes)"

# Check if Forge from a previous run is still active
if [[ -f "$FORGE_GUARD" ]]; then
  FORGE_AGE=$(($(date +%s) - $(sed 's/ .*//' < "$FORGE_GUARD")))
  if (( FORGE_AGE < 2700 )); then
    log "Forge from previous run still active (${FORGE_AGE}s), skipping this run."
    echo "Ralph clippy: Forge still running from previous session (${FORGE_AGE}s)"
    exit 0
  fi
fi

# Fetch issues once per loop iteration
fetch_issues() {
  gh issue list --repo "$REPO" --state open \
    --json number,title,body,labels,createdAt --limit 100
}

dispatch_forge() {
  local ISSUE_NUM="$1"
  local IS_GOALS_GAP="${2:-false}"
  local GOAL_TEXT="${3:-}"
  local FORGE_BRIEF LABEL

  if [[ "$IS_GOALS_GAP" == "false" ]]; then
    local ISSUE_TITLE ISSUE_BODY ENRICHMENT
    ISSUE_TITLE=$(gh issue view "$ISSUE_NUM" --repo "$REPO" --json title --jq '.title')
    ISSUE_BODY=$(gh issue view "$ISSUE_NUM" --repo "$REPO" --json body --jq -r '.body // ""')
    ENRICHMENT=$(gh issue view "$ISSUE_NUM" --repo "$REPO" --json comments \
      --jq '[.comments[] | select(.body | contains("Enrichment Analysis"))] | last.body' 2>/dev/null || echo "")

    FORGE_BRIEF=$(mktemp)
    cat > "$FORGE_BRIEF" << BRIEF_EOF
## Forge Brief — Issue #$ISSUE_NUM: $ISSUE_TITLE

### Issue Description
$ISSUE_BODY

### Enrichment Analysis
$ENRICHMENT

### Repo Context
- Working branch: $BRANCH (origin/$BRANCH)
- Docs: docs/GOALS.md, docs/ARCHITECTURE.md
- Tests: src/test/ (npm test)
- SKILL: $WORKSPACE/SKILL/SKILL.md

### Instructions
1. Read docs/GOALS.md and docs/ARCHITECTURE.md
2. Implement the fix/feature for issue #$ISSUE_NUM
3. Write or update tests in src/test/
4. Update SKILL/SKILL.md and README.md if behaviour changed
5. Commit to branch $BRANCH and push
6. Do NOT merge to main
7. When done: apply "implemented" label and remove "in-progress" from issue #$ISSUE_NUM
8. Reply with a concise summary of changes
BRIEF_EOF
  else
    FORGE_BRIEF=$(mktemp)
    cat > "$FORGE_BRIEF" << BRIEF_EOF
## Forge Brief — GOALS Gap-Fill

### Target
$GOAL_TEXT

### Source
docs/GOALS.md

### Context
- Issue queue is empty — gap-filling from GOALS.md
- Read docs/GOALS.md and docs/ARCHITECTURE.md first
- Repo: $REPO, branch: $BRANCH

### Steps
1. Read docs/GOALS.md and docs/ARCHITECTURE.md
2. Understand what this goal requires
3. Implement the feature or improvement
4. Write or update tests in src/test/
5. Update SKILL/SKILL.md if behaviour changed
6. Commit and push to $BRANCH
7. Mark the GOALS item as done (replace "- [ ]" with "- [x]") in a separate commit
8. Reply with summary of changes
BRIEF_EOF
  fi

  if [[ "$IS_GOALS_GAP" == "true" ]]; then
    log "Dispatching Forge for GOALS gap-fill"
  else
    log "Dispatching Forge for :#$ISSUE_NUM"
  fi

  if [[ "$IS_GOALS_GAP" == "false" ]]; then
    gh issue edit "$ISSUE_NUM" --repo "$REPO" --add-label "in-progress" 2>/dev/null || true
  fi

  LABEL="forge-clippy-$(date +%s)"
  sessions_spawn runtime="subagent" agentId="forge" mode="run" timeoutSeconds="3600" \
    task="$(cat "$FORGE_BRIEF")" label="$LABEL" 2>/dev/null &
  if [[ "$IS_GOALS_GAP" == "true" ]]; then
    echo "$(date +%s) pid=$! label=$LABEL issue=goals-gap goals_gap=1" > "$FORGE_GUARD"
  else
    echo "$(date +%s) pid=$! label=$LABEL issue=$ISSUE_NUM" > "$FORGE_GUARD"
  fi

  rm -f "$FORGE_BRIEF"
  log "Forge dispatched (label=$LABEL)"
}

# ── Main loop: keep working while budget allows ────────────────────────────────
DISPATCH_COUNT=0

while true; do
  BUDGET=$(remaining_budget)
  if (( BUDGET < MIN_BUDGET )); then
    log "Budget exhausted (${BUDGET}s left, need >${MIN_BUDGET}s). Stopping."
    break
  fi

  # Check if Forge from this run is still active
  if [[ -f "$FORGE_GUARD" ]]; then
    FORGE_AGE=$(($(date +%s) - $(sed 's/ .*//' < "$FORGE_GUARD")))
    if (( FORGE_AGE < 2700 )); then
      log "Forge still running (${FORGE_AGE}s), waiting for completion."
      break
    fi
  fi

  # Python analysis
  ANALYSIS=$(fetch_issues | python3 -c '
import json, sys
from datetime import datetime, timezone

data = json.loads(sys.stdin.read())

def label_names(labels):
    return [l["name"] for l in labels]

def queue_prio(labels):
    for n in label_names(labels):
        if n.startswith("Queue:"):
            return n.split(":",1)[1]
    return None

enriched     = [i for i in data if "enriched" in label_names(i["labels"])]
unenriched   = [i for i in data if "enriched" not in label_names(i["labels"])]
needs_triage = [i for i in enriched if not queue_prio(i["labels"])]

counts = {"Critical":0,"High":0,"Medium":0,"Low":0}
for i in data:
    p = queue_prio(i["labels"])
    if p in counts:
        counts[p] += 1

# Forge candidate: oldest enriched+prioritized, not in-progress
forge_cand = None
for p in ("Critical","High","Medium","Low"):
    for i in sorted(enriched, key=lambda x: x["number"]):
        if queue_prio(i["labels"]) == p:
            if "in-progress" not in label_names(i["labels"]):
                forge_cand = i["number"]
                break
    if forge_cand:
        break

print(f"OPEN={len(data)}")
print(f"ENRICHED={len(enriched)}")
print(f"UNENRICHED={len(unenriched)}")
print(f"NEEDS_TRIAGE={len(needs_triage)}")
print(f"FORGE_CANDIDATE={forge_cand or ''}")
')

  while IFS='=' read -r key value; do
    declare "$key=$value"
  done <<< "$ANALYSIS"

  log "Loop: open=$OPEN enriched=$ENRICHED unenriched=$UNENRICHED forge_cand=${FORGE_CANDIDATE:-none} budget=${BUDGET}s"

  # Scholar: always worth doing if unenriched exist
  if (( UNENRICHED > 0 )); then
    log "Unenriched issues — Scholar should enrich"
  fi

  # Forge dispatch: enriched + prioritized + not in-progress
  if [[ -n "$FORGE_CANDIDATE" ]]; then
    dispatch_forge "$FORGE_CANDIDATE"
    DISPATCH_COUNT=$((DISPATCH_COUNT + 1))
    continue
  fi

  # GOALS gap-fill: no forge candidate AND no unenriched
  if [[ -z "$FORGE_CANDIDATE" && "$UNENRICHED" == "0" ]]; then
    if [[ -f "$GOALS_FILE" ]]; then
      FIRST_UNDONE=$(grep -n "^- \[ \]" "$GOALS_FILE" 2>/dev/null | head -1 || true)
      if [[ -n "$FIRST_UNDONE" ]]; then
        LINE_NUM=$(echo "$FIRST_UNDONE" | cut -d: -f1)
        GOAL_TEXT=$(sed -n "${LINE_NUM}s/- \[ \] //p" "$GOALS_FILE")
        log "GOALS gap-fill: $GOAL_TEXT"
        dispatch_forge "" "true" "$GOAL_TEXT"
        DISPATCH_COUNT=$((DISPATCH_COUNT + 1))
        continue
      fi
    fi
  fi

  # Nothing actionable
  log "Nothing to dispatch. Breaking."
  break
done

# ── Summary ─────────────────────────────────────────────────────────────────
log "Loop done. Forges dispatched: $DISPATCH_COUNT"

if (( DISPATCH_COUNT > 0 )); then
  echo "Ralph clippy: $DISPATCH_COUNT Forge dispatch$([[ $DISPATCH_COUNT -gt 1 ]] && echo "es") | open=$OPEN enriched=$ENRICHED"
elif (( UNENRICHED > 0 )); then
  echo "Ralph clippy: $UNENRICHED unenriched (Scholar) | open=$OPEN"
elif [[ -n "${NEEDS_TRIAGE:-}" ]] && (( NEEDS_TRIAGE > 0 )); then
  echo "Ralph clippy: $NEEDS_TRIAGE enriched need Queue labels | open=$OPEN"
else
  echo "Ralph clippy: all clear | open=$OPEN enriched=$ENRICHED"
fi

log "Run complete: $(date -Iminutes)"
log "---"
