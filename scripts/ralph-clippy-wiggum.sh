#!/bin/bash
# Ralph Wiggum loop for clippy — the repo that notices dead birds
# Runs every 15 min, checks for issues needing enrichment, stale issues, blockers
set -euo pipefail

REPO="markus-lassfolk/clippy"
BRANCH="dev"
WORKSPACE="/home/markus/.openclaw/workspace/markus-lassfolk-clippy"
GOALS_FILE="$WORKSPACE/docs/GOALS.md"
LOG_DIR="$HOME/.openclaw/workspace/ralph/logs"
TODAY=$(date +%Y-%m-%d)
LOG_FILE="$LOG_DIR/clippy-$(date +%Y-%m-%d).md"
GUARD_FILE="$HOME/.openclaw/workspace/tmp/guards/ralph-clippy-wiggum.txt"

mkdir -p "$LOG_DIR"
mkdir -p "$(dirname "$GUARD_FILE")"

# Guard: skip if ran within 10 min
if [[ -f "$GUARD_FILE" ]]; then
  LAST=$(cat "$GUARD_FILE")
  NOW=$(date +%s)
  AGE=$((NOW - LAST))
  if (( AGE < 600 )); then
    echo "Ralph clippy: guard active (${AGE}s ago, min 600s)"
    exit 0
  fi
fi
echo "$(date +%s)" > "$GUARD_FILE"

cd "$WORKSPACE"

echo "Ralph clippy Wiggum run: $(date -Iminutes)" >> "$LOG_FILE"
echo "" >> "$LOG_FILE"

# 1. How many open issues total?
OPEN=$(gh issue list --repo "$REPO" --state open --json number --jq 'length')
echo "Open issues: $OPEN" >> "$LOG_FILE"

# 2. How many enriched?
ENRICHED=$(gh issue list --repo "$REPO" --state open --label enriched --json number --jq 'length')
echo "Enriched: $ENRICHED" >> "$LOG_FILE"

# 3. How many unenriched? (needs Scholar attention)
UNENRICHED=$(gh issue list --repo "$REPO" --state open --json number,labels --jq '.[] | select(.labels[].name != "enriched") | .number' 2>/dev/null | wc -l || echo 0)
echo "Unenriched (needs Scholar): $UNENRICHED" >> "$LOG_FILE"

# 4. Queue priority breakdown
for PRIORITY in Critical High Medium Low; do
  COUNT=$(gh issue list --repo "$REPO" --state open --label "Queue:$PRIORITY" --json number --jq 'length' 2>/dev/null || echo 0)
  echo "Queue:$PRIORITY: $COUNT" >> "$LOG_FILE"
done

# 5. Oldest unenriched issues (more than 24h since created, no enriched label)
OLD_UNENRICHED=$(gh issue list --repo "$REPO" --state open --json number,title,createdAt,labels \
  --jq '.[] | select(.labels[] | .name == "enriched" | not) | select((now - (.createdAt | fromdateiso8601) > 86400)) | "\(.number) \(.title) — created \(.createdAt[0:10])"' 2>/dev/null || true)
if [[ -n "$OLD_UNENRICHED" ]]; then
  echo "Stale unenriched issues (>24h, needs Scholar):" >> "$LOG_FILE"
  echo "$OLD_UNENRICHED" >> "$LOG_FILE"
fi

# 6. Enriched but no Queue label (needs priority assignment)
UNPRIORITIZED=$(gh issue list --repo "$REPO" --state open --json number,title,labels \
  --jq '.[] | select(.labels[] | .name == "enriched") | select(.labels[] | startswith("Queue:")) == false | "\(.number) \(.title)"' 2>/dev/null || true)
if [[ -n "$UNPRIORITIZED" ]]; then
  echo "Enriched but unprioritized (needs Queue label):" >> "$LOG_FILE"
  echo "$UNPRIORITIZED" >> "$LOG_FILE"
fi

# 7. Security/Reliability bugs (highest urgency — flag if unassigned)
BUGS_URGENT=$(gh issue list --repo "$REPO" --state open --label bug --json number,title,labels \
  --jq '.[] | select(.labels[] | .name == "enriched") | select(.labels[] | .name == "security" or .name == "reliability") | "\(.number) \(.title)"' 2>/dev/null || true)
if [[ -n "$BUGS_URGENT" ]]; then
  echo "Security/Reliability bugs (enriched, priority):" >> "$LOG_FILE"
  echo "$BUGS_URGENT" >> "$LOG_FILE"
fi

# 8. Dev branch PR status
PR_STATE=$(gh pr view "$BRANCH" --repo "$REPO" --json state,headRefName 2>/dev/null || echo "{}")
PR_OPEN=$(echo "$PR_STATE" | jq -r '.state' 2>/dev/null || echo "unknown")
echo "Dev PR state: $PR_OPEN (branch: $BRANCH)" >> "$LOG_FILE"

# 9. GOALS.md — what's the current focus area?
if [[ -f "$GOALS_FILE" ]]; then
  FOCUS=$(grep "^## Part" "$GOALS_FILE" | head -3 || true)
  echo "GOALS focus:" >> "$LOG_FILE"
  echo "$FOCUS" >> "$LOG_FILE"
fi

echo "Run complete: $(date -Iminutes)" >> "$LOG_FILE"
echo "---" >> "$LOG_FILE"

# Summary for stdout (captured by cron delivery)
echo "Ralph clippy: open=$OPEN enriched=$ENRICHED unenriched=$UNENRICHED pr=$PR_OPEN"
