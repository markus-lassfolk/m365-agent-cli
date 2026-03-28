#!/bin/bash
# Ralph Wiggum loop for clippy — the repo that notices dead birds
# Runs every 15 min. Monitors, enriches, and IMPLEMENTS.
set -euo pipefail

REPO="markus-lassfolk/clippy"
BRANCH="dev"
WORKSPACE="/home/markus/.openclaw/workspace/markus-lassfolk-clippy"
GOALS_FILE="$WORKSPACE/docs/GOALS.md"
ARCH_FILE="$WORKSPACE/docs/ARCHITECTURE.md"
LOG_DIR="$HOME/.openclaw/workspace/ralph/logs"
TODAY=$(date +%Y-%m-%d)
LOG_FILE="$LOG_DIR/clippy-$(date +%Y-%m-%d).md"
GUARD_FILE="$HOME/.openclaw/workspace/tmp/guards/ralph-clippy-wiggum.ms"
FORGE_GUARD="$HOME/.openclaw/workspace/tmp/guards/ralph-clippy-forge-running.txt"

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

log() {
  echo "$1" >> "$LOG_FILE"
}

log "Ralph clippy Wiggum run: $(date -Iminutes)"

# ── 1. Repo & issue stats ────────────────────────────────────────────────────
OPEN=$(gh issue list --repo "$REPO" --state open --json number --jq 'length')
ENRICHED=$(gh issue list --repo "$REPO" --state open --label enriched --json number --jq 'length')
UNENRICHED=$(gh issue list --repo "$REPO" --state open --json number,labels --jq '[.[] | select([.labels[].name] | index("enriched") | not)] | length' 2>/dev/null || echo 0)
log "Open: $OPEN | Enriched: $ENRICHED | Unenriched: $UNENRICHED"

# ── 2. Priority breakdown ────────────────────────────────────────────────────
for PRIORITY in Critical High Medium Low; do
  COUNT=$(gh issue list --repo "$REPO" --state open --label "Queue:$PRIORITY" --json number --jq 'length' 2>/dev/null || echo 0)
  log "Queue:$PRIORITY: $COUNT"
done

# ── 3. Enriched but unprioritized ──────────────────────────────────────────
UNPRIORITIZED=$(gh issue list --repo "$REPO" --state open --json number,title,labels \
  --jq '.[] | select([.labels[].name] | index("enriched")) | select([.labels[].name | startswith("Queue:")] | any | not) | "\(.number) \(.title)"' 2>/dev/null || true)
if [[ -n "$UNPRIORITIZED" ]]; then
  log "Enriched but unprioritized:"
  log "$UNPRIORITIZED"
fi

# ── 4. Stale unenriched ─────────────────────────────────────────────────────
OLD_UNENRICHED=$(gh issue list --repo "$REPO" --state open --json number,title,createdAt,labels \
  --jq '.[] | select([.labels[].name] | index("enriched") | not) | select((now - (.createdAt | fromdateiso8601) > 86400)) | "\(.number) \(.title) — created \(.createdAt[0:10])"' 2>/dev/null || true)
if [[ -n "$OLD_UNENRICHED" ]]; then
  log "Stale unenriched (>24h, needs Scholar):"
  log "$OLD_UNENRICHED"
fi

# ── 5. Security / reliability bugs ───────────────────────────────────────────
BUGS_URGENT=$(gh issue list --repo "$REPO" --state open --label bug --json number,title,labels \
  --jq '.[] | select(.labels[] | .name == "bug") | select(.labels[] | .name == "security" or .name == "reliability") | "\(.number) \(.title)"' 2>/dev/null || true)
if [[ -n "$BUGS_URGENT" ]]; then
  log "Security/Reliability bugs:"
  log "$BUGS_URGENT"
fi

# ── 6. Dev PR status ────────────────────────────────────────────────────────
PR_STATE=$(gh pr view "$BRANCH" --repo "$REPO" --json state,title,headRefName,additions,deletions 2>/dev/null || echo "{}")
PR_OPEN=$(echo "$PR_STATE" | jq -r '.state' 2>/dev/null || echo "unknown")
log "Dev PR: $PR_OPEN (branch=$BRANCH)"

# ── 7. Forge already running? ───────────────────────────────────────────────
if [[ -f "$FORGE_GUARD" ]]; then
  FORGE_AGE=$(($(date +%s) - $(cat "$FORGE_GUARD")))
  log "Forge guard age: ${FORGE_AGE}s"
  if (( FORGE_AGE < 2700 )); then
    log "Forge still active (${FORGE_AGE}s), skipping dispatch."
    echo "Ralph clippy: open=$OPEN enriched=$ENRICHED unenriched=$UNENRICHED | Forge active (${FORGE_AGE}s) — skipping"
    exit 0
  fi
  log "Forge guard stale (${FORGE_AGE}s > 45min), will consider dispatch."
fi

# ── 8. Scholar dispatch if unenriched issues exist ─────────────────────────
SCHOLAR_NEEDED=false
if (( UNENRICHED > 0 )); then
  log "Unenriched issues found — Scholar should enrich them."
  SCHOLAR_NEEDED=true
fi

# ── 9. Forge dispatch: highest-priority enriched+prioritized issue ─────────
FORGE_DISPATCHED=false
ISSUE_TO_WORK=""

if [[ "$SCHOLAR_NEEDED" == "false" ]]; then
  # No unenriched — try to find a Forge-able issue
  # Priority: Critical > High > Medium > Low
  for PRIORITY in Critical High Medium Low; do
    ISSUES=$(gh issue list --repo "$REPO" --state open --label "Queue:$PRIORITY" --label enriched \
      --json number,title,body,labels --jq '.[] | select(.labels[] | .name != "in-progress" and .name != "enhancement" or true) | .number' 2>/dev/null || true)
    if [[ -n "$ISSUES" ]]; then
      # Take the oldest (lowest number) issue
      ISSUE_TO_WORK=$(echo "$ISSUES" | head -1)
      break
    fi
  done

  if [[ -n "$ISSUE_TO_WORK" ]]; then
    ISSUE_TITLE=$(gh issue view "$ISSUE_TO_WORK" --repo "$REPO" --json title --jq '.title')
    ISSUE_BODY=$(gh issue view "$ISSUE_TO_WORK" --repo "$REPO" --json body --jq -r '.body // ""')
    ISSUE_LABELS=$(gh issue view "$ISSUE_TO_WORK" --repo "$REPO" --json labels --jq '.labels[].name | select(. != "enriched") | @json' -r)
    ENRICHMENT_COMMENT=$(gh issue view "$ISSUE_TO_WORK" --repo "$REPO" --comments --json body --jq '[.comments[] | select(.body | contains("Enrichment Analysis"))] | last.body' 2>/dev/null || echo "")

    log "Dispatching Forge for issue #$ISSUE_TO_WORK: $ISSUE_TITLE"

    # Mark in-progress
    gh issue edit "$ISSUE_TO_WORK" --repo "$REPO" --add-label "in-progress" 2>/dev/null || true

    # Write forge brief
    FORGE_BRIEF=$(mktemp)
    cat > "$FORGE_BRIEF" << BRIEF_EOF
## Forge Brief — Issue #$ISSUE_TO_WORK: $ISSUE_TITLE

### Issue Description
$ISSUE_BODY

### Enrichment Analysis
$ENRICHMENT_COMMENT

### Labels
$ISSUE_LABELS

### Repo Context
- Working branch: $BRANCH (tracked as origin/$BRANCH)
- PR open to main: $PR_OPEN
- Docs: docs/GOALS.md, docs/ARCHITECTURE.md
- Tests: src/test/ (run: npm test)
- Clippy SKILL: /home/markus/.openclaw/workspace/markus-lassfolk-clippy/SKILL/SKILL.md

### Instructions
1. Read docs/GOALS.md and docs/ARCHITECTURE.md for context
2. Implement the fix/feature for issue #$ISSUE_TO_WORK
3. Write or update tests in src/test/
4. Update SKILL/SKILL.md and README.md if behaviour changed
5. Commit to branch $BRANCH and push
6. Do NOT merge to main
7. When done: apply the "implemented" label to issue #$ISSUE_TO_WORK and remove "in-progress"
8. Reply with a concise summary of what was changed
BRIEF_EOF

    # Spawn Forge (async — don't wait)
    cd "$WORKSPACE"
    sessions_spawn_label="forge-clippy-$(date +%s)" # unique per dispatch
    sessions_spawn runtime="subagent" agentId="forge" mode="run" timeoutSeconds="3600" \
      task="$(cat "$FORGE_BRIEF")" label="$sessions_spawn_label" 2>/dev/null &
    FORGE_PID=$!

    # Write guard with PID
    echo "$(date +%s)" > "$FORGE_GUARD"
    echo "pid=$FORGE_PID label=$sessions_spawn_label issue=$ISSUE_TO_WORK" >> "$FORGE_GUARD"

    rm -f "$FORGE_BRIEF"
    FORGE_DISPATCHED=true
    log "Forge dispatched (pid=$FORGE_PID, label=$sessions_spawn_label, issue=#$ISSUE_TO_WORK)"

  else
    log "No enriched+prioritized issues to dispatch. Checking GOALS for gap-filling work..."
  fi
fi

# ── 10. GOALS gap-filling when issue queue is empty ─────────────────────────
if [[ "$FORGE_DISPATCHED" == "false" && "$SCHOLAR_NEEDED" == "false" ]]; then
  log "Issue queue empty — scanning GOALS.md for gap-filling work"

  if [[ -f "$GOALS_FILE" ]]; then
    # Look for unchecked items in GOALS.md
    UNDONE=$(grep -n "^- \[ \]" "$GOALS_FILE" 2>/dev/null || true)
    if [[ -n "$UNDONE" ]]; then
      log "Unchecked GOALS items found:"
      log "$UNDONE"
      # Pick the first undone item
      FIRST_LINE=$(echo "$UNDONE" | head -1 | cut -d: -f1)
      # Extract the goal text (everything after "- [ ] ")
      GOAL_TEXT=$(sed -n "${FIRST_LINE}s/- \[ \] //p" "$GOALS_FILE")
      log "Next gap-fill target: $GOAL_TEXT"

      # Build a brief for Forge to research + implement
      FORGE_BRIEF=$(mktemp)
      cat > "$FORGE_BRIEF" << BRIEF_EOF
## Forge Brief — GOALS Gap-Fill: $GOAL_TEXT

### What to do
This item appears in docs/GOALS.md as an unfinished goal. Research the gap and implement it.

### Context
- Issue queue is empty — this is gap-filling work
- Read docs/GOALS.md and docs/ARCHITECTURE.md for full context
- Repo: $REPO, branch: $BRANCH

### Steps
1. Read docs/GOALS.md and docs/ARCHITECTURE.md
2. Understand what this goal requires
3. Identify which files need changes
4. Implement the feature or improvement
5. Write or update tests
6. Update SKILL/SKILL.md if behaviour changed
7. Commit and push to $BRANCH
8. Mark the GOALS item as done (replace "- [ ]" with "- [x]") in a separate commit
9. Reply with summary of changes
BRIEF_EOF

      sessions_spawn runtime="subagent" agentId="forge" mode="run" timeoutSeconds="3600" \
        task="$(cat "$FORGE_BRIEF")" label="forge-clippy-goals-$(date +%s)" 2>/dev/null &
      FORGE_PID=$!

      echo "$(date +%s)" > "$FORGE_GUARD"
      echo "pid=$FORGE_PID label=forge-clippy-goals-$(date +%s) goals_gap=1" >> "$FORGE_GUARD"

      rm -f "$FORGE_BRIEF"
      log "Forge dispatched for GOALS gap-fill (pid=$FORGE_PID)"
      FORGE_DISPATCHED=true
    else
      log "All GOALS items are checked off. Clippy is fully GOALS-aligned! 🦊"
    fi
  fi
fi

# ── Summary ─────────────────────────────────────────────────────────────────
if [[ "$FORGE_DISPATCHED" == "true" ]]; then
  echo "Ralph clippy: Forge dispatched | open=$OPEN enriched=$ENRICHED"
elif [[ "$SCHOLAR_NEEDED" == "true" ]]; then
  echo "Ralph clippy: $UNENRICHED unenriched issues (Scholar should run) | open=$OPEN"
else
  echo "Ralph clippy: all clear | open=$OPEN enriched=$ENRICHED"
fi

log "Run complete: $(date -Iminutes)"
log "---"
