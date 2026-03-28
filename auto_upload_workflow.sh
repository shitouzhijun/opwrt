#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$HOME/.openclaw/workspace"
BRANCH="${SYNC_BRANCH:-openclaw-sync}"
LOG_FILE="$REPO_DIR/.git-auto-sync.log"
SSH_CMD="ssh -o BatchMode=yes -o StrictHostKeyChecking=accept-new"

cd "$REPO_DIR"

git add -A

if git diff --cached --quiet; then
  exit 0
fi

COMMIT_MSG="openclaw auto sync: $(date '+%Y-%m-%d %H:%M:%S %z')"
if ! git commit -m "$COMMIT_MSG" >/dev/null 2>&1; then
  exit 0
fi

if ! timeout 180s env GIT_SSH_COMMAND="$SSH_CMD" git pull --rebase origin "$BRANCH" >>"$LOG_FILE" 2>&1; then
  printf '%s WARN pull failed for branch %s\n' "$(date '+%Y-%m-%d %H:%M:%S %z')" "$BRANCH" >>"$LOG_FILE"
fi

if ! timeout 180s env GIT_SSH_COMMAND="$SSH_CMD" git push origin "HEAD:$BRANCH" >>"$LOG_FILE" 2>&1; then
  printf '%s ERROR push failed for branch %s\n' "$(date '+%Y-%m-%d %H:%M:%S %z')" "$BRANCH" >>"$LOG_FILE"
  exit 1
fi
