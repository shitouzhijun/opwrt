#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$HOME/.openclaw/workspace"
BRANCH="${SYNC_BRANCH:-master}"
LOG_FILE="$REPO_DIR/.git-auto-sync.log"

cd "$REPO_DIR"

git add -A

if git diff --cached --quiet; then
  exit 0
fi

COMMIT_MSG="openclaw auto sync: $(date '+%Y-%m-%d %H:%M:%S %z')"
if ! git commit -m "$COMMIT_MSG" >/dev/null 2>&1; then
  exit 0
fi

git pull --rebase origin "$BRANCH" >>"$LOG_FILE" 2>&1 || true
git push origin "$BRANCH" >>"$LOG_FILE" 2>&1
