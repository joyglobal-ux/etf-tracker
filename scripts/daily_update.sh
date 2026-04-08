#!/bin/bash
# ETF Tracker — 매일 자동 수집 + git push
# launchd 또는 cron에서 호출

set -e

REPO_DIR="$HOME/code/etf-tracker"
LOG="$REPO_DIR/logs/cron.log"
mkdir -p "$REPO_DIR/logs"

{
  echo "========== $(date '+%Y-%m-%d %H:%M:%S') =========="

  cd "$REPO_DIR"
  git pull --quiet

  # Python 경로 (시스템 또는 brew)
  PYTHON=$(command -v python3)
  $PYTHON kosdaq_etf_tracker.py

  git add data/ kosdaq_etf_report.html
  if git diff --staged --quiet; then
    echo "No changes to commit"
  else
    git commit -m "chore: daily ETF snapshot $(date '+%Y-%m-%d')"
    git push
    echo "Pushed successfully"
  fi

  echo "Done."
} >> "$LOG" 2>&1
