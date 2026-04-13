#!/bin/bash
# ETF Tracker 자동 업데이트 + GitHub push
# cron: 평일 18:30 실행

ETF_DIR="/Users/a12668/Library/Mobile Documents/com~apple~CloudDocs/Claude_sharing/etf_tracker"
PYTHON="/usr/bin/python3"
LOG="$ETF_DIR/logs/cron.log"

echo "==============================" >> "$LOG"
echo "[$(date '+%Y-%m-%d %H:%M:%S')] Starting auto update" >> "$LOG"

# 1. ETF 데이터 수집 + 리포트 생성
cd "$ETF_DIR"
$PYTHON kosdaq_etf_tracker.py >> "$LOG" 2>&1

if [ $? -ne 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ERROR: Python script failed" >> "$LOG"
    exit 1
fi

# 2. Git pull (개인 맥북 커밋 동기화) + commit + push
cd "$ETF_DIR"
DATE=$(date '+%Y-%m-%d')

git pull origin main --no-rebase --no-edit >> "$LOG" 2>&1

git add data/ kosdaq_etf_report.html >> "$LOG" 2>&1
git commit -m "Daily update: $DATE" >> "$LOG" 2>&1

if [ $? -eq 0 ]; then
    git push origin main >> "$LOG" 2>&1
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Push complete" >> "$LOG"
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] No changes to commit" >> "$LOG"
fi
