#!/bin/bash
# 회사 Mac에서 1회 실행: launchd 등록
# 사용법: cd ~/code/etf-tracker && bash scripts/install_launchd.sh

PLIST_NAME="com.jay.etf-tracker"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_NAME}.plist"
SCRIPT_PATH="$HOME/code/etf-tracker/scripts/daily_update.sh"

chmod +x "$SCRIPT_PATH"

cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${PLIST_NAME}</string>
    <key>ProgramArguments</key>
    <array>
        <string>/bin/bash</string>
        <string>${SCRIPT_PATH}</string>
    </array>
    <key>StartCalendarInterval</key>
    <array>
        <dict><key>Weekday</key><integer>1</integer><key>Hour</key><integer>18</integer><key>Minute</key><integer>30</integer></dict>
        <dict><key>Weekday</key><integer>2</integer><key>Hour</key><integer>18</integer><key>Minute</key><integer>30</integer></dict>
        <dict><key>Weekday</key><integer>3</integer><key>Hour</key><integer>18</integer><key>Minute</key><integer>30</integer></dict>
        <dict><key>Weekday</key><integer>4</integer><key>Hour</key><integer>18</integer><key>Minute</key><integer>30</integer></dict>
        <dict><key>Weekday</key><integer>5</integer><key>Hour</key><integer>18</integer><key>Minute</key><integer>30</integer></dict>
    </array>
    <key>StandardOutPath</key>
    <string>${HOME}/code/etf-tracker/logs/launchd_out.log</string>
    <key>StandardErrorPath</key>
    <string>${HOME}/code/etf-tracker/logs/launchd_err.log</string>
</dict>
</plist>
EOF

# 기존 등록 해제 후 재등록
launchctl bootout gui/$(id -u) "$PLIST_PATH" 2>/dev/null || true
launchctl bootstrap gui/$(id -u) "$PLIST_PATH"

echo "✅ launchd 등록 완료: 평일 18:30 자동 실행"
echo "   plist: $PLIST_PATH"
echo "   로그:  ~/code/etf-tracker/logs/cron.log"
echo ""
echo "수동 테스트: bash ~/code/etf-tracker/scripts/daily_update.sh"
