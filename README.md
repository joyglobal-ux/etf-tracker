# 코스닥 액티브 ETF 트래커

TIME 코스닥액티브 (0162Y0) / KoAct 코스닥액티브 (0163Y0) 구성종목 일별 변화 추적

## 설치
```bash
cd ~/Desktop/투자/etf_tracker
pip install -r requirements.txt
```

## 실행
```bash
# 오늘 데이터 수집 + 리포트 생성
python3 kosdaq_etf_tracker.py

# 특정 날짜 조회
python3 kosdaq_etf_tracker.py --date 2026-03-28

# 강제 재수집 (캐시 무시)
python3 kosdaq_etf_tracker.py --force

# 리포트 열기
open kosdaq_etf_report.html
```

## 자동화 (cron)
```bash
# 평일 오후 6시 30분 자동 실행 (장 마감 + T+1 공시 후)
crontab -e
# 아래 줄 추가:
30 18 * * 1-5 cd /Users/jay/Desktop/투자/etf_tracker && /usr/bin/python3 kosdaq_etf_tracker.py >> logs/cron.log 2>&1
```

## 데이터 소스 우선순위
1. Naver Finance mobile API (JSON)
2. Naver Finance HTML 스크래핑
3. KRX 데이터 포털 (공식, OTP 기반)

공식 공시:
- TIME: https://timefolio.co.kr/etf/fund_disclosure_list.php
- KoAct: https://www.samsungactive.co.kr/etf/view.do?id=2ETFU6

## 파일 구조
```
etf_tracker/
├── kosdaq_etf_tracker.py   # 메인 스크립트
├── requirements.txt
├── kosdaq_etf_report.html  # 생성된 리포트 (daily)
├── data/
│   ├── 2026-03-30_0162Y0.json  # TIME 일별 스냅샷
│   └── 2026-03-30_0163Y0.json  # KoAct 일별 스냅샷
└── logs/
    ├── tracker.log
    └── cron.log
```
