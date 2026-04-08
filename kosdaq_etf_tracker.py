#!/usr/bin/env python3
"""
코스닥 액티브 ETF 구성종목 변화 트래커
────────────────────────────────────────
TIME 코스닥액티브  (0162Y0) — 타임폴리오자산운용
KoAct 코스닥액티브 (0163Y0) — 삼성액티브자산운용

데이터 소스:
  TIME  → https://timeetf.co.kr/pdf_excel.php?idx=24&pdfDate=YYYY-MM-DD  (xlsx)
  KoAct → https://www.samsungactive.co.kr/api/v1/product/etf-pdf/2ETFU6.do?gijunYMD=YYYYMMDD  (JSON)

실행: python3 kosdaq_etf_tracker.py
출력: kosdaq_etf_report.html
"""

import requests
import json
import os
import sys
import re
import io
import time
import zipfile
import argparse
import logging
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False

# ─────────────────────────────────────────
# 설정
# ─────────────────────────────────────────

ETF_CONFIG = {
    "TIME": {
        "code": "0162Y0",
        "isin": "KR70162Y0008",
        "name": "TIME 코스닥액티브",
        "manager": "타임폴리오자산운용",
        "fee": "0.80%",
        "color": "#2c7be5",
        "source": "timeetf",
        "source_idx": 24,
        "disclosure_url": "https://timeetf.co.kr/m11_view.php?idx=24",
    },
    "KoAct": {
        "code": "0163Y0",
        "isin": "KR70163Y0007",
        "name": "KoAct 코스닥액티브",
        "manager": "삼성액티브자산운용",
        "fee": "0.50%",
        "color": "#e63757",
        "source": "samsung",
        "source_fid": "2ETFU6",
        "api_base": "https://www.samsungactive.co.kr/api/v1/product/etf-pdf",
        "referer": "https://www.samsungactive.co.kr/etf/view.do?id=2ETFU6",
        "disclosure_url": "https://www.samsungactive.co.kr/etf/view.do?id=2ETFU6",
    },
    "KODEX신재생": {
        "code": "385510",
        "name": "KODEX 신재생에너지액티브",
        "manager": "삼성자산운용",
        "fee": "0.50%",
        "color": "#4caf80",
        "source": "samsung",
        "source_fid": "2ETFE5",
        "api_base": "https://www.samsungfund.com/api/v1/kodex/product-pdf",
        "referer": "https://www.samsungfund.com/etf/product/view.do?id=2ETFE5",
        "disclosure_url": "https://www.samsungfund.com/etf/product/view.do?id=2ETFE5",
    },
    "TIME_AI": {
        "code": "456600",
        "name": "TIME 글로벌AI인공지능액티브",
        "manager": "타임폴리오자산운용",
        "fee": "0.80%",
        "color": "#ab47bc",
        "source": "timeetf",
        "source_idx": 6,
        "disclosure_url": "https://timeetf.co.kr/m11_view.php?idx=6",
    },
}

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"
DATA_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/html, */*",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}

XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Naver Finance ETF API (ETF 시세/순자산/거래대금 일괄 조회)
NAVER_ETF_API = "https://finance.naver.com/api/sise/etfItemList.nhn?etfType=0&targetColumn=market_sum&sortOrder=desc"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOGS_DIR / "tracker.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────
# ETF 시세/순자산/거래대금 (Naver Finance)
# ─────────────────────────────────────────

def fetch_etf_metadata() -> dict:
    """Naver Finance ETF API에서 4개 ETF의 시세/순자산/거래대금을 일괄 조회.
    Returns: {etf_code: {nowVal, changeRate, marketSum, amount, nav}} 형태."""
    our_codes = {cfg["code"] for cfg in ETF_CONFIG.values()}
    try:
        resp = requests.get(NAVER_ETF_API, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            log.warning(f"  Naver ETF API: HTTP {resp.status_code}")
            return {}
        data = resp.json()
        items = data.get("result", {}).get("etfItemList", [])
        result = {}
        for item in items:
            code = item.get("itemcode", "")
            if code in our_codes:
                result[code] = {
                    "nowVal": item.get("nowVal"),           # 현재가
                    "changeVal": item.get("changeVal"),     # 전일 대비 변동
                    "changeRate": item.get("changeRate"),    # 등락률(%)
                    "nav": item.get("nav"),                 # 기준가(NAV)
                    "amount": item.get("amonut"),           # 거래대금 (백만원)
                    "marketSum": item.get("marketSum"),      # 순자산총액 (억원)
                    "quant": item.get("quant"),             # 거래량
                }
        if result:
            log.info(f"  📡 Naver ETF 시세 조회 완료: {len(result)}개 ETF")
        return result
    except Exception as e:
        log.warning(f"  Naver ETF API 조회 실패: {e}")
        return {}


# ─────────────────────────────────────────
# 데이터 수집 — TIME ETF (timeetf.co.kr)
# ─────────────────────────────────────────

def _parse_xlsx_holdings(content: bytes) -> list[dict]:
    """xlsx 바이너리에서 종목 리스트 추출 (zipfile + ElementTree)"""
    with zipfile.ZipFile(io.BytesIO(content)) as z:
        names = z.namelist()

        # 공유 문자열 로드
        strings: list[str] = []
        if "xl/sharedStrings.xml" in names:
            with z.open("xl/sharedStrings.xml") as f:
                root = ET.parse(f).getroot()
                for si in root.iter(f"{{{XLSX_NS}}}si"):
                    t = "".join(x.text or "" for x in si.iter(f"{{{XLSX_NS}}}t"))
                    strings.append(t)

        if len(strings) <= 5:
            return []  # 헤더만 있고 데이터 없음

        # sheet1 파싱
        sheet_files = [n for n in names if "worksheets/sheet" in n]
        if not sheet_files:
            return []
        with z.open(sheet_files[0]) as f:
            root = ET.parse(f).getroot()

        rows = []
        for row_el in root.iter(f"{{{XLSX_NS}}}row"):
            row_data = []
            for c in row_el.iter(f"{{{XLSX_NS}}}c"):
                t_attr = c.get("t", "")
                v_el = c.find(f"{{{XLSX_NS}}}v")
                if v_el is None:
                    row_data.append("")
                    continue
                val = v_el.text or ""
                if t_attr == "s":
                    idx = int(val)
                    val = strings[idx] if idx < len(strings) else val
                row_data.append(val)
            rows.append(row_data)

        # 헤더 행 건너뛰고 파싱 (col: 종목코드, 종목명, 수량, 평가금액, 비중%)
        holdings = []
        for row in rows[1:]:
            if len(row) < 5:
                continue
            code = row[0].strip()
            name = row[1].strip()
            weight_str = row[4].strip()
            try:
                weight = float(weight_str)
                if name and weight > 0:
                    holdings.append({"name": name, "code": code, "weight": weight})
            except ValueError:
                continue
        return holdings


def fetch_timeetf_excel(date_str: str, idx: int = 24) -> list[dict]:
    """TIME ETF: timeetf.co.kr Excel(xlsx) 다운로드"""
    url = f"https://timeetf.co.kr/pdf_excel.php?idx={idx}&pdfDate={date_str}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
        if resp.status_code != 200 or len(resp.content) < 1000:
            log.debug(f"  timeetf Excel: {resp.status_code} / {len(resp.content)} bytes")
            return []
        holdings = _parse_xlsx_holdings(resp.content)
        if holdings:
            log.info(f"  ✅ timeetf.co.kr idx={idx} ({date_str}): {len(holdings)}종목")
        return sorted(holdings, key=lambda x: x["weight"], reverse=True)
    except Exception as e:
        log.debug(f"  timeetf Excel failed ({date_str}): {e}")
        return []


# ─────────────────────────────────────────
# 데이터 수집 — KoAct ETF (Samsung Active API)
# ─────────────────────────────────────────

def fetch_samsung_api(date_str: str, api_base: str, fid: str, referer: str) -> list[dict]:
    """삼성 계열 JSON API (samsungactive.co.kr, samsungfund.com 공통 스키마)"""
    gijun = date_str.replace("-", "")
    url = f"{api_base}/{fid}.do?gijunYMD={gijun}"
    hdrs = {**HEADERS, "Referer": referer}
    try:
        resp = requests.get(url, headers=hdrs, timeout=15)
        if resp.status_code != 200:
            log.debug(f"  Samsung API ({fid}): HTTP {resp.status_code}")
            return []
        data = resp.json()
        pdf = data.get("pdf", {})
        actual_date = pdf.get("gijunYMD", gijun)
        lst = pdf.get("list", [])

        holdings = []
        for item in lst:
            ratio_str = item.get("ratio", "")
            if not ratio_str:          # 현금성 자산 등 제외
                continue
            name = item.get("secNm", "").strip()
            code = item.get("itmNo", "").strip()
            try:
                weight = float(ratio_str)
                if name and weight > 0:
                    holdings.append({"name": name, "code": code, "weight": weight})
            except ValueError:
                continue

        if holdings:
            log.info(
                f"  ✅ Samsung API {fid} (기준일 {actual_date}): {len(holdings)}종목"
            )
        return sorted(holdings, key=lambda x: x["weight"], reverse=True)
    except Exception as e:
        log.debug(f"  Samsung API failed ({fid}, {date_str}): {e}")
        return []


# ─────────────────────────────────────────
# 통합 수집 (날짜 자동 후퇴)
# ─────────────────────────────────────────

def _prev_weekdays(date_str: str, n: int = 5) -> list[str]:
    """date_str 포함 이전 평일 n개 반환"""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    result = []
    while len(result) < n:
        if dt.weekday() < 5:  # 월~금
            result.append(dt.strftime("%Y-%m-%d"))
        dt -= timedelta(days=1)
    return result


def fetch_holdings(etf_key: str, date_str: str = None) -> list[dict]:
    """지정 날짜(또는 최근 영업일) 데이터를 수집. 성공 시 즉시 반환."""
    cfg = ETF_CONFIG[etf_key]
    target = date_str or datetime.now().strftime("%Y-%m-%d")
    log.info(f"\n[{etf_key}] {cfg['name']} 수집 중... (요청일: {target})")

    candidates = _prev_weekdays(target, n=5)

    source = cfg.get("source")
    if source == "timeetf":
        idx = cfg.get("source_idx", 24)
        for d in candidates:
            holdings = fetch_timeetf_excel(d, idx=idx)
            if holdings:
                return holdings
    elif source == "samsung":
        for d in candidates:
            holdings = fetch_samsung_api(
                d, cfg["api_base"], cfg["source_fid"], cfg["referer"]
            )
            if holdings:
                return holdings

    log.warning(
        f"  ❌ {etf_key}: 데이터 수집 실패. 공시 페이지 확인: {cfg['disclosure_url']}"
    )
    return []


# ─────────────────────────────────────────
# 데이터 저장/로드
# ─────────────────────────────────────────

def data_path(etf_key: str, date_str: str) -> Path:
    return DATA_DIR / f"{date_str}_{ETF_CONFIG[etf_key]['code']}.json"


def save_holdings(etf_key: str, holdings: list, date_str: str):
    path = data_path(etf_key, date_str)
    payload = {
        "date": date_str,
        "etf": etf_key,
        "code": ETF_CONFIG[etf_key]["code"],
        "fetched_at": datetime.now().isoformat(),
        "holdings": holdings,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    log.info(f"  💾 저장: {path.name} ({len(holdings)}종목)")


def load_holdings(etf_key: str, date_str: str) -> list:
    path = data_path(etf_key, date_str)
    if not path.exists():
        return []
    with open(path, encoding="utf-8") as f:
        return json.load(f).get("holdings", [])


def get_prev_date(etf_key: str, before: str) -> Optional[str]:
    """before 이전 데이터 파일 중 가장 최신 날짜 반환"""
    code = ETF_CONFIG[etf_key]["code"]
    files = sorted(DATA_DIR.glob(f"*_{code}.json"), reverse=True)
    for f in files:
        fdate = f.name[:10]
        if fdate < before:
            return fdate
    return None


def get_nearest_date(etf_key: str, target: str, not_after: str = None) -> Optional[str]:
    """target에 가장 가까운 데이터 날짜 (앞뒤 모두 탐색). not_after 이전만 허용."""
    code = ETF_CONFIG[etf_key]["code"]
    dates = sorted(f.name[:10] for f in DATA_DIR.glob(f"*_{code}.json"))
    if not dates:
        return None
    # not_after 이전 날짜만 필터링
    if not_after:
        dates = [d for d in dates if d < not_after]
    if not dates:
        return None
    # target에 가장 가까운 날짜
    return min(dates, key=lambda d: abs(
        (datetime.strptime(d, "%Y-%m-%d") - datetime.strptime(target, "%Y-%m-%d")).days
    ))


def get_all_dates(etf_key: str) -> list[str]:
    """해당 ETF의 전체 데이터 날짜 목록 (오래된 순)"""
    code = ETF_CONFIG[etf_key]["code"]
    return sorted(f.name[:10] for f in DATA_DIR.glob(f"*_{code}.json"))


# 비교 기간 정의
PERIODS = [
    ("1d",  "전일 대비",  1),
    ("1w",  "1주 전",     7),
    ("1m",  "1개월 전",  30),
]


# ─────────────────────────────────────────
# 변화 분석
# ─────────────────────────────────────────

def diff_holdings(prev: list, curr: list) -> dict:
    prev_map = {h["name"]: h for h in prev}
    curr_map = {h["name"]: h for h in curr}
    new, removed, increased, decreased, unchanged = [], [], [], [], []

    for name, h in curr_map.items():
        if name not in prev_map:
            new.append(h)
        else:
            delta = round(h["weight"] - prev_map[name]["weight"], 3)
            entry = {**h, "prev_weight": prev_map[name]["weight"], "delta": delta}
            if delta > 0.1:
                increased.append(entry)
            elif delta < -0.1:
                decreased.append(entry)
            else:
                unchanged.append(entry)

    for name, h in prev_map.items():
        if name not in curr_map:
            removed.append({**h, "prev_weight": h["weight"]})

    return {
        "new":       sorted(new, key=lambda x: x["weight"], reverse=True),
        "removed":   sorted(removed, key=lambda x: x["weight"], reverse=True),
        "increased": sorted(increased, key=lambda x: x["delta"], reverse=True),
        "decreased": sorted(decreased, key=lambda x: x["delta"]),
        "unchanged": sorted(unchanged, key=lambda x: x["weight"], reverse=True),
    }


# ─────────────────────────────────────────
# HTML 생성
# ─────────────────────────────────────────

def _rows(items, style):
    if not items:
        return '<tr><td colspan="4" class="empty">없음</td></tr>'
    rows = []
    for h in items:
        w = h.get("weight", 0)
        name = h.get("name", "")
        code = h.get("code", "")
        prev_w = h.get("prev_weight")
        delta = h.get("delta")
        code_link = (
            f'<a href="https://finance.naver.com/item/main.nhn?code={code}" '
            f'target="_blank" class="stock-link">{name}</a>'
            if re.match(r"^\d{6}$", code or "")
            else name
        )
        if style == "new":
            rows.append(
                f'<tr class="row-new"><td><span class="badge new">NEW</span>{code_link}</td>'
                f'<td class="num green">{w:.2f}%</td><td class="num">—</td>'
                f'<td class="num green">+{w:.2f}%</td></tr>'
            )
        elif style == "removed":
            rows.append(
                f'<tr class="row-removed"><td><span class="badge out">OUT</span>{name}</td>'
                f'<td class="num red">—</td><td class="num">{prev_w:.2f}%</td>'
                f'<td class="num red">-{prev_w:.2f}%</td></tr>'
            )
        elif style in ("increased", "decreased"):
            d_cls = "green" if delta > 0 else "orange"
            d_str = f"+{delta:.2f}%" if delta > 0 else f"{delta:.2f}%"
            rows.append(
                f'<tr><td>{code_link}</td><td class="num">{w:.2f}%</td>'
                f'<td class="num muted">{prev_w:.2f}%</td>'
                f'<td class="num {d_cls} bold">{d_str}</td></tr>'
            )
        else:
            rows.append(
                f'<tr><td>{code_link}</td><td class="num">{w:.2f}%</td>'
                f'<td class="num muted">{prev_w:.2f}%</td><td></td></tr>'
            )
    return "\n".join(rows)


def _full_rows(curr, prev_map):
    rows = []
    for i, h in enumerate(curr, 1):
        w = h.get("weight", 0)
        code = h.get("code", "")
        name = h.get("name", "")
        prev = prev_map.get(name, {})
        pw = prev.get("weight")
        code_link = (
            f'<a href="https://finance.naver.com/item/main.nhn?code={code}" '
            f'target="_blank" class="stock-link">{name}</a>'
            if re.match(r"^\d{6}$", code or "")
            else name
        )
        if pw is not None:
            delta = w - pw
            d_str = f"+{delta:.2f}%" if delta > 0 else f"{delta:.2f}%"
            d_cls = "green" if delta > 0.1 else ("orange" if delta < -0.1 else "muted")
        else:
            d_str = "NEW"
            d_cls = "green bold"
        pw_str = f"{pw:.2f}%" if pw is not None else "—"
        rows.append(
            f'<tr><td class="num muted">{i}</td><td>{code_link}</td>'
            f'<td class="num bold">{w:.2f}%</td>'
            f'<td class="num muted">{pw_str}</td>'
            f'<td class="num {d_cls}">{d_str}</td></tr>'
        )
    return "\n".join(rows)


def _build_diff_sections(diff: dict) -> str:
    """diff 결과를 HTML section 블록으로 변환"""
    sections = []
    if diff["new"]:
        sections.append(f"""
        <div class="section">
          <div class="sec-label green">🟢 신규 진입 — {len(diff['new'])}종목</div>
          <table><tr><th>종목</th><th class="num">현재</th><th class="num">이전</th><th class="num">변화</th></tr>
          {_rows(diff['new'], 'new')}</table>
        </div>""")
    if diff["removed"]:
        sections.append(f"""
        <div class="section">
          <div class="sec-label red">🔴 완전 청산 — {len(diff['removed'])}종목</div>
          <table><tr><th>종목</th><th class="num">현재</th><th class="num">이전</th><th class="num">변화</th></tr>
          {_rows(diff['removed'], 'removed')}</table>
        </div>""")
    if diff["increased"]:
        sections.append(f"""
        <div class="section">
          <div class="sec-label lgreen">📈 비중 확대 — {len(diff['increased'])}종목 (상위 15)</div>
          <table><tr><th>종목</th><th class="num">현재</th><th class="num">이전</th><th class="num">변화</th></tr>
          {_rows(diff['increased'][:15], 'increased')}</table>
        </div>""")
    if diff["decreased"]:
        sections.append(f"""
        <div class="section">
          <div class="sec-label orange">📉 비중 축소 — {len(diff['decreased'])}종목 (상위 15)</div>
          <table><tr><th>종목</th><th class="num">현재</th><th class="num">이전</th><th class="num">변화</th></tr>
          {_rows(diff['decreased'][:15], 'decreased')}</table>
        </div>""")
    if not any([diff["new"], diff["removed"], diff["increased"], diff["decreased"]]):
        sections.append('<div class="no-change">변화 없음</div>')
    return "".join(sections)


def _format_amount(val):
    """백만원 → 억원 또는 조원 단위 포맷"""
    if val is None:
        return "—"
    if val >= 1_000_000:
        return f"{val / 1_000_000:.1f}조"
    if val >= 10_000:
        return f"{val / 100:.0f}억"
    if val >= 1_000:
        return f"{val / 100:.1f}억"
    return f"{val:,.0f}백만"


def _format_market_sum(val):
    """억원 단위 포맷"""
    if val is None:
        return "—"
    if val >= 10_000:
        return f"{val / 10_000:.2f}조"
    return f"{val:,}억"


def generate_html(results: dict, today: str, etf_meta: dict = None) -> Path:
    etf_meta = etf_meta or {}
    panels_html = []
    for etf_key, data in results.items():
        cfg = ETF_CONFIG[etf_key]
        curr = data["curr"]
        periods_data = data["periods"]  # {period_key: {diff, prev, prev_date}}

        # 탭 버튼 + 기간별 콘텐츠
        tab_buttons = []
        tab_contents = []
        first_diff = None
        for i, (pkey, plabel, _days) in enumerate(PERIODS):
            pd = periods_data.get(pkey, {})
            diff = pd.get("diff")
            prev_date = pd.get("prev_date")
            active = "active" if i == 0 else ""

            if diff and prev_date:
                tab_buttons.append(
                    f'<button class="tab-btn {active}" '
                    f'onclick="switchTab(\'{etf_key}\',\'{pkey}\')"'
                    f' data-period="{pkey}">{plabel}'
                    f'<span class="tab-date">{prev_date}</span></button>'
                )
                chip_bar = (
                    f'<div class="chip-bar">'
                    f'<span class="chip c-new">신규 +{len(diff["new"])}</span>'
                    f'<span class="chip c-out">청산 -{len(diff["removed"])}</span>'
                    f'<span class="chip c-up">확대 ↑{len(diff["increased"])}</span>'
                    f'<span class="chip c-dn">축소 ↓{len(diff["decreased"])}</span>'
                    f'</div>'
                )
                sections_html = _build_diff_sections(diff)
                display = "block" if i == 0 else "none"
                tab_contents.append(
                    f'<div class="tab-content" id="tab_{etf_key}_{pkey}" '
                    f'style="display:{display}">{chip_bar}{sections_html}</div>'
                )
                if first_diff is None:
                    first_diff = diff
            else:
                tab_buttons.append(
                    f'<button class="tab-btn {active} disabled" '
                    f'data-period="{pkey}">{plabel}'
                    f'<span class="tab-date">데이터 없음</span></button>'
                )
                display = "block" if i == 0 else "none"
                all_dates = get_all_dates(etf_key)
                start = all_dates[0] if all_dates else "—"
                tab_contents.append(
                    f'<div class="tab-content" id="tab_{etf_key}_{pkey}" '
                    f'style="display:{display}">'
                    f'<div class="no-change">아직 충분한 데이터가 없습니다<br>'
                    f'<span class="muted" style="font-size:0.8rem">'
                    f'수집 시작: {start} · 데이터가 쌓이면 자동 활성화</span></div></div>'
                )

        # 전체 종목 테이블은 전일 대비 기준
        default_prev = periods_data.get("1d", {}).get("prev", [])
        prev_map = {h["name"]: h for h in default_prev}
        full_rows = _full_rows(curr, prev_map)

        # ETF 시세 메타데이터
        meta = etf_meta.get(cfg["code"], {})
        meta_html = ""
        if meta:
            price = meta.get("nowVal")
            chg_rate = meta.get("changeRate")
            mkt_sum = meta.get("marketSum")
            amount = meta.get("amount")

            price_str = f"{price:,}" if price else "—"
            chg_cls = "green" if (chg_rate or 0) > 0 else ("red" if (chg_rate or 0) < 0 else "muted")
            chg_sign = "+" if (chg_rate or 0) > 0 else ""
            chg_str = f"{chg_sign}{chg_rate:.2f}%" if chg_rate is not None else "—"
            mkt_str = _format_market_sum(mkt_sum)
            amt_str = _format_amount(amount)

            meta_html = f"""
            <div class="etf-meta-bar">
              <div class="meta-item"><span class="meta-label">현재가</span><span class="meta-val">{price_str}</span><span class="meta-chg {chg_cls}">{chg_str}</span></div>
              <div class="meta-item"><span class="meta-label">순자산</span><span class="meta-val">{mkt_str}</span></div>
              <div class="meta-item"><span class="meta-label">거래대금</span><span class="meta-val">{amt_str}</span></div>
            </div>"""

        panel = f"""
        <div class="panel">
          <div class="panel-header" style="border-top:3px solid {cfg['color']}">
            <div class="panel-name" style="color:{cfg['color']}">{cfg['name']}</div>
            <div class="panel-meta">{cfg['manager']} · {cfg['code']} · 운보수 {cfg['fee']}</div>
            <div class="panel-meta muted">{len(curr)}종목 보유 · 기준일: {today}</div>
            {meta_html}
          </div>
          <div class="tab-bar" id="tabs_{etf_key}">
            {"".join(tab_buttons)}
          </div>
          {"".join(tab_contents)}
          <div class="toggle-wrap">
            <button onclick="toggle('full_{etf_key}', {len(curr)})" id="btn_{etf_key}" class="toggle-btn">
              ▼ 전체 {len(curr)}종목 보기
            </button>
            <table id="full_{etf_key}" class="full-table" style="display:none">
              <tr><th>#</th><th>종목</th><th class="num">비중</th><th class="num">전일</th><th class="num">변화</th></tr>
              {full_rows}
            </table>
          </div>
        </div>"""
        panels_html.append(panel)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>코스닥 액티브 ETF 트래커 — {today}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#0d0d1a;color:#e0e0e0;font-family:'Apple SD Gothic Neo','Pretendard',sans-serif;line-height:1.7;padding:20px}}
.container{{max-width:1440px;margin:0 auto}}
h1{{text-align:center;font-size:1.7rem;color:#fff;padding:30px 0 6px}}
.subtitle{{text-align:center;color:#888;font-size:0.9rem;margin-bottom:6px}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:20px}}
@media(max-width:960px){{.grid{{grid-template-columns:1fr}}}}
.panel{{background:#12122a;border-radius:12px;overflow:hidden}}
.panel-header{{padding:18px 20px 12px}}
.panel-name{{font-size:1.2rem;font-weight:800;margin-bottom:3px}}
.panel-meta{{font-size:0.82rem;color:#888;margin-bottom:2px}}
.tab-bar{{display:flex;border-bottom:1px solid #1e1e3a}}
.tab-btn{{flex:1;background:none;border:none;color:#666;padding:10px 8px 8px;font-size:0.82rem;font-weight:600;cursor:pointer;border-bottom:2px solid transparent;transition:all .15s;display:flex;flex-direction:column;align-items:center;gap:2px}}
.tab-btn:hover{{color:#aaa;background:#151530}}
.tab-btn.active{{color:#fff;border-bottom-color:#3498db}}
.tab-btn.disabled{{color:#333;cursor:default}}
.tab-btn.disabled:hover{{background:none;color:#333}}
.tab-date{{font-size:0.68rem;color:#555;font-weight:400}}
.tab-btn.active .tab-date{{color:#3498db}}
.chip-bar{{display:flex;gap:8px;flex-wrap:wrap;padding:10px 20px 14px}}
.chip{{padding:3px 10px;border-radius:20px;font-size:0.75rem;font-weight:700}}
.c-new{{background:#1e4d2b;color:#4caf80}}
.c-out{{background:#4d1e1e;color:#ef5350}}
.c-up{{background:#1a3d1a;color:#66bb6a}}
.c-dn{{background:#3d2a1a;color:#ffa726}}
.c-total{{background:#1e1e3a;color:#7986cb}}
.section{{padding:14px 20px;border-bottom:1px solid #111125}}
.sec-label{{font-size:0.72rem;font-weight:700;letter-spacing:0.06em;padding:3px 8px;border-radius:4px;display:inline-block;margin-bottom:10px}}
.sec-label.green{{background:#1e4d2b;color:#4caf80}}
.sec-label.red{{background:#4d1e1e;color:#ef5350}}
.sec-label.lgreen{{background:#1a3d1a;color:#66bb6a}}
.sec-label.orange{{background:#3d2a1a;color:#ffa726}}
table{{width:100%;border-collapse:collapse;font-size:0.83rem;margin-top:4px}}
th{{color:#555;font-weight:600;padding:5px 6px;text-align:left;border-bottom:1px solid #1e1e3a;font-size:0.76rem}}
td{{padding:4px 6px;border-bottom:1px solid #111125;vertical-align:middle}}
tr:hover td{{background:#1a1a40}}
.num{{text-align:right;font-variant-numeric:tabular-nums}}
.bold{{font-weight:700}}
.green{{color:#4caf80}}
.red{{color:#ef5350}}
.orange{{color:#ffa726}}
.muted{{color:#555}}
.badge{{font-size:0.65rem;font-weight:800;padding:1px 5px;border-radius:3px;margin-right:4px}}
.badge.new{{background:#1e4d2b;color:#4caf80}}
.badge.out{{background:#4d1e1e;color:#ef5350}}
.row-new td{{background:#111e17}}
.row-removed td{{background:#1e1111;opacity:.8}}
.empty{{text-align:center;color:#444;padding:10px;font-style:italic}}
.no-change{{text-align:center;color:#444;padding:20px;font-style:italic}}
.toggle-wrap{{padding:14px 20px}}
.toggle-btn{{background:#1e1e3a;border:1px solid #2d2d4e;color:#aaa;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:0.82rem;width:100%}}
.toggle-btn:hover{{background:#2d2d4e;color:#fff}}
.full-table{{margin-top:10px}}
.stock-link{{color:#e0e0e0;text-decoration:none}}
.stock-link:hover{{color:#3498db;text-decoration:underline}}
.source{{text-align:center;color:#333;font-size:0.75rem;margin-top:28px;padding-bottom:20px}}
.source a{{color:#3498db}}
.etf-meta-bar{{display:flex;gap:16px;margin-top:10px;padding:8px 12px;background:#0d0d1a;border-radius:8px}}
.meta-item{{display:flex;align-items:baseline;gap:5px}}
.meta-label{{font-size:0.7rem;color:#555;font-weight:600}}
.meta-val{{font-size:0.95rem;color:#fff;font-weight:700;font-variant-numeric:tabular-nums}}
.meta-chg{{font-size:0.8rem;font-weight:700;margin-left:2px}}
.refresh-hint{{text-align:center;background:#12122a;border-radius:8px;padding:12px;margin-bottom:20px;font-size:0.82rem;color:#888}}
.refresh-hint code{{background:#1e1e3a;padding:2px 6px;border-radius:3px;color:#7ecfff;font-size:0.78rem}}
</style>
</head>
<body>
<div class="container">
  <h1>📊 액티브 ETF 트래커</h1>
  <p class="subtitle">{" · ".join(f'{c["name"]} ({c["code"]})' for c in ETF_CONFIG.values())}</p>
  <p class="subtitle">기준일: <strong style="color:#fff">{today}</strong></p>
  <div class="refresh-hint">
    🔄 매일 오후 6시 이후: <code>python3 kosdaq_etf_tracker.py</code>
    &nbsp;·&nbsp; 특정일: <code>python3 kosdaq_etf_tracker.py --date 2026-03-28</code>
    &nbsp;·&nbsp; 강제 재수집: <code>python3 kosdaq_etf_tracker.py --force</code>
  </div>
  <div class="grid">
    {"".join(panels_html)}
  </div>
  <p class="source">
    📡 출처: {" · ".join(f'<a href="{c["disclosure_url"]}" target="_blank">{c["name"]}</a>' for c in ETF_CONFIG.values())}<br>
    ⚠️ 액티브 ETF는 익일 공시 원칙 (T+1). 비중은 전일 종가 기준.
  </p>
</div>
<script>
function switchTab(etf, period){{
  var tabs=document.querySelectorAll('#tabs_'+etf+' .tab-btn');
  tabs.forEach(function(t){{
    t.classList.remove('active');
    if(t.dataset.period===period && !t.classList.contains('disabled'))t.classList.add('active');
  }});
  var contents=document.querySelectorAll('[id^="tab_'+etf+'_"]');
  contents.forEach(function(c){{c.style.display='none'}});
  var target=document.getElementById('tab_'+etf+'_'+period);
  if(target)target.style.display='block';
}}
function toggle(id, cnt){{
  var el=document.getElementById(id);
  var btn=document.getElementById('btn_'+id.split('_')[1]);
  var show=el.style.display==='none';
  el.style.display=show?'table':'none';
  btn.textContent=show?'▲ 접기':'▼ 전체 '+cnt+'종목 보기';
}}
</script>
</body>
</html>"""
    out = BASE_DIR / "kosdaq_etf_report.html"
    out.write_text(html, encoding="utf-8")
    log.info(f"\n✅ 리포트 생성: {out}")
    return out


# ─────────────────────────────────────────
# 메인
# ─────────────────────────────────────────

def run(target_date: str = None, force_fetch: bool = False):
    today = target_date or datetime.now().strftime("%Y-%m-%d")
    log.info(f"\n{'='*55}")
    log.info(f"코스닥 액티브 ETF 트래커 — {today}")
    log.info(f"{'='*55}")

    # ETF 시세/순자산/거래대금 조회 (Naver Finance)
    etf_meta = fetch_etf_metadata()

    results = {}
    for etf_key in ETF_CONFIG:
        # 오늘 데이터 로드 or 수집
        curr = load_holdings(etf_key, today)
        if not curr or force_fetch:
            curr = fetch_holdings(etf_key, today)
            if curr:
                save_holdings(etf_key, curr, today)
            else:
                log.warning(f"⚠️  {etf_key}: 데이터 없음. 빈 상태로 리포트 생성.")
                curr = []
        else:
            log.info(f"  📦 {etf_key}: 캐시 사용 ({len(curr)}종목)")

        # 다중 기간 비교
        periods_data = {}
        today_dt = datetime.strptime(today, "%Y-%m-%d")
        used_dates = set()
        for pkey, plabel, days in PERIODS:
            target_dt = today_dt - timedelta(days=days)
            target_str = target_dt.strftime("%Y-%m-%d")
            if pkey == "1d":
                prev_date = get_prev_date(etf_key, today)
            else:
                prev_date = get_nearest_date(etf_key, target_str, not_after=today)
            # 이미 다른 기간에서 쓴 날짜면 중복 표시 방지
            if prev_date and prev_date in used_dates:
                prev_date = None
            if prev_date:
                used_dates.add(prev_date)
            prev = load_holdings(etf_key, prev_date) if prev_date else []
            diff = diff_holdings(prev, curr) if prev else None
            periods_data[pkey] = {
                "prev_date": prev_date,
                "prev": prev,
                "diff": diff,
            }
            if prev_date and diff:
                log.info(
                    f"  📊 {plabel} ({prev_date}): "
                    f"신규 {len(diff['new'])} | 청산 {len(diff['removed'])} | "
                    f"확대 {len(diff['increased'])} | 축소 {len(diff['decreased'])}"
                )
            else:
                log.info(f"  ℹ️  {plabel}: 비교 데이터 없음")

        results[etf_key] = {"curr": curr, "periods": periods_data}

    out = generate_html(results, today, etf_meta=etf_meta)
    print(f"\n🎉 완료! 브라우저에서 열기:")
    print(f"   open '{out}'")
    return out


def main():
    parser = argparse.ArgumentParser(description="코스닥 액티브 ETF 트래커")
    parser.add_argument("--date", type=str, help="조회 날짜 (YYYY-MM-DD, 기본값: 오늘)")
    parser.add_argument("--force", action="store_true", help="캐시 무시하고 강제 재수집")
    args = parser.parse_args()
    run(target_date=args.date, force_fetch=args.force)


if __name__ == "__main__":
    main()
