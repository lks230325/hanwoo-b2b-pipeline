#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[근본한우 B2B 인텔리전스] 축산물품질평가원(EKAPEPIA) 공식 한우 도매 지육시세 시계열 수집기 (2024.01.01 ~ 현재)
- 수집 대상: 한우전체(T), 한우거세(3), 한우암소(1)
- 수집 등급: 전국평균(avg), 1++등급(g1pp), 1+등급(g1p), 1등급(g1), 2등급(g2), 3등급(g3)
- 저장 경로: output/hanwoo_carcass_history_2024.json
"""

from __future__ import annotations

import json
import logging
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger("HanwooPipeline")

EKAPEPIA_PERIOD_URL = "https://www.ekapepia.com/v3/price/auction/period/auctionPrice/period.do"
HISTORY_JSON_PATH = Path(__file__).resolve().parent / "output" / "hanwoo_carcass_history_2024.json"


def _parse_int(txt: str) -> int:
    m = re.search(r"[\d,]+", txt or "")
    return int(m.group(0).replace(",", "")) if m else 0


def _iso_week_key(date_str: str) -> str:
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    iso_y, iso_w, _ = dt.isocalendar()
    return f"{iso_y}-W{iso_w:02d}"


def _fetch_single_chunk(
    year: int, start_date: str, end_date: str, sc2: str
) -> Tuple[str, Dict[str, Dict[str, int]]]:
    """단일 연도 구간 및 축종(T=전체, 3=거세, 1=암소) 일별 등급별 경락단가 수집"""
    ua_headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    url = (
        f"{EKAPEPIA_PERIOD_URL}?searchYn=Y&searchType=day"
        f"&startYYMMDD={start_date}&endYYMMDD={end_date}"
        f"&searchStartDate={start_date}&searchEndDate={end_date}"
        f"&searchCondition=0000&searchCondition1=1&searchCondition2={sc2}&searchCondition3=2"
    )
    out: Dict[str, Dict[str, int]] = {}
    try:
        resp = requests.get(url, headers=ua_headers, timeout=25, verify=False)
        resp.encoding = "utf-8"
        tbl = BeautifulSoup(resp.text, "html.parser").find("table")
        if not tbl:
            return sc2, out
        for tr in tbl.find_all("tr")[2:]:
            cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
            if len(cells) < 10:
                continue
            m = re.search(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일", cells[0] or "")
            if not m:
                continue
            d_str = f"{year}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
            g1pp = _parse_int(cells[1])
            g1p = _parse_int(cells[2])
            g1 = _parse_int(cells[3])
            g2 = _parse_int(cells[4])
            g3 = _parse_int(cells[5])
            avg_p = _parse_int(cells[9])
            if avg_p > 0 or g1pp > 0 or g1p > 0:
                out[d_str] = {
                    "avg": avg_p,
                    "g1pp": g1pp,
                    "g1p": g1p,
                    "g1": g1,
                    "g2": g2,
                    "g3": g3,
                }
    except Exception as exc:
        logger.warning("EKAPEPIA 한우 기간별 경락가 조회 오류 (%s~%s, sc2=%s): %s", start_date, end_date, sc2, exc)
    return sc2, out


def sync_hanwoo_carcass_history(force_full: bool = False) -> Dict[str, Any]:
    """2024-01-01 ~ 현재까지 한우 지육시세(전체/거세/암소 × 등급별) 누적 동기화"""
    HISTORY_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    existing_by_date: Dict[str, Dict[str, Any]] = {}
    if HISTORY_JSON_PATH.exists() and not force_full:
        try:
            prev = json.loads(HISTORY_JSON_PATH.read_text(encoding="utf-8"))
            for row in prev.get("daily_rows", []):
                if row.get("d"):
                    existing_by_date[row["d"]] = row
        except Exception:
            existing_by_date = {}

    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")

    # 기존 데이터가 충분히 있으면 최근 45일만 증분 조회, 없으면 2024-01-01부터 반기 단위 병렬 조회
    if len(existing_by_date) >= 600 and not force_full:
        cur_year = now.year
        chunks = [
            (cur_year, f"{cur_year}-01-01", today_str),
        ]
    else:
        chunks = [
            (2024, "2024-01-01", "2024-06-30"),
            (2024, "2024-07-01", "2024-12-31"),
            (2025, "2025-01-01", "2025-06-30"),
            (2025, "2025-07-01", "2025-12-31"),
            (2026, "2026-01-01", "2026-06-30"),
            (2026, "2026-07-01", today_str),
        ]

    tasks = []
    for yr, s_dt, e_dt in chunks:
        if s_dt > today_str:
            continue
        e_clamped = min(e_dt, today_str)
        for sc2 in ("T", "3", "1"):
            tasks.append((yr, s_dt, e_clamped, sc2))

    raw_by_sex: Dict[str, Dict[str, Dict[str, int]]] = {"T": {}, "3": {}, "1": {}}
    with ThreadPoolExecutor(max_workers=9) as ex:
        futs = [ex.submit(_fetch_single_chunk, yr, s_dt, e_dt, sc2) for yr, s_dt, e_dt, sc2 in tasks]
        for fut in as_completed(futs):
            sc2, res_map = fut.result()
            raw_by_sex[sc2].update(res_map)

    all_dates = sorted(set(raw_by_sex["T"].keys()) | set(raw_by_sex["3"].keys()) | set(raw_by_sex["1"].keys()) | set(existing_by_date.keys()))
    daily_rows: List[Dict[str, Any]] = []

    for d_str in all_dates:
        if d_str < "2024-01-01":
            continue
        t_obj = raw_by_sex["T"].get(d_str)
        s_obj = raw_by_sex["3"].get(d_str)
        c_obj = raw_by_sex["1"].get(d_str)

        if not t_obj and not s_obj and not c_obj and d_str in existing_by_date:
            daily_rows.append(existing_by_date[d_str])
            continue

        prev_row = existing_by_date.get(d_str, {})
        t_data = t_obj or prev_row.get("all") or {"avg": 0, "g1pp": 0, "g1p": 0, "g1": 0, "g2": 0, "g3": 0}
        s_data = s_obj or prev_row.get("steer") or t_data
        c_data = c_obj or prev_row.get("cow") or t_data

        if (t_data.get("avg", 0) <= 0) and (s_data.get("avg", 0) <= 0) and (c_data.get("avg", 0) <= 0):
            continue

        daily_rows.append({
            "d": d_str,
            "y": int(d_str[:4]),
            "m": d_str[:7],
            "w": _iso_week_key(d_str),
            "all": t_data,
            "steer": s_data,
            "cow": c_data,
        })

    daily_rows.sort(key=lambda x: x["d"])
    payload = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "min_date": daily_rows[0]["d"] if daily_rows else "2024-01-02",
        "max_date": daily_rows[-1]["d"] if daily_rows else today_str,
        "total_days": len(daily_rows),
        "daily_rows": daily_rows,
    }
    HISTORY_JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    logger.info(
        "EKAPEPIA 한우 도매 지육시세 시계열 동기화 완료: %s ~ %s (총 %d거래일)",
        payload["min_date"],
        payload["max_date"],
        payload["total_days"],
    )
    return payload


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    res = sync_hanwoo_carcass_history(force_full=True)
    print(f"Synced {res['total_days']} trading days ({res['min_date']} ~ {res['max_date']})")
    if res["daily_rows"]:
        print("Sample first day:", res["daily_rows"][0])
        print("Sample latest day:", res["daily_rows"][-1])
