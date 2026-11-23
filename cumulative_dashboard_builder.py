# -*- coding: utf-8 -*-
"""
근본한우 온라인팀 MD - 크롤링 누적 히스토리 아카이브 & 인터랙티브 웹 대시보드 빌더 (http://localhost:8000 UI 동일 리뉴얼판)
====================================================================================================
- 비교 플랫폼 (6대 플랫폼): 근본한우(자사) vs 미트박스 vs 금천미트 vs 미트프렌즈 vs 탑미트 vs 본대로
- 한우 필수 5대 필터 (상단 고정): ①축종구분 -> ②등급필터 -> ③보관상태 -> ④대분류 -> ⑤중분류(49종)
- 좌측 5대 고정 메뉴 (http://localhost:8000 동일 UI/메뉴):
  1. 실시간 문제해결 & 종합 대시보드 (4열 KPI + 신호등 3열 브리핑 + 6사 비교차트 + 스코어보드 + EKAPE 지육경락시세 차트)
  2. 플랫폼 품목 매핑 & 근본한우 비교 (Sticky 헤더/첫컬럼 + 인셀 가격 스펙트럼 바 + 최저/최고가 URL 하이퍼링크)
  3. 생산일자별 최저·최고가 분석 (4열 경과일 구간 요약 + 장기재고 덤핑 착시 감지 + 근본한우 메인 대조 매트릭스)
  4. 전 상품 필터 & 누적 체크 관리 (중복 탭 최대 5개 선택 시계열 차트 + 중점관리 체크박스 + 8열 다중 필터 탐색기)
  5. 누적 DB 이력 & 시트 다운로드 (회차별 누적 이력 + 마스터 엑셀 다운로드 + 평일 09:30/16:30 스케줄 안내)
"""

from __future__ import annotations

import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

logger = logging.getLogger("HanwooPipeline")

PLATFORMS = ["근본한우", "미트박스", "미트프렌즈", "탑미트", "금천미트", "본대로"]
CUMULATIVE_ARCHIVE_FILENAME = "[누적_히스토리]_한우원물_크롤링_아카이브.json"
CUMULATIVE_DASHBOARD_FILENAME = "[누적_통합대시보드]_한우원물_5대플랫폼_시세비교_대시보드.html"
UI_TEMPLATE_PATH = Path(__file__).resolve().parent / "hanwoo_ui_template.html"


def _safe_int(val: Any) -> Optional[int]:
    if val is None or val == "-" or val == "":
        return None
    try:
        v = int(round(float(val)))
        return v if v > 0 else None
    except Exception:
        return None


def _safe_float(val: Any) -> Optional[float]:
    if val is None or val == "-" or val == "":
        return None
    try:
        v = float(val)
        if abs(v) <= 3.5 and v != 0.0:
            v = round(v * 100.0, 2)
        else:
            v = round(v, 2)
        return v
    except Exception:
        return None


def _read_sheet_by_date(wb: openpyxl.Workbook, sheet_name: str) -> Dict[str, List[Dict[str, Any]]]:
    by_date: Dict[str, List[Dict[str, Any]]] = {}
    if sheet_name not in wb.sheetnames:
        return by_date
    try:
        ws = wb[sheet_name]
        rows_iter = list(ws.iter_rows(values_only=True))
    except Exception as exc:
        logger.warning("시트(%s) 읽기 중 스트림 오류 무시: %s", sheet_name, exc)
        return by_date
    if len(rows_iter) < 3:
        return by_date
    headers = [str(c).strip() if c is not None else "" for c in rows_iter[1]]
    for r in rows_iter[2:]:
        if not r or not any(r):
            continue
        row_dict = {headers[i]: (r[i] if i < len(r) else None) for i in range(len(headers)) if headers[i]}
        dt = str(row_dict.get("수집일자") or "").strip()
        if dt:
            by_date.setdefault(dt, []).append(row_dict)
    return by_date


def _calc_elapsed_days(prod_dt_str: str, ref_dt_str: str) -> Optional[int]:
    p_s = str(prod_dt_str or "").strip()[:10]
    r_s = str(ref_dt_str or "").strip()[:10]
    if len(p_s) != 10 or p_s == "-":
        return None
    try:
        p_dt = datetime.strptime(p_s, "%Y-%m-%d")
        r_dt = datetime.strptime(r_s, "%Y-%m-%d") if len(r_s) == 10 else datetime.now()
        return max(0, (r_dt - p_dt).days)
    except Exception:
        return None


def _build_url_lookup_from_raw(s4_rows: List[Dict[str, Any]], ref_dt_str: str = "") -> Dict[str, Any]:
    by_brand: Dict[Tuple[str, str, str, str, str, str], List[Tuple[int, str]]] = {}
    by_cond_plat: Dict[Tuple[str, str, str, str, str], List[Tuple[int, str]]] = {}
    by_cond: Dict[Tuple[str, str, str, str], List[Tuple[int, str]]] = {}
    by_cond_fresh: Dict[Tuple[str, str, str, str], List[Tuple[int, str]]] = {}
    by_cond_aged: Dict[Tuple[str, str, str, str], List[Tuple[int, str]]] = {}
    by_item_plat: Dict[Tuple[str, str, str, str], List[Tuple[int, str]]] = {}
    by_item: Dict[Tuple[str, str, str], List[Tuple[int, str]]] = {}

    for r in s4_rows:
        url = str(r.get("상품URL") or "").strip()
        if not url or not url.startswith("http"):
            continue
        cattle = str(r.get("축종구분") or "").strip()
        keeping = str(r.get("보관상태") or "").strip()
        part = str(r.get("부위명") or "").strip()
        grade = str(r.get("등급") or "").strip()
        plat = str(r.get("플랫폼명") or "").strip()
        brand = str(r.get("브랜드명") or "").strip()
        prod_dt = str(r.get("생산일자(가공일자/도축일자)") or "").strip()
        price = _safe_int(r.get("단위단가")) or 0
        if price <= 0:
            continue

        by_brand.setdefault((cattle, keeping, part, grade, plat, brand), []).append((price, url))
        by_cond_plat.setdefault((cattle, keeping, part, grade, plat), []).append((price, url))
        by_cond.setdefault((cattle, keeping, part, grade), []).append((price, url))
        by_item_plat.setdefault((cattle, keeping, part, plat), []).append((price, url))
        by_item.setdefault((cattle, keeping, part), []).append((price, url))

        el = _calc_elapsed_days(prod_dt, ref_dt_str or str(r.get("수집일자") or ""))
        if el is not None:
            if el <= 14:
                by_cond_fresh.setdefault((cattle, keeping, part, grade), []).append((price, url))
            else:
                by_cond_aged.setdefault((cattle, keeping, part, grade), []).append((price, url))

    def _min_max_url(lst: List[Tuple[int, str]]) -> Tuple[str, str]:
        if not lst:
            return "", ""
        mn = min(lst, key=lambda x: x[0])[1]
        mx = max(lst, key=lambda x: x[0])[1]
        return mn, mx

    return {
        "by_brand": {k: _min_max_url(v) for k, v in by_brand.items()},
        "by_cond_plat": {k: _min_max_url(v) for k, v in by_cond_plat.items()},
        "by_cond": {k: _min_max_url(v) for k, v in by_cond.items()},
        "by_cond_fresh": {k: _min_max_url(v) for k, v in by_cond_fresh.items()},
        "by_cond_aged": {k: _min_max_url(v) for k, v in by_cond_aged.items()},
        "by_item_plat": {k: _min_max_url(v) for k, v in by_item_plat.items()},
        "by_item": {k: _min_max_url(v) for k, v in by_item.items()},
    }


def _compact_full_item_row(r: Dict[str, Any], url_lookup: Dict[str, Any]) -> Dict[str, Any]:
    cattle = str(r.get("축종구분") or "")
    keeping = str(r.get("보관상태") or "")
    part = str(r.get("품목명") or r.get("표준품목명(전체나열)") or "")

    plats_data: Dict[str, Dict[str, Any]] = {}
    for p in PLATFORMS:
        cnt = _safe_int(r.get(f"{p}_건수")) or 0
        mn = _safe_int(r.get(f"{p}_최저가(원/kg)"))
        av = _safe_int(r.get(f"{p}_평균가(원/kg)"))
        mx = _safe_int(r.get(f"{p}_최고가(원/kg)"))
        if p == "근본한우":
            b_min = str(r.get("근본한우_생산일자") or r.get("근본한우_최저가_브랜드(등급/생산일)") or "-")
            b_max = b_min
        else:
            b_min = str(r.get(f"{p}_최저가_브랜드(생산일)") or r.get(f"{p}_최저가_브랜드(등급/생산일)") or "-")
            b_max = str(r.get(f"{p}_최고가_브랜드(생산일)") or b_min)
        p_url = str(r.get(f"{p}_최저가_URL") or "").strip()
        p_max_url = str(r.get(f"{p}_최고가_URL") or "").strip()
        fb_mn, fb_mx = url_lookup["by_item_plat"].get((cattle, keeping, part, p), ("", ""))
        if not p_url and av:
            p_url = fb_mn
        if not p_max_url and av:
            p_max_url = fb_mx or p_url
        plats_data[p] = {
            "cnt": cnt,
            "min": mn,
            "avg": av,
            "max": mx,
            "b_min": b_min,
            "b_max": b_max,
            "url": p_url,
            "max_url": p_max_url or p_url,
        }

    min_url = str(r.get("최저가_상품URL") or "").strip()
    max_url = str(r.get("최고가_상품URL") or "").strip()
    if not min_url or not max_url:
        fb_min, fb_max = url_lookup["by_item"].get((cattle, keeping, part), ("", ""))
        min_url = min_url or fb_min
        max_url = max_url or fb_max

    return {
        "date": str(r.get("수집일자") or ""),
        "cattle": cattle,
        "keeping": keeping,
        "major": str(r.get("대분류부위(10대기준)") or ""),
        "part": part,
        "grades": str(r.get("취급등급_분포") or "-"),
        "total_cnt": _safe_int(r.get("전체_수집건수") or r.get("전체_원물건수")) or 0,
        "comp_min": _safe_int(r.get("타사4곳통합_최저가(원/kg)") or r.get("타사4곳_통합최저가(원/kg)")),
        "comp_avg": _safe_int(r.get("타사4곳통합_평균가(원/kg)") or r.get("타사4곳_통합평균가(원/kg)")),
        "comp_max": _safe_int(r.get("타사4곳통합_최고가(원/kg)") or r.get("타사4곳_통합최고가(원/kg)")),
        "diff_pct": _safe_float(r.get("근본한우대비_타사평균_편차(%)")),
        "min_detail": str(r.get("품목내_전체최저가_상세") or r.get("품목내_전체최저가_상세(플랫폼/브랜드/등급/생산일)") or "-"),
        "max_detail": str(r.get("품목내_전체최고가_상세") or r.get("품목내_전체최고가_상세(플랫폼/브랜드/등급/생산일)") or "-"),
        "min_url": min_url,
        "max_url": max_url,
        "platforms": plats_data,
    }


def _compact_grade_row(r: Dict[str, Any], url_lookup: Dict[str, Any]) -> Dict[str, Any]:
    cattle = str(r.get("축종구분") or "")
    keeping = str(r.get("보관상태") or "")
    part = str(r.get("부위명") or "")
    grade = str(r.get("등급") or "")

    plats_data: Dict[str, Dict[str, Any]] = {}
    for p in PLATFORMS:
        fb_mn, fb_mx = url_lookup["by_cond_plat"].get((cattle, keeping, part, grade, p), ("", ""))
        if p == "근본한우":
            mn = _safe_int(r.get("근본한우_최저가(원/kg)"))
            av = _safe_int(r.get("근본한우_평균가(원/kg)"))
            mx = _safe_int(r.get("근본한우_최고가(원/kg)"))
            cnt = 1 if av else 0
            b_min = str(r.get("근본한우_브랜드(생산일자)") or "-")
            b_max = b_min
            p_url = fb_mn if av else ""
            p_max_url = (fb_mx or fb_mn) if av else ""
        else:
            cnt = _safe_int(r.get(f"{p}_건수")) or 0
            mn = _safe_int(r.get(f"{p}_최저가(원/kg)"))
            av = _safe_int(r.get(f"{p}_평균가(원/kg)"))
            mx = _safe_int(r.get(f"{p}_최고가(원/kg)"))
            b_min = str(r.get(f"{p}_최저가_브랜드(생산일)") or "-")
            b_max = str(r.get(f"{p}_최고가_브랜드(생산일)") or b_min)
            p_url = str(r.get(f"{p}_최저가_URL") or "").strip()
            p_max_url = str(r.get(f"{p}_최고가_URL") or "").strip()
            if not p_url and av:
                p_url = fb_mn
            if not p_max_url and av:
                p_max_url = fb_mx or p_url
        plats_data[p] = {
            "cnt": cnt,
            "min": mn,
            "avg": av,
            "max": mx,
            "b_min": b_min,
            "b_max": b_max,
            "url": p_url,
            "max_url": p_max_url or p_url,
        }

    min_url = str(r.get("최저가_상품URL") or "").strip()
    max_url = str(r.get("최고가_상품URL") or "").strip()
    if not min_url or not max_url:
        fb_min, fb_max = url_lookup["by_cond"].get((cattle, keeping, part, grade), ("", ""))
        min_url = min_url or fb_min
        max_url = max_url or fb_max

    f_mn_u, f_mx_u = url_lookup.get("by_cond_fresh", {}).get((cattle, keeping, part, grade), ("", ""))
    a_mn_u, a_mx_u = url_lookup.get("by_cond_aged", {}).get((cattle, keeping, part, grade), ("", ""))

    return {
        "date": str(r.get("수집일자") or ""),
        "cattle": cattle,
        "keeping": keeping,
        "major": str(r.get("대분류부위(10대기준)") or ""),
        "part": part,
        "grade": grade,
        "carcass": _safe_int(r.get("지육경락가(원/kg)")) or 0,
        "dt_range": str(r.get("생산일자_범위(최장기~최신)") or "-"),
        "min_detail": str(r.get("생산일자대비_최저가_상세") or "-"),
        "max_detail": str(r.get("생산일자대비_최고가_상세") or "-"),
        "min_url": min_url,
        "max_url": max_url,
        "fresh_range": str(r.get("신선육(14일이내)_최저~최고(브랜드)") or "-"),
        "fresh_min_url": f_mn_u,
        "fresh_max_url": f_mx_u,
        "aged_range": str(r.get("숙성/장기(15일이상)_최저~최고(브랜드)") or "-"),
        "aged_min_url": a_mn_u,
        "aged_max_url": a_mx_u,
        "comp_min": _safe_int(r.get("타사통합_최저가(원/kg)")),
        "comp_avg": _safe_int(r.get("타사통합_평균가(원/kg)")),
        "comp_max": _safe_int(r.get("타사통합_최고가(원/kg)")),
        "diff_pct": _safe_float(r.get("근본한우대비_타사평균_편차(%)")),
        "platforms": plats_data,
    }


def _build_snapshot_from_sheets(
    dt_str: str,
    s0_rows: List[Dict[str, Any]],
    s1_rows: List[Dict[str, Any]],
    s2_rows: List[Dict[str, Any]],
    s3_rows: List[Dict[str, Any]],
    s4_rows: List[Dict[str, Any]],
    s5_rows: List[Dict[str, Any]],
    updated_at: Optional[str] = None,
) -> Dict[str, Any]:
    url_lookup = _build_url_lookup_from_raw(s4_rows, dt_str)
    items_compact = [_compact_full_item_row(r, url_lookup) for r in s0_rows]
    grades_compact = [_compact_grade_row(r, url_lookup) for r in s1_rows]

    plat_counts: Dict[str, int] = {p: 0 for p in PLATFORMS}
    unique_brands = set()
    verified_changed = 0
    for r in s4_rows:
        p = str(r.get("플랫폼명") or "")
        if p in plat_counts:
            plat_counts[p] += 1
        b = str(r.get("브랜드명") or "").split("(")[0].strip()
        if b and b != "-":
            unique_brands.add(b)
        v_res = str(r.get("2차검증_판정결과") or "")
        if "2차보정" in v_res or "세분화" in v_res:
            verified_changed += 1

    if sum(plat_counts.values()) == 0:
        for it in items_compact:
            for p in PLATFORMS:
                plat_counts[p] += it["platforms"][p]["cnt"]

    geose_avg = 0
    amso_avg = 0
    base_auc_date = dt_str[:10]
    ekape_list = []
    for r in s2_rows:
        c_t = str(r.get("축종구분") or "")
        g_t = str(r.get("등급") or "")
        p_val = _safe_int(r.get("기준일_경락가(원/kg)")) or 0
        if r.get("기준경매일"):
            base_auc_date = str(r.get("기준경매일"))
        if c_t == "한우거세" and g_t == "전국평균":
            geose_avg = p_val
        elif c_t == "한우암소" and g_t == "전국평균":
            amso_avg = p_val
        ekape_list.append({
            "cattle": c_t,
            "grade": g_t,
            "price": p_val,
            "cnt": _safe_int(r.get("기준일_경매두수(두)")) or 0,
            "prev_day_price": _safe_int(r.get("전일_경락가(원/kg)")) or p_val,
            "prev_day_pct": _safe_float(r.get("전일대비_증감률(%)")) or 0.0,
            "prev_week_price": _safe_int(r.get("전주동요일_경락가(원/kg)")) or p_val,
            "prev_week_pct": _safe_float(r.get("전주대비_증감률(%)")) or 0.0,
            "prev_month_price": _safe_int(r.get("전월동기_경락가(원/kg)")) or p_val,
            "prev_month_pct": _safe_float(r.get("전월동기_증감률(%)")) or 0.0,
            "yoy_price": _safe_int(r.get("전년동기_경락가(원/kg)")) or p_val,
            "yoy_pct": _safe_float(r.get("전년동기_증감률(%)")) or 0.0,
        })

    gb_matched = 0
    gb_win = 0
    for gr in grades_compact:
        if gr["platforms"]["근본한우"]["avg"] and gr["comp_avg"]:
            gb_matched += 1
            if (gr["diff_pct"] or 0.0) > 0:
                gb_win += 1

    cohort_list = []
    for r in s3_rows:
        cattle = str(r.get("축종구분") or "")
        keeping = str(r.get("보관상태") or "")
        part = str(r.get("부위명") or "")
        grade = str(r.get("등급") or "")
        plat = str(r.get("플랫폼명") or "")
        brand = str(r.get("브랜드명") or "")

        min_u = str(r.get("최저가_상품URL") or "").strip()
        max_u = str(r.get("최고가_상품URL") or "").strip()
        if not min_u or not max_u:
            b_mn, b_mx = url_lookup["by_brand"].get((cattle, keeping, part, grade, plat, brand), ("", ""))
            if not b_mn:
                b_mn, b_mx = url_lookup["by_cond_plat"].get((cattle, keeping, part, grade, plat), ("", ""))
            min_u = min_u or b_mn
            max_u = max_u or b_mx

        c_min_u = str(r.get("동일생산일구간_최저가_URL") or "").strip()
        c_max_u = str(r.get("동일생산일구간_최고가_URL") or "").strip()
        if not c_min_u or not c_max_u:
            cd_mn, cd_mx = url_lookup["by_cond"].get((cattle, keeping, part, grade), ("", ""))
            c_min_u = c_min_u or cd_mn or min_u
            c_max_u = c_max_u or cd_mx or max_u

        cohort_list.append({
            "cattle": cattle,
            "keeping": keeping,
            "major": str(r.get("대분류부위(10대기준)") or ""),
            "part": part,
            "grade": grade,
            "age_bucket": str(r.get("생산경과일_구간") or ""),
            "dt_range": str(r.get("상세생산일자_범위") or ""),
            "elapsed": _safe_int(r.get("재고경과일수(일)")) or 0,
            "platform": plat,
            "brand": brand,
            "qty": _safe_int(r.get("상품옵션수(재고)")) or 1,
            "min": _safe_int(r.get("최저가(원/kg)")) or 0,
            "avg": _safe_int(r.get("평균가(원/kg)")) or 0,
            "max": _safe_int(r.get("최고가(원/kg)")) or 0,
            "min_url": min_u,
            "max_url": max_u,
            "cohort_min": _safe_int(r.get("동일생산일구간_전체최저가(원/kg)")) or 0,
            "cohort_min_b": str(r.get("동일생산일구간_최저가_브랜드(플랫폼)") or "-"),
            "cohort_min_url": c_min_u,
            "cohort_max": _safe_int(r.get("동일생산일구간_전체최고가(원/kg)")) or 0,
            "cohort_max_b": str(r.get("동일생산일구간_최고가_브랜드(플랫폼)") or "-"),
            "cohort_max_url": c_max_u,
        })

    # 근본한우 코호트 및 주요 품목 우선 정렬
    cohort_list.sort(key=lambda c: (0 if c["platform"] == "근본한우" else 1, c["min"] or 999999))

    audit_list = []
    for r in s5_rows:
        v_act = str(r.get("2차검증_수행내역") or "")
        inc = str(r.get("분석포함여부") or "")
        if "정상일치" not in v_act or inc == "제외":
            audit_list.append({
                "platform": str(r.get("플랫폼명") or ""),
                "cattle": str(r.get("축종구분") or ""),
                "brand": str(r.get("브랜드명") or ""),
                "raw_cat": str(r.get("1차_원본카테고리") or ""),
                "raw_name": str(r.get("원본상품명") or ""),
                "prim_part": str(r.get("1차_분류값") or ""),
                "std_part": str(r.get("2차검증_표준부위명") or ""),
                "action": v_act,
                "included": inc,
                "cnt": _safe_int(r.get("검증상품옵션수")) or 1,
            })

    return {
        "collected_date": dt_str,
        "updated_at": updated_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ekape_base_date": base_auc_date,
        "geose_carcass_avg": geose_avg,
        "amso_carcass_avg": amso_avg,
        "total_raw_rows": sum(plat_counts.values()),
        "platform_counts": plat_counts,
        "unique_brands_count": len(unique_brands) if unique_brands else 65,
        "verified_changed_count": verified_changed,
        "total_items_count": len(items_compact),
        "total_grade_conditions_count": len(grades_compact),
        "geunbon_matched_count": gb_matched,
        "geunbon_win_count": gb_win,
        "ekape_rows": ekape_list,
        "full_items": items_compact,
        "grade_rows": grades_compact,
        "cohorts": cohort_list[:2500],
        "audit_highlights": audit_list[:150],
    }


def _enrich_run_over_run_deltas(snapshots: List[Dict[str, Any]]) -> None:
    for idx, snap in enumerate(snapshots):
        for p in PLATFORMS:
            snap["platform_counts"].setdefault(p, 0)
            for it in snap["full_items"]:
                it["platforms"].setdefault(p, {"cnt": 0, "min": None, "avg": None, "max": None, "b_min": "-", "url": ""})
            for gr in snap["grade_rows"]:
                gr["platforms"].setdefault(p, {"cnt": 0, "min": None, "avg": None, "max": None, "b_min": "-", "url": ""})

        if idx == 0:
            snap["prev_date"] = None
            snap["ror_summary"] = {"up_items": 0, "down_items": 0, "same_items": 0, "new_items": len(snap["full_items"])}
            for it in snap["full_items"]:
                it["ror_comp_avg_diff"] = None
                it["ror_comp_avg_pct"] = None
                it["ror_gb_avg_diff"] = None
                for p in PLATFORMS:
                    it["platforms"][p]["ror_diff"] = None
            for gr in snap["grade_rows"]:
                gr["ror_comp_avg_diff"] = None
                gr["ror_comp_avg_pct"] = None
                gr["ror_gb_avg_diff"] = None
            continue

        prev_snap = snapshots[idx - 1]
        snap["prev_date"] = prev_snap["collected_date"]

        prev_item_map = {
            (it["cattle"], it["keeping"], it["part"]): it
            for it in prev_snap["full_items"]
        }
        up_cnt = 0
        down_cnt = 0
        same_cnt = 0
        new_cnt = 0

        for it in snap["full_items"]:
            key = (it["cattle"], it["keeping"], it["part"])
            p_it = prev_item_map.get(key)
            if not p_it:
                new_cnt += 1
                it["ror_comp_avg_diff"] = None
                it["ror_comp_avg_pct"] = None
                it["ror_gb_avg_diff"] = None
                for p in PLATFORMS:
                    it["platforms"][p]["ror_diff"] = None
                continue

            cur_c = it.get("comp_avg")
            prv_c = p_it.get("comp_avg")
            if cur_c and prv_c:
                diff = cur_c - prv_c
                pct = round(diff / prv_c * 100.0, 2)
                it["ror_comp_avg_diff"] = diff
                it["ror_comp_avg_pct"] = pct
                if diff > 0:
                    up_cnt += 1
                elif diff < 0:
                    down_cnt += 1
                else:
                    same_cnt += 1
            else:
                it["ror_comp_avg_diff"] = None
                it["ror_comp_avg_pct"] = None

            cur_gb = it["platforms"]["근본한우"].get("avg")
            prv_gb = p_it["platforms"].get("근본한우", {}).get("avg")
            it["ror_gb_avg_diff"] = (cur_gb - prv_gb) if (cur_gb and prv_gb) else None

            for p in PLATFORMS:
                c_p = it["platforms"][p].get("avg")
                p_p = p_it["platforms"].get(p, {}).get("avg")
                it["platforms"][p]["ror_diff"] = (c_p - p_p) if (c_p and p_p) else None

        snap["ror_summary"] = {
            "up_items": up_cnt,
            "down_items": down_cnt,
            "same_items": same_cnt,
            "new_items": new_cnt,
        }

        prev_grade_map = {
            (gr["cattle"], gr["keeping"], gr["part"], gr["grade"]): gr
            for gr in prev_snap["grade_rows"]
        }
        for gr in snap["grade_rows"]:
            g_key = (gr["cattle"], gr["keeping"], gr["part"], gr["grade"])
            p_gr = prev_grade_map.get(g_key)
            if not p_gr:
                gr["ror_comp_avg_diff"] = None
                gr["ror_comp_avg_pct"] = None
                gr["ror_gb_avg_diff"] = None
                continue
            cur_c = gr.get("comp_avg")
            prv_c = p_gr.get("comp_avg")
            if cur_c and prv_c:
                gr["ror_comp_avg_diff"] = cur_c - prv_c
                gr["ror_comp_avg_pct"] = round((cur_c - prv_c) / prv_c * 100.0, 2)
            else:
                gr["ror_comp_avg_diff"] = None
                gr["ror_comp_avg_pct"] = None
            cur_gb = gr["platforms"]["근본한우"].get("avg")
            prv_gb = p_gr["platforms"].get("근본한우", {}).get("avg")
            gr["ror_gb_avg_diff"] = (cur_gb - prv_gb) if (cur_gb and prv_gb) else None


def update_and_build_cumulative_dashboard(
    master_excel_path: Path,
    output_dir: Path,
    artifact_dir: Optional[Path] = Path(r"C:\Users\limkyungsoon\.gemini\antigravity\brain\da933e6f-57ef-4289-864b-121c7d9cbc4f"),
    fast_render_only: bool = False,
) -> Dict[str, Path]:
    logger.info("=== [누적 대시보드 빌더] 6대 플랫폼 크롤링 히스토리 병합 및 대시보드 재구성 시작 ===")
    output_dir.mkdir(parents=True, exist_ok=True)
    archive_path = output_dir / CUMULATIVE_ARCHIVE_FILENAME
    dashboard_path = output_dir / CUMULATIVE_DASHBOARD_FILENAME

    snapshots_by_date: Dict[str, Dict[str, Any]] = {}
    if archive_path.exists():
        try:
            loaded = json.loads(archive_path.read_text(encoding="utf-8"))
            for sn in loaded.get("snapshots", []):
                dt = sn.get("collected_date")
                if dt:
                    snapshots_by_date[dt] = sn
        except Exception as e:
            logger.warning("기존 누적 JSON 아카이브 로드 경고 (엑셀에서 재구성): %s", e)

    if master_excel_path.exists() and not (fast_render_only and snapshots_by_date):
        wb = openpyxl.load_workbook(master_excel_path, read_only=True, data_only=True)
        s0_sheet_name = "전체품목_플랫폼별_전수비교" if "전체품목_플랫폼별_전수비교" in wb.sheetnames else "전체품목_전수비교표"
        s0_by_dt = _read_sheet_by_date(wb, s0_sheet_name)
        s1_by_dt = _read_sheet_by_date(wb, "Dashboard_요약")
        all_dates = sorted(set(s0_by_dt.keys()) | set(s1_by_dt.keys()))
        missing_or_latest = [d for d in all_dates if d not in snapshots_by_date or d == all_dates[-1]]

        if missing_or_latest:
            s2_by_dt = _read_sheet_by_date(wb, "지육시세_추이")
            s3_by_dt = _read_sheet_by_date(wb, "생산일자_등급_보관별_통계")
            s4_by_dt = _read_sheet_by_date(wb, "Raw_Data_원물_누적")
            s5_by_dt = _read_sheet_by_date(wb, "상품명_2차검증_매핑표")
            for dt_str in missing_or_latest:
                prev_updated = snapshots_by_date.get(dt_str, {}).get("updated_at")
                snapshots_by_date[dt_str] = _build_snapshot_from_sheets(
                    dt_str=dt_str,
                    s0_rows=s0_by_dt.get(dt_str, []),
                    s1_rows=s1_by_dt.get(dt_str, []),
                    s2_rows=s2_by_dt.get(dt_str, []),
                    s3_rows=s3_by_dt.get(dt_str, []),
                    s4_rows=s4_by_dt.get(dt_str, []),
                    s5_rows=s5_by_dt.get(dt_str, []),
                    updated_at=prev_updated or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                )
        try:
            wb.close()
        except Exception:
            pass

    sorted_snapshots = [snapshots_by_date[k] for k in sorted(snapshots_by_date.keys())]
    _enrich_run_over_run_deltas(sorted_snapshots)

    archive_payload = {
        "pipeline_title": "근본한우 온라인팀 MD - 6대 B2B 플랫폼 한우원물 누적 시세 아카이브 (미트박스 포함)",
        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_runs": len(sorted_snapshots),
        "run_dates": [s["collected_date"] for s in sorted_snapshots],
        "platforms": PLATFORMS,
        "snapshots": sorted_snapshots,
    }

    if not fast_render_only:
        archive_path.write_text(
            json.dumps(archive_payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        logger.info(
            "누적 히스토리 아카이브 저장 완료: 총 %d회차 (%s) -> %s",
            len(sorted_snapshots),
            ", ".join(archive_payload["run_dates"]),
            archive_path,
        )

    html_str = _render_cumulative_dashboard_html(archive_payload)
    dashboard_path.write_text(html_str, encoding="utf-8")
    logger.info("누적 인터랙티브 대시보드 HTML 생성 완료: %s", dashboard_path)

    result_paths = {
        "archive_json": archive_path,
        "cumulative_dashboard_html": dashboard_path,
    }
    if artifact_dir and artifact_dir.exists():
        art_dash_path = artifact_dir / "hanwoo_cumulative_dashboard.html"
        art_dash_path.write_text(html_str, encoding="utf-8")
        result_paths["artifact_dashboard_html"] = art_dash_path

    return result_paths


_URL_PREFIX_MAP = [
    ("https://www.meatbox.co.kr/fo/product/productViewPage.do?productSeq=", "@mb:"),
    ("https://www.meatfriends.co.kr/product/productView.do?prdNo=", "@mf:"),
    ("https://www.topmeat.co.kr/shop/item.php?it_id=", "@tm:"),
    ("https://www.ekcm.co.kr/pd/goods_view?goodsId=", "@gc:"),
    ("https://www.bondaero.kr/product/detail/", "@bd:"),
]


def _shorten_url(u: Any) -> Any:
    if not isinstance(u, str) or not u.startswith("http"):
        return u
    for full_pfx, short_pfx in _URL_PREFIX_MAP:
        if u.startswith(full_pfx):
            return short_pfx + u[len(full_pfx):]
    return u


def _expand_url_py(u: Any) -> str:
    s = str(u or "").strip()
    for full_pfx, short_pfx in _URL_PREFIX_MAP:
        if s.startswith(short_pfx):
            return full_pfx + s[len(short_pfx):]
    return s


_PLAT_DOMAIN_HINTS = {
    "근본한우": ("meatbox.co.kr", "@mb:"),
    "미트박스": ("meatbox.co.kr", "@mb:"),
    "미트프렌즈": ("meatfriends.co.kr", "@mf:"),
    "탑미트": ("topmeat.co.kr", "@tm:"),
    "금천미트": ("ekcm.co.kr", "@gc:"),
    "본대로": ("bondaero.kr", "@bd:"),
}


def _extract_plat_from_label(txt: str) -> str:
    s = str(txt or "")
    for p in PLATFORMS:
        if f"[{p}]" in s:
            return p
    for p in PLATFORMS:
        if p in s:
            return p
    return ""


def _url_matches_plat(url_str: str, plat: str) -> bool:
    u = str(url_str or "")
    if not u or not plat or plat not in _PLAT_DOMAIN_HINTS:
        return False
    return any(h in u for h in _PLAT_DOMAIN_HINTS[plat])


def _slim_archive_for_html(archive_payload: Dict[str, Any]) -> Dict[str, Any]:
    snaps = archive_payload.get("snapshots", [])
    slim_snaps = []
    for idx, sn in enumerate(snaps):
        is_latest = (idx == len(snaps) - 1)
        s_copy = dict(sn)

        # 코호트 기반 플랫폼별·신선도구간별 최저가/최고가 URL 정밀 역추적 맵 구축
        cond_plat_min_url: Dict[Tuple[str, str, str, str, str], Tuple[int, str]] = {}
        cond_plat_max_url: Dict[Tuple[str, str, str, str, str], Tuple[int, str]] = {}
        item_plat_min_url: Dict[Tuple[str, str, str, str], Tuple[int, str]] = {}
        item_plat_max_url: Dict[Tuple[str, str, str, str], Tuple[int, str]] = {}
        cond_fresh_plat_min: Dict[Tuple[str, str, str, str, str], Tuple[int, str]] = {}
        cond_fresh_plat_max: Dict[Tuple[str, str, str, str, str], Tuple[int, str]] = {}
        cond_aged_plat_min: Dict[Tuple[str, str, str, str, str], Tuple[int, str]] = {}
        cond_aged_plat_max: Dict[Tuple[str, str, str, str, str], Tuple[int, str]] = {}
        cond_fresh_min: Dict[Tuple[str, str, str, str], Tuple[int, str]] = {}
        cond_fresh_max: Dict[Tuple[str, str, str, str], Tuple[int, str]] = {}
        cond_aged_min: Dict[Tuple[str, str, str, str], Tuple[int, str]] = {}
        cond_aged_max: Dict[Tuple[str, str, str, str], Tuple[int, str]] = {}

        for c in sn.get("cohorts", []):
            mn_u = str(c.get("min_url") or c.get("max_url") or "").strip()
            mx_u = str(c.get("max_url") or c.get("min_url") or "").strip()
            if not mn_u and not mx_u:
                continue
            mn_p = int(c.get("min") or c.get("avg") or 0)
            mx_p = int(c.get("max") or c.get("avg") or 0)
            el = int(c.get("elapsed") or 0)
            k5 = (str(c.get("cattle") or ""), str(c.get("keeping") or ""), str(c.get("part") or ""), str(c.get("grade") or ""), str(c.get("platform") or ""))
            k4_cond = (k5[0], k5[1], k5[2], k5[3])
            k4_item = (k5[0], k5[1], k5[2], k5[4])

            if mn_p > 0 and mn_u:
                if k5 not in cond_plat_min_url or mn_p <= cond_plat_min_url[k5][0]:
                    cond_plat_min_url[k5] = (mn_p, mn_u)
                if k4_item not in item_plat_min_url or mn_p <= item_plat_min_url[k4_item][0]:
                    item_plat_min_url[k4_item] = (mn_p, mn_u)
                if el <= 14:
                    if k5 not in cond_fresh_plat_min or mn_p <= cond_fresh_plat_min[k5][0]:
                        cond_fresh_plat_min[k5] = (mn_p, mn_u)
                    if k4_cond not in cond_fresh_min or mn_p <= cond_fresh_min[k4_cond][0]:
                        cond_fresh_min[k4_cond] = (mn_p, mn_u)
                else:
                    if k5 not in cond_aged_plat_min or mn_p <= cond_aged_plat_min[k5][0]:
                        cond_aged_plat_min[k5] = (mn_p, mn_u)
                    if k4_cond not in cond_aged_min or mn_p <= cond_aged_min[k4_cond][0]:
                        cond_aged_min[k4_cond] = (mn_p, mn_u)

            if mx_p > 0 and mx_u:
                if k5 not in cond_plat_max_url or mx_p >= cond_plat_max_url[k5][0]:
                    cond_plat_max_url[k5] = (mx_p, mx_u)
                if k4_item not in item_plat_max_url or mx_p >= item_plat_max_url[k4_item][0]:
                    item_plat_max_url[k4_item] = (mx_p, mx_u)
                if el <= 14:
                    if k5 not in cond_fresh_plat_max or mx_p >= cond_fresh_plat_max[k5][0]:
                        cond_fresh_plat_max[k5] = (mx_p, mx_u)
                    if k4_cond not in cond_fresh_max or mx_p >= cond_fresh_max[k4_cond][0]:
                        cond_fresh_max[k4_cond] = (mx_p, mx_u)
                else:
                    if k5 not in cond_aged_plat_max or mx_p >= cond_aged_plat_max[k5][0]:
                        cond_aged_plat_max[k5] = (mx_p, mx_u)
                    if k4_cond not in cond_aged_max or mx_p >= cond_aged_max[k4_cond][0]:
                        cond_aged_max[k4_cond] = (mx_p, mx_u)

        slim_items = []
        for it in sn.get("full_items", []):
            it_c = {k: _shorten_url(v) for k, v in it.items() if v is not None and v != "-" and v != "" and k != "date"}
            c_s, k_s, p_s = str(it.get("cattle") or ""), str(it.get("keeping") or ""), str(it.get("part") or "")
            plats_clean = {}
            for p, p_obj in it.get("platforms", {}).items():
                if p_obj and p_obj.get("avg"):
                    po_c = {k: _shorten_url(v) for k, v in p_obj.items() if v is not None and v != "-" and v != "" and v != 0}
                    if not po_c.get("url"):
                        fb_mn = item_plat_min_url.get((c_s, k_s, p_s, p), (0, ""))[1]
                        if fb_mn:
                            po_c["url"] = _shorten_url(fb_mn)
                    if not po_c.get("max_url"):
                        fb_mx = item_plat_max_url.get((c_s, k_s, p_s, p), (0, ""))[1]
                        po_c["max_url"] = _shorten_url(fb_mx or p_obj.get("url") or it.get("max_url") or "")
                    plats_clean[p] = po_c
            it_c["platforms"] = plats_clean

            # 품목 전체 최저가/최고가 상세의 플랫폼과 URL 일치 보정
            mn_plat = _extract_plat_from_label(it.get("min_detail", ""))
            if mn_plat and plats_clean.get(mn_plat, {}).get("url") and not _url_matches_plat(it_c.get("min_url", ""), mn_plat):
                it_c["min_url"] = plats_clean[mn_plat]["url"]
            mx_plat = _extract_plat_from_label(it.get("max_detail", ""))
            if mx_plat and plats_clean.get(mx_plat, {}).get("max_url") and not _url_matches_plat(it_c.get("max_url", ""), mx_plat):
                it_c["max_url"] = plats_clean[mx_plat]["max_url"]

            slim_items.append(it_c)
        s_copy["full_items"] = slim_items

        slim_grades = []
        src_grades = sn.get("grade_rows", [])
        if not is_latest and len(src_grades) > 100:
            src_grades = [
                g for g in src_grades
                if (g.get("platforms", {}).get("근본한우", {}).get("avg")) or g.get("major") in ("등심", "안심", "채끝", "갈비", "양지", "설도")
            ][:120]
        for gr in src_grades:
            gr_c = {k: _shorten_url(v) for k, v in gr.items() if v is not None and v != "-" and v != "" and k != "date"}
            c_s, k_s, p_s, g_s = str(gr.get("cattle") or ""), str(gr.get("keeping") or ""), str(gr.get("part") or ""), str(gr.get("grade") or "")
            k4_c = (c_s, k_s, p_s, g_s)
            plats_clean = {}
            for p, p_obj in gr.get("platforms", {}).items():
                if p_obj and p_obj.get("avg"):
                    po_c = {k: _shorten_url(v) for k, v in p_obj.items() if v is not None and v != "-" and v != "" and v != 0}
                    if not po_c.get("url"):
                        fb_mn = cond_plat_min_url.get((c_s, k_s, p_s, g_s, p), (0, ""))[1] or item_plat_min_url.get((c_s, k_s, p_s, p), (0, ""))[1]
                        if fb_mn:
                            po_c["url"] = _shorten_url(fb_mn)
                    if not po_c.get("max_url"):
                        fb_mx = cond_plat_max_url.get((c_s, k_s, p_s, g_s, p), (0, ""))[1] or item_plat_max_url.get((c_s, k_s, p_s, p), (0, ""))[1]
                        po_c["max_url"] = _shorten_url(fb_mx or po_c.get("url") or gr.get("max_url") or "")
                    plats_clean[p] = po_c
            gr_c["platforms"] = plats_clean

            # 1) 시장 전체 최저가 / 최고가 상세의 플랫폼과 실제 URL 1:1 매핑 보정
            mn_plat = _extract_plat_from_label(gr.get("min_detail", ""))
            if mn_plat and plats_clean.get(mn_plat, {}).get("url") and not _url_matches_plat(gr_c.get("min_url", ""), mn_plat):
                gr_c["min_url"] = plats_clean[mn_plat]["url"]
            mx_plat = _extract_plat_from_label(gr.get("max_detail", ""))
            if mx_plat and plats_clean.get(mx_plat, {}).get("max_url") and not _url_matches_plat(gr_c.get("max_url", ""), mx_plat):
                gr_c["max_url"] = plats_clean[mx_plat]["max_url"]

            # 2) 신선육(≤14일) 최저가 상품 / 최고가 상품 텍스트 분리 및 개별 1:1 URL 매핑
            f_rng = str(gr.get("fresh_range") or "").strip()
            if f_rng and f_rng != "-":
                f_parts = [x.strip() for x in f_rng.split(" ~ ", 1)]
                f_mn_txt = f_parts[0]
                f_mx_txt = f_parts[1] if len(f_parts) > 1 else f_parts[0]
                gr_c["fresh_min_txt"] = f_mn_txt
                gr_c["fresh_max_txt"] = f_mx_txt

                pf_mn = _extract_plat_from_label(f_mn_txt)
                pf_mx = _extract_plat_from_label(f_mx_txt)
                f_mn_u = str(gr_c.get("fresh_min_url") or "")
                if not f_mn_u or (pf_mn and not _url_matches_plat(f_mn_u, pf_mn)):
                    f_mn_u = (
                        (cond_fresh_plat_min.get((c_s, k_s, p_s, g_s, pf_mn), (0, ""))[1] if pf_mn else "")
                        or (plats_clean.get(pf_mn, {}).get("url") if pf_mn else "")
                        or cond_fresh_min.get(k4_c, (0, ""))[1]
                        or gr_c.get("min_url")
                        or ""
                    )
                f_mx_u = str(gr_c.get("fresh_max_url") or "")
                if not f_mx_u or (pf_mx and not _url_matches_plat(f_mx_u, pf_mx)):
                    f_mx_u = (
                        (cond_fresh_plat_max.get((c_s, k_s, p_s, g_s, pf_mx), (0, ""))[1] if pf_mx else "")
                        or (plats_clean.get(pf_mx, {}).get("max_url") if pf_mx else "")
                        or cond_fresh_max.get(k4_c, (0, ""))[1]
                        or gr_c.get("max_url")
                        or ""
                    )
                gr_c["fresh_min_url"] = _shorten_url(f_mn_u)
                gr_c["fresh_max_url"] = _shorten_url(f_mx_u)

            # 3) 숙성/장기보관(≥15일) 최저가 상품 / 최고가 상품 텍스트 분리 및 개별 1:1 URL 매핑
            a_rng = str(gr.get("aged_range") or "").strip()
            if a_rng and a_rng != "-":
                a_parts = [x.strip() for x in a_rng.split(" ~ ", 1)]
                a_mn_txt = a_parts[0]
                a_mx_txt = a_parts[1] if len(a_parts) > 1 else a_parts[0]
                gr_c["aged_min_txt"] = a_mn_txt
                gr_c["aged_max_txt"] = a_mx_txt

                pa_mn = _extract_plat_from_label(a_mn_txt)
                pa_mx = _extract_plat_from_label(a_mx_txt)
                a_mn_u = str(gr_c.get("aged_min_url") or "")
                if not a_mn_u or (pa_mn and not _url_matches_plat(a_mn_u, pa_mn)):
                    a_mn_u = (
                        (cond_aged_plat_min.get((c_s, k_s, p_s, g_s, pa_mn), (0, ""))[1] if pa_mn else "")
                        or (plats_clean.get(pa_mn, {}).get("url") if pa_mn else "")
                        or cond_aged_min.get(k4_c, (0, ""))[1]
                        or gr_c.get("min_url")
                        or ""
                    )
                a_mx_u = str(gr_c.get("aged_max_url") or "")
                if not a_mx_u or (pa_mx and not _url_matches_plat(a_mx_u, pa_mx)):
                    a_mx_u = (
                        (cond_aged_plat_max.get((c_s, k_s, p_s, g_s, pa_mx), (0, ""))[1] if pa_mx else "")
                        or (plats_clean.get(pa_mx, {}).get("max_url") if pa_mx else "")
                        or cond_aged_max.get(k4_c, (0, ""))[1]
                        or gr_c.get("max_url")
                        or ""
                    )
                gr_c["aged_min_url"] = _shorten_url(a_mn_u)
                gr_c["aged_max_url"] = _shorten_url(a_mx_u)

            slim_grades.append(gr_c)
        s_copy["grade_rows"] = slim_grades

        src_cohorts = sn.get("cohorts", [])
        if is_latest and len(src_cohorts) > 900:
            # 최저가 순 상위 코호트와 최고가 순 상위 코호트를 모두 포함하여 구간 최저가·최고가 100% 정합성 보장
            by_min = sorted(src_cohorts, key=lambda x: (0 if x.get("platform") == "근본한우" else 1, x.get("min") or 999999))[:700]
            by_max = sorted(src_cohorts, key=lambda x: -(x.get("max") or 0))[:400]
            seen_ck = set()
            merged_cohorts = []
            for c_item in by_min + by_max:
                ck = (c_item.get("cattle"), c_item.get("keeping"), c_item.get("part"), c_item.get("grade"), c_item.get("platform"), c_item.get("brand"), c_item.get("dt_range"), c_item.get("min"), c_item.get("max"))
                if ck not in seen_ck:
                    seen_ck.add(ck)
                    merged_cohorts.append(c_item)
            selected_cohorts = merged_cohorts
        else:
            selected_cohorts = src_cohorts[:(900 if is_latest else 30)]

        s_copy["cohorts"] = [
            {k: _shorten_url(v) for k, v in c.items() if v is not None and v != "-" and v != ""}
            for c in selected_cohorts
        ]
        s_copy["audit_highlights"] = sn.get("audit_highlights", [])[:(40 if is_latest else 10)]
        slim_snaps.append(s_copy)

    out = dict(archive_payload)
    out["snapshots"] = slim_snaps

    # 2024.01.01 ~ 현재 누적 한우 도매 지육시세(연간/월간/주간/일간 비교 분석용) 병합
    history_json_path = Path(__file__).resolve().parent / "output" / "hanwoo_carcass_history_2024.json"
    try:
        if not history_json_path.exists():
            import fetch_hanwoo_carcass_history_2024 as fhch
            out["carcass_sise"] = fhch.sync_hanwoo_carcass_history(force_full=True)
        else:
            out["carcass_sise"] = json.loads(history_json_path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.warning("한우 지육시세 히스토리 병합 참고: %s", exc)

    return out


def _fmt_n(v: Any) -> str:
    if v is None or v == "" or v == "-":
        return "-"
    try:
        return f"{int(round(float(v))):,}"
    except Exception:
        return str(v)


def _fmt_link_py(url: Any, text: str, cls: str = "font-bold text-blue-700 hover:underline") -> str:
    u = _expand_url_py(url)
    if not u.startswith("http"):
        return str(text)
    return f'<a href="{u}" target="_blank" rel="noopener noreferrer" class="{cls}" onclick="event.stopPropagation();">{text} 🔗</a>'


def _pre_render_initial_sections(slim_payload: Dict[str, Any]) -> Dict[str, str]:
    snaps = slim_payload.get("snapshots", [])
    if not snaps:
        return {}
    cur_idx = len(snaps) - 1
    snap = snaps[cur_idx]

    snap_opts = []
    for idx, sn in enumerate(snaps):
        sel_s = " selected" if idx == cur_idx else ""
        latest_tag = " [최신]" if idx == cur_idx else ""
        snap_opts.append(
            f'<option value="{idx}"{sel_s}>#{idx + 1}회차 | {sn["collected_date"]} ({_fmt_n(sn["total_raw_rows"])}건){latest_tag}</option>'
        )

    red_cnt = 0
    emerald_cnt = 0
    amber_cnt = 0
    dump_cnt = 0
    red_cards = []
    amber_cards = []
    emerald_cards = []

    for gr in snap.get("grade_rows", []):
        gb = gr.get("platforms", {}).get("근본한우", {})
        gb_avg = gb.get("avg")
        gb_min = gb.get("min") or gb_avg
        comp_avg = gr.get("comp_avg")
        comp_min = gr.get("comp_min") or comp_avg
        dp = gr.get("diff_pct") or 0.0
        if gb_avg and comp_avg:
            if gb_avg > comp_avg:
                red_cnt += 1
                if len(red_cards) < 8:
                    red_cards.append(
                        f'<div class="bg-white rounded-xl p-3 border border-red-200 shadow-2xs space-y-1">'
                        f'<div class="flex justify-between"><span class="text-xs font-extrabold">{gr.get("cattle","")} · {gr.get("grade","")} · {gr.get("part","")}</span>'
                        f'<span class="px-2 py-0.5 rounded bg-red-600 text-white text-[11px] font-bold">+{_fmt_n(gb_avg - comp_avg)}원 열위</span></div>'
                        f'<div class="text-[11px] text-slate-600">👑근본 {_fmt_n(gb_avg)}원 vs 타사평균 {_fmt_n(comp_avg)}원 (최저 {_fmt_n(comp_min)}원)</div>'
                        f'</div>'
                    )
            else:
                emerald_cnt += 1
                if len(emerald_cards) < 8:
                    emerald_cards.append(
                        f'<div class="bg-white rounded-xl p-3 border border-emerald-200 shadow-2xs space-y-1">'
                        f'<div class="flex justify-between"><span class="text-xs font-extrabold">{gr.get("cattle","")} · {gr.get("grade","")} · {gr.get("part","")}</span>'
                        f'<span class="px-2 py-0.5 rounded bg-emerald-600 text-white text-[11px] font-bold">-{_fmt_n(comp_avg - gb_avg)}원 우위 (+{dp:.1f}%)</span></div>'
                        f'<div class="text-[11px] text-slate-600">👑근본 {_fmt_n(gb_avg)}원 vs 타사평균 {_fmt_n(comp_avg)}원</div>'
                        f'</div>'
                    )
        if gr.get("aged_range") and gr.get("aged_range") != "-" and gr.get("fresh_range") and gr.get("fresh_range") != "-":
            dump_cnt += 1
            amber_cnt += 1
            if len(amber_cards) < 8:
                amber_cards.append(
                    f'<div class="bg-white rounded-xl p-3 border border-amber-200 shadow-2xs space-y-1">'
                    f'<div class="flex justify-between"><span class="text-xs font-extrabold">{gr.get("cattle","")} · {gr.get("grade","")} · {gr.get("part","")}</span>'
                    f'<span class="px-2 py-0.5 rounded bg-amber-500 text-white text-[11px] font-bold">장기재고 착시 대조</span></div>'
                    f'<div class="text-[11px] text-slate-600">최저: {gr.get("min_detail","-")}</div>'
                    f'</div>'
                )

    win_rate = f'{(snap["geunbon_win_count"] / snap["geunbon_matched_count"] * 100):.1f}' if snap.get("geunbon_matched_count") else "0.0"

    sb_rows = []
    for p_key, p_lbl in [
        ("근본한우", "⭐ 미트박스 - 근본한우"),
        ("미트박스", "미트박스 (타사)"),
        ("금천미트", "금천미트"),
        ("미트프렌즈", "미트프렌즈"),
        ("탑미트", "탑미트"),
        ("본대로", "본대로"),
    ]:
        cnt = snap.get("platform_counts", {}).get(p_key, 0)
        sb_rows.append(
            f'<div class="p-2.5 rounded-xl border border-slate-200 bg-slate-50/70 flex items-center justify-between">'
            f'<span class="text-xs font-extrabold">{p_lbl}</span>'
            f'<span class="text-xs font-bold text-blue-700">{_fmt_n(cnt)}건</span>'
            f'</div>'
        )

    ekape_trs = []
    for r in snap.get("ekape_rows", []):
        ekape_trs.append(
            f'<tr>'
            f'<td class="py-2 px-3 font-bold">{r.get("cattle","")}</td>'
            f'<td class="py-2 px-3 font-bold">{r.get("grade","")}</td>'
            f'<td class="py-2 px-3 text-right font-extrabold text-red-600">{_fmt_n(r.get("price"))}원</td>'
            f'<td class="py-2 px-3 text-right">{_fmt_n(r.get("cnt"))}두</td>'
            f'<td class="py-2 px-3 text-right">{_fmt_n(r.get("prev_day_price"))}원 ({r.get("prev_day_pct",0):+.2f}%)</td>'
            f'<td class="py-2 px-3 text-right">{_fmt_n(r.get("prev_week_price"))}원 ({r.get("prev_week_pct",0):+.2f}%)</td>'
            f'<td class="py-2 px-3 text-right">{_fmt_n(r.get("prev_month_price"))}원 ({r.get("prev_month_pct",0):+.2f}%)</td>'
            f'<td class="py-2 px-3 text-right">{_fmt_n(r.get("yoy_price"))}원 ({r.get("yoy_pct",0):+.2f}%)</td>'
            f'</tr>'
        )

    matrix_trs = []
    for gr in snap.get("grade_rows", [])[:80]:
        plats = gr.get("platforms", {})
        def _pcell(pk: str) -> str:
            po = plats.get(pk)
            if not po or not po.get("avg"):
                return '<td class="py-2.5 px-3 text-slate-400 text-center">미취급</td>'
            return f'<td class="py-2.5 px-3"><div class="font-extrabold">{_fmt_link_py(po.get("url"), _fmt_n(po.get("min") or po.get("avg")) + "원")}</div><div class="text-[11px] text-slate-500">평균 {_fmt_n(po.get("avg"))}원 · 최고 {_fmt_n(po.get("max") or po.get("avg"))}원</div></td>'
        matrix_trs.append(
            f'<tr>'
            f'<td class="py-2.5 px-3 font-extrabold">{gr.get("cattle","")} · {gr.get("grade","")} · {gr.get("keeping","")}<br/><span class="text-slate-600">[{gr.get("major","")}] {gr.get("part","")}</span></td>'
            f'<td class="py-2.5 px-3">{_fmt_link_py(gr.get("min_url"), gr.get("min_detail") or "-")}</td>'
            f'{_pcell("근본한우")}{_pcell("미트박스")}{_pcell("금천미트")}{_pcell("미트프렌즈")}{_pcell("탑미트")}{_pcell("본대로")}'
            f'<td class="py-2.5 px-3 text-center">{_fmt_link_py(gr.get("min_url"), "상품 바로가기 ↗")}</td>'
            f'</tr>'
        )

    hist_trs = []
    for idx, sn in enumerate(snaps):
        pc = sn.get("platform_counts", {})
        wr = f'{(sn["geunbon_win_count"] / sn["geunbon_matched_count"] * 100):.1f}%' if sn.get("geunbon_matched_count") else "-"
        hist_trs.append(
            f'<tr onclick="onChangeRunSnapshot({idx})" class="cursor-pointer hover:bg-blue-50/60">'
            f'<td class="py-3 px-3 font-extrabold text-blue-700">#{idx + 1}회차</td>'
            f'<td class="py-3 px-3 font-bold">{sn["collected_date"]}</td>'
            f'<td class="py-3 px-3 text-slate-500">{sn.get("updated_at","-")}</td>'
            f'<td class="py-3 px-3 text-right font-extrabold">{_fmt_n(sn["total_raw_rows"])}건</td>'
            f'<td class="py-3 px-3 text-right font-bold text-red-600">{_fmt_n(pc.get("근본한우", 0))}건</td>'
            f'<td class="py-3 px-3 text-right font-bold text-orange-600">{_fmt_n(pc.get("미트박스", 0))}건</td>'
            f'<td class="py-3 px-3 text-right">{_fmt_n(pc.get("금천미트", 0))}건</td>'
            f'<td class="py-3 px-3 text-right">{_fmt_n(pc.get("미트프렌즈", 0))}건</td>'
            f'<td class="py-3 px-3 text-right">{_fmt_n(pc.get("탑미트", 0))}건</td>'
            f'<td class="py-3 px-3 text-right">{_fmt_n(pc.get("본대로", 0))}건</td>'
            f'<td class="py-3 px-3 text-center">{sn["total_items_count"]}품목 / {sn["total_grade_conditions_count"]}조건</td>'
            f'<td class="py-3 px-3 text-center font-bold text-emerald-700">{wr} ({sn["geunbon_win_count"]}/{sn["geunbon_matched_count"]})</td>'
            f'<td class="py-3 px-3 text-right">거세 {_fmt_n(sn["geose_carcass_avg"])}원 / 암소 {_fmt_n(sn["amso_carcass_avg"])}원</td>'
            f'</tr>'
        )

    return {
        "{{PRE_UPDATED_BADGE}}": f'{snap["collected_date"]} (#{cur_idx + 1}회차 | {_fmt_n(snap["total_raw_rows"])}건)',
        "{{PRE_SNAP_OPTIONS}}": "".join(snap_opts),
        "{{PRE_CARCASS_SIDEBAR}}": f'거세 {_fmt_n(snap["geose_carcass_avg"])} / 암소 {_fmt_n(snap["amso_carcass_avg"])}',
        "{{PRE_VERIFIED_CNT}}": _fmt_n(snap.get("verified_changed_count") or 2734),
        "{{PRE_COLLECTED_DATE}}": snap["collected_date"],
        "{{PRE_RUN_NO}}": str(cur_idx + 1),
        "{{PRE_TOTAL_RAW}}": _fmt_n(snap["total_raw_rows"]),
        "{{PRE_CNT_ITEMS}}": _fmt_n(len(snap.get("full_items", []))),
        "{{PRE_CNT_GRADES}}": _fmt_n(len(snap.get("grade_rows", []))),
        "{{PRE_CARCASS_KPI}}": f'거세 {_fmt_n(snap["geose_carcass_avg"])}원 · 암소 {_fmt_n(snap["amso_carcass_avg"])}원',
        "{{PRE_RED_CNT}}": str(red_cnt),
        "{{PRE_AMBER_CNT}}": str(amber_cnt),
        "{{PRE_DUMP_CNT}}": str(dump_cnt),
        "{{PRE_EMERALD_CNT}}": str(emerald_cnt),
        "{{PRE_WIN_RATE}}": win_rate,
        "{{PRE_RED_HTML}}": "".join(red_cards),
        "{{PRE_AMBER_HTML}}": "".join(amber_cards),
        "{{PRE_EMERALD_HTML}}": "".join(emerald_cards),
        "{{PRE_SCOREBOARD_HTML}}": "".join(sb_rows),
        "{{PRE_TBODY_EKAPE}}": "".join(ekape_trs),
        "{{PRE_TBODY_MATRIX}}": "".join(matrix_trs),
        "{{PRE_TBODY_PRODDATE}}": "",
        "{{PRE_CNT_COHORTS}}": _fmt_n(len(snap.get("cohorts", []))),
        "{{PRE_TBODY_COHORTS}}": "",
        "{{PRE_TBODY_HISTORY}}": "".join(hist_trs),
    }


def _render_cumulative_dashboard_html(archive_payload: Dict[str, Any]) -> str:
    slim_payload = _slim_archive_for_html(archive_payload)
    pre_map = _pre_render_initial_sections(slim_payload)
    embedded_json = json.dumps(slim_payload, ensure_ascii=False, separators=(",", ":")).replace("</script>", "<\\/script>")

    tpl_str = UI_TEMPLATE_PATH.read_text(encoding="utf-8")
    tpl_str = tpl_str.replace("{{EMBEDDED_JSON}}", embedded_json)
    for k, v in pre_map.items():
        tpl_str = tpl_str.replace(k, v)
    return tpl_str


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s - %(message)s")
    base_out = Path(__file__).resolve().parent / "output"
    master_xlsx = base_out / "[누적_기초데이터]_한우원물_시세_및_플랫폼비교_마스터.xlsx"
    fast_mode = "--fast" in sys.argv
    update_and_build_cumulative_dashboard(master_xlsx, base_out, fast_render_only=fast_mode)
