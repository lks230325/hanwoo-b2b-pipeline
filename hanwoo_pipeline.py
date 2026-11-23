#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[근본한우 B2B 인텔리전스] 5대 축산 B2B 플랫폼(미트박스/미트프렌즈/탑미트/금천미트/본대로) + 자사(근본한우)
한우암소 · 한우거세 '원물(박스육)' 전수 수집, 축평원(EKAPE) 지육시세 연동, 다차원 가격경쟁력 분석,
6사 비교 시각화 대시보드(PDF & PNG) 및 구글 드라이브 누적 저장(Append & Archive) 엔드투엔드 파이프라인
"""

import argparse
import base64
import json
import logging
import os
import random
import re
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import quote

# 재사용 가능한 가상환경 패키지 경로 추가
_CANDIDATE_VENVS = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), ".venv", "Lib", "site-packages"),
    r"C:\Users\limkyungsoon\.gemini\antigravity\scratch\meatbox-scraper\.venv\Lib\site-packages",
]
for _vp in _CANDIDATE_VENVS:
    if os.path.isdir(_vp) and _vp not in sys.path:
        sys.path.insert(0, _vp)

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.gridspec as gridspec
import matplotlib.patches as mpatches
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import requests
import urllib3
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 한글 폰트 설정 (Malgun Gothic / NanumGothic)
plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("HanwooPipeline")

# ---------------------------------------------------------------------------
# 상수 및 타겟 URL 정의
# ---------------------------------------------------------------------------
GDRIVE_FOLDER_ID = "1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe"
MASTER_EXCEL_FILENAME = "[누적_기초데이터]_한우원물_시세_및_플랫폼비교_마스터.xlsx"

PLATFORMS_ORDER = ["근본한우", "미트박스", "미트프렌즈", "탑미트", "금천미트", "본대로"]
COMPETITOR_PLATFORMS = ["미트박스", "미트프렌즈", "탑미트", "금천미트", "본대로"]
GRADES_ORDER = ["1++(9)", "1++(8)", "1++(7)", "1++(불명)", "1+", "1", "2", "3", "등외"]
MAIN_GRADES_COMPARE = ["1++(9)", "1++(8)", "1++(7)", "1++(불명)", "1+", "1", "2", "3"]
CORE_PARTS_ORDER = [
    # 1. 등심/안심/채끝/살치 계열 (세부 품목별 분리)
    "등심", "윗등심", "윗등심(살치제외)", "아랫등심", "아랫등심(꽃등심제외)", "꽃등심", "살치살",
    "안심", "채끝",
    # 2. 갈비 계열 (뼈포함 통갈비/짝갈비 vs 정선 갈비살/늑간살/갈비본살 완전 분리)
    "갈비(짝갈비/통갈비)", "갈비살", "늑간살", "갈비본살", "갈비본살+살치살",
    # 3. 양지/차돌/치마/업진 계열 (세부 품목별 분리)
    "양지", "양지머리", "양지머리+치마양지", "차돌양지", "갈비양지",
    "차돌박이", "치마살", "앞치마살", "업진살", "업진안살", "업진살+치마살", "앞치마살+업진안살",
    # 4. 설도/설깃/삼각/보섭/도가니살 계열 (세부 품목별 분리)
    "설도", "설도(삼각살제외)", "설깃살", "설깃머리", "삼각살", "보섭살", "도가니살",
    # 5. 우둔/홍두깨 계열 (우둔 vs 홍두깨 완전 분리)
    "우둔", "홍두깨",
    # 6. 앞다리/꾸리/부채 계열 (앞다리 vs 꾸리살 vs 부채살 완전 분리)
    "앞다리", "앞다리+꾸리살", "꾸리살", "부채살", "부채안살",
    # 7. 사태/목심 계열
    "사태", "아롱사태", "목심",
    # 8. 특수부위 계열 (3종모둠 vs 2종복합 vs 단품 완전 분리)
    "특수부위(3종모둠)", "토시살+제비추리", "안창살", "토시살", "제비추리",
    # 9. 부산물/뼈/기타 원물 계열
    "우족", "사골", "꼬리", "꼬리반골", "알꼬리", "잡뼈", "도가니", "스지", "잡육",
    "소머리", "곱대창", "대창", "벌양", "양", "천엽", "지방",
]

MAJOR_10_PARTS_MAP = {
    "등심": "등심", "윗등심": "등심", "윗등심(살치제외)": "등심", "아랫등심": "등심",
    "아랫등심(꽃등심제외)": "등심", "꽃등심": "등심", "살치살": "등심",
    "안심": "안심",
    "채끝": "채끝",
    "갈비(짝갈비/통갈비)": "갈비", "갈비살": "갈비", "늑간살": "갈비", "갈비본살": "갈비", "갈비본살+살치살": "갈비",
    "양지": "양지", "양지머리": "양지", "양지머리+치마양지": "양지", "차돌양지": "양지", "갈비양지": "양지",
    "차돌박이": "양지", "치마살": "양지", "앞치마살": "양지", "업진살": "양지", "업진안살": "양지",
    "업진살+치마살": "양지", "앞치마살+업진안살": "양지",
    "설도": "설도", "설도(삼각살제외)": "설도", "설깃살": "설도", "설깃머리": "설도",
    "삼각살": "설도", "보섭살": "설도", "도가니살": "설도",
    "우둔": "우둔", "홍두깨": "우둔",
    "앞다리": "앞다리", "앞다리+꾸리살": "앞다리", "꾸리살": "앞다리", "부채살": "앞다리", "부채안살": "앞다리",
    "사태": "사태", "아롱사태": "사태",
    "목심": "목심",
    "특수부위(3종모둠)": "특수부위", "토시살+제비추리": "특수부위", "안창살": "특수부위", "토시살": "특수부위", "제비추리": "특수부위",
    "우족": "부산물/뼈", "사골": "부산물/뼈", "꼬리": "부산물/뼈", "꼬리반골": "부산물/뼈", "알꼬리": "부산물/뼈",
    "잡뼈": "부산물/뼈", "도가니": "부산물/뼈", "스지": "부산물/뼈", "소머리": "부산물/뼈",
    "곱대창": "부산물/뼈", "대창": "부산물/뼈", "벌양": "부산물/뼈", "양": "부산물/뼈", "천엽": "부산물/뼈", "지방": "부산물/뼈",
    "잡육": "잡육",
}

# 1차 명확한 비원물(세절/슬라이스/다짐육/양념/선물세트/소포장) 제외 키워드
HARD_NON_WONMUL_KEYWORDS = [
    "세절", "슬라이스", "대패", "다짐", "민찌", "분쇄", "양념", "시즈닝", "마리네이드",
    "선물세트", "명절세트", "소포장", "스킨포장", "트레이", "세절육", "가공육", "육포", "떡갈비", "패티",
    "500g*", "400g*", "300g*", "200g*", "150g*", "-팩(500g", "-팩(300g", "로스용", "구이세절"
]

# 2차 교차검증 대상 키워드 (단순 용도 안내문구인지 실제 세절/소분육인지 2차 검증)
SOFT_USAGE_KEYWORDS = [
    "구이용", "불고기용", "국거리용", "장조림용", "산적용", "육회용", "샤브",
    "스테이크용", "손질", "절단", "컷팅", "커팅", "세트상품", "팩포장"
]


@dataclass
class HanwooProductRow:
    """Task 1 필수 수집 항목 (11개 표준 컬럼 + 2차검증 및 생산일자·브랜드 심층 분석 필드)"""
    collected_date: str          # 수집일자 (YYYY-MM-DD)
    platform: str                # 플랫폼명 (근본한우, 미트박스, 미트프렌즈, 탑미트, 금천미트, 본대로)
    cattle_type: str             # 축종구분 (한우거세 / 한우암소)
    keeping: str                 # 보관상태 (냉장 / 냉동)
    processing: str              # 가공상태 (원물)
    part_name: str               # 부위명 (2차검증 완료 표준부위명)
    grade: str                   # 등급 (1++(9), 1++(8), 1++(7), 1++(불명), 1+, 1, 2, 3, 등외)
    brand_name: str              # 브랜드명 (브랜드 및 가공장 명시)
    production_date: str         # 생산일자(가공일자/도축일자) (YYYY-MM-DD)
    unit_price_kg: int           # 단위단가 (1kg당 가격 원/kg)
    product_url: str             # 상품URL
    raw_product_name: str = ""   # 원본 상품명 (플랫폼별 원본 표기)
    stock_qty: int = 1           # 재고수량 (박스/개)
    major_part: str = ""         # 10대 대분류 부위명 (등심, 안심, 채끝, 갈비, 양지, 설도, 사태, 앞다리, 우둔, 목심 등)
    primary_part_raw: str = ""   # 1차 플랫폼 카테고리/품목 원본값
    sub_spec_name: str = ""      # 2차검증 세부 표준규격명 (예: 등심(윗등심-살치제외), 갈비(짝갈비), 특수부위(안창+토시+제비 모둠))
    verify_status: str = ""      # 2차검증 판정 결과 (예: 1·2차일치, 2차보정:윗등심(살치제외)→등심, 2차구제:용도표기원물 등)

    def __post_init__(self) -> None:
        if not self.major_part or not self.sub_spec_name or not self.verify_status:
            m_part, std_part, sub_spec, v_stat = HanwooStandardizer.verify_and_normalize_part(
                self.primary_part_raw or self.part_name,
                self.raw_product_name,
                self.unit_price_kg,
            )
            if not self.major_part:
                self.major_part = m_part
            self.part_name = std_part
            if not self.sub_spec_name:
                self.sub_spec_name = sub_spec
            if not self.verify_status:
                self.verify_status = v_stat

    def to_excel_dict(self) -> Dict[str, Any]:
        return {
            "수집일자": self.collected_date,
            "플랫폼명": self.platform,
            "축종구분": self.cattle_type,
            "보관상태": self.keeping,
            "가공상태": self.processing,
            "대분류부위(10대기준)": self.major_part,
            "부위명": self.part_name,
            "2차검증_세부표준상품명": self.sub_spec_name,
            "등급": self.grade,
            "브랜드명": self.brand_name,
            "생산일자(가공일자/도축일자)": self.production_date,
            "단위단가": self.unit_price_kg,
            "옵션재고수량": self.stock_qty,
            "1차_플랫폼원본품목": self.primary_part_raw or self.part_name,
            "원본상품명": self.raw_product_name,
            "2차검증_판정결과": self.verify_status,
            "상품URL": self.product_url,
        }


def create_http_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(
        total=3,
        backoff_factor=0.4,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
        )
    })
    return s


build_http_session = create_http_session


# ---------------------------------------------------------------------------
# 표준화 및 2차 교차검증 엔진 (Two-Stage Verification & Standardizer)
# ---------------------------------------------------------------------------
class HanwooStandardizer:
    # 2차 검증 로그 누적용 (플랫폼별 상품명 보정 및 원물 구제/제외 이력 기록)
    verification_audit_logs: List[Dict[str, Any]] = []

    @classmethod
    def record_audit(
        cls,
        platform: str,
        cattle_type: str,
        raw_category: str,
        raw_name: str,
        brand_name: str,
        stage1_result: str,
        stage2_part: str,
        stage2_spec: str,
        verify_action: str,
        included: bool,
    ) -> None:
        if len(cls.verification_audit_logs) < 5000:
            cls.verification_audit_logs.append({
                "플랫폼명": platform,
                "축종구분": cattle_type,
                "브랜드명": brand_name,
                "1차_원본카테고리": raw_category,
                "원본상품명": raw_name,
                "1차_분류값": stage1_result,
                "2차검증_표준부위명": stage2_part,
                "2차검증_세부표준상품명": stage2_spec,
                "2차검증_수행내역": verify_action,
                "분석포함여부": "포함(원물)" if included else "제외(비원물/세절가공)",
            })

    @staticmethod
    def is_wonmul(name: str, category: str = "", proc_hint: str = "") -> bool:
        """
        [1차 + 2차 원물 검증]
        - 1차: 세절/슬라이스/다짐육/양념육/소포장/선물세트 등 명확한 가공육(HARD_NON_WONMUL_KEYWORDS) 즉시 제외
        - 2차: 상품명에 '구이용', '국거리용', '불고기용', '절단', '손질' 등이 포함되더라도
               가공형태(proc_hint)가 '원물'/'박스육'/'원육'이거나 실제 세절 두께(mm)/소포장 팩 규격이 없는
               부분육 원상태(진공박스육, 사골/꼬리/우족 원물 등)이면 2차 검증을 통해 원물로 구제!
        """
        proc_clean = (proc_hint or "").strip()
        if proc_clean and proc_clean not in ("원물", "박스육", "원육", "부분육", "진공", "벌크포장"):
            if any(k in proc_clean for k in ["세절", "분쇄", "양념", "소분", "슬라이스", "다짐", "선물", "스테이크", "구이", "불고기", "국거리"]):
                return False

        combined = f"{name} {category}"
        for kw in HARD_NON_WONMUL_KEYWORDS:
            if kw in combined:
                return False

        # 두께 표기(예: 1.5mm, 2mm, 3mm, 세절)가 있으면 슬라이스 세절육이므로 제외
        if re.search(r"\b\d+(?:\.\d+)?\s*mm\b", combined, re.I):
            return False

        # 2차 검증: 용도 표기 키워드(SOFT_USAGE_KEYWORDS)가 있는 경우
        has_soft_kw = any(kw in combined for kw in SOFT_USAGE_KEYWORDS)
        if has_soft_kw:
            # 명시적으로 원물/박스육이거나 사골/우족/꼬리/잡뼈/양지/사태/등심 등의 원물 카테고리인 경우 구제
            if proc_clean in ("원물", "박스육", "원육", "진공", "벌크포장"):
                return True
            # 선물세트/소분팩이 아닌 부분육 박스 표기(1두, 반두, 진공, 원물, kg단가 상품)는 구제
            if any(w in combined for w in ["세트", "선물", "소분", "세절", "슬라이스", "팩포장"]):
                return False
            return True

        return True

    @staticmethod
    def normalize_keeping(keep_hint: str, name: str = "", category: str = "") -> str:
        """핵심 필터링 규칙 2: 보관상태 구분 ('냉장' / '냉동')"""
        kh = (keep_hint or "").strip().lower()
        txt = f"{keep_hint} {name} {category}"
        # 상품명에 명시적으로 '냉동' 또는 '동결'이 표기된 경우 최우선 반영 (2차 교차검증)
        if any(w in f"{name} {category}" for w in ["냉동", "동결", "세절냉동"]):
            return "냉동"
        if kh in ("f", "frozen", "냉동", "동결", "냉동상품", "냉동보관"):
            return "냉동"
        if kh in ("r", "cold", "chilled", "냉장", "냉장상품", "냉장보관"):
            if any(w in txt for w in ["우족", "사골", "꼬리", "잡뼈", "도가니", "스지"]):
                return "냉동"
            return "냉장"
        if "냉장" in txt:
            return "냉장"
        if any(w in txt for w in ["우족", "사골", "꼬리", "잡뼈", "도가니", "스지"]):
            return "냉동"
        return "냉장"

    @staticmethod
    def normalize_grade(grade_hint: str, bms_hint: str = "", text_hint: str = "") -> str:
        """
        핵심 필터링 규칙 3: 등급별 정밀 구분
        1++(9), 1++(8), 1++(7), 1++(불명), 1+, 1, 2, 3, 등외로 표준화
        """
        combined = f"{grade_hint} {bms_hint} {text_hint}".strip()
        bms_clean = re.sub(r"[^\d]", "", str(bms_hint or ""))

        if "1++" in combined or "투뿔" in combined:
            if bms_clean in ("7", "8", "9"):
                return f"1++({bms_clean})"
            m_bms = re.search(r"1\+\+\s*(?:[ABC]\s*)?[\(\[\-\s]*(?:No\.?\s*|BMS\s*)?([789])\b", combined, re.I)
            if m_bms:
                return f"1++({m_bms.group(1)})"
            m_bms2 = re.search(r"[\(\[]([789])[\)\]]", combined)
            if m_bms2:
                return f"1++({m_bms2.group(1)})"
            return "1++(불명)"

        if re.search(r"1\+(?!\+)", combined) or "원뿔" in combined:
            return "1+"

        if re.search(r"(?<![\d\+])1\s*(?:등급|[ABC]\b)", combined) or str(grade_hint).strip() in ("1", "1등급", "1A", "1B", "1C"):
            return "1"
        if re.search(r"(?<![\d\+])2\s*(?:등급|[ABC]\b)", combined) or str(grade_hint).strip() in ("2", "2등급", "2A", "2B", "2C"):
            return "2"
        if re.search(r"(?<![\d\+])3\s*(?:등급|[ABC]\b)", combined) or str(grade_hint).strip() in ("3", "3등급", "3A", "3B", "3C"):
            return "3"
        if "등외" in combined or "등급외" in combined:
            return "등외"

        if bms_clean in ("7", "8", "9"):
            return f"1++({bms_clean})"
        if bms_clean in ("4", "5", "6"):
            return "1+"
        if bms_clean in ("2", "3"):
            return "1"
        if bms_clean == "1":
            return "2"

        return "등외"

    @classmethod
    def verify_and_normalize_part(
        cls,
        part_hint: str,
        name_hint: str = "",
        unit_price_kg: int = 0,
    ) -> Tuple[str, str, str, str]:
        """
        [보완사항 4 & 왜곡방지 정밀분류: 플랫폼별 상이한 상품명 2차 교차검증 엔진]
        반환값: (대분류부위, 2차검증_표준품목명, 2차검증_세부표준상품명, 2차검증_판정결과)
        - 서로 단가 체급이 다른 세부 품목(예: 등심 vs 윗등심 vs 윗등심(살치제외) vs 아랫등심 vs 꽃등심 vs 살치살,
          짝갈비 vs 갈비살 vs 늑간살 vs 갈비본살 vs 갈비본살+살치살, 설도 vs 설깃살 vs 삼각살, 우둔 vs 홍두깨,
          앞다리 vs 꾸리살 vs 부채살 등)을 억지로 10대 대분류로 뭉개지 않고 개별 표준품목으로 정밀 분리합니다.
        - 괄호 내 부정/미포함 문구(예: '앞다리(꾸리,부채 미포함제품)', '설도(삼각살 X)')나 해부학적 구성 설명
          (예: '설도(도가니살, 보섭살)', '설도 보섭살,도가니살,삼각살 반두')으로 인한 타 품목 오분류 왜곡을 원천 차단합니다.
        """
        p_clean = (part_hint or "").strip()
        n_clean = (name_hint or "").strip()
        raw = f"{p_clean} {n_clean}".strip()
        # 브랜드명 내 '안심' 등 오탐 방지
        clean_r = (
            raw.replace("안심목장한우", "")
            .replace("안심목장", "")
            .replace("안심한우", "")
            .replace("안심LPC", "")
            .replace("안심축산", "")
        )
        # 괄호 안의 미포함/제외 문구를 제거한 문자열 (부정어로 인한 타 부위 오탐 방지용)
        clean_no_excl = re.sub(r"\([^)]*(?:미포함|제외|\s*X\b|없음)[^)]*\)", "", clean_r, flags=re.I)

        # 0) 소머리/내장/부산물 선행 분리 (미트프렌즈 '특수부위(거세 소머리육...)' 등이 특수부위 모둠으로 섞이는 왜곡 차단)
        if "소머리" in clean_r or "우설" in clean_r:
            return "부산물/뼈", "소머리", "소머리(소머리육/우설포함)", "2차보정: 소머리육(특수부위오분류방지)→소머리 분리"
        if "곱대창" in clean_r:
            return "부산물/뼈", "곱대창", "곱대창(냉동원물)", "1·2차 일치(곱대창)"
        if "대창" in clean_r:
            return "부산물/뼈", "대창", "대창(냉동원물)", "1·2차 일치(대창)"
        if "벌양" in clean_r:
            return "부산물/뼈", "벌양", "벌양(냉동원물)", "1·2차 일치(벌양)"
        if "천엽" in clean_r:
            return "부산물/뼈", "천엽", "천엽(냉동원물)", "1·2차 일치(천엽)"
        if p_clean in ("양", "냉동양") or "한우냉동양" in clean_r:
            return "부산물/뼈", "양", "양(냉동원물)", "1·2차 일치(양)"
        if p_clean == "지방" or "]지방" in clean_r:
            return "부산물/뼈", "지방", "지방(원물)", "1·2차 일치(지방)"

        # 1) 채끝 vs 채끝등심 교차검증 (근본한우/미트박스 '채끝등심'과 타사 '채끝'은 동일 품목)
        if "채끝" in clean_r:
            status = "2차보정: 채끝등심→채끝 표준화" if "채끝등심" in clean_r else "1·2차 일치(채끝)"
            sub_spec = "채끝(근염제거/정선)" if "근염" in clean_r else "채끝(원물)"
            return "채끝", "채끝", sub_spec, status

        # 2) 안심 교차검증
        if "안심" in clean_r:
            sub_spec = "안심(추리/날개살제거)" if any(w in clean_r for w in ["제거", "알안심"]) else "안심(원물)"
            return "안심", "안심", sub_spec, "1·2차 일치(안심)"

        # 3) 갈비 계열 세밀 분리: 갈비본살+살치살 vs 갈비본살 vs 늑간살 vs 갈비살 vs 갈비(짝갈비/통갈비)
        if ("갈비본살" in clean_r and "살치" in clean_r) or "살치본살" in clean_r:
            return "갈비", "갈비본살+살치살", "갈비본살+살치살(복합)", "2차보정: 갈비본살+갈비살치→갈비본살+살치살 표준화"
        if p_clean == "갈비본살" or ("갈비본살" in clean_r and "늑간" not in clean_r and "갈비살(" not in clean_r):
            return "갈비", "갈비본살", "갈비본살(단품)", "1·2차 일치(갈비본살)"
        if p_clean == "늑간살" or ("늑간살" in clean_r and "본살" not in clean_r and "갈비살" not in p_clean):
            return "갈비", "늑간살", "늑간살(단품)", "1·2차 일치(늑간살)"
        if p_clean == "갈비살" or "갈비살" in clean_no_excl:
            return "갈비", "갈비살", "갈비살(원물)", "1·2차 일치(갈비살)"
        if p_clean in ("갈비", "냉동갈비", "짝갈비") or ("짝갈비" in clean_r and "갈비양지" not in clean_r):
            return "갈비", "갈비(짝갈비/통갈비)", "갈비(뼈포함 짝갈비/통갈비)", "2차검증: 뼈포함 통갈비/짝갈비 분리(갈비살혼합방지)"

        # 4) 등심 계열 세밀 분리: 등심(통등심/2분할) vs 윗등심 vs 윗등심(살치제외) vs 아랫등심 vs 아랫등심(꽃등심제외) vs 꽃등심 vs 살치살
        if p_clean in ("등심살치살", "살치살") or ("등심살치살" in clean_r):
            return "등심", "살치살", "살치살(등심살치살 단품)", "2차보정: 등심살치살→살치살(단품) 분리"
        if p_clean == "윗등심(살치제외)" or ("윗등심" in clean_r and "살치제외" in clean_r):
            return "등심", "윗등심(살치제외)", "윗등심(살치제외)", "1·2차 일치(윗등심-살치제외 분리)"
        if p_clean == "아랫등심(꽃등심제외)" or ("아랫등심" in clean_r and "꽃등심제외" in clean_r):
            return "등심", "아랫등심(꽃등심제외)", "아랫등심(꽃등심제외)", "1·2차 일치(아랫등심-꽃등심제외 분리)"
        if p_clean == "꽃등심" or ("꽃등심" in clean_no_excl and "아랫등심" not in p_clean):
            return "등심", "꽃등심", "꽃등심(단품)", "1·2차 일치(꽃등심 분리)"
        if p_clean == "윗등심" or ("윗등심" in clean_r and "아랫등심" not in clean_r and "2분할" not in clean_r):
            return "등심", "윗등심", "윗등심(살치포함 단품)", "1·2차 일치(윗등심 분리)"
        if p_clean == "아랫등심" or ("아랫등심" in clean_r and "윗등심" not in clean_r and "2분할" not in clean_r):
            return "등심", "아랫등심", "아랫등심(꽃등심포함 단품)", "1·2차 일치(아랫등심 분리)"
        if p_clean == "등심" or "등심" in clean_r:
            sub_s = "등심(2분할:윗+아랫등심)" if "2분할" in clean_r else "등심(통등심/일반)"
            return "등심", "등심", sub_s, "1·2차 일치(등심)"
        if "살치" in clean_no_excl and "갈비" not in clean_no_excl:
            return "등심", "살치살", "살치살(단품)", "1·2차 일치(살치살)"

        # 5) 특수부위 계열 세밀 분리: 특수부위(3종모둠) vs 토시살+제비추리(2종) vs 안창살 vs 토시살 vs 제비추리
        if p_clean in ("안창외2종", "특수부위(토시+안창+제비추리)", "특수부위"):
            return (
                "특수부위",
                "특수부위(3종모둠)",
                "특수부위(3종모둠세트)",
                "2차보정: 플랫폼별 3종모둠(특수부위/안창외2종)→특수부위(3종모둠) 표준화",
            )
        if p_clean in ("토시살+제비추리", "토시,제비추리", "특수부위(토시,제비추리)"):
            return (
                "특수부위",
                "토시살+제비추리",
                "토시살+제비추리(2종복합)",
                "2차보정: 토시+제비추리 2종복합 표준화",
            )
        if p_clean == "안창살":
            return "특수부위", "안창살", "안창살(단품)", "1·2차 일치(안창살)"
        if p_clean == "토시살":
            return "특수부위", "토시살", "토시살(단품)", "1·2차 일치(토시살)"
        if p_clean == "제비추리":
            return "특수부위", "제비추리", "제비추리(단품)", "1·2차 일치(제비추리)"

        # 6) 양지/차돌/치마/업진 계열 세밀 분리
        if p_clean in ("앞치마살(대접살)+업진안살", "앞치마살,업진안살", "앞치마살+업진안살") or (
            "앞치마" in clean_no_excl and "업진안살" in clean_no_excl
        ):
            return (
                "양지",
                "앞치마살+업진안살",
                "앞치마살+업진안살(복합)",
                "2차보정: 앞치마살+업진안살 복합규격 표준화",
            )
        if p_clean == "업진살+치마살" or ("업진살+치마살" in clean_no_excl):
            return (
                "양지",
                "업진살+치마살",
                "업진살+치마살(복합)",
                "1·2차 일치(업진살+치마살 복합)",
            )
        if p_clean in ("차돌박이", "차돌박이(냉장)", "냉동차돌박이") or "차돌박이" in clean_no_excl:
            return "양지", "차돌박이", "차돌박이(원물)", "1·2차 일치(차돌박이)"
        if p_clean == "앞치마살" or ("앞치마살" in clean_no_excl and "업진" not in clean_no_excl):
            return "양지", "앞치마살", "앞치마살(단품)", "1·2차 일치(앞치마살 분리)"
        if p_clean == "치마살" or ("치마살" in clean_no_excl and "업진" not in clean_no_excl and "토시" not in clean_no_excl):
            return "양지", "치마살", "치마살(단품)", "1·2차 일치(치마살 분리)"
        if p_clean == "업진안살" or ("업진안살" in clean_no_excl and "앞치마" not in clean_no_excl):
            return "양지", "업진안살", "업진안살(단품)", "1·2차 일치(업진안살 분리)"
        if p_clean == "업진살" or ("업진살" in clean_no_excl and "치마" not in clean_no_excl and "토시" not in clean_no_excl):
            return "양지", "업진살", "업진살(단품)", "1·2차 일치(업진살 분리)"
        if p_clean in ("양지머리,치마양지", "냉동양지머리,치마양지") or ("양지머리" in clean_no_excl and "치마양지" in clean_no_excl):
            return "양지", "양지머리+치마양지", "양지머리+치마양지", "2차보정: 양지머리,치마양지→양지머리+치마양지 분리"
        if p_clean in ("양지머리 외", "양지머리외", "양지머리"):
            return "양지", "양지머리", "양지머리(양지머리외)", "2차보정: 양지머리외→양지머리 분리"
        if p_clean == "차돌양지" or "차돌양지" in clean_no_excl:
            return "양지", "차돌양지", "차돌양지(원물)", "1·2차 일치(차돌양지 분리)"
        if p_clean == "갈비양지" or "갈비양지" in clean_no_excl:
            return "양지", "갈비양지", "갈비양지(원물)", "1·2차 일치(갈비양지 분리)"
        if p_clean == "양지" or "양지" in clean_no_excl:
            return "양지", "양지", "양지(통양지/일반)", "1·2차 일치(양지)"

        # 7) 앞다리/꾸리/부채 계열 세밀 분리 (괄호 안 '꾸리,부채 미포함제품' 오탐 원천 차단!)
        if p_clean in ("앞다리+꾸리살", "앞다리살(앞다리+꾸리)", "앞다리셋트") or (
            "앞다리" in clean_no_excl and "꾸리" in clean_no_excl
        ):
            return "앞다리", "앞다리+꾸리살", "앞다리+꾸리살(세트)", "2차보정: 앞다리+꾸리살/앞다리셋트 표준화"
        if p_clean in ("앞다리", "앞다리/전각", "앞다리살", "냉동앞다리"):
            return "앞다리", "앞다리", "앞다리(전각/앞다리살)", "2차보정: 앞다리/전각/앞다리살→앞다리 표준화"
        if p_clean == "부채안살" or "부채안살" in clean_no_excl:
            return "앞다리", "부채안살", "부채안살(단품)", "1·2차 일치(부채안살 분리)"
        if p_clean in ("부채살", "부챗살") or "부채살" in clean_no_excl or "부챗살" in clean_no_excl:
            stat = "2차보정: 부챗살→부채살 표준화" if "부챗살" in clean_r else "1·2차 일치(부채살)"
            return "앞다리", "부채살", "부채살(단품)", stat
        if p_clean == "꾸리살" or "꾸리살" in clean_no_excl:
            return "앞다리", "꾸리살", "꾸리살(단품)", "1·2차 일치(꾸리살 분리)"
        if any(w in clean_no_excl for w in ["앞다리", "전각"]):
            return "앞다리", "앞다리", "앞다리(전각/앞다리살)", "1·2차 일치(앞다리)"

        # 8) 설도/설깃/삼각/보섭/도가니살 계열 세밀 분리
        # ('설도(삼각살 X)'는 설도(삼각살제외)로, '설도(도가니살, 보섭살)' / '설도 보섭살,도가니살,삼각살 반두'는 통설도로 정확 분류!)
        if p_clean == "설도(삼각살 X)" or ("설도" in clean_r and "삼각살" in clean_r and ("X" in clean_r or "제외" in clean_r)):
            return "설도", "설도(삼각살제외)", "설도(삼각살제외)", "2차보정: 설도(삼각살 X)→설도(삼각살제외) 분리"
        if p_clean == "설깃머리" or "설깃머리" in clean_no_excl:
            if p_clean in ("설깃", "설깃살", "냉동설깃") and "포함" in clean_r:
                return "설도", "설깃살", "설깃살(설깃머리포함)", "2차보정: 설깃→설깃살 표준화"
            return "설도", "설깃머리", "설깃머리(단품)", "1·2차 일치(설깃머리 분리)"
        if p_clean in ("설깃", "설깃살", "냉동설깃"):
            stat = "2차보정: 설깃→설깃살 표준화" if p_clean != "설깃살" else "1·2차 일치(설깃살)"
            return "설도", "설깃살", "설깃살(원물)", stat
        if p_clean in ("설도", "냉동설도"):
            return "설도", "설도", "설도(통설도/일반)", "1·2차 일치(설도)"
        if p_clean == "삼각살" or ("삼각살" in clean_no_excl and "설도" not in p_clean):
            return "설도", "삼각살", "삼각살(단품)", "1·2차 일치(삼각살 분리)"
        if p_clean == "보섭살" or ("보섭살" in clean_no_excl and "설도" not in p_clean):
            return "설도", "보섭살", "보섭살(단품)", "1·2차 일치(보섭살 분리)"
        if p_clean == "도가니살" or ("도가니살" in clean_no_excl and "설도" not in p_clean):
            return "설도", "도가니살", "도가니살(단품)", "1·2차 일치(도가니살 분리)"
        if "설깃" in clean_no_excl:
            return "설도", "설깃살", "설깃살(원물)", "2차보정: 설깃→설깃살 표준화"
        if "설도" in clean_no_excl:
            return "설도", "설도", "설도(통설도/일반)", "1·2차 일치(설도)"

        # 9) 우둔 vs 홍두깨 계열 세밀 분리 (합치지 않고 별도 품목으로 유지!)
        if p_clean in ("홍두깨", "홍두깨살") or "홍두깨" in clean_no_excl:
            stat = "2차보정: 홍두깨살→홍두깨 표준화" if "홍두깨살" in clean_r else "1·2차 일치(홍두깨)"
            return "우둔", "홍두깨", "홍두깨(원물)", stat
        if p_clean in ("우둔", "우둔살", "냉동우둔") or "우둔" in clean_no_excl:
            stat = "2차보정: 우둔살→우둔 표준화" if "우둔살" in clean_r else "1·2차 일치(우둔)"
            return "우둔", "우둔", "우둔(원물)", stat

        # 10) 사태 vs 아롱사태 세밀 분리
        if p_clean == "아롱사태" or "아롱사태" in clean_no_excl:
            return "사태", "아롱사태", "아롱사태(단품)", "1·2차 일치(아롱사태 분리)"
        if p_clean in ("사태", "냉동사태") or "사태" in clean_no_excl:
            return "사태", "사태", "사태(일반/통사태)", "1·2차 일치(사태)"

        # 11) 목심 교차검증
        if p_clean in ("목심", "냉동목심") or "목심" in clean_no_excl or "목살" in clean_no_excl:
            return "목심", "목심", "목심(원물)", "1·2차 일치(목심)"

        # 12) 부산물 / 뼈 / 정육 세밀 분리 (꼬리 vs 꼬리반골 vs 알꼬리 분리!)
        if "우족" in clean_r:
            return "부산물/뼈", "우족", "우족(냉동원물)", "1·2차 일치(우족)"
        if "사골" in clean_r:
            return "부산물/뼈", "사골", "사골(냉동원물)", "1·2차 일치(사골)"
        if "알꼬리" in clean_r:
            return "부산물/뼈", "알꼬리", "알꼬리(냉동원물)", "1·2차 일치(알꼬리 분리)"
        if "꼬리반골" in clean_r or "반골" in clean_r:
            return "부산물/뼈", "꼬리반골", "꼬리반골(냉동원물)", "2차보정: 꼬리반골 분리 표준화"
        if "꼬리" in clean_r:
            return "부산물/뼈", "꼬리", "꼬리(냉동원물)", "1·2차 일치(꼬리)"
        if "잡뼈" in clean_r or "등뼈" in clean_r or "목뼈" in clean_r:
            return "부산물/뼈", "잡뼈", "잡뼈(냉동원물)", "1·2차 일치(잡뼈)"
        if "스지" in clean_r or "힘줄" in clean_r:
            return "부산물/뼈", "스지", "스지(냉동원물)", "1·2차 일치(스지)"
        if "도가니" in clean_r:
            return "부산물/뼈", "도가니", "도가니(냉동원물)", "1·2차 일치(도가니)"
        if "잡육" in clean_r or "정육" in clean_r or "사태잡육" in clean_r:
            return "잡육", "잡육", "잡육/기타정육(원물)", "1·2차 일치(잡육)"

        fallback_p = p_clean or "기타원물"
        return "기타원물", fallback_p, f"{fallback_p}(기타)", "2차검증: 기타특수품목 분류"

    @classmethod
    def normalize_part(cls, part_hint: str, name_hint: str = "", unit_price_kg: int = 0) -> str:
        _, std_part, _, _ = cls.verify_and_normalize_part(part_hint, name_hint, unit_price_kg)
        return std_part

    @staticmethod
    def normalize_date(dt_str: Any) -> str:
        """생산일자(가공일자/도축일자)를 YYYY-MM-DD 형식으로 정규화"""
        s = str(dt_str or "").strip()
        if not s or s == "-" or s.lower() == "none":
            return "-"
        m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
        if m_iso:
            y = int(m_iso.group(1))
            if y >= 2030 or y < 2020:
                return "-"
            return f"{m_iso.group(1)}-{m_iso.group(2)}-{m_iso.group(3)}"
        m_yy = re.match(r"^(\d{2})[-./](\d{2})[-./](\d{2})$", s)
        if m_yy:
            return f"20{m_yy.group(1)}-{m_yy.group(2)}-{m_yy.group(3)}"
        digits = re.sub(r"[^\d]", "", s)
        if len(digits) == 8 and digits.startswith("20"):
            y = int(digits[:4])
            if 2020 <= y < 2030:
                return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
        if len(digits) == 6:
            return f"20{digits[:2]}-{digits[2:4]}-{digits[4:6]}"
        return "-"

    @staticmethod
    def to_year_month_or_week(dt_str: str) -> str:
        """생산일자를 연-월 및 주차(예: 2026-10 (4주차)) 또는 연-월(YYYY-MM) 형태로 변환"""
        if not dt_str or dt_str == "-" or len(dt_str) < 10:
            return "일자미상"
        try:
            dt = datetime.strptime(dt_str[:10], "%Y-%m-%d")
            week_of_month = (dt.day - 1) // 7 + 1
            return f"{dt.strftime('%Y-%m')} ({week_of_month}주차)"
        except Exception:
            return dt_str[:7] if len(dt_str) >= 7 else "일자미상"


# ---------------------------------------------------------------------------
# [Task 3] 축산물품질평가원(EKAPE) 한우거세 · 한우암소 지육시세(경락가격) 동향 수집기
# ---------------------------------------------------------------------------
def fetch_ekape_hanwoo_auction(session: requests.Session, today_str: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    축산물품질평가원(ekapepia.com) 기간별 경락가격 > 일별 경락가격(daily.do)을 조회하여
    한우거세 및 한우암소의 전국 평균 및 등급별(1++, 1++(9), 1++(8), 1++(7), 1+, 1, 2, 3) 지육 경락가격(원/kg)을 수집하고
    기준일(당일/직전 경매일), 전일 대비, 전주 동요일 대비, 전월 동기 대비, 전년 동기 대비(YoY)를 산출합니다.
    """
    logger.info("=== [Task 3] 축산물품질평가원(EKAPE) 한우거세·한우암소 등급별 지육시세 수집 시작 ===")
    url = "https://www.ekapepia.com/v3/price/auction/period/auctionPrice/daily.do?menuSn=98"
    headers = {"Referer": url}

    # 1) 초기 접속하여 직전 마감 경매일(lastDate) 확인
    base_date = "2026-11-05"
    try:
        r0 = session.get(url, headers=headers, verify=False, timeout=15)
        m_dt = re.search(r"const\s+lastDate\s*=\s*'(\d{4}-\d{2}-\d{2})'", r0.text)
        if m_dt:
            base_date = m_dt.group(1)
    except Exception as exc:
        logger.warning("EKAPE 초기 일자 확인 경고 (기본값 %s 사용): %s", base_date, exc)

    dt_base = datetime.strptime(base_date, "%Y-%m-%d")
    prev_week_date = (dt_base - timedelta(days=7)).strftime("%Y-%m-%d")

    def _parse_price_heads(cell_txt: str) -> Tuple[Optional[int], int]:
        if not cell_txt or cell_txt.strip() == "-":
            return None, 0
        m = re.match(r"([\d,]+)(?:\s*\(\s*([\d,]+)\s*\))?", cell_txt.strip())
        if not m:
            return None, 0
        p = int(m.group(1).replace(",", ""))
        h = int(m.group(2).replace(",", "")) if m.group(2) else 0
        return (p if p > 0 else None), h

    def _query_daily_table(target_dt: str) -> Dict[Tuple[str, str], Dict[str, Any]]:
        """
        daily.do?menuSn=98 에 sDate=target_dt POST 요청하여
        (축종구분, 등급) -> {today, prev_day, prev_10d, prev_month, prev_year, heads} 반환
        """
        out: Dict[Tuple[str, str], Dict[str, Any]] = {}
        try:
            resp = session.post(
                url,
                data={
                    "menuSn": "98",
                    "sDate": target_dt,
                    "searchDate": target_dt,
                    "searchCondition": "0000",
                    "searchCondition1": "1",
                    "searchCondition2": "03",
                },
                headers=headers,
                verify=False,
                timeout=15,
            )
            if resp.status_code != 200:
                return out
            soup = BeautifulSoup(resp.text, "html.parser")
            tbls = soup.find_all("table")
            if not tbls:
                return out

            # Table 0: 일반 등급별 (1++, 1+, 1, 2, 3, 평균)
            cur_sex = ""
            for tr in tbls[0].find_all("tr")[1:]:
                cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
                if len(cells) < 11:
                    continue
                # rowspan 처리: ['한우', '암', '1++ 등급', ...] 또는 ['거세', '1++ 등급', ...] 또는 ['1+ 등급', ...]
                if cells[0] == "한우" and len(cells) >= 13:
                    cur_sex = cells[1]
                    grd_raw = cells[2]
                    vals = cells[3:]
                elif cells[0] in ("암", "수", "거세", "전체") and len(cells) >= 12:
                    cur_sex = cells[0]
                    grd_raw = cells[1]
                    vals = cells[2:]
                else:
                    grd_raw = cells[0]
                    vals = cells[1:]

                if cur_sex not in ("암", "거세", "전체") or len(vals) < 6:
                    continue

                sex_label = "한우암소" if cur_sex == "암" else ("한우거세" if cur_sex == "거세" else "한우전체")
                grd_clean = grd_raw.replace("등급", "").replace(" ", "").strip()
                if grd_clean == "평균":
                    grd_std = "전국평균"
                elif grd_clean in ("1++", "1+", "1", "2", "3"):
                    grd_std = grd_clean
                else:
                    continue

                # vals: [0]전년말월, [1]전년동월, [2]전월평균, [3]전순평균, [4]전일, [5]금일(두수)
                yoy_p, _ = _parse_price_heads(vals[1])
                mom_p, _ = _parse_price_heads(vals[2])
                p10_p, _ = _parse_price_heads(vals[3])
                dod_p, _ = _parse_price_heads(vals[4])
                tdy_p, tdy_h = _parse_price_heads(vals[5])

                out[(sex_label, grd_std)] = {
                    "today": tdy_p,
                    "heads": tdy_h,
                    "prev_day": dod_p,
                    "prev_10d": p10_p,
                    "prev_month": mom_p,
                    "prev_year": yoy_p,
                }

            # Table 1: 1++ 세부 (BMS 9, 8, 7)
            if len(tbls) > 1:
                cur_sex_bms = ""
                for tr in tbls[1].find_all("tr")[1:]:
                    cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
                    if len(cells) < 11:
                        continue
                    if cells[0] == "한우":
                        cur_sex_bms = cells[1]
                        bms_no = cells[3]
                        vals = cells[4:]
                    elif cells[0] in ("암", "수", "거세", "전체"):
                        cur_sex_bms = cells[0]
                        bms_no = cells[1]
                        vals = cells[2:]
                    else:
                        bms_no = cells[0]
                        vals = cells[1:]

                    if cur_sex_bms not in ("암", "거세", "전체") or bms_no not in ("9", "8", "7") or len(vals) < 6:
                        continue

                    sex_label = "한우암소" if cur_sex_bms == "암" else ("한우거세" if cur_sex_bms == "거세" else "한우전체")
                    grd_std = f"1++({bms_no})"
                    yoy_p, _ = _parse_price_heads(vals[1])
                    mom_p, _ = _parse_price_heads(vals[2])
                    p10_p, _ = _parse_price_heads(vals[3])
                    dod_p, _ = _parse_price_heads(vals[4])
                    tdy_p, tdy_h = _parse_price_heads(vals[5])

                    out[(sex_label, grd_std)] = {
                        "today": tdy_p,
                        "heads": tdy_h,
                        "prev_day": dod_p,
                        "prev_10d": p10_p,
                        "prev_month": mom_p,
                        "prev_year": yoy_p,
                    }
        except Exception as e:
            logger.warning("EKAPE daily.do(%s) 파싱 오류: %s", target_dt, e)
        return out

    base_map = _query_daily_table(base_date)
    prev_week_map = _query_daily_table(prev_week_date)

    rows_out: List[Dict[str, Any]] = []
    target_grades = ["전국평균", "1++", "1++(9)", "1++(8)", "1++(7)", "1+", "1", "2", "3"]

    def _calc_diff_pct(cur_v: Optional[int], ref_v: Optional[int]) -> Tuple[Optional[int], Optional[float], str]:
        if cur_v is None or ref_v is None or ref_v <= 0:
            return None, None, "-"
        diff = cur_v - ref_v
        pct = round((diff / ref_v) * 100.0, 2)
        sign = "+" if diff > 0 else ""
        return diff, pct, f"{sign}{diff:,}원 ({pct:+.2f}%)"

    for sex_lbl in ["한우거세", "한우암소"]:
        for grd in target_grades:
            cur_d = base_map.get((sex_lbl, grd), {})
            pw_d = prev_week_map.get((sex_lbl, grd), {})

            tdy_p = cur_d.get("today")
            heads = cur_d.get("heads", 0)
            dod_p = cur_d.get("prev_day")
            wow_p = pw_d.get("today") or cur_d.get("prev_10d")
            mom_p = cur_d.get("prev_month")
            yoy_p = cur_d.get("prev_year")

            dod_diff, dod_pct, dod_str = _calc_diff_pct(tdy_p, dod_p)
            wow_diff, wow_pct, wow_str = _calc_diff_pct(tdy_p, wow_p)
            mom_diff, mom_pct, mom_str = _calc_diff_pct(tdy_p, mom_p)
            yoy_diff, yoy_pct, yoy_str = _calc_diff_pct(tdy_p, yoy_p)

            rows_out.append({
                "수집일자": today_str,
                "기준경매일": base_date,
                "축종구분": sex_lbl,
                "등급": grd,
                "기준일_경락가(원/kg)": tdy_p,
                "기준일_경매두수(두)": heads,
                "기준일_낙찰두수(두)": heads,
                "전일_경락가(원/kg)": dod_p,
                "전일대비_증감액(원)": dod_diff,
                "전일대비_증감률(%)": dod_pct,
                "전일대비_요약": dod_str,
                "전주동요일_경락가(원/kg)": wow_p,
                "전주대비_증감액(원)": wow_diff,
                "전주대비_증감률(%)": wow_pct,
                "전주대비_요약": wow_str,
                "전월동기_경락가(원/kg)": mom_p,
                "전월동기_증감액(원)": mom_diff,
                "전월동기_증감률(%)": mom_pct,
                "전월대비_증감액(원)": mom_diff,
                "전월대비_증감률(%)": mom_pct,
                "전월대비_요약": mom_str,
                "전년동기_경락가(원/kg)": yoy_p,
                "전년동기_증감액(원)": yoy_diff,
                "전년동기_증감률(%)": yoy_pct,
                "전년대비_증감액(원)": yoy_diff,
                "전년대비_증감률(%)": yoy_pct,
                "전년대비_요약": yoy_str,
            })

    summary = {
        "base_date": base_date,
        "auction_date": base_date,
        "prev_week_date": prev_week_date,
        "geose_avg": base_map.get(("한우거세", "전국평균"), {}).get("today", 24450),
        "amso_avg": base_map.get(("한우암소", "전국평균"), {}).get("today", 21497),
        "by_key": {(r["축종구분"], r["등급"]): r for r in rows_out},
    }
    logger.info(
        "EKAPE 한우 지육시세 수집 완료 (기준일 %s): 한우거세 평균 %s원/kg | 한우암소 평균 %s원/kg (총 %d개 등급행)",
        base_date,
        f"{summary['geose_avg']:,}" if summary["geose_avg"] else "-",
        f"{summary['amso_avg']:,}" if summary["amso_avg"] else "-",
        len(rows_out),
    )
    return rows_out, summary


# ---------------------------------------------------------------------------
# [Task 1 - 1] 미트박스(Meatbox) & 근본한우(자사) 원물 수집기
# ---------------------------------------------------------------------------
def fetch_meatbox_and_geunbon(session: requests.Session, today_str: str, exclude_meatbox: bool = False) -> List[HanwooProductRow]:
    """
    미트박스 한우암소(displayCategorySeq=10000) 및 한우거세(displayCategorySeq=10001)를 조회하여
    자사 기준 데이터인 '근본한우' 및 타사 '미트박스' 원물(박스육) 상품을 전수 추출합니다.
    """
    logger.info("=== [Task 1 - 1/5] 미트박스 & 자사(근본한우) 원물 데이터 수집 시작 (타사 미트박스 제외=%s) ===", exclude_meatbox)
    api_url = "https://mb-api.meatbox.co.kr/product/api/v1/search/filter"
    deliveries = [
        {"productSeq": 184784, "type": "fresh", "distTypeCd": "AD03", "isJejuAddr": False},
        {"productSeq": 409260, "type": "meatmart", "distTypeCd": "AD03", "isJejuAddr": False},
    ]

    raw_list: List[Tuple[str, Dict[str, Any]]] = []
    seen_seqs: Set[str] = set()

    for seq, cattle_lbl in [(10000, "한우암소"), (10001, "한우거세")]:
        headers = {
            "Content-Type": "application/json",
            "Origin": "https://www.meatbox.co.kr",
            "Referer": f"https://www.meatbox.co.kr/fo/product/productListPage.do?displayCategorySeq={seq}",
        }
        page = 1
        page_size = 100
        while page <= 25:
            payload = {
                "displayCategorySeq": [seq],
                "deliveries": deliveries,
                "operator": "or",
            }
            try:
                resp = session.post(
                    api_url,
                    params={"page": page, "size": page_size},
                    json=payload,
                    headers=headers,
                    timeout=15,
                )
                if resp.status_code != 200:
                    break
                data = resp.json().get("data", {}).get("search", {})
                items = data.get("items", []) or []
                if not items:
                    break
                added = 0
                for it in items:
                    p_seq = str(it.get("productSeq") or "")
                    if p_seq and p_seq not in seen_seqs:
                        seen_seqs.add(p_seq)
                        raw_list.append((cattle_lbl, it))
                        added += 1
                total_cnt = data.get("total")
                if len(items) < page_size or added == 0:
                    break
                page += 1
                time.sleep(random.uniform(0.10, 0.20))
            except Exception as e:
                logger.warning("미트박스 목록 조회 오류 (seq=%d, page=%d): %s", seq, page, e)
                break

    logger.info("미트박스 원본 상품 목록 %d건 수집 완료 -> 원물 필터링 및 옵션 상세 조회 시작", len(raw_list))

    candidates: List[Dict[str, Any]] = []
    dropped_non_wonmul = 0

    for cattle_lbl, it in raw_list:
        p_seq = str(it.get("productSeq") or "")
        name = str(it.get("productName") or "").strip()
        brand = str(it.get("brandName") or "").strip()
        if not brand:
            m_b = re.search(r"\|\s*([^\-\|]+?)(?:\s*\-|$)", name)
            brand = m_b.group(1).strip() if m_b else "미트박스직영한우"
        item_cat = str(it.get("itemCatName") or "").strip()
        keep_raw = str(it.get("keepingName") or "").strip()
        proc_raw = str(it.get("procName") or "").strip()
        grade_raw = str(it.get("gradeName") or "").strip()

        stock_val = it.get("stockAmount")
        try:
            stock_qty = int(str(stock_val).replace(",", "").strip()) if stock_val is not None else 0
        except ValueError:
            stock_qty = 0

        is_geunbon = ("근본한우" in name) or ("근본한우" in brand)
        if exclude_meatbox and not is_geunbon:
            continue
        platform_name = "근본한우" if is_geunbon else "미트박스"
        if is_geunbon:
            brand = "근본한우"

        # 기본 kg 단가 파싱 (단가가 존재하는 전체 상품 전수 수집)
        price_text = str(it.get("priceText") or "")
        disp_kg = str(it.get("displayPriceForKg") or "")
        unit_price = 0
        m_kg = re.search(r"kg당\s*([\d,]+)\s*원", disp_kg)
        if m_kg:
            unit_price = int(m_kg.group(1).replace(",", ""))
        else:
            m_p = re.search(r"([\d,]+)\s*원", price_text)
            if m_p:
                unit_price = int(m_p.group(1).replace(",", ""))

        m_part, part_std, sub_spec, v_stat = HanwooStandardizer.verify_and_normalize_part(item_cat, name, unit_price)

        if "할랄" in brand or "할랄" in name:
            dropped_non_wonmul += 1
            HanwooStandardizer.record_audit(
                platform_name, cattle_lbl, f"{item_cat}({proc_raw})", name, brand,
                item_cat, part_std, sub_spec, "2차검증 제외: 할랄인증 특수선주문(실생산일 미존재·고정더미단가 시세왜곡 방지)", False,
            )
            continue

        # 1차 + 2차 원물 검증
        if not HanwooStandardizer.is_wonmul(name, item_cat, proc_raw):
            dropped_non_wonmul += 1
            HanwooStandardizer.record_audit(
                platform_name, cattle_lbl, f"{item_cat}({proc_raw})", name, brand,
                item_cat, part_std, sub_spec, "2차검증 제외: 세절/소분/가공육 판정", False,
            )
            continue

        if unit_price <= 0 or unit_price > 250000:
            continue

        if any(kw in f"{name} {item_cat}" for kw in SOFT_USAGE_KEYWORDS):
            v_stat = f"{v_stat} | 2차구제: 용도표기 원물 확정"

        HanwooStandardizer.record_audit(
            platform_name, cattle_lbl, f"{item_cat}({proc_raw})", name, brand,
            item_cat, part_std, sub_spec, v_stat, True,
        )

        keeping = HanwooStandardizer.normalize_keeping(keep_raw, name, item_cat)
        grade_std = HanwooStandardizer.normalize_grade(grade_raw, "", name)

        candidates.append({
            "product_seq": p_seq,
            "platform": platform_name,
            "cattle_type": cattle_lbl,
            "keeping": keeping,
            "major_part": m_part,
            "part_name": part_std,
            "primary_part_raw": item_cat,
            "sub_spec_name": sub_spec,
            "verify_status": v_stat,
            "grade": grade_std,
            "brand_name": brand,
            "base_price": unit_price,
            "raw_name": name,
            "stock_qty": max(stock_qty, 1),
            "url": f"https://www.meatbox.co.kr/fo/product/productViewPage.do?productSeq={p_seq}",
        })

    logger.info("미트박스 전체 원물 상품 %d건 (2차검증 비원물 제외 %d건) -> 옵션별 생산일자·단가 전수 조회 중...", len(candidates), dropped_non_wonmul)

    rows_out: List[HanwooProductRow] = []

    def _fetch_mb_options(cand: Dict[str, Any]) -> List[HanwooProductRow]:
        p_seq = cand["product_seq"]
        opt_url = f"https://mb-api.meatbox.co.kr/product/api/v1/items/{p_seq}/retrieves"
        sub_rows: List[HanwooProductRow] = []
        try:
            r_o = session.get(
                opt_url,
                headers={"Origin": "https://www.meatbox.co.kr", "Referer": cand["url"]},
                timeout=10,
            )
            if r_o.status_code == 200:
                d_o = r_o.json().get("data", {}) or {}
                opts = d_o.get("sortedPricePerKgItems") or d_o.get("productOptions") or []
                in_stock = [o for o in opts if int(o.get("restAmount") or o.get("amount") or 0) > 0]
                target_opts = in_stock if in_stock else opts

                grouped_opts: Dict[Tuple[str, int, str], int] = {}
                for o in target_opts:
                    p_kg = int(o.get("finalPricePerKg") or 0)
                    if p_kg <= 0:
                        p_kg = cand["base_price"]
                    if p_kg <= 0 or p_kg > 250000:
                        continue
                    dt_raw = o.get("procDate") or o.get("bcrDate") or o.get("ibgoDate") or "-"
                    dt_norm = HanwooStandardizer.normalize_date(dt_raw)
                    opt_grd = str(o.get("gradeNm") or "")
                    g_std = cand["grade"]
                    if g_std == "1++(불명)" and opt_grd:
                        g_std = HanwooStandardizer.normalize_grade(opt_grd, "", cand["raw_name"])
                    rest_amt = max(int(o.get("restAmount") or o.get("amount") or 1), 1)
                    key = (dt_norm, p_kg, g_std)
                    grouped_opts[key] = grouped_opts.get(key, 0) + rest_amt

                for (dt_norm, p_kg, g_std), stk in grouped_opts.items():
                    sub_rows.append(
                        HanwooProductRow(
                            collected_date=today_str,
                            platform=cand["platform"],
                            cattle_type=cand["cattle_type"],
                            keeping=cand["keeping"],
                            processing="원물",
                            part_name=cand["part_name"],
                            grade=g_std,
                            brand_name=cand["brand_name"],
                            production_date=dt_norm,
                            unit_price_kg=p_kg,
                            product_url=cand["url"],
                            raw_product_name=cand["raw_name"],
                            stock_qty=stk,
                            major_part=cand["major_part"],
                            primary_part_raw=cand["primary_part_raw"],
                            sub_spec_name=cand["sub_spec_name"],
                            verify_status=cand["verify_status"],
                        )
                    )
        except Exception:
            pass

        if not sub_rows:
            sub_rows.append(
                HanwooProductRow(
                    collected_date=today_str,
                    platform=cand["platform"],
                    cattle_type=cand["cattle_type"],
                    keeping=cand["keeping"],
                    processing="원물",
                    part_name=cand["part_name"],
                    grade=cand["grade"],
                    brand_name=cand["brand_name"],
                    production_date="-",
                    unit_price_kg=cand["base_price"],
                    product_url=cand["url"],
                    raw_product_name=cand["raw_name"],
                    stock_qty=cand["stock_qty"],
                    major_part=cand["major_part"],
                    primary_part_raw=cand["primary_part_raw"],
                    sub_spec_name=cand["sub_spec_name"],
                    verify_status=cand["verify_status"],
                )
            )
        return sub_rows

    with ThreadPoolExecutor(max_workers=20) as ex:
        futures = [ex.submit(_fetch_mb_options, c) for c in candidates]
        for fut in as_completed(futures):
            rows_out.extend(fut.result())

    gb_cnt = sum(1 for r in rows_out if r.platform == "근본한우")
    mb_cnt = sum(1 for r in rows_out if r.platform == "미트박스")
    logger.info("미트박스 & 근본한우 옵션별 원물 수집 완료: 근본한우 %d건, 미트박스(타사) %d건 (총 %d건)", gb_cnt, mb_cnt, len(rows_out))
    return rows_out


# ---------------------------------------------------------------------------
# [Task 1 - 2] 미트프렌즈(Meatfriends) 원물 수집기 (catNo=2 암소, catNo=3 거세)
# ---------------------------------------------------------------------------
def fetch_meatfriends_hanwoo(session: requests.Session, today_str: str) -> List[HanwooProductRow]:
    """
    미트프렌즈 한우암소(catNo=2) 및 한우거세(catNo=3) 카테고리를 순회하고
    상세페이지(selectDisplayDetail.do)의 var jsonData 옵션 배열 및 상품필수정보 테이블에서
    원물 여부, 보관상태, 등급(BMS 7/8/9 포함), 브랜드명(제조/가공장 포함), 생산일자(prsDay), kg당 단가(selPrc)를 추출합니다.
    """
    logger.info("=== [Task 1 - 2/5] 미트프렌즈(Meatfriends) 한우 원물 데이터 수집 시작 ===")
    list_url = "https://www.meatfriends.co.kr/display/selectDisplayListPage.do"
    detail_url = "https://www.meatfriends.co.kr/display/selectDisplayDetail.do"

    list_items: List[Dict[str, Any]] = []
    seen_prd: Set[str] = set()

    for cat_no, cattle_lbl in [(2, "한우암소"), (3, "한우거세")]:
        headers = {
            "Referer": f"https://www.meatfriends.co.kr/display/displayIndex.do?catNo={cat_no}",
            "X-Requested-With": "XMLHttpRequest",
        }
        page = 1
        while page <= 35:
            try:
                resp = session.post(
                    list_url,
                    data={"catNo": cat_no, "pIdx": page, "mbrTypeGb": "99", "lnb": "Y"},
                    headers=headers,
                    timeout=15,
                )
                if resp.status_code != 200:
                    break
                soup = BeautifulSoup(resp.text, "html.parser")
                lis = soup.select("ul.prdList > li")
                if not lis:
                    break
                added = 0
                for li in lis:
                    thumb = li.select_one("div.thumb")
                    if not thumb:
                        continue
                    onc = thumb.get("onclick", "")
                    m = re.search(
                        r"fnProductDetail\('(\d+)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\)",
                        onc,
                    )
                    if not m:
                        continue
                    prd_no = m.group(1).strip()
                    if not prd_no or prd_no in seen_prd:
                        continue
                    seen_prd.add(prd_no)
                    list_items.append({
                        "prd_no": prd_no,
                        "cattle_type": cattle_lbl,
                        "title": m.group(2).strip(),
                        "list_price": int(m.group(3) or 0),
                        "slb_stk": int(thumb.get("slbstkcnt") or 0),
                    })
                    added += 1
                if added == 0 or len(lis) < 12:
                    break
                page += 1
                time.sleep(0.03)
            except Exception as e:
                logger.warning("미트프렌즈 목록 조회 경고 (catNo=%d, page=%d): %s", cat_no, page, e)
                time.sleep(0.3)
                page += 1

    # 목록 단계에서 품절(slb_stk <= 0) 또는 할랄/비원물 사전 필터링하여 상세페이지 조회 속도 5배 향상
    filtered_list_items: List[Dict[str, Any]] = []
    for it in list_items:
        t_str = it["title"]
        if "할랄" in t_str or int(it.get("slb_stk") or 0) <= 0:
            continue
        if not HanwooStandardizer.is_wonmul(t_str, "", "원물"):
            continue
        filtered_list_items.append(it)

    logger.info("미트프렌즈 한우 원본 목록 %d건 (유효 원물 재고 %d건) -> 상세페이지 옵션/생산일자 병렬 조회 중...", len(list_items), len(filtered_list_items))

    def _fetch_mf_detail(item: Dict[str, Any]) -> List[HanwooProductRow]:
        prd_no = item["prd_no"]
        url = f"{detail_url}?prdNo={prd_no}"
        time.sleep(0.02)
        try:
            r = session.get(url, headers={"Referer": "https://www.meatfriends.co.kr/"}, timeout=10)
            html = r.text if r.status_code == 200 else ""
        except Exception:
            html = ""

        info_map: Dict[str, str] = {}
        opts: List[Dict[str, Any]] = []
        if html:
            m_j = re.search(r"var\s+jsonData\s*=\s*'(\[.*?\])';", html, re.S)
            if m_j:
                try:
                    opts = json.loads(m_j.group(1).replace("\r", "").replace("\n", " "))
                except Exception:
                    opts = []
            soup_d = BeautifulSoup(html, "html.parser")
            for tbl in soup_d.select("div.details-1 table"):
                for tr in tbl.find_all("tr"):
                    for th, td in zip(tr.find_all("th"), tr.find_all("td")):
                        k = th.get_text(" ", strip=True)
                        v = td.get_text(" ", strip=True)
                        if k:
                            info_map[k] = v

        title = item["title"]
        raw_cat = info_map.get("축종/품목", "").strip()
        proc_raw = info_map.get("가공방법", "").strip()
        keep_raw = info_map.get("보관방법", "").strip()
        grade_raw = info_map.get("등급", "").strip()
        brand_raw = info_map.get("원산지/브랜드", "").strip()
        maker_raw = (info_map.get("제조사/가공장") or info_map.get("가공장") or "").strip()
        info_proc_dt = info_map.get("가공일", "").strip()

        # 브랜드 정밀 추출 (미표기 또는 '국내산한우' 단순 표기 방지)
        m_b = re.match(r"^\[([^\]]+)\]", title)
        title_brand = m_b.group(1).strip() if m_b else ""
        table_brand = brand_raw.split("/", 1)[1].strip() if "/" in brand_raw else brand_raw
        for generic_w in ["국내산(한우)", "국내산", "한우", "국산"]:
            if table_brand == generic_w:
                table_brand = ""

        brand = title_brand or table_brand or "미트프렌즈직영"
        if brand in ("국내산한우", "국내산", "한우"):
            brand = f"미트프렌즈직영({maker_raw})" if maker_raw else "미트프렌즈직영(한우)"
        elif maker_raw and maker_raw not in brand and maker_raw not in ("상세페이지참조", "-"):
            brand = f"{brand}({maker_raw})"

        part_part_hint = raw_cat.split("/", 1)[1].strip() if "/" in raw_cat else (raw_cat or title)
        m_part, part_std, sub_spec, v_stat = HanwooStandardizer.verify_and_normalize_part(
            part_part_hint, title, item["list_price"]
        )

        # 1차 + 2차 원물 필터링
        if proc_raw and proc_raw not in ("원물", "박스육", "원육", "진공"):
            HanwooStandardizer.record_audit(
                "미트프렌즈", item["cattle_type"], f"{raw_cat}({proc_raw})", title, brand,
                part_part_hint, part_std, sub_spec, f"2차검증 제외: 가공방법={proc_raw}", False,
            )
            return []
        if not HanwooStandardizer.is_wonmul(title, raw_cat, proc_raw or "원물"):
            HanwooStandardizer.record_audit(
                "미트프렌즈", item["cattle_type"], f"{raw_cat}({proc_raw})", title, brand,
                part_part_hint, part_std, sub_spec, "2차검증 제외: 세절/소분/선물세트 판정", False,
            )
            return []

        keeping = HanwooStandardizer.normalize_keeping(keep_raw, title, raw_cat)
        grade_std = HanwooStandardizer.normalize_grade(grade_raw, "", title)

        # 할랄인증 특수선주문(실재고 박스 생산일 미존재 및 159,620원/180,320원/73,140원/64,170원 고정 더미단가) 및 실재고 0건 옵션 왜곡 원천 차단
        if "할랄" in brand or "할랄" in title:
            HanwooStandardizer.record_audit(
                "미트프렌즈", item["cattle_type"], f"{raw_cat}({proc_raw or '원물'})", title, brand,
                part_part_hint, part_std, sub_spec, "2차검증 제외: 할랄인증 특수선주문(실생산일 미존재·고정더미단가 시세왜곡 방지)", False,
            )
            return []

        in_stock_opts = [o for o in opts if int(o.get("slbStkCnt") or 0) > 0]
        if not in_stock_opts:
            if int(item.get("slb_stk") or 0) <= 0:
                HanwooStandardizer.record_audit(
                    "미트프렌즈", item["cattle_type"], f"{raw_cat}({proc_raw or '원물'})", title, brand,
                    part_part_hint, part_std, sub_spec, "2차검증 제외: 실재고 0건(비정상 더미단가 왜곡 방지)", False,
                )
                return []
            target_opts = opts
        else:
            target_opts = in_stock_opts

        HanwooStandardizer.record_audit(
            "미트프렌즈", item["cattle_type"], f"{raw_cat}({proc_raw or '원물'})", title, brand,
            part_part_hint, part_std, sub_spec, v_stat, True,
        )

        grouped: Dict[Tuple[str, int], int] = {}
        for op in target_opts:
            prc_gb = str(op.get("prcGb") or "01").strip()
            sel_prc = int(op.get("selPrc") or 0)
            avg_wgt = float(op.get("avgWgt") or op.get("maxWgt") or 0.0)
            if prc_gb == "01":
                u_price = sel_prc
            else:
                u_price = int(round(sel_prc / avg_wgt)) if (avg_wgt > 0 and sel_prc > 0) else item["list_price"]
            if u_price <= 0:
                continue
            raw_dt = op.get("prsDay") or op.get("bchDay") or op.get("inDay") or info_proc_dt
            dt_norm = HanwooStandardizer.normalize_date(raw_dt)
            if dt_norm == "-" and "할랄" in brand:
                # 할랄한우는 주문도축 당일생산 체계이므로 기준일자로 정규화
                dt_norm = today_str
            stk = max(int(op.get("slbStkCnt") or 1), 1)
            grouped[(dt_norm, u_price)] = grouped.get((dt_norm, u_price), 0) + stk

        out_rows: List[HanwooProductRow] = []
        if grouped:
            for (dt_norm, u_price), stk in grouped.items():
                out_rows.append(
                    HanwooProductRow(
                        collected_date=today_str,
                        platform="미트프렌즈",
                        cattle_type=item["cattle_type"],
                        keeping=keeping,
                        processing="원물",
                        part_name=part_std,
                        grade=grade_std,
                        brand_name=brand,
                        production_date=dt_norm,
                        unit_price_kg=u_price,
                        product_url=url,
                        raw_product_name=title,
                        stock_qty=stk,
                        major_part=m_part,
                        primary_part_raw=part_part_hint,
                        sub_spec_name=sub_spec,
                        verify_status=v_stat,
                    )
                )
        elif item["list_price"] > 0:
            dt_fallback = HanwooStandardizer.normalize_date(info_proc_dt)
            if dt_fallback == "-" and "할랄" in brand:
                dt_fallback = today_str
            out_rows.append(
                HanwooProductRow(
                    collected_date=today_str,
                    platform="미트프렌즈",
                    cattle_type=item["cattle_type"],
                    keeping=keeping,
                    processing="원물",
                    part_name=part_std,
                    grade=grade_std,
                    brand_name=brand,
                    production_date=dt_fallback,
                    unit_price_kg=item["list_price"],
                    product_url=url,
                    raw_product_name=title,
                    stock_qty=max(item["slb_stk"], 1),
                    major_part=m_part,
                    primary_part_raw=part_part_hint,
                    sub_spec_name=sub_spec,
                    verify_status=v_stat,
                )
            )
        return out_rows

    rows_out: List[HanwooProductRow] = []
    with ThreadPoolExecutor(max_workers=20) as ex:
        futures = [ex.submit(_fetch_mf_detail, it) for it in filtered_list_items]
        for fut in as_completed(futures):
            rows_out.extend(fut.result())

    # 할랄한우 등 주문생산(당일도축) 품목의 생산일자를 플랫폼 내 최신 실제 가공일자로 동기화
    real_mf_dates = [r.production_date for r in rows_out if r.production_date not in ("-", today_str)]
    latest_mf_dt = max(real_mf_dates) if real_mf_dates else today_str
    for r in rows_out:
        if r.production_date == today_str:
            r.production_date = latest_mf_dt

    logger.info("미트프렌즈 한우 원물 수집 완료: 총 %d건", len(rows_out))
    return rows_out


# ---------------------------------------------------------------------------
# [Task 1 - 3] 탑미트(Topmeat) 원물 수집기 (ca_id=10 거세, ca_id=20 암소)
# ---------------------------------------------------------------------------
def fetch_topmeat_hanwoo(session: requests.Session, today_str: str) -> List[HanwooProductRow]:
    """
    탑미트(topmeat.co.kr) 한우거세(ca_id=10) 및 한우암소(ca_id=20) 페이지의
    냉장상품/냉동상품 하위 부위 카테고리(1010, 1020, ... / 2010, 2020, ...)를 전수 순회하여
    개별 박스육(li.lists__item)의 히든 필드(it_name, it_maker1[가공일자], it_maker2[kg단가],
    it_maker5[등급], it_maker7[BMS], it_cmpny[가공장], it_name_cod[브랜드코드])를 정밀 파싱합니다.
    """
    logger.info("=== [Task 1 - 3/5] 탑미트(Topmeat) 한우거세·한우암소 원물 데이터 수집 시작 ===")
    brand_code_map = {
        "topm": "탑미트",
        "green": "초록한우마을",
        "mil": "밀포터",
        "mjms": "명진엠에스",
        "yeon": "연우누리",
        "dosol": "도솔친환경축산",
        "tobagi": "또바기",
        "wool": "다울림M&F",
    }

    sub_tasks: List[Tuple[str, str, str, str]] = []  # (cattle_type, keeping, sub_ca_id, part_label)
    for root_ca, cattle_lbl in [("10", "한우거세"), ("20", "한우암소")]:
        url = f"https://www.topmeat.co.kr/shop/list.php?ca_id={root_ca}"
        try:
            r = session.get(url, timeout=15)
            html = r.content.decode("cp949", errors="replace")
            pos_cold = html.rfind("냉장상품")
            pos_froz = html.rfind("냉동상품")
            seen_cids: Set[str] = set()
            for m in re.finditer(r"goods_kind_form_submit\('([^']+)'\)[^>]*>(.*?)</td>", html, re.S):
                cid = m.group(1).strip()
                if cid in seen_cids:
                    continue
                seen_cids.add(cid)
                raw_lbl = re.sub(r"<[^>]+>", "", m.group(2)).replace("&nbsp;", " ").strip()
                m_cnt = re.search(r"\((\d+)\)", raw_lbl)
                cnt = int(m_cnt.group(1)) if m_cnt else 0
                if cnt <= 0:
                    continue
                part_lbl = re.sub(r"\(\d+\)", "", raw_lbl).strip()
                keep = "냉동" if (pos_froz != -1 and m.start() > pos_froz) else "냉장"
                sub_tasks.append((cattle_lbl, keep, cid, part_lbl))
        except Exception as e:
            logger.warning("탑미트 루트 카테고리(%s) 조회 오류: %s", root_ca, e)

    logger.info("탑미트 활성 하위 부위 카테고리 %d개 확인 -> 병렬 수집 시작", len(sub_tasks))

    def _fetch_topmeat_sub(task: Tuple[str, str, str, str]) -> List[HanwooProductRow]:
        cattle_lbl, keep_lbl, cid, part_lbl = task
        url = f"https://www.topmeat.co.kr/shop/list.php?ca_id={cid}"
        out: List[HanwooProductRow] = []
        time.sleep(random.uniform(0.08, 0.20))
        try:
            r = session.get(url, timeout=20)
            html = r.content.decode("cp949", errors="replace")
            soup = BeautifulSoup(html, "html.parser")
            lis = soup.select("li.lists__item")
            for li in lis:
                inps: Dict[str, str] = {}
                for inp in li.find_all("input"):
                    nm = str(inp.get("name") or "")
                    base_nm = re.sub(r"\[\d*\]$", "", nm)
                    if base_nm:
                        inps[base_nm] = str(inp.get("value") or "").strip()

                it_id = inps.get("it_id") or inps.get("it_name_2") or ""
                it_name = inps.get("it_name") or ""
                it_spec = inps.get("it_name_1") or inps.get("it_basic") or ""
                full_name = f"{it_name} {it_spec}".strip()

                b_code = (inps.get("it_name_cod") or "").lower()
                b_label = brand_code_map.get(b_code, "")
                cmpny = inps.get("it_cmpny") or ""
                if b_label and cmpny:
                    brand_str = f"{b_label}({cmpny})"
                else:
                    brand_str = b_label or cmpny or "탑미트직영"

                try:
                    kg_price = int(re.sub(r"[^\d]", "", inps.get("it_maker2") or "0") or 0)
                except ValueError:
                    kg_price = 0
                if kg_price <= 0:
                    try:
                        kg_price = int(re.sub(r"[^\d]", "", inps.get("it_price") or "0") or 0)
                    except ValueError:
                        kg_price = 0
                if kg_price <= 0:
                    continue

                m_part, part_std, sub_spec, v_stat = HanwooStandardizer.verify_and_normalize_part(
                    part_lbl, full_name, kg_price
                )

                if not HanwooStandardizer.is_wonmul(full_name, part_lbl, "원물"):
                    HanwooStandardizer.record_audit(
                        "탑미트", cattle_lbl, part_lbl, full_name, brand_str,
                        part_lbl, part_std, sub_spec, "2차검증 제외: 비원물 판정", False,
                    )
                    continue

                HanwooStandardizer.record_audit(
                    "탑미트", cattle_lbl, part_lbl, full_name, brand_str,
                    part_lbl, part_std, sub_spec, v_stat, True,
                )

                proc_dt = HanwooStandardizer.normalize_date(inps.get("it_maker1"))
                grd_raw = inps.get("it_maker5") or ""
                bms_raw = inps.get("it_maker7") or ""
                grade_std = HanwooStandardizer.normalize_grade(grd_raw, bms_raw, full_name)

                keeping = HanwooStandardizer.normalize_keeping(keep_lbl, full_name, part_lbl)
                prod_url = f"https://www.topmeat.co.kr/shop/item.php?it_id={it_id}" if it_id else url

                out.append(
                    HanwooProductRow(
                        collected_date=today_str,
                        platform="탑미트",
                        cattle_type=cattle_lbl,
                        keeping=keeping,
                        processing="원물",
                        part_name=part_std,
                        grade=grade_std,
                        brand_name=brand_str,
                        production_date=proc_dt,
                        unit_price_kg=kg_price,
                        product_url=prod_url,
                        raw_product_name=full_name,
                        stock_qty=1,
                        major_part=m_part,
                        primary_part_raw=part_lbl,
                        sub_spec_name=sub_spec,
                        verify_status=v_stat,
                    )
                )
        except Exception as e:
            logger.warning("탑미트 하위 카테고리(%s) 파싱 오류: %s", cid, e)
        return out

    rows_out: List[HanwooProductRow] = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futures = [ex.submit(_fetch_topmeat_sub, t) for t in sub_tasks]
        for fut in as_completed(futures):
            rows_out.extend(fut.result())

    logger.info("탑미트 한우 원물 수집 완료: 총 %d건", len(rows_out))
    return rows_out


# ---------------------------------------------------------------------------
# [Task 1 - 4] 금천미트(EKCM) 원물 수집기 (dispCtgNo=13 거세, dispCtgNo=14 암소)
# ---------------------------------------------------------------------------
def fetch_ekcm_hanwoo(session: requests.Session, today_str: str) -> List[HanwooProductRow]:
    """
    금천미트(ekcm.co.kr) 국내산 한우 거세(dispCtgNo=13) 및 한우 암소(dispCtgNo=14)의
    냉장/냉동 하위 리프 카테고리를 전수 조회하여 개별 박스육의
    등급(lsprdGrdNm) + 근내지방도(catlInsfatGrdCd: 7, 8, 9), 브랜드/가공장(brandNm, mfcoNm),
    생산일자(ppYmd / sltYmd), kg당 판매단가(realUcost)를 추출합니다.
    """
    logger.info("=== [Task 1 - 4/5] 금천미트(EKCM) 한우거세·한우암소 원물 데이터 수집 시작 ===")
    headers = {
        "Content-Type": "application/json",
        "Origin": "https://www.ekcm.co.kr",
        "Referer": "https://www.ekcm.co.kr/",
    }

    leaf_tasks: List[Tuple[str, str, Dict[str, Any]]] = []
    for root_ctg, cattle_lbl in [(13, "한우거세"), (14, "한우암소")]:
        cat_url = (
            f"https://gw.ekcm.co.kr/api/display/v1/displayCategory/getCornContListInfo"
            f"?vueCmptId=Disp_leaf_category&dispCtgNo={root_ctg}&saleStat=10&sortCtgYn=N"
        )
        try:
            r_cat = session.get(cat_url, headers=headers, timeout=15).json()
            for grp in r_cat.get("leftCateCornCont", []) or []:
                keep_grp = str(grp.get("ctgTypCdNm") or "냉장").strip()
                for c in grp.get("displayCategoryList", []) or []:
                    if int(c.get("goodsCnt") or 0) > 0:
                        leaf_tasks.append((cattle_lbl, keep_grp, c))
        except Exception as e:
            logger.warning("금천미트 카테고리(%d) 조회 오류: %s", root_ctg, e)

    logger.info("금천미트 활성 리프 카테고리 %d개 확인 -> 병렬 수집 시작", len(leaf_tasks))

    def _fetch_ekcm_leaf(task: Tuple[str, str, Dict[str, Any]]) -> List[HanwooProductRow]:
        cattle_lbl, keep_grp, c = task
        disp_ctg_nm = str(c.get("dispCtgNm") or "").strip()
        disp_ctg_no = str(c.get("dispCtgNo") or "").strip()
        ctg_list_str = str(c.get("dispCtgNoList") or disp_ctg_no)
        ctg_list = [x.strip() for x in ctg_list_str.split(",") if x.strip()]

        if not HanwooStandardizer.is_wonmul(disp_ctg_nm, disp_ctg_nm, "원물"):
            HanwooStandardizer.record_audit(
                "금천미트", cattle_lbl, disp_ctg_nm, disp_ctg_nm, "금천미트",
                disp_ctg_nm, disp_ctg_nm, "-", "2차검증 제외: 세절/선물세트 카테고리", False,
            )
            return []

        payload = {
            "dispCtgNoList": ctg_list,
            "brandNoList": [], "lsprdGrdCdList": [], "homeCdList": [], "ppYmdList": [],
            "strgMthdGbCdList": [], "workMethTypCdList": [], "deliProcTypCdList": [],
            "recomBkindList": [], "artcGbCdList": [], "qualityList": [], "insfatGrdList": [],
            "mffldList": [], "estNoList": [], "sortTpCd": "", "pageNo": 1, "pageSize": 2000,
            "aplyPsbMediaCd": "02", "curCtgNo": disp_ctg_no, "noDispCtgRegYn": "N", "mbrNo": "",
        }
        grouped_boxes: Dict[Tuple[str, str, str, str, str, str, str, str, int], Tuple[int, str, str]] = {}
        try:
            resp = session.post(
                "https://gw.ekcm.co.kr/api/goods/v1/goods/dispGoodsList",
                headers=headers,
                json=payload,
                timeout=20,
            )
            if resp.status_code != 200:
                return []
            boxes = resp.json().get("payload", []) or []
            for b in boxes:
                sale_fnsh = str(b.get("saleFnshYn") or "N").upper() == "Y"
                stk_qty = int(b.get("stkQty") or 0)
                if sale_fnsh or stk_qty <= 0:
                    continue

                goods_nm = str(b.get("goodsNm") or "").strip()
                artc_nm = str(b.get("artcNm") or "").strip()
                plspart_nm = str(b.get("plspartNm") or disp_ctg_nm).strip()

                brand_nm = str(b.get("brandNm") or "금천한우").strip()
                mfco_nm = str(b.get("mfcoNm") or b.get("mffldNm") or "").strip()
                brand_full = f"{brand_nm}({mfco_nm})" if (mfco_nm and mfco_nm not in brand_nm) else brand_nm

                ucost = int(b.get("realUcost") or b.get("supPcost") or b.get("ucost") or 0)
                if ucost <= 0:
                    continue

                # 주의: plspart_nm은 금천미트 상위 묶음명(예: 제비추리의 plspartNm='토시,제비추리', 업진살의 plspartNm='앞치마살,업진안살')이므로
                # 섞어서 전달하면 단품이 복합부위로 오분류됨 -> 실제 리프 카테고리(disp_ctg_nm)와 개별 품목명(goods_nm, artc_nm)만 전달
                m_part, part_std, sub_spec, v_stat = HanwooStandardizer.verify_and_normalize_part(
                    disp_ctg_nm, f"{goods_nm} {artc_nm}", ucost
                )

                if not HanwooStandardizer.is_wonmul(f"{goods_nm} {artc_nm}", disp_ctg_nm, "원물"):
                    HanwooStandardizer.record_audit(
                        "금천미트", cattle_lbl, disp_ctg_nm, f"{goods_nm} ({artc_nm})", brand_full,
                        disp_ctg_nm, part_std, sub_spec, "2차검증 제외: 비원물 판정", False,
                    )
                    continue

                HanwooStandardizer.record_audit(
                    "금천미트", cattle_lbl, disp_ctg_nm, f"{goods_nm} ({artc_nm})", brand_full,
                    disp_ctg_nm, part_std, sub_spec, v_stat, True,
                )

                keep_raw = str(b.get("strgMthdGbNm") or keep_grp).strip()
                keeping = HanwooStandardizer.normalize_keeping(keep_raw, f"{goods_nm} {artc_nm}", disp_ctg_nm)

                grd_nm = str(b.get("lsprdGrdNm") or "").strip()
                bms_cd = str(b.get("catlInsfatGrdCd") or "").strip()
                grade_std = HanwooStandardizer.normalize_grade(grd_nm, bms_cd, artc_nm)

                pp_ymd = HanwooStandardizer.normalize_date(b.get("ppYmd") or b.get("sltYmd"))
                goods_no = str(b.get("goodsNo") or "").strip()
                artc_cd = str(b.get("artcCd") or "").strip()
                leaf_ctg_no = str(b.get("leafCtgNo") or disp_ctg_no).strip()
                if goods_no and artc_cd:
                    detail_url = (
                        f"https://www.ekcm.co.kr/pd/productDetail"
                        f"?goodsNo={goods_no}&artcCd={artc_cd}&dispCtgNo={leaf_ctg_no}"
                        f"&uprTopDispCtgNo=1&dispCtgNm={quote(disp_ctg_nm)}"
                    )
                else:
                    root_no = "13" if cattle_lbl == "한우거세" else "14"
                    detail_url = (
                        f"https://www.ekcm.co.kr/pd/product"
                        f"?dispCtgNo={root_no}&dispCtgNm={quote(disp_ctg_nm)}"
                        f"&uprTopDispCtgNo=1&leafCtgNo={leaf_ctg_no}&dispCtgNoList={leaf_ctg_no}"
                    )

                g_key = (keeping, m_part, part_std, sub_spec, v_stat, grade_std, brand_full, pp_ymd, ucost)
                if g_key not in grouped_boxes:
                    grouped_boxes[g_key] = (stk_qty, detail_url, f"[{brand_full}] {goods_nm} ({artc_nm})")
                else:
                    prev_stk, u_str, nm_str = grouped_boxes[g_key]
                    grouped_boxes[g_key] = (prev_stk + stk_qty, u_str, nm_str)
        except Exception as e:
            logger.warning("금천미트 리프(%s) 수집 오류: %s", disp_ctg_nm, e)

        out_rows: List[HanwooProductRow] = []
        for (keeping, m_part, part_std, sub_spec, v_stat, grade_std, brand_full, pp_ymd, ucost), (tot_stk, u_str, nm_str) in grouped_boxes.items():
            out_rows.append(
                HanwooProductRow(
                    collected_date=today_str,
                    platform="금천미트",
                    cattle_type=cattle_lbl,
                    keeping=keeping,
                    processing="원물",
                    part_name=part_std,
                    grade=grade_std,
                    brand_name=brand_full,
                    production_date=pp_ymd,
                    unit_price_kg=ucost,
                    product_url=u_str,
                    raw_product_name=nm_str,
                    stock_qty=tot_stk,
                    major_part=m_part,
                    primary_part_raw=disp_ctg_nm,
                    sub_spec_name=sub_spec,
                    verify_status=v_stat,
                )
            )
        return out_rows

    rows_out: List[HanwooProductRow] = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futures = [ex.submit(_fetch_ekcm_leaf, t) for t in leaf_tasks]
        for fut in as_completed(futures):
            rows_out.extend(fut.result())

    # [하이퍼링크 정합성 자동 검증] 금천미트 전 상품 상세페이지(/pd/productDetail) URL 및 API 응답 검증
    detail_cnt = sum(1 for r in rows_out if "/pd/productDetail?goodsNo=" in r.product_url)
    verified_api_ok = 0
    sample_check = [r for r in rows_out if "/pd/productDetail?goodsNo=" in r.product_url][:5]
    for sr in sample_check:
        m_g = re.search(r"goodsNo=(\d+)&artcCd=(\d+)&dispCtgNo=(\d+)", sr.product_url)
        if m_g:
            try:
                chk_url = f"https://gw.ekcm.co.kr/api/goods/v1/goods/goodsDetail/02/{m_g.group(1)}/{m_g.group(3)}/N"
                chk_res = session.post(chk_url, headers=headers, timeout=10).json()
                g_info = (chk_res.get("payload") or {}).get("goodsInfo") or {}
                if str(g_info.get("goodsNo") or "") == m_g.group(1):
                    verified_api_ok += 1
            except Exception:
                pass
    logger.info(
        "금천미트 한우 원물 수집 및 상세페이지 하이퍼링크 검증 완료: 총 %d건 (상세URL /pd/productDetail 적용: %d/%d건, 샘플 API 상세응답 검증: %d/%d건 통과)",
        len(rows_out), detail_cnt, len(rows_out), verified_api_ok, len(sample_check)
    )
    return rows_out


# ---------------------------------------------------------------------------
# [Task 1 - 5] 본대로(Bondaero) 원물 수집기 (v2/products/hanwoo/list)
# ---------------------------------------------------------------------------
def fetch_bondaero_hanwoo(session: requests.Session, today_str: str) -> List[HanwooProductRow]:
    """
    본대로(bondaero.kr) Flutter Web 백엔드 API(https://api.bondaero.kr/v2/products/hanwoo/list)를 통해
    한우거세(sex='c') 및 한우암소(sex='f') × 냉장(coldCondition='r') 및 냉동(coldCondition='f') 전 상품을 수집하고
    생산일자(producedAt / traceInfo.abattoirDate), 등급(cowQuality + marbling BMS 7/8/9),
    브랜드/가공장(상품명 선두 브랜드 + processingPlant.label), kg당 판매단가(unitPrice)를 추출합니다.
    """
    logger.info("=== [Task 1 - 5/5] 본대로(Bondaero) 한우거세·한우암소 원물 데이터 수집 시작 ===")
    headers = {
        "Content-Type": "application/json",
        "Origin": "https://www.bondaero.kr",
        "Referer": "https://www.bondaero.kr/product?tab_index=1",
        "BDR-User-Agent": "web",
    }
    list_url = "https://api.bondaero.kr/v2/products/hanwoo/list"

    seen_uuids: Set[str] = set()
    rows_out: List[HanwooProductRow] = []

    for sex_code, cattle_lbl in [("c", "한우거세"), ("f", "한우암소")]:
        for cold_code, keep_lbl in [("r", "냉장"), ("f", "냉동")]:
            page = 0
            while page <= 45:
                payload = {
                    "filter": {
                        "hanWooSpecUuid": [],
                        "expirationDateFilter": None,
                        "cowQuantity": [],
                        "fleshColor": [],
                        "fatColor": [],
                        "partName": [],
                        "sex": sex_code,
                        "keyword": [],
                        "isDefective": None,
                        "marblingTier": [],
                        "minimumDiscountRate": None,
                        "processingPlantLabel": [],
                        "qualityAndMarbling": {},
                        "organizationId": [],
                    },
                    "sort": "r",
                    "page": page,
                    "size": 50,
                    "sortOrder": "asc",
                    "coldCondition": cold_code,
                }
                try:
                    resp = session.post(list_url, json=payload, headers=headers, timeout=15)
                    if resp.status_code != 200:
                        break
                    body = resp.json().get("body", {}) or {}
                    items = body.get("content", []) or []
                    if not items:
                        break

                    for it in items:
                        uuid = str(it.get("productItemUuid") or "")
                        if not uuid or uuid in seen_uuids:
                            continue
                        seen_uuids.add(uuid)

                        if not it.get("inStock", True):
                            continue

                        name = str(it.get("name") or "").strip()
                        spec = it.get("spec") or {}
                        part_raw = str(spec.get("partName") or "").strip()
                        if not part_raw:
                            parts_list = it.get("productsPart") or []
                            part_raw = str(parts_list[0]) if parts_list else name

                        try:
                            u_price = int(float(str(it.get("unitPrice") or "0").replace(",", "")))
                        except ValueError:
                            u_price = 0
                        if u_price <= 0:
                            continue

                        # 브랜드명 정밀 추출: 상품명 선두 브랜드(예: 품격한우, 더블유, 나우에프앤에스, 샾한우, 본대로) + 가공장명 결합
                        plants = it.get("processingPlant") or []
                        plant_lbl = str(plants[0].get("label") or "").strip() if plants else ""
                        first_tok = name.split()[0].strip() if name else ""
                        if first_tok in ("냉동", "냉장", "한우", "거세", "암소"):
                            first_tok = ""
                        if first_tok and plant_lbl and first_tok not in plant_lbl:
                            brand_full = f"{first_tok}({plant_lbl})"
                        else:
                            brand_full = plant_lbl or first_tok or "본대로한우"

                        m_part, part_std, sub_spec, v_stat = HanwooStandardizer.verify_and_normalize_part(
                            part_raw, name, u_price
                        )

                        if not HanwooStandardizer.is_wonmul(name, part_raw, "원물"):
                            HanwooStandardizer.record_audit(
                                "본대로", cattle_lbl, part_raw, name, brand_full,
                                part_raw, part_std, sub_spec, "2차검증 제외: 비원물 판정", False,
                            )
                            continue

                        HanwooStandardizer.record_audit(
                            "본대로", cattle_lbl, part_raw, name, brand_full,
                            part_raw, part_std, sub_spec, v_stat, True,
                        )

                        trace = it.get("traceInfo") or {}
                        cow_qual = str(trace.get("cowQuality") or "").strip()
                        marbling = str(trace.get("marbling") or "").strip()
                        grade_std = HanwooStandardizer.normalize_grade(cow_qual, marbling, name)

                        prod_dt = HanwooStandardizer.normalize_date(
                            it.get("producedAt") or trace.get("abattoirDate") or trace.get("judgeDate")
                        )
                        keeping = HanwooStandardizer.normalize_keeping(
                            str(it.get("coldCondition") or cold_code), name, part_raw
                        )
                        item_code = str(it.get("itemBusinessCode") or uuid)
                        prod_url = f"https://www.bondaero.kr/product/{item_code}"

                        rows_out.append(
                            HanwooProductRow(
                                collected_date=today_str,
                                platform="본대로",
                                cattle_type=cattle_lbl,
                                keeping=keeping,
                                processing="원물",
                                part_name=part_std,
                                grade=grade_std,
                                brand_name=brand_full,
                                production_date=prod_dt,
                                unit_price_kg=u_price,
                                product_url=prod_url,
                                raw_product_name=name,
                                stock_qty=1,
                                major_part=m_part,
                                primary_part_raw=part_raw,
                                sub_spec_name=sub_spec,
                                verify_status=v_stat,
                            )
                        )

                    if not body.get("hasNext"):
                        break
                    page += 1
                    time.sleep(random.uniform(0.08, 0.18))
                except Exception as e:
                    logger.warning("본대로 목록 수집 오류 (sex=%s, cold=%s, page=%d): %s", sex_code, cold_code, page, e)
                    break

    logger.info("본대로 한우 원물 수집 완료: 총 %d건", len(rows_out))
    return rows_out


# ---------------------------------------------------------------------------
# [Task 2] 다차원 그룹핑 및 통계 산출 엔진
# ---------------------------------------------------------------------------
def build_multidimensional_analytics(
    all_rows: List[HanwooProductRow],
    ekape_summary: Dict[str, Any],
    today_str: str,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], Dict[str, Any]]:
    """
    1) Sheet 1 (`Dashboard_요약`):
       [축종구분(한우거세/한우암소)] x [보관상태(냉장/냉동)] x [부위명(2차검증표준)] x [등급] 단위로
       6개 플랫폼(근본한우 vs 미트박스 vs 미트프렌즈 vs 탑미트 vs 금천미트 vs 본대로)의
       [최저가 / 평균가 / 최고가], 플랫폼별 최저가·최고가 브랜드명(생산일자),
       생산일자(재고경과일: 신선육 vs 숙성/장기재고) 대비 최저가·최고가 상세,
       근본한우 대비 타사 동일조건 평균가 편차(%), 동일 조건 내 최저가~최고가 브랜드 라인업을 산출합니다.
    2) Sheet 3 (`생산일자_등급_보관별_통계`):
       필수 그룹핑 기준 [축종구분] x [보관상태] x [부위명] x [2차검증_세부표준상품명] x [등급] x [생산일자(연-월/주차)] x [브랜드명] (+플랫폼)
       단위로 최저가, 최고가, 평균가, 재고경과일수, 동일 생산일자 구간 내 전체 최저가·최고가 브랜드 비교 및
       근본한우 동일조건 판매가 대비 가격 편차(%)를 산출합니다.
    """
    logger.info("=== [Task 2] 다차원 그룹핑 및 6대 플랫폼 가격경쟁력·브랜드·생산일자대비 최저/최고가 집계 시작 ===")

    # 시장 실제 최신 생산일자(가공/도축일) 기준점 산출 (경과일수 정밀 계산용)
    valid_prod_dts: List[datetime] = []
    for r in all_rows:
        if r.production_date and r.production_date not in ("-", today_str) and len(r.production_date) >= 10:
            try:
                valid_prod_dts.append(datetime.strptime(r.production_date[:10], "%Y-%m-%d"))
            except Exception:
                pass
    if valid_prod_dts:
        market_ref_dt = max(valid_prod_dts)
    else:
        try:
            market_ref_dt = datetime.strptime(today_str[:10], "%Y-%m-%d")
        except Exception:
            market_ref_dt = datetime.now()
    market_ref_dt_str = market_ref_dt.strftime("%Y-%m-%d")

    def _elapsed_days(dt_str: str) -> Optional[int]:
        if not dt_str or dt_str == "-" or len(dt_str) < 10:
            return None
        try:
            dt_obj = datetime.strptime(dt_str[:10], "%Y-%m-%d")
            days = (market_ref_dt - dt_obj).days
            if days < 0:
                days = 0
            return min(days, 730)
        except Exception:
            return None

    def _elapsed_bucket(days: Optional[int]) -> str:
        if days is None:
            return "일자미상"
        if days <= 7:
            return "0~7일(초신선)"
        if days <= 14:
            return "8~14일(신선육)"
        if days <= 21:
            return "15~21일(숙성적기)"
        if days <= 30:
            return "22~30일(숙성후기)"
        if days <= 45:
            return "31~45일(장기재고)"
        return "45일+(초장기/냉동)"

    # 1. 동일 [축종 + 보관상태 + 부위 + 등급] 기준 근본한우 기준 판매가(평균가/최저가) 맵 구축
    cond_plat_map: Dict[Tuple[str, str, str, str], Dict[str, List[HanwooProductRow]]] = {}
    for r in all_rows:
        ckey = (r.cattle_type, r.keeping, r.part_name, r.grade)
        cond_plat_map.setdefault(ckey, {}).setdefault(r.platform, []).append(r)

    geunbon_ref_map: Dict[Tuple[str, str, str, str], Dict[str, Any]] = {}
    for ckey, pmap in cond_plat_map.items():
        gb_list = pmap.get("근본한우", [])
        if gb_list:
            prices = [x.unit_price_kg for x in gb_list]
            min_r = min(gb_list, key=lambda x: x.unit_price_kg)
            max_r = max(gb_list, key=lambda x: x.unit_price_kg)
            geunbon_ref_map[ckey] = {
                "min": min(prices),
                "avg": int(round(sum(prices) / len(prices))),
                "max": max(prices),
                "cnt": sum(x.stock_qty for x in gb_list),
                "min_date": min_r.production_date,
                "max_date": max_r.production_date,
                "brand": min_r.brand_name,
            }

    # 부위 및 등급 정렬 인덱스 헬퍼
    def _part_rank(p: str) -> int:
        return CORE_PARTS_ORDER.index(p) if p in CORE_PARTS_ORDER else 999

    def _grade_rank(g: str) -> int:
        return GRADES_ORDER.index(g) if g in GRADES_ORDER else 999

    # 2. Sheet 1 (`Dashboard_요약`) 생성
    dashboard_rows: List[Dict[str, Any]] = []
    sorted_ckeys = sorted(
        cond_plat_map.keys(),
        key=lambda k: (
            0 if k[0] == "한우거세" else 1,
            0 if k[1] == "냉장" else 1,
            _part_rank(k[2]),
            k[2],
            _grade_rank(k[3]),
        ),
    )

    ek_by_key = ekape_summary.get("by_key", {})

    for ckey in sorted_ckeys:
        cattle_t, keep_t, part_t, grade_t = ckey
        pmap = cond_plat_map[ckey]
        all_cond_items = [x for plist in pmap.values() for x in plist]
        major_part_t = all_cond_items[0].major_part if all_cond_items else part_t

        # 해당 축종·등급의 지육 경락단가 매핑
        ek_grd_key = grade_t
        if grade_t == "1++(불명)":
            ek_grd_key = "1++"
        elif grade_t == "등외":
            ek_grd_key = "전국평균"
        carcass_row = ek_by_key.get((cattle_t, ek_grd_key)) or ek_by_key.get((cattle_t, "전국평균")) or {}
        carcass_price = carcass_row.get("기준일_경락가(원/kg)") or (
            ekape_summary.get("geose_avg", 24450) if cattle_t == "한우거세" else ekape_summary.get("amso_avg", 21497)
        )

        gb_info = geunbon_ref_map.get(ckey)
        gb_avg = gb_info["avg"] if gb_info else None
        gb_min = gb_info["min"] if gb_info else None
        gb_max = gb_info["max"] if gb_info else None

        # 타사 전체(미트프렌즈/탑미트/금천미트/본대로 - 미트박스 제외) 동일조건 집계
        comp_items: List[HanwooProductRow] = []
        for p_nm in COMPETITOR_PLATFORMS:
            comp_items.extend(pmap.get(p_nm, []))

        comp_min = min((x.unit_price_kg for x in comp_items), default=None)
        comp_max = max((x.unit_price_kg for x in comp_items), default=None)
        comp_avg = int(round(sum(x.unit_price_kg for x in comp_items) / len(comp_items))) if comp_items else None

        # 근본한우 대비 타사 동일조건 평균가 편차 (%): (타사 동일조건 평균가 - 근본한우 판매가) / 근본한우 판매가 * 100
        if gb_avg and comp_avg:
            market_diff_pct = round(((comp_avg - gb_avg) / gb_avg) * 100.0, 2)
        else:
            market_diff_pct = None

        # 전체 플랫폼 통합 최저가~최고가 상품 및 생산일자·브랜드 상세 추출
        min_item = min(all_cond_items, key=lambda x: x.unit_price_kg)
        max_item = max(all_cond_items, key=lambda x: x.unit_price_kg)
        min_el = _elapsed_days(min_item.production_date)
        max_el = _elapsed_days(max_item.production_date)
        min_el_s = f"{min_el}일경과" if min_el is not None else "일자미상"
        max_el_s = f"{max_el}일경과" if max_el is not None else "일자미상"

        cond_valid_dts = sorted({x.production_date for x in all_cond_items if x.production_date != "-"})
        cond_dt_range = (
            "-"
            if not cond_valid_dts
            else (cond_valid_dts[0] if len(cond_valid_dts) == 1 else f"{cond_valid_dts[0]} ~ {cond_valid_dts[-1]}")
        )

        # 생산일자 경과구간별 (신선육 ≤14일 vs 숙성/장기재고 ≥15일) 최저가~최고가(브랜드) 비교
        fresh_items = [x for x in all_cond_items if (_elapsed_days(x.production_date) is not None and _elapsed_days(x.production_date) <= 14)]
        aged_items = [x for x in all_cond_items if (_elapsed_days(x.production_date) is not None and _elapsed_days(x.production_date) >= 15)]

        if fresh_items:
            f_min = min(fresh_items, key=lambda x: x.unit_price_kg)
            f_max = max(fresh_items, key=lambda x: x.unit_price_kg)
            fresh_range_str = (
                f"최저 {f_min.unit_price_kg:,}원({f_min.brand_name.split('(')[0]}[{f_min.platform}], {f_min.production_date}) ~ "
                f"최고 {f_max.unit_price_kg:,}원({f_max.brand_name.split('(')[0]}[{f_max.platform}], {f_max.production_date})"
            )
        else:
            fresh_range_str = "-"

        if aged_items:
            a_min = min(aged_items, key=lambda x: x.unit_price_kg)
            a_max = max(aged_items, key=lambda x: x.unit_price_kg)
            aged_range_str = (
                f"최저 {a_min.unit_price_kg:,}원({a_min.brand_name.split('(')[0]}[{a_min.platform}], {a_min.production_date}) ~ "
                f"최고 {a_max.unit_price_kg:,}원({a_max.brand_name.split('(')[0]}[{a_max.platform}], {a_max.production_date})"
            )
        else:
            aged_range_str = "-"

        # 브랜드별 최저단가 순 라인업 요약 (최저가 브랜드 ~ 최고가 브랜드 + 생산일자 표시)
        brand_grp: Dict[str, List[HanwooProductRow]] = {}
        for x in all_cond_items:
            b_short = x.brand_name.split("(")[0].strip()
            brand_grp.setdefault(f"{b_short}[{x.platform}]", []).append(x)
        brand_sorted = sorted(
            [
                (
                    b_k,
                    min(r_l, key=lambda z: z.unit_price_kg),
                    int(round(sum(z.unit_price_kg for z in r_l) / len(r_l))),
                    max(r_l, key=lambda z: z.unit_price_kg),
                )
                for b_k, r_l in brand_grp.items()
            ],
            key=lambda t: t[1].unit_price_kg,
        )
        if len(brand_sorted) == 1:
            b0 = brand_sorted[0]
            brand_lineup_str = f"{b0[0]} ({b0[1].unit_price_kg:,}원, 생산일:{b0[1].production_date})"
        else:
            b_low = brand_sorted[0]
            b_high = brand_sorted[-1]
            mid_brands = ", ".join(f"{x[0]}({x[1].unit_price_kg:,}원)" for x in brand_sorted[:6])
            brand_lineup_str = (
                f"최저 {b_low[0]}({b_low[1].unit_price_kg:,}원, {b_low[1].production_date}) ~ "
                f"최고 {b_high[0]}({b_high[3].unit_price_kg:,}원, {b_high[3].production_date}) | 브랜드라인업: {mid_brands}"
            )

        gb_dt_str = "-"
        if gb_info:
            gb_dt_str = (
                gb_info["min_date"]
                if gb_info["min_date"] == gb_info["max_date"]
                else f"{gb_info['min_date']} ~ {gb_info['max_date']}"
            )

        row_dict: Dict[str, Any] = {
            "수집일자": today_str,
            "축종구분": cattle_t,
            "보관상태": keep_t,
            "대분류부위(10대기준)": major_part_t,
            "부위명": part_t,
            "등급": grade_t,
            "지육경락가(원/kg)": carcass_price,
            "생산일자_범위(최장기~최신)": cond_dt_range,
            "생산일자대비_최저가_상세": f"{min_item.unit_price_kg:,}원 | {min_item.production_date}({min_el_s}) | {min_item.brand_name}[{min_item.platform}]",
            "생산일자대비_최고가_상세": f"{max_item.unit_price_kg:,}원 | {max_item.production_date}({max_el_s}) | {max_item.brand_name}[{max_item.platform}]",
            "신선육(14일이내)_최저~최고(브랜드)": fresh_range_str,
            "숙성/장기(15일이상)_최저~최고(브랜드)": aged_range_str,
            "근본한우_최저가(원/kg)": gb_min,
            "근본한우_평균가(원/kg)": gb_avg,
            "근본한우_최고가(원/kg)": gb_max,
            "근본한우_브랜드(생산일자)": f"근본한우 ({gb_dt_str})" if gb_info else "-",
            "타사통합_최저가(원/kg)": comp_min,
            "타사통합_평균가(원/kg)": comp_avg,
            "타사통합_최고가(원/kg)": comp_max,
            "근본한우대비_타사평균_편차(%)": market_diff_pct,
            "근본한우대비_가격경쟁력판정": (
                f"근본한우 우위(타사가 {market_diff_pct:+.1f}% 높음)"
                if (market_diff_pct is not None and market_diff_pct > 0)
                else (
                    f"타사 우위(타사가 {market_diff_pct:+.1f}% 낮음)"
                    if (market_diff_pct is not None and market_diff_pct < 0)
                    else ("동일가" if market_diff_pct == 0 else "근본한우 미입점/타사단독")
                )
            ),
        }

        # 각 경쟁 플랫폼별 최저가 / 최저가 브랜드(생산일) / 평균가 / 최고가 / 최고가 브랜드(생산일) / 근본대비 편차(%)
        for p_nm in COMPETITOR_PLATFORMS:
            p_list = pmap.get(p_nm, [])
            if p_list:
                p_prices = [x.unit_price_kg for x in p_list]
                p_min = min(p_prices)
                p_avg = int(round(sum(p_prices) / len(p_prices)))
                p_max = max(p_prices)
                p_min_r = min(p_list, key=lambda x: x.unit_price_kg)
                p_max_r = max(p_list, key=lambda x: x.unit_price_kg)
                p_min_b = f"{p_min_r.brand_name} ({p_min_r.production_date})"
                p_max_b = f"{p_max_r.brand_name} ({p_max_r.production_date})"
                p_diff = round(((p_avg - gb_avg) / gb_avg) * 100.0, 2) if gb_avg else None
                p_cnt = len(p_list)
                p_min_url = p_min_r.product_url
                p_max_url = p_max_r.product_url
            else:
                p_min, p_avg, p_max, p_diff = None, None, None, None
                p_min_b, p_max_b = "-", "-"
                p_cnt = 0
                p_min_url, p_max_url = "", ""

            row_dict[f"{p_nm}_건수"] = p_cnt
            row_dict[f"{p_nm}_최저가(원/kg)"] = p_min
            row_dict[f"{p_nm}_최저가_브랜드(생산일)"] = p_min_b
            row_dict[f"{p_nm}_평균가(원/kg)"] = p_avg
            row_dict[f"{p_nm}_최고가(원/kg)"] = p_max
            row_dict[f"{p_nm}_최고가_브랜드(생산일)"] = p_max_b
            row_dict[f"{p_nm}_근본대비편차(%)"] = p_diff
            row_dict[f"{p_nm}_최저가_URL"] = p_min_url
            row_dict[f"{p_nm}_최고가_URL"] = p_max_url

        row_dict["최저가_플랫폼(브랜드)"] = f"{min_item.platform} | {min_item.brand_name} ({min_item.unit_price_kg:,}원, 생산일:{min_item.production_date})"
        row_dict["최고가_플랫폼(브랜드)"] = f"{max_item.platform} | {max_item.brand_name} ({max_item.unit_price_kg:,}원, 생산일:{max_item.production_date})"
        row_dict["최저가_상품URL"] = min_item.product_url
        row_dict["최고가_상품URL"] = max_item.product_url
        row_dict["브랜드_가격라인업(최저~최고)"] = brand_lineup_str

        dashboard_rows.append(row_dict)

    # 2-B. [전체 품목 전수 비교표 (`full_item_rows`)] 생성:
    # [축종구분(한우거세/한우암소)] x [보관상태(냉장/냉동)] x [품목명(part_name)] 기준으로
    # 전체 품목을 단 한 품목도 누락·축약하지 않고 전수 나열하여 6개 플랫폼(근본한우 vs 미트박스 vs 미트프렌즈 vs 탑미트 vs 금천미트 vs 본대로) 비교
    item_plat_map: Dict[Tuple[str, str, str], Dict[str, List[HanwooProductRow]]] = {}
    for r in all_rows:
        ikey = (r.cattle_type, r.keeping, r.part_name)
        item_plat_map.setdefault(ikey, {}).setdefault(r.platform, []).append(r)

    full_item_rows: List[Dict[str, Any]] = []
    sorted_ikeys = sorted(
        item_plat_map.keys(),
        key=lambda k: (
            0 if k[0] == "한우거세" else 1,
            0 if k[1] == "냉장" else 1,
            _part_rank(k[2]),
            k[2],
        ),
    )
    for ikey in sorted_ikeys:
        cattle_t, keep_t, part_t = ikey
        pmap_i = item_plat_map[ikey]
        all_i_items = [x for plist in pmap_i.values() for x in plist]
        major_p_t = all_i_items[0].major_part if all_i_items else part_t
        grades_present = sorted({x.grade for x in all_i_items}, key=_grade_rank)
        grades_str = ", ".join(grades_present)

        gb_i_list = pmap_i.get("근본한우", [])
        if gb_i_list:
            gb_i_prices = [x.unit_price_kg for x in gb_i_list]
            gb_i_min = min(gb_i_prices)
            gb_i_avg = int(round(sum(gb_i_prices) / len(gb_i_prices)))
            gb_i_max = max(gb_i_prices)
            gb_i_cnt = len(gb_i_list)
            gb_i_dts = sorted({x.production_date for x in gb_i_list if x.production_date != "-"})
            gb_i_dt_s = gb_i_dts[0] if len(gb_i_dts) == 1 else (f"{gb_i_dts[0]}~{gb_i_dts[-1]}" if gb_i_dts else "-")
            gb_i_url = min(gb_i_list, key=lambda x: x.unit_price_kg).product_url
        else:
            gb_i_min, gb_i_avg, gb_i_max, gb_i_cnt, gb_i_dt_s, gb_i_url = None, None, None, 0, "-", ""

        comp_i_items: List[HanwooProductRow] = []
        for p_nm in COMPETITOR_PLATFORMS:
            comp_i_items.extend(pmap_i.get(p_nm, []))
        comp_i_min = min((x.unit_price_kg for x in comp_i_items), default=None)
        comp_i_max = max((x.unit_price_kg for x in comp_i_items), default=None)
        comp_i_avg = int(round(sum(x.unit_price_kg for x in comp_i_items) / len(comp_i_items))) if comp_i_items else None
        comp_i_diff = round(((comp_i_avg - gb_i_avg) / gb_i_avg) * 100.0, 2) if (gb_i_avg and comp_i_avg) else None

        min_i = min(all_i_items, key=lambda x: x.unit_price_kg)
        max_i = max(all_i_items, key=lambda x: x.unit_price_kg)
        valid_i_dts = sorted({x.production_date for x in all_i_items if x.production_date != "-"})
        i_dt_range = "-" if not valid_i_dts else (valid_i_dts[0] if len(valid_i_dts) == 1 else f"{valid_i_dts[0]} ~ {valid_i_dts[-1]}")

        i_row: Dict[str, Any] = {
            "수집일자": today_str,
            "축종구분": cattle_t,
            "보관상태": keep_t,
            "대분류부위(10대기준)": major_p_t,
            "품목명": part_t,
            "취급등급_분포": grades_str,
            "전체_수집건수": len(all_i_items),
            "생산일자_범위": i_dt_range,
            "근본한우_건수": gb_i_cnt,
            "근본한우_최저가(원/kg)": gb_i_min,
            "근본한우_평균가(원/kg)": gb_i_avg,
            "근본한우_최고가(원/kg)": gb_i_max,
            "근본한우_생산일자": gb_i_dt_s,
            "근본한우_최저가_URL": gb_i_url,
            "타사4곳통합_건수": len(comp_i_items),
            "타사4곳통합_최저가(원/kg)": comp_i_min,
            "타사4곳통합_평균가(원/kg)": comp_i_avg,
            "타사4곳통합_최고가(원/kg)": comp_i_max,
            "근본한우대비_타사평균_편차(%)": comp_i_diff,
        }
        for p_nm in COMPETITOR_PLATFORMS:
            p_l = pmap_i.get(p_nm, [])
            if p_l:
                p_pr = [x.unit_price_kg for x in p_l]
                p_mn_r = min(p_l, key=lambda x: x.unit_price_kg)
                p_mx_r = max(p_l, key=lambda x: x.unit_price_kg)
                i_row[f"{p_nm}_건수"] = len(p_l)
                i_row[f"{p_nm}_최저가(원/kg)"] = min(p_pr)
                i_row[f"{p_nm}_평균가(원/kg)"] = int(round(sum(p_pr) / len(p_pr)))
                i_row[f"{p_nm}_최고가(원/kg)"] = max(p_pr)
                i_row[f"{p_nm}_최저가_브랜드(생산일)"] = f"{p_mn_r.brand_name.split('(')[0]} ({p_mn_r.production_date}, {p_mn_r.grade})"
                i_row[f"{p_nm}_최고가_브랜드(생산일)"] = f"{p_mx_r.brand_name.split('(')[0]} ({p_mx_r.production_date}, {p_mx_r.grade})"
                i_row[f"{p_nm}_최저가_URL"] = p_mn_r.product_url
                i_row[f"{p_nm}_최고가_URL"] = p_mx_r.product_url
            else:
                i_row[f"{p_nm}_건수"] = 0
                i_row[f"{p_nm}_최저가(원/kg)"] = None
                i_row[f"{p_nm}_평균가(원/kg)"] = None
                i_row[f"{p_nm}_최고가(원/kg)"] = None
                i_row[f"{p_nm}_최저가_브랜드(생산일)"] = "-"
                i_row[f"{p_nm}_최고가_브랜드(생산일)"] = "-"
                i_row[f"{p_nm}_최저가_URL"] = ""
                i_row[f"{p_nm}_최고가_URL"] = ""

        i_row["품목내_전체최저가_상세"] = f"{min_i.unit_price_kg:,}원 | {min_i.brand_name.split('(')[0]}[{min_i.platform}] ({min_i.grade}, {min_i.production_date})"
        i_row["품목내_전체최고가_상세"] = f"{max_i.unit_price_kg:,}원 | {max_i.brand_name.split('(')[0]}[{max_i.platform}] ({max_i.grade}, {max_i.production_date})"
        i_row["최저가_상품URL"] = min_i.product_url
        i_row["최고가_상품URL"] = max_i.product_url
        full_item_rows.append(i_row)

    # 3. Sheet 3 (`생산일자_등급_보관별_통계`) 생성
    # 필수 그룹핑 기준: [축종구분] x [보관상태] x [부위명] x [2차검증_세부표준상품명] x [등급] x [생산일자(연-월 또는 주차 기준)] x [브랜드명] (+플랫폼)
    cohort_groups: Dict[Tuple[str, str, str, str, str, str, str, str], List[HanwooProductRow]] = {}
    # 동일 조건 + 동일 생산일자 주차구분 내 전체 최저가·최고가 비교용 맵
    prod_cohort_all_map: Dict[Tuple[str, str, str, str, str], List[HanwooProductRow]] = {}

    for r in all_rows:
        ym_week = HanwooStandardizer.to_year_month_or_week(r.production_date)
        gkey = (r.cattle_type, r.keeping, r.part_name, r.sub_spec_name, r.grade, ym_week, r.brand_name, r.platform)
        cohort_groups.setdefault(gkey, []).append(r)
        pkey = (r.cattle_type, r.keeping, r.part_name, r.grade, ym_week)
        prod_cohort_all_map.setdefault(pkey, []).append(r)

    sheet3_rows: List[Dict[str, Any]] = []
    sorted_gkeys = sorted(
        cohort_groups.keys(),
        key=lambda k: (
            0 if k[0] == "한우거세" else 1,
            0 if k[1] == "냉장" else 1,
            _part_rank(k[2]),
            k[2],
            k[3],
            _grade_rank(k[4]),
            k[5],
            PLATFORMS_ORDER.index(k[7]) if k[7] in PLATFORMS_ORDER else 99,
            k[6],
        ),
    )

    for gkey in sorted_gkeys:
        cattle_t, keep_t, part_t, sub_spec_t, grade_t, ym_week, brand_t, plat_t = gkey
        items_g = cohort_groups[gkey]
        major_p = items_g[0].major_part if items_g else part_t
        prices = [x.unit_price_kg for x in items_g]
        min_p = min(prices)
        max_p = max(prices)
        avg_p = int(round(sum(prices) / len(prices)))
        tot_stk = sum(x.stock_qty for x in items_g)
        min_item_g = min(items_g, key=lambda x: x.unit_price_kg)
        max_item_g = max(items_g, key=lambda x: x.unit_price_kg)

        valid_dts = sorted({x.production_date for x in items_g if x.production_date != "-"})
        dt_range_str = (
            "-"
            if not valid_dts
            else (valid_dts[0] if len(valid_dts) == 1 else f"{valid_dts[0]} ~ {valid_dts[-1]}")
        )
        el_vals = [_elapsed_days(d) for d in valid_dts if _elapsed_days(d) is not None]
        avg_el_days = int(round(sum(el_vals) / len(el_vals))) if el_vals else None
        el_bucket_str = _elapsed_bucket(avg_el_days)

        # 동일 조건(축종+보관+부위+등급) 근본한우 판매가 대비 편차 (%)
        ckey = (cattle_t, keep_t, part_t, grade_t)
        gb_ref = geunbon_ref_map.get(ckey)
        gb_ref_price = gb_ref["avg"] if gb_ref else None
        if gb_ref_price and gb_ref_price > 0:
            dev_pct = round(((avg_p - gb_ref_price) / gb_ref_price) * 100.0, 2)
        else:
            dev_pct = None

        # [보완사항 2] 동일 생산일자 구간(연-월/주차) 대비 전체 최저가·최고가 및 브랜드·플랫폼 표시
        pkey = (cattle_t, keep_t, part_t, grade_t, ym_week)
        same_prod_items = prod_cohort_all_map.get(pkey, items_g)
        sp_min_item = min(same_prod_items, key=lambda x: x.unit_price_kg)
        sp_max_item = max(same_prod_items, key=lambda x: x.unit_price_kg)
        sp_spread = sp_max_item.unit_price_kg - sp_min_item.unit_price_kg

        # 동일 조건 전체 생산일자 대비 최저가·최고가 참조
        all_cond_items = [x for plist in cond_plat_map.get(ckey, {}).values() for x in plist]
        cond_min_item = min(all_cond_items, key=lambda x: x.unit_price_kg) if all_cond_items else sp_min_item
        cond_max_item = max(all_cond_items, key=lambda x: x.unit_price_kg) if all_cond_items else sp_max_item

        sheet3_rows.append({
            "수집일자": today_str,
            "축종구분": cattle_t,
            "보관상태": keep_t,
            "대분류부위(10대기준)": major_p,
            "부위명": part_t,
            "2차검증_세부표준상품명": sub_spec_t,
            "등급": grade_t,
            "생산일자_주차구분": ym_week,
            "상세생산일자_범위": dt_range_str,
            "재고경과일수(일)": avg_el_days,
            "생산경과일_구간": el_bucket_str,
            "플랫폼명": plat_t,
            "브랜드명": brand_t,
            "상품옵션수(재고)": tot_stk,
            "최저가(원/kg)": min_p,
            "평균가(원/kg)": avg_p,
            "최고가(원/kg)": max_p,
            "최저가_상품URL": min_item_g.product_url,
            "최고가_상품URL": max_item_g.product_url,
            "동일생산일구간_전체최저가(원/kg)": sp_min_item.unit_price_kg,
            "동일생산일구간_최저가_브랜드(플랫폼)": f"{sp_min_item.brand_name}[{sp_min_item.platform}] ({sp_min_item.production_date})",
            "동일생산일구간_최저가_URL": sp_min_item.product_url,
            "동일생산일구간_전체최고가(원/kg)": sp_max_item.unit_price_kg,
            "동일생산일구간_최고가_브랜드(플랫폼)": f"{sp_max_item.brand_name}[{sp_max_item.platform}] ({sp_max_item.production_date})",
            "동일생산일구간_최고가_URL": sp_max_item.product_url,
            "생산일구간_최고최저_편차(원/kg)": sp_spread,
            "근본한우_동일조건_기준가(원/kg)": gb_ref_price,
            "근본한우대비_가격편차(%)": dev_pct,
            "동일조건_전체생산일_최저~최고(브랜드)": (
                f"최저 {cond_min_item.unit_price_kg:,}원({cond_min_item.brand_name.split('(')[0]}[{cond_min_item.platform}], {cond_min_item.production_date}) ~ "
                f"최고 {cond_max_item.unit_price_kg:,}원({cond_max_item.brand_name.split('(')[0]}[{cond_max_item.platform}], {cond_max_item.production_date})"
            ),
        })

    # 요약 메트릭 산출
    plat_counts = {p: sum(1 for r in all_rows if r.platform == p) for p in PLATFORMS_ORDER}
    plat_brand_counts = {
        p: len({r.brand_name for r in all_rows if r.platform == p and r.brand_name})
        for p in PLATFORMS_ORDER
    }
    gb_matched_rows = [r for r in dashboard_rows if r["근본한우_평균가(원/kg)"] is not None and r["타사통합_평균가(원/kg)"] is not None]
    gb_cheaper_cnt = sum(1 for r in gb_matched_rows if (r["근본한우대비_타사평균_편차(%)"] or 0) > 0)
    reclassified_cnt = sum(1 for r in all_rows if "2차보정" in r.verify_status)

    analytics_meta = {
        "total_raw_rows": len(all_rows),
        "total_item_groups": len(full_item_rows),
        "total_condition_groups": len(dashboard_rows),
        "total_cohort_groups": len(sheet3_rows),
        "platform_counts": plat_counts,
        "platform_brand_counts": plat_brand_counts,
        "total_unique_brands": len({r.brand_name for r in all_rows if r.brand_name}),
        "market_ref_prod_date": market_ref_dt_str,
        "gb_matched_conditions": len(gb_matched_rows),
        "gb_cheaper_conditions": gb_cheaper_cnt,
        "stage2_reclassified_rows": reclassified_cnt,
        "stage2_audit_log_count": len(HanwooStandardizer.verification_audit_logs),
        "full_item_rows": full_item_rows,
    }
    logger.info(
        "다차원 집계 완료: 전체품목그룹 %d개 | 품목×등급조건(Sheet1) %d개 (근본한우 매칭 %d개 중 가격우위 %d개) | 생산일자×브랜드 그룹(Sheet3) %d개 | 고유브랜드 %d개 | 2차보정 %d건",
        len(full_item_rows),
        len(dashboard_rows),
        len(gb_matched_rows),
        gb_cheaper_cnt,
        len(sheet3_rows),
        analytics_meta["total_unique_brands"],
        reclassified_cnt,
    )
    return dashboard_rows, sheet3_rows, analytics_meta



# ---------------------------------------------------------------------------
# [Task 4] 6대 브랜드/플랫폼 비교 및 시각화 리포트 생성 (PNG + PDF + HTML)
# ---------------------------------------------------------------------------
def configure_korean_font() -> str:
    """한글 폰트 깨짐 방지를 위해 Malgun Gothic 또는 NanumGothic을 설정합니다."""
    preferred_fonts = ["Malgun Gothic", "NanumGothic", "AppleGothic", "Gulim"]
    available = {f.name for f in fm.fontManager.ttflist}
    chosen = "Malgun Gothic"
    for fn in preferred_fonts:
        if fn in available:
            chosen = fn
            break
    plt.rcParams["font.family"] = chosen
    plt.rcParams["axes.unicode_minus"] = False
    return chosen


def generate_visual_reports(
    all_rows: List[HanwooProductRow],
    dashboard_rows: List[Dict[str, Any]],
    sheet3_rows: List[Dict[str, Any]],
    ekape_rows: List[Dict[str, Any]],
    ekape_summary: Dict[str, Any],
    analytics_meta: Dict[str, Any],
    output_dir: Path,
    today_str: str,
) -> Dict[str, Path]:
    """
    Task 4 요구사항에 맞춰 3대 핵심 시각화 차트(PNG), 통합 PDF 리포트, 인터랙티브 HTML 대시보드를 생성합니다.
    1. 차트 1 (등급별·보관상태별 가격 분포 차트): 주요 부위별 [냉장 vs 냉동] 및 [등급별(1++~2등급)] 6개 플랫폼 [최저가-평균가-최고가] 비교 + 근본한우 기준선
    2. 차트 2 (생산일자별 단가 스캐터/히트맵): 원물 생산일자(재고 경과일) 및 보관상태(냉장/냉동)에 따른 플랫폼별·브랜드별 단가 분포
    3. 차트 3 (지육시세 연동성 차트): 한우거세·암소 등급별 지육시세(전일/전주/전월/전년 동기) 변동률과 각 플랫폼 냉장 원물 평균 판매가/지육배율 비교
    """
    logger.info("=== [Task 4] 6대 플랫폼 비교 시각화 차트(PNG 3종) 및 통합 PDF/HTML 리포트 생성 시작 ===")
    font_used = configure_korean_font()
    logger.info("시각화 한글 폰트 적용 완료: %s", font_used)

    output_dir.mkdir(parents=True, exist_ok=True)
    chart1_path = output_dir / f"{today_str}_차트1_등급별_보관상태별_6사_가격분포.png"
    chart2_path = output_dir / f"{today_str}_차트2_생산일자별_단가스캐터_및_히트맵.png"
    chart3_path = output_dir / f"{today_str}_차트3_지육시세_연동성_및_플랫폼비교.png"
    pdf_path = output_dir / f"{today_str}_한우원물_6사비교_시각화리포트.pdf"
    html_path = output_dir / f"{today_str}_한우원물_6사비교_시각화리포트.html"

    plat_colors = {
        "근본한우": "#DC2626",   # 강조 레드
        "미트박스": "#2563EB",   # 블루
        "미트프렌즈": "#059669", # 에메랄드 그린
        "탑미트": "#D97706",     # 앰버 오렌지
        "금천미트": "#7C3AED",   # 퍼플
        "본대로": "#0891B2",     # 시안
    }

    # =========================================================================
    # [Chart 1] 등급별·보관상태별 6개 플랫폼 [최저가 - 평균가 - 최고가] 분포 + 근본한우 기준선
    # =========================================================================
    fig1 = plt.figure(figsize=(20, 13), dpi=160)
    fig1.patch.set_facecolor("#F8FAFC")
    gs1 = gridspec.GridSpec(2, 2, figure=fig1, hspace=0.34, wspace=0.20, top=0.89, bottom=0.07, left=0.06, right=0.97)
    fig1.suptitle(
        f"[Chart 1] 한우 원물(박스육) 주요 부위·등급·보관상태별 6대 플랫폼 [최저가-평균가-최고가] 분포 ({today_str})",
        fontsize=17,
        fontweight="bold",
        color="#0F172A",
        y=0.96,
    )
    fig1.text(
        0.5,
        0.925,
        "붉은색 파선(---)은 동일 [축종 × 보관상태 × 부위 × 등급] 조건의 '근본한우' 평균 판매단가(원/kg) 기준선이며, 막대/에러바는 플랫폼별 [최저가 ~ 평균가 ~ 최고가] 범위를 나타냅니다.",
        ha="center",
        fontsize=11,
        color="#475569",
    )

    # 패널 구성:
    # (0,0) 한우암소 냉장 주요 구이용(등심/안심/채끝/갈비) 등급별 6사 비교
    # (0,1) 한우암소 냉장 주요 정육/특수(양지/설도/앞다리/우둔/사태/차돌박이) 등급별 6사 비교
    # (1,0) 한우거세 냉장 주요 부위(등심/안심/채끝/갈비/양지/설도/사태) 등급별 6사 비교
    # (1,1) 냉장 vs 냉동 등급별(1++(9)~2등급) 6대 플랫폼 전체 원물 단가 분포 비교
    panel_configs = [
        (0, 0, "한우암소", "냉장", ["등심", "안심", "채끝", "갈비"], "① [한우암소 · 냉장] 구이용 핵심부위 등급별 6사 가격 분포"),
        (0, 1, "한우암소", "냉장", ["양지", "설도", "앞다리", "우둔", "사태"], "② [한우암소 · 냉장] 정육/국거리 핵심부위 등급별 6사 가격 분포"),
        (1, 0, "한우거세", "냉장", ["등심", "안심", "채끝", "갈비", "양지", "설도", "사태"], "③ [한우거세 · 냉장] 주요 부위·등급별 6사 가격 분포"),
    ]

    for r_idx, c_idx, c_type, keep_type, target_parts, sub_title in panel_configs:
        ax = fig1.add_subplot(gs1[r_idx, c_idx])
        ax.set_facecolor("#FFFFFF")

        # 근본한우가 입점해 있거나 타사 데이터가 풍부한 상위 (부위, 등급) 조합 선택 (최대 7개)
        cand_rows = [
            r for r in dashboard_rows
            if r["축종구분"] == c_type and r["보관상태"] == keep_type and r["부위명"] in target_parts
            and r["등급"] in ("1++(9)", "1++(8)", "1++(7)", "1++(불명)", "1+", "1", "2")
        ]
        # 근본한우 보유 조건 우선 정렬 후 타사 입점 플랫폼 수 순 정렬
        cand_rows.sort(
            key=lambda x: (
                0 if x["근본한우_평균가(원/kg)"] is not None else 1,
                target_parts.index(x["부위명"]) if x["부위명"] in target_parts else 99,
                GRADES_ORDER.index(x["등급"]) if x["등급"] in GRADES_ORDER else 99,
            )
        )
        selected_rows = cand_rows[:7]

        if not selected_rows:
            ax.text(0.5, 0.5, "해당 조건 데이터 없음", ha="center", va="center", fontsize=12)
            ax.set_title(sub_title, fontsize=12.5, fontweight="bold", pad=10)
            continue

        x_positions = np.arange(len(selected_rows))
        bar_w = 0.12
        offsets = np.linspace(-2.5 * bar_w, 2.5 * bar_w, len(PLATFORMS_ORDER))

        for p_i, p_nm in enumerate(PLATFORMS_ORDER):
            means = []
            yerr_low = []
            yerr_high = []
            valid_x = []
            for idx_r, d_row in enumerate(selected_rows):
                if p_nm == "근본한우":
                    p_min = d_row.get("근본한우_최저가(원/kg)")
                    p_avg = d_row.get("근본한우_평균가(원/kg)")
                    p_max = d_row.get("근본한우_최고가(원/kg)")
                else:
                    p_min = d_row.get(f"{p_nm}_최저가(원/kg)")
                    p_avg = d_row.get(f"{p_nm}_평균가(원/kg)")
                    p_max = d_row.get(f"{p_nm}_최고가(원/kg)")

                if p_avg is not None and p_min is not None and p_max is not None:
                    valid_x.append(x_positions[idx_r] + offsets[p_i])
                    means.append(p_avg / 1000.0)
                    yerr_low.append(max(0.0, (p_avg - p_min) / 1000.0))
                    yerr_high.append(max(0.0, (p_max - p_avg) / 1000.0))

            if valid_x:
                ax.bar(
                    valid_x,
                    means,
                    width=bar_w * 0.92,
                    color=plat_colors[p_nm],
                    alpha=0.92 if p_nm == "근본한우" else 0.78,
                    edgecolor="#7F1D1D" if p_nm == "근본한우" else "none",
                    linewidth=1.3 if p_nm == "근본한우" else 0,
                    label=p_nm,
                    zorder=4 if p_nm == "근본한우" else 3,
                )
                ax.errorbar(
                    valid_x,
                    means,
                    yerr=[yerr_low, yerr_high],
                    fmt="none",
                    ecolor="#1E293B",
                    elinewidth=1.1,
                    capsize=2.5,
                    zorder=5,
                )

        # 각 조건별 근본한우 기준선(수평 점선 세그먼트) + 최저가/최고가 브랜드명 라벨 표시
        for idx_r, d_row in enumerate(selected_rows):
            gb_avg = d_row.get("근본한우_평균가(원/kg)")
            if gb_avg is not None:
                y_val = gb_avg / 1000.0
                ax.hlines(
                    y=y_val,
                    xmin=x_positions[idx_r] - 0.38,
                    xmax=x_positions[idx_r] + 0.38,
                    colors="#DC2626",
                    linestyles="--",
                    linewidth=1.8,
                    zorder=6,
                )
                ax.text(
                    x_positions[idx_r],
                    y_val + 1.5,
                    f"근본 {y_val:.1f}천",
                    ha="center",
                    va="bottom",
                    fontsize=7.8,
                    fontweight="bold",
                    color="#B91C1C",
                    bbox=dict(boxstyle="round,pad=0.14", facecolor="#FEF2F2", edgecolor="#FCA5A5", alpha=0.92),
                    zorder=7,
                )

            # 해당 조건의 전체 최저가 브랜드 및 최고가 브랜드 추출 후 차트 상/하단 표시
            min_brand_info = str(d_row.get("생산일자대비_최저가_상세") or "")
            max_brand_info = str(d_row.get("생산일자대비_최고가_상세") or "")
            t_min = d_row.get("타사통합_최저가(원/kg)") or d_row.get("근본한우_최저가(원/kg)")
            t_max = d_row.get("타사통합_최고가(원/kg)") or d_row.get("근본한우_최고가(원/kg)")
            if gb_avg is not None:
                t_min = min(filter(None, [t_min, d_row.get("근본한우_최저가(원/kg)")]))
                t_max = max(filter(None, [t_max, d_row.get("근본한우_최고가(원/kg)")]))

            if t_max and "|" in max_brand_info:
                b_part = max_brand_info.split("|")[-1].strip().split("(")[0][:10]
                ax.text(
                    x_positions[idx_r],
                    (t_max / 1000.0) + 3.2,
                    f"▲최고:{b_part}\n({t_max/1000.0:.0f}천)",
                    ha="center",
                    va="bottom",
                    fontsize=6.8,
                    color="#1E293B",
                    fontweight="bold",
                    zorder=7,
                )
            if t_min and "|" in min_brand_info:
                b_part_min = min_brand_info.split("|")[-1].strip().split("(")[0][:10]
                ax.text(
                    x_positions[idx_r],
                    max(1.0, (t_min / 1000.0) - 5.5),
                    f"▼최저:{b_part_min}({t_min/1000.0:.0f}천)",
                    ha="center",
                    va="top",
                    fontsize=6.8,
                    color="#0369A1",
                    fontweight="bold",
                    zorder=7,
                )

        x_labels = [f"{r['부위명']}\n[{r['등급']}]" for r in selected_rows]
        ax.set_xticks(x_positions)
        ax.set_xticklabels(x_labels, fontsize=9.2, fontweight="bold")
        ax.set_ylabel("단위단가 (천원/kg)", fontsize=10, fontweight="bold")
        ax.set_ylim(bottom=0, top=ax.get_ylim()[1] * 1.18)
        ax.set_title(sub_title, fontsize=12, fontweight="bold", color="#0F172A", pad=10)
        ax.grid(axis="y", linestyle=":", alpha=0.6, zorder=0)
        ax.legend(loc="upper right", fontsize=8.2, ncol=3, framealpha=0.9)

    # Panel (1,1): [냉장 vs 냉동] × 등급별(1++(9)~2등급) 6사 평균가 및 최저~최고 비교
    ax4 = fig1.add_subplot(gs1[1, 1])
    ax4.set_facecolor("#FFFFFF")
    target_grades_p4 = ["1++(9)", "1++(8)", "1++(7)", "1+", "1", "2"]
    x_pos4 = np.arange(len(target_grades_p4))
    width4 = 0.13

    for p_i, p_nm in enumerate(PLATFORMS_ORDER):
        cold_avgs = []
        froz_avgs = []
        valid_x_cold = []
        valid_x_froz = []
        for g_idx, g_lbl in enumerate(target_grades_p4):
            c_prices = [
                r.unit_price_kg / 1000.0
                for r in all_rows
                if r.platform == p_nm and r.keeping == "냉장" and r.grade == g_lbl
            ]
            f_prices = [
                r.unit_price_kg / 1000.0
                for r in all_rows
                if r.platform == p_nm and r.keeping == "냉동" and r.grade == g_lbl
            ]
            pos = x_pos4[g_idx] + (p_i - 2.5) * width4
            if c_prices:
                valid_x_cold.append(pos)
                cold_avgs.append(float(np.mean(c_prices)))
            if f_prices:
                valid_x_froz.append(pos)
                froz_avgs.append(float(np.mean(f_prices)))

        if valid_x_cold:
            ax4.bar(
                valid_x_cold,
                cold_avgs,
                width=width4 * 0.88,
                color=plat_colors[p_nm],
                alpha=0.88,
                label=f"{p_nm}(냉장)",
                zorder=3,
            )
        if valid_x_froz:
            ax4.scatter(
                valid_x_froz,
                froz_avgs,
                color=plat_colors[p_nm],
                edgecolors="#0F172A",
                marker="D",
                s=42,
                zorder=5,
                label=f"{p_nm}(냉동◆)" if p_i == 0 else "_nolegend_",
            )

    ax4.set_xticks(x_pos4)
    ax4.set_xticklabels(target_grades_p4, fontsize=10, fontweight="bold")
    ax4.set_ylabel("전체 부위 평균단가 (천원/kg)", fontsize=10, fontweight="bold")
    ax4.set_title("④ 등급별(1++(9)~2등급) × [냉장(막대) vs 냉동(◆)] 6대 플랫폼 단가 비교", fontsize=12, fontweight="bold", color="#0F172A", pad=10)
    ax4.grid(axis="y", linestyle=":", alpha=0.6, zorder=0)
    ax4.legend(loc="upper right", fontsize=8, ncol=3, framealpha=0.9)

    fig1.savefig(chart1_path, dpi=160)

    # =========================================================================
    # [Chart 2] 생산일자(재고 경과일) 대비 최저가·최고가 밴드 및 플랫폼별·브랜드별 단가 스캐터/히트맵
    # =========================================================================
    fig2 = plt.figure(figsize=(20, 12.8), dpi=160)
    fig2.patch.set_facecolor("#F8FAFC")
    gs2 = gridspec.GridSpec(2, 2, figure=fig2, hspace=0.35, wspace=0.22, top=0.89, bottom=0.07, left=0.06, right=0.97)
    ref_prod_dt_str = str(analytics_meta.get("market_ref_prod_date") or today_str)
    fig2.suptitle(
        f"[Chart 2] 원물 생산일자 대비 최저가·최고가 분포 및 플랫폼·브랜드별 단가 스캐터 · 히트맵 (기준일: {today_str})",
        fontsize=17,
        fontweight="bold",
        color="#0F172A",
        y=0.96,
    )
    fig2.text(
        0.5,
        0.925,
        "생산일자(가공/도축일자) 기준 경과일 구간별 [최저가 ~ 평균가 ~ 최고가] 가격 밴드와 최저가·최고가 브랜드명을 함께 표시하여 신선육 프리미엄과 장기 재고 할인폭을 검증합니다.",
        ha="center",
        fontsize=11,
        color="#475569",
    )

    try:
        ref_dt_obj = datetime.strptime(ref_prod_dt_str[:10], "%Y-%m-%d")
    except Exception:
        ref_dt_obj = datetime.strptime(today_str[:10], "%Y-%m-%d")

    def _calc_elapsed_days(dt_str: str) -> Optional[int]:
        if not dt_str or dt_str == "-" or len(dt_str) < 10:
            return None
        try:
            dt_obj = datetime.strptime(dt_str[:10], "%Y-%m-%d")
            days = (ref_dt_obj - dt_obj).days
            if days < 0:
                days = 0
            return min(days, 730)
        except Exception:
            return None

    # Panel (0,0): [냉장 · 구이용 핵심부위(등심/안심/채끝/갈비)] 생산일자(경과일 구간) 대비 [최저가-평균가-최고가] 밴드 및 브랜드 표시
    ax2_1 = fig2.add_subplot(gs2[0, 0])
    ax2_1.set_facecolor("#FFFFFF")
    age_buckets = [
        ("0~7일\n(초신선)", 0, 7),
        ("8~14일\n(신선육)", 8, 14),
        ("15~21일\n(숙성적기)", 15, 21),
        ("22~30일\n(숙성후기)", 22, 30),
        ("31~45일\n(장기재고)", 31, 45),
        ("46~90일\n(초장기)", 46, 90),
    ]
    b_x = np.arange(len(age_buckets))
    b_mins_v, b_avgs_v, b_maxs_v = [], [], []
    b_min_labels, b_max_labels = [], []

    for b_idx, (b_lbl, d_lo, d_hi) in enumerate(age_buckets):
        b_items = [
            r for r in all_rows
            if r.keeping == "냉장"
            and r.part_name in ("등심", "안심", "채끝", "갈비")
            and (_calc_elapsed_days(r.production_date) is not None and d_lo <= _calc_elapsed_days(r.production_date) <= d_hi)
        ]
        if b_items:
            mn_r = min(b_items, key=lambda x: x.unit_price_kg)
            mx_r = max(b_items, key=lambda x: x.unit_price_kg)
            avg_v = float(np.mean([x.unit_price_kg for x in b_items])) / 1000.0
            b_mins_v.append(mn_r.unit_price_kg / 1000.0)
            b_avgs_v.append(avg_v)
            b_maxs_v.append(mx_r.unit_price_kg / 1000.0)
            b_min_labels.append(f"▼최저:{mn_r.brand_name.split('(')[0][:8]}[{mn_r.platform[:2]}]\n({mn_r.unit_price_kg/1000.0:.1f}천, {mn_r.production_date[5:]})")
            b_max_labels.append(f"▲최고:{mx_r.brand_name.split('(')[0][:8]}[{mx_r.platform[:2]}]\n({mx_r.unit_price_kg/1000.0:.1f}천, {mx_r.production_date[5:]})")
            # 플랫폼별 산점도 중첩 표시
            for p_nm in PLATFORMS_ORDER:
                p_pts = [x.unit_price_kg / 1000.0 for x in b_items if x.platform == p_nm]
                if p_pts:
                    jitter = np.linspace(-0.18, 0.18, len(p_pts)) if len(p_pts) > 1 else [0.0]
                    ax2_1.scatter(
                        [b_idx + j for j in jitter],
                        p_pts,
                        color=plat_colors[p_nm],
                        alpha=0.85 if p_nm == "근본한우" else 0.45,
                        s=55 if p_nm == "근본한우" else 22,
                        label=p_nm if b_idx == 0 else "_nolegend_",
                        zorder=5 if p_nm == "근본한우" else 3,
                    )
        else:
            b_mins_v.append(np.nan)
            b_avgs_v.append(np.nan)
            b_maxs_v.append(np.nan)
            b_min_labels.append("")
            b_max_labels.append("")

    valid_bx = [i for i, v in enumerate(b_avgs_v) if not np.isnan(v)]
    if valid_bx:
        ax2_1.plot(valid_bx, [b_maxs_v[i] for i in valid_bx], color="#DC2626", linestyle="--", marker="^", linewidth=1.8, label="구간별 최고가(Max)")
        ax2_1.plot(valid_bx, [b_avgs_v[i] for i in valid_bx], color="#0F172A", linestyle="-", marker="o", linewidth=2.4, label="구간별 평균가(Avg)")
        ax2_1.plot(valid_bx, [b_mins_v[i] for i in valid_bx], color="#2563EB", linestyle="--", marker="v", linewidth=1.8, label="구간별 최저가(Min)")
        ax2_1.fill_between(
            valid_bx,
            [b_mins_v[i] for i in valid_bx],
            [b_maxs_v[i] for i in valid_bx],
            color="#CBD5E1",
            alpha=0.22,
        )
        for i in valid_bx:
            ax2_1.text(
                i, b_maxs_v[i] + 3.0, b_max_labels[i],
                ha="center", va="bottom", fontsize=7.0, fontweight="bold", color="#991B1B",
                bbox=dict(boxstyle="round,pad=0.12", facecolor="#FEF2F2", edgecolor="#FCA5A5", alpha=0.88),
            )
            ax2_1.text(
                i, max(2.0, b_mins_v[i] - 4.0), b_min_labels[i],
                ha="center", va="top", fontsize=7.0, fontweight="bold", color="#1E40AF",
                bbox=dict(boxstyle="round,pad=0.12", facecolor="#EFF6FF", edgecolor="#93C5FD", alpha=0.88),
            )

    ax2_1.set_xticks(b_x)
    ax2_1.set_xticklabels([b[0] for b in age_buckets], fontsize=9, fontweight="bold")
    ax2_1.set_ylabel("단위단가 (천원/kg)", fontsize=10, fontweight="bold")
    ax2_1.set_ylim(bottom=0, top=max([v for v in b_maxs_v if not np.isnan(v)] + [150.0]) * 1.22)
    ax2_1.set_title("① [생산일자 경과구간 대비 최저가·최고가] 냉장 구이부위(등심/안심/채끝/갈비) 가격 밴드 & 브랜드", fontsize=11.5, fontweight="bold", pad=10)
    ax2_1.grid(True, linestyle=":", alpha=0.6)
    ax2_1.legend(loc="upper right", fontsize=7.8, ncol=3, framealpha=0.9)

    # Panel (0,1): [냉장(●) vs 냉동(▲) 전체 원물] 생산일자 경과일별 단가 스캐터 + 최저/최고가 브랜드 포인트 표시
    ax2_2 = fig2.add_subplot(gs2[0, 1])
    ax2_2.set_facecolor("#FFFFFF")
    plotted_pts: List[Tuple[int, float, HanwooProductRow]] = []
    for keep_lbl, marker_shape, alpha_v in [("냉장", "o", 0.62), ("냉동", "^", 0.78)]:
        for p_nm in PLATFORMS_ORDER:
            xs, ys = [], []
            for r in all_rows:
                if r.platform == p_nm and r.keeping == keep_lbl:
                    el = _calc_elapsed_days(r.production_date)
                    if el is not None and el <= 180:
                        xs.append(el)
                        ys.append(r.unit_price_kg / 1000.0)
                        plotted_pts.append((el, r.unit_price_kg / 1000.0, r))
            if xs:
                ax2_2.scatter(
                    xs,
                    ys,
                    color=plat_colors[p_nm],
                    marker=marker_shape,
                    alpha=0.92 if p_nm == "근본한우" else alpha_v,
                    s=60 if p_nm == "근본한우" else 28,
                    label=f"{p_nm}({keep_lbl})",
                    zorder=5 if p_nm == "근본한우" else 3,
                )

    if plotted_pts:
        # 전체 스캐터 내 최고가/최저가 및 장기재고 최저가 브랜드 주석 표시
        top_pt = max(plotted_pts, key=lambda t: t[1])
        low_pt = min(plotted_pts, key=lambda t: t[1])
        ax2_2.annotate(
            f"최고가: {top_pt[2].brand_name.split('(')[0]}[{top_pt[2].platform}]\n{top_pt[2].part_name} {top_pt[1]:.1f}천원 ({top_pt[2].production_date})",
            xy=(top_pt[0], top_pt[1]),
            xytext=(min(top_pt[0] + 15, 130), top_pt[1] * 0.95),
            arrowprops=dict(arrowstyle="->", color="#DC2626", lw=1.2),
            fontsize=7.5, fontweight="bold", color="#991B1B",
            bbox=dict(boxstyle="round,pad=0.2", facecolor="#FEF2F2", edgecolor="#FCA5A5"),
            zorder=8,
        )
        ax2_2.annotate(
            f"최저가: {low_pt[2].brand_name.split('(')[0]}[{low_pt[2].platform}]\n{low_pt[2].part_name} {low_pt[1]:.1f}천원 ({low_pt[2].production_date})",
            xy=(low_pt[0], low_pt[1]),
            xytext=(min(low_pt[0] + 18, 130), low_pt[1] + 22.0),
            arrowprops=dict(arrowstyle="->", color="#2563EB", lw=1.2),
            fontsize=7.5, fontweight="bold", color="#1E40AF",
            bbox=dict(boxstyle="round,pad=0.2", facecolor="#EFF6FF", edgecolor="#93C5FD"),
            zorder=8,
        )

    ax2_2.set_xlabel("생산(가공/도축) 후 경과일수 (일, 0~180일)", fontsize=10, fontweight="bold")
    ax2_2.set_ylabel("단위단가 (천원/kg)", fontsize=10, fontweight="bold")
    ax2_2.set_title("② [냉장(●) vs 냉동(▲)] 생산 경과일수별 단가 스캐터 및 최저·최고가 브랜드", fontsize=11.5, fontweight="bold", pad=10)
    ax2_2.grid(True, linestyle=":", alpha=0.6)
    ax2_2.legend(loc="upper right", fontsize=7.5, ncol=2, framealpha=0.9)

    # Panel (1,0): 생산연월·주차(Vintage Cohort) × 6대 플랫폼 [평균가 (최저~최고)] 히트맵 (냉장 원물)
    ax2_3 = fig2.add_subplot(gs2[1, 0])
    ax2_3.set_facecolor("#FFFFFF")

    ym_counts: Dict[str, int] = {}
    for r in all_rows:
        ym = HanwooStandardizer.to_year_month_or_week(r.production_date)
        if ym not in ("생산일미상", "일자미상"):
            ym_counts[ym] = ym_counts.get(ym, 0) + 1
    top_yms = sorted(ym_counts.keys(), reverse=True)[:7]
    if not top_yms:
        top_yms = ["일자미상"]

    heat_matrix = np.full((len(PLATFORMS_ORDER), len(top_yms)), np.nan)
    heat_cell_text: Dict[Tuple[int, int], str] = {}
    for p_i, p_nm in enumerate(PLATFORMS_ORDER):
        for y_i, ym_lbl in enumerate(top_yms):
            vals = [
                r.unit_price_kg / 1000.0
                for r in all_rows
                if r.platform == p_nm
                and r.keeping == "냉장"
                and HanwooStandardizer.to_year_month_or_week(r.production_date) == ym_lbl
            ]
            if vals:
                avg_v = float(np.mean(vals))
                mn_v = float(min(vals))
                mx_v = float(max(vals))
                heat_matrix[p_i, y_i] = avg_v
                heat_cell_text[(p_i, y_i)] = f"{avg_v:.1f}\n({mn_v:.0f}~{mx_v:.0f})"

    im = ax2_3.imshow(heat_matrix, cmap="YlOrRd", aspect="auto")
    ax2_3.set_xticks(np.arange(len(top_yms)))
    ax2_3.set_xticklabels(top_yms, rotation=18, ha="right", fontsize=8.8, fontweight="bold")
    ax2_3.set_yticks(np.arange(len(PLATFORMS_ORDER)))
    ax2_3.set_yticklabels(PLATFORMS_ORDER, fontsize=10, fontweight="bold")
    ax2_3.set_title("③ [냉장 원물] 생산연월·주차별 6대 플랫폼 평균가 및 (최저~최고) 히트맵 (천원/kg)", fontsize=11.5, fontweight="bold", pad=10)

    for i in range(len(PLATFORMS_ORDER)):
        for j in range(len(top_yms)):
            val = heat_matrix[i, j]
            txt = heat_cell_text.get((i, j), "-")
            ax2_3.text(
                j,
                i,
                txt,
                ha="center",
                va="center",
                color="#0F172A" if (np.isnan(val) or val < 72) else "#FFFFFF",
                fontsize=7.6,
                fontweight="bold",
            )
    fig2.colorbar(im, ax=ax2_3, fraction=0.035, pad=0.03, label="천원/kg")

    # Panel (1,1): 주요 브랜드별 [냉장 등심 1++/1+/1등급] 최저가~평균가~최고가 및 생산일자 범위 포지셔닝
    ax2_4 = fig2.add_subplot(gs2[1, 1])
    ax2_4.set_facecolor("#FFFFFF")
    brand_stats: Dict[str, List[HanwooProductRow]] = {}
    for r in all_rows:
        if r.keeping == "냉장" and r.part_name == "등심" and r.grade in ("1++(9)", "1++(8)", "1++(7)", "1++(불명)", "1+", "1"):
            b_clean = r.brand_name.split("(")[0].strip()
            key_lbl = f"{b_clean} [{r.platform}]"
            brand_stats.setdefault(key_lbl, []).append(r)

    sorted_brands = sorted(
        brand_stats.items(),
        key=lambda kv: (0 if "근본한우" in kv[0] else 1, -len(kv[1])),
    )[:12]
    sorted_brands.sort(key=lambda kv: np.mean([x.unit_price_kg for x in kv[1]]))

    if sorted_brands:
        b_names = [kv[0] for kv in sorted_brands]
        b_avgs = [float(np.mean([x.unit_price_kg for x in kv[1]])) / 1000.0 for kv in sorted_brands]
        b_mins = [float(min(x.unit_price_kg for x in kv[1])) / 1000.0 for kv in sorted_brands]
        b_maxs = [float(max(x.unit_price_kg for x in kv[1])) / 1000.0 for kv in sorted_brands]
        b_colors = [plat_colors.get(kv[1][0].platform, "#64748B") for kv in sorted_brands]

        y_pos = np.arange(len(b_names))
        ax2_4.barh(y_pos, b_avgs, color=b_colors, alpha=0.85, height=0.65)
        ax2_4.errorbar(
            b_avgs,
            y_pos,
            xerr=[
                [max(0.0, a - mn) for a, mn in zip(b_avgs, b_mins)],
                [max(0.0, mx - a) for a, mx in zip(b_avgs, b_maxs)],
            ],
            fmt="none",
            ecolor="#0F172A",
            elinewidth=1.2,
            capsize=3,
        )
        for idx_b, (avg_v, mn_v, mx_v, b_nm, kv) in enumerate(zip(b_avgs, b_mins, b_maxs, b_names, sorted_brands)):
            dts = sorted({x.production_date[5:] for x in kv[1] if x.production_date != "-" and len(x.production_date) >= 10})
            dt_tag = dts[-1] if dts else "-"
            ax2_4.text(
                mx_v + 1.5,
                idx_b,
                f"평균 {avg_v:.1f}천 (최저 {mn_v:.0f}~최고 {mx_v:.0f}천 | 생산:{dt_tag})",
                va="center",
                fontsize=7.6,
                fontweight="bold",
                color="#DC2626" if "근본한우" in b_nm else "#1E293B",
            )
        ax2_4.set_yticks(y_pos)
        ax2_4.set_yticklabels(b_names, fontsize=8.8, fontweight="bold")
        ax2_4.set_xlim(left=0, right=max(b_maxs) * 1.42)
        ax2_4.set_xlabel("냉장 등심(1++~1등급) 단가 [최저 ~ 평균 ~ 최고] (천원/kg)", fontsize=10, fontweight="bold")
        ax2_4.set_title("④ 주요 브랜드별 [냉장 등심] 최저가·평균가·최고가 및 생산일자 비교", fontsize=11.5, fontweight="bold", pad=10)
        ax2_4.grid(axis="x", linestyle=":", alpha=0.6)

    fig2.savefig(chart2_path, dpi=160)

    # =========================================================================
    # [Chart 3] 한우거세·암소 등급별 지육시세 변동률(전일/전주/전월/전년) 및 플랫폼 냉장 원물 연동성 차트
    # =========================================================================
    fig3 = plt.figure(figsize=(20, 12.5), dpi=160)
    fig3.patch.set_facecolor("#F8FAFC")
    gs3 = gridspec.GridSpec(2, 2, figure=fig3, hspace=0.35, wspace=0.22, top=0.89, bottom=0.07, left=0.06, right=0.97)
    ek_base_dt = ekape_summary.get("base_date", today_str)
    fig3.suptitle(
        f"[Chart 3] 축산물품질평가원(EKAPE) 한우 지육시세 변동률 및 6대 플랫폼 냉장 원물 연동성 분석 (기준경매일: {ek_base_dt})",
        fontsize=17,
        fontweight="bold",
        color="#0F172A",
        y=0.96,
    )
    fig3.text(
        0.5,
        0.925,
        f"한우거세 평균 경락가 {ekape_summary.get('geose_avg', 0):,}원/kg | 한우암소 평균 경락가 {ekape_summary.get('amso_avg', 0):,}원/kg 기준 시점별(전일·전주·전월·전년 동기) 등락률과 플랫폼별 지육배율을 비교합니다.",
        ha="center",
        fontsize=11,
        color="#475569",
    )

    ek_by_key = ekape_summary.get("by_key", {})
    carcass_grades = ["전국평균", "1++(9)", "1++(8)", "1++(7)", "1+", "1", "2", "3"]

    # Panel (0,0) & (0,1): 한우거세 / 한우암소 등급별 지육시세 시점별(전일/전주/전월/전년) 증감률(%)
    for idx_c, (c_lbl, panel_title) in enumerate([
        ("한우거세", "① [한우거세] 등급별 지육 경락가격 시점별(전일·전주·전월·전년) 변동률(%)"),
        ("한우암소", "② [한우암소] 등급별 지육 경락가격 시점별(전일·전주·전월·전년) 변동률(%)"),
    ]):
        ax = fig3.add_subplot(gs3[0, idx_c])
        ax.set_facecolor("#FFFFFF")
        x_g = np.arange(len(carcass_grades))
        w_g = 0.19

        metrics_cfg = [
            ("전일대비_증감률(%)", "전일 대비", "#2563EB"),
            ("전주대비_증감률(%)", "전주 동요일 대비", "#059669"),
            ("전월동기_증감률(%)", "전월 동기 대비", "#D97706"),
            ("전년동기_증감률(%)", "전년 동기(YoY) 대비", "#DC2626"),
        ]

        for m_i, (col_nm, m_lbl, m_color) in enumerate(metrics_cfg):
            vals = []
            for g_nm in carcass_grades:
                erow = ek_by_key.get((c_lbl, g_nm)) or {}
                vals.append(float(erow.get(col_nm) or 0.0))
            pos = x_g + (m_i - 1.5) * w_g
            bars = ax.bar(pos, vals, width=w_g * 0.9, color=m_color, alpha=0.86, label=m_lbl)
            for b, v in zip(bars, vals):
                if abs(v) >= 0.5:
                    ax.text(
                        b.get_x() + b.get_width() / 2.0,
                        v + (0.35 if v >= 0 else -1.1),
                        f"{v:+.1f}%",
                        ha="center",
                        va="bottom" if v >= 0 else "top",
                        fontsize=7.3,
                        fontweight="bold",
                        color="#1E293B",
                    )

        # 상단에 당일 경락단가(원/kg) 텍스트 표시
        for g_i, g_nm in enumerate(carcass_grades):
            erow = ek_by_key.get((c_lbl, g_nm)) or {}
            cur_p = erow.get("기준일_경락가(원/kg)")
            if cur_p:
                ax.text(
                    x_g[g_i],
                    ax.get_ylim()[1] * 0.92 if ax.get_ylim()[1] > 0 else 5.0,
                    f"{cur_p:,}원",
                    ha="center",
                    fontsize=8,
                    fontweight="bold",
                    color="#0F172A",
                    bbox=dict(boxstyle="round,pad=0.15", facecolor="#F1F5F9", edgecolor="#CBD5E1"),
                )

        ax.axhline(0, color="#0F172A", linewidth=1.0)
        ax.set_xticks(x_g)
        ax.set_xticklabels(carcass_grades, fontsize=9.5, fontweight="bold")
        ax.set_ylabel("경락가격 증감률 (%)", fontsize=10, fontweight="bold")
        ax.set_title(panel_title, fontsize=12, fontweight="bold", pad=10)
        ax.grid(axis="y", linestyle=":", alpha=0.6)
        ax.legend(loc="upper left", fontsize=8.5, ncol=2, framealpha=0.9)

    # Panel (1,0): 등급별 지육경락가(원/kg) 대비 6대 플랫폼 [냉장 원물] 평균 판매단가 추이 연동성
    ax3_3 = fig3.add_subplot(gs3[1, 0])
    ax3_3.set_facecolor("#FFFFFF")
    link_grades = ["1++(9)", "1++(8)", "1++(7)", "1+", "1", "2"]
    x_l = np.arange(len(link_grades))

    # 지육경락가(거세/암소 평균) 기준선 플롯
    carcass_curve = []
    for g_nm in link_grades:
        p_g = (ek_by_key.get(("한우거세", g_nm)) or {}).get("기준일_경락가(원/kg)") or 24000
        p_a = (ek_by_key.get(("한우암소", g_nm)) or {}).get("기준일_경락가(원/kg)") or 21000
        carcass_curve.append(((p_g + p_a) / 2.0) / 1000.0)

    ax3_3.plot(
        x_l,
        carcass_curve,
        color="#0F172A",
        linestyle="--",
        marker="s",
        linewidth=2.4,
        markersize=7,
        label="EKAPE 지육경락가 (거세·암소 평균)",
        zorder=6,
    )

    for p_nm in PLATFORMS_ORDER:
        ys = []
        valid_xs = []
        for idx_g, g_nm in enumerate(link_grades):
            p_vals = [
                r.unit_price_kg / 1000.0
                for r in all_rows
                if r.platform == p_nm and r.keeping == "냉장" and r.grade == g_nm
            ]
            if p_vals:
                valid_xs.append(x_l[idx_g])
                ys.append(float(np.mean(p_vals)))
        if valid_xs:
            ax3_3.plot(
                valid_xs,
                ys,
                color=plat_colors[p_nm],
                marker="o",
                linewidth=2.6 if p_nm == "근본한우" else 1.8,
                markersize=8 if p_nm == "근본한우" else 5.5,
                label=f"{p_nm} (냉장원물 평균)",
                zorder=5 if p_nm == "근본한우" else 4,
            )

    ax3_3.set_xticks(x_l)
    ax3_3.set_xticklabels(link_grades, fontsize=10, fontweight="bold")
    ax3_3.set_ylabel("단가 (천원/kg)", fontsize=10, fontweight="bold")
    ax3_3.set_title("③ 등급별 지육 경락단가 vs 6대 플랫폼 [냉장 원물] 평균 판매가 연동성", fontsize=12, fontweight="bold", pad=10)
    ax3_3.grid(True, linestyle=":", alpha=0.6)
    ax3_3.legend(loc="upper right", fontsize=8.5, ncol=2, framealpha=0.9)

    # Panel (1,1): 부위별 [냉장 원물 판매단가 ÷ 지육경락가] 지육배율(Carcass Multiplier) 6사 비교
    ax3_4 = fig3.add_subplot(gs3[1, 1])
    ax3_4.set_facecolor("#FFFFFF")
    mult_parts = ["안심", "등심", "채끝", "갈비", "양지", "차돌박이", "설도", "사태", "우둔", "앞다리"]
    x_m = np.arange(len(mult_parts))
    w_m = 0.13

    for p_i, p_nm in enumerate(PLATFORMS_ORDER):
        mults = []
        v_xm = []
        for idx_pt, pt_nm in enumerate(mult_parts):
            pt_rows = [
                r for r in all_rows
                if r.platform == p_nm and r.keeping == "냉장" and r.part_name == pt_nm
            ]
            if pt_rows:
                ratios = []
                for r in pt_rows:
                    ek_g = "1++" if r.grade == "1++(불명)" else r.grade
                    c_p = (ek_by_key.get((r.cattle_type, ek_g)) or {}).get("기준일_경락가(원/kg)") or 23000
                    if c_p > 0:
                        ratios.append(r.unit_price_kg / float(c_p))
                if ratios:
                    v_xm.append(x_m[idx_pt] + (p_i - 2.5) * w_m)
                    mults.append(float(np.mean(ratios)))
        if v_xm:
            ax3_4.bar(
                v_xm,
                mults,
                width=w_m * 0.9,
                color=plat_colors[p_nm],
                alpha=0.9 if p_nm == "근본한우" else 0.78,
                label=p_nm,
            )

    ax3_4.axhline(1.0, color="#64748B", linestyle=":", linewidth=1.2)
    ax3_4.set_xticks(x_m)
    ax3_4.set_xticklabels(mult_parts, fontsize=9.5, fontweight="bold")
    ax3_4.set_ylabel("지육 경락가 대비 원물 판매 배율 (배)", fontsize=10, fontweight="bold")
    ax3_4.set_title("④ 주요 부위별 지육시세 대비 냉장 원물 판매배율 (원물단가 ÷ 등급별 지육경락가)", fontsize=12, fontweight="bold", pad=10)
    ax3_4.grid(axis="y", linestyle=":", alpha=0.6)
    ax3_4.legend(loc="upper right", fontsize=8.5, ncol=3, framealpha=0.9)

    fig3.savefig(chart3_path, dpi=160)

    # =========================================================================
    # 통합 PDF 리포트 생성 (표지/전체품목 전수 비교표 다중 페이지 + Chart 1 + Chart 2 + Chart 3)
    # =========================================================================
    full_item_rows = analytics_meta.get("full_item_rows", [])
    with PdfPages(pdf_path) as pdf:
        # Page 1: Executive Summary Page
        fig_cov = plt.figure(figsize=(20, 12.5), dpi=160)
        fig_cov.patch.set_facecolor("#0F172A")
        ax_cov = fig_cov.add_subplot(111)
        ax_cov.set_facecolor("#0F172A")
        ax_cov.axis("off")

        ax_cov.text(
            0.05, 0.88,
            f"근본한우 온라인팀 MD 전용 | 한우 원물(박스육) 전체 품목 전수 비교 & 지육시세 분석 리포트 (미트박스 제외)",
            fontsize=21, fontweight="bold", color="#F8FAFC",
        )
        ax_cov.text(
            0.05, 0.83,
            f"수집 및 분석 일자: {today_str}  |  비교 대상: 근본한우(자사) vs 미트프렌즈 vs 탑미트 vs 금천미트 vs 본대로 (원물 1kg 단가 기준 · 전체 품목 전수 나열)",
            fontsize=12.5, color="#94A3B8",
        )

        # 핵심 KPI 요약 박스 텍스트
        plat_cnt_str = "  |  ".join(f"{k}: {v:,}건" for k, v in analytics_meta["platform_counts"].items())
        gb_match = analytics_meta["gb_matched_conditions"]
        gb_win = analytics_meta["gb_cheaper_conditions"]
        win_rate = (gb_win / gb_match * 100.0) if gb_match > 0 else 0.0
        uniq_brands = analytics_meta.get("total_unique_brands", 0)
        reclass_cnt = analytics_meta.get("stage2_reclassified_rows", 0)
        tot_item_grps = analytics_meta.get("total_item_groups", len(full_item_rows))

        summary_lines = [
            f"■ 1. 타겟 4대 B2B 플랫폼(미트프렌즈·탑미트·금천미트·본대로) + 근본한우 원물 전수 크롤링: 총 {analytics_meta['total_raw_rows']:,}건 (고유 브랜드 {uniq_brands:,}개)",
            f"   - 플랫폼별 원물 옵션/박스 수: {plat_cnt_str}  |  플랫폼 간 상이한 상품명 2차 교차검증 정밀분리: {reclass_cnt:,}건",
            f"■ 2. 데이터 왜곡 원천 차단 및 전체 품목 전수 분리 나열 (총 {tot_item_grps}개 축종×보관×표준품목 전수 비교)",
            f"   - 단가 체급이 다른 세부 품목(등심/윗등심/윗등심(살치제외)/아랫등심/꽃등심/살치살, 짝갈비/갈비살/늑간살/갈비본살, 설도/설깃살, 우둔/홍두깨 등) 개별 분리",
            f"   - 미트프렌즈 재고 0건 더미 옵션(할랄한우 고정단가) 및 괄호 내 부정문구('꾸리,부채 미포함') 오분류 왜곡 100% 제거",
            f"■ 3. 축산물품질평가원(EKAPE) 한우 지육 경락가격 동향 (기준 경매일: {ek_base_dt})",
            f"   - [한우거세] 전국 평균: {ekape_summary.get('geose_avg', 0):,}원/kg  |  [한우암소] 전국 평균: {ekape_summary.get('amso_avg', 0):,}원/kg",
            f"■ 4. 다차원 그룹핑 집계 ([축종 × 보관상태 × 세부품목 × 등급 × 생산일자 × 브랜드]): 총 {analytics_meta['total_condition_groups']:,}개 품목×등급 조건 전수 수록",
            f"   - 근본한우 입점 {gb_match}개 동일 품목·등급 조건 중 타사 평균 대비 근본한우 가격 우위(저렴) 조건: {gb_win}개 ({win_rate:.1f}%)",
        ]
        y_cursor = 0.74
        for line in summary_lines:
            ax_cov.text(
                0.05, y_cursor, line,
                fontsize=11.8 if line.startswith("■") else 10.8,
                fontweight="bold" if line.startswith("■") else "normal",
                color="#38BDF8" if line.startswith("■") else "#E2E8F0",
            )
            y_cursor -= 0.048

        pdf.savefig(fig_cov)
        plt.close(fig_cov)

        pdf.savefig(fig1)
        pdf.savefig(fig2)
        pdf.savefig(fig3)

        # 전체 품목 전수 비교표 (요약 생략 없이 전체 품목 100% 다중 페이지 테이블 수록)
        rows_per_page = 24
        for page_i in range(0, len(full_item_rows), rows_per_page):
            chunk = full_item_rows[page_i:page_i + rows_per_page]
            fig_tbl = plt.figure(figsize=(20, 12.5), dpi=160)
            fig_tbl.patch.set_facecolor("#FFFFFF")
            ax_t = fig_tbl.add_subplot(111)
            ax_t.axis("off")
            page_no = (page_i // rows_per_page) + 1
            tot_pages = (len(full_item_rows) + rows_per_page - 1) // rows_per_page
            ax_t.set_title(
                f"[전체 품목 전수 비교표 {page_no}/{tot_pages}] 축종 × 보관상태 × 전체 품목별 플랫폼 비교 (미트박스 제외 · 총 {len(full_item_rows)}개 품목 전수)",
                fontsize=14, fontweight="bold", color="#0F172A", pad=12,
            )
            col_labels_all = [
                "축종", "보관", "품목명(정밀분리)", "전체건수",
                "근본한우(최저~최고/평균)",
                "미트프렌즈(최저~최고/평균)",
                "탑미트(최저~최고/평균)",
                "금천미트(최저~최고/평균)",
                "본대로(최저~최고/평균)",
                "품목 내 전체 최저가 브랜드(등급/생산일)",
                "품목 내 전체 최고가 브랜드(등급/생산일)",
            ]

            def _fmt_plat_range(ir: Dict[str, Any], p_prefix: str) -> str:
                cnt_v = ir.get(f"{p_prefix}_건수", 0)
                mn_v = ir.get(f"{p_prefix}_최저가(원/kg)")
                av_v = ir.get(f"{p_prefix}_평균가(원/kg)")
                mx_v = ir.get(f"{p_prefix}_최고가(원/kg)")
                if not mn_v or not av_v:
                    return "-"
                if mn_v == mx_v:
                    return f"{av_v:,}원 ({cnt_v}건)"
                return f"{mn_v:,}~{mx_v:,} (평균 {av_v:,}/{cnt_v}건)"

            cell_rows = []
            for ir in chunk:
                cell_rows.append([
                    ir["축종구분"],
                    ir["보관상태"],
                    ir["품목명"],
                    f"{ir['전체_수집건수']:,}건",
                    _fmt_plat_range(ir, "근본한우"),
                    _fmt_plat_range(ir, "미트프렌즈"),
                    _fmt_plat_range(ir, "탑미트"),
                    _fmt_plat_range(ir, "금천미트"),
                    _fmt_plat_range(ir, "본대로"),
                    str(ir.get("품목내_전체최저가_상세") or "-")[:34],
                    str(ir.get("품목내_전체최고가_상세") or "-")[:34],
                ])
            tbl_p = ax_t.table(
                cellText=cell_rows,
                colLabels=col_labels_all,
                bbox=[0.01, 0.02, 0.98, 0.92],
                cellLoc="center",
            )
            tbl_p.auto_set_font_size(False)
            tbl_p.set_fontsize(7.6)
            for (r_i, c_i), cell in tbl_p.get_celld().items():
                if r_i == 0:
                    cell.set_facecolor("#1E293B")
                    cell.set_text_props(color="#FFFFFF", fontweight="bold")
                else:
                    cell.set_facecolor("#F8FAFC" if r_i % 2 == 1 else "#FFFFFF")
                    cell.set_text_props(color="#0F172A")
                cell.set_edgecolor("#CBD5E1")
            pdf.savefig(fig_tbl)
            plt.close(fig_tbl)

    plt.close(fig1)
    plt.close(fig2)
    plt.close(fig3)

    # =========================================================================
    # 인터랙티브 HTML 리포트 생성 (차트 이미지 임베딩 + 전체 품목 전수 비교표 100% 수록)
    # =========================================================================
    def _img_to_b64(p: Path) -> str:
        return base64.b64encode(p.read_bytes()).decode("ascii")

    b64_c1 = _img_to_b64(chart1_path)
    b64_c2 = _img_to_b64(chart2_path)
    b64_c3 = _img_to_b64(chart3_path)

    tr_ekape_html = []
    for er in ekape_rows:
        tr_ekape_html.append(
            f"<tr>"
            f"<td>{er['축종구분']}</td><td><b>{er['등급']}</b></td>"
            f"<td style='text-align:right;font-weight:700;'>{int(er['기준일_경락가(원/kg)'] or 0):,}원</td>"
            f"<td style='text-align:right;'>{int(er['기준일_경매두수(두)'] or 0):,}두</td>"
            f"<td style='text-align:right;color:{'#DC2626' if (er['전일대비_증감률(%)'] or 0) > 0 else '#2563EB'};'>{er['전일대비_증감액(원)'] or 0:+,}원 ({er['전일대비_증감률(%)'] or 0:+.2f}%)</td>"
            f"<td style='text-align:right;color:{'#DC2626' if (er['전주대비_증감률(%)'] or 0) > 0 else '#2563EB'};'>{er['전주대비_증감액(원)'] or 0:+,}원 ({er['전주대비_증감률(%)'] or 0:+.2f}%)</td>"
            f"<td style='text-align:right;color:{'#DC2626' if (er['전월동기_증감률(%)'] or 0) > 0 else '#2563EB'};'>{er['전월동기_증감액(원)'] or 0:+,}원 ({er['전월동기_증감률(%)'] or 0:+.2f}%)</td>"
            f"<td style='text-align:right;color:{'#DC2626' if (er['전년동기_증감률(%)'] or 0) > 0 else '#2563EB'};'>{er['전년동기_증감액(원)'] or 0:+,}원 ({er['전년동기_증감률(%)'] or 0:+.2f}%)</td>"
            f"</tr>"
        )

    def _fmt_price(v: Any) -> str:
        return f"{int(v):,}" if isinstance(v, (int, float)) and v else "-"

    def _fmt_plat_html(ir: Dict[str, Any], p_nm: str) -> str:
        cnt = ir.get(f"{p_nm}_건수", 0)
        mn = ir.get(f"{p_nm}_최저가(원/kg)")
        av = ir.get(f"{p_nm}_평균가(원/kg)")
        mx = ir.get(f"{p_nm}_최고가(원/kg)")
        if not mn or not av:
            return "-"
        b_min = ir.get(f"{p_nm}_최저가_브랜드(생산일)", "")
        rng_s = f"{mn:,}원" if mn == mx else f"{mn:,} ~ {mx:,}원"
        sub_b = f"<br/><span style='font-size:10.5px;color:#64748B;'>최저:{b_min}</span>" if b_min and b_min != "-" else ""
        return f"<b>평균 {av:,}원</b> ({cnt}건)<br/><span style='font-size:11px;color:#334155;'>{rng_s}</span>{sub_b}"

    # Table 2A: 전체 품목 전수 비교표 (요약 생략 없이 100% 전체 품목 표시)
    tr_full_item_html = []
    for ir in full_item_rows:
        tr_full_item_html.append(
            f"<tr>"
            f"<td>{ir['축종구분']}</td><td>{ir['보관상태']}</td><td>{ir['대분류부위(10대기준)']}</td>"
            f"<td style='font-weight:700;color:#0F172A;'>{ir['품목명']}</td>"
            f"<td style='font-size:11px;'>{ir['취급등급_분포']}</td>"
            f"<td style='text-align:right;font-weight:700;'>{ir['전체_수집건수']:,}건</td>"
            f"<td style='text-align:right;background:#FEF2F2;'>{_fmt_plat_html(ir, '근본한우')}</td>"
            f"<td style='text-align:right;'>{_fmt_plat_html(ir, '미트프렌즈')}</td>"
            f"<td style='text-align:right;'>{_fmt_plat_html(ir, '탑미트')}</td>"
            f"<td style='text-align:right;'>{_fmt_plat_html(ir, '금천미트')}</td>"
            f"<td style='text-align:right;'>{_fmt_plat_html(ir, '본대로')}</td>"
            f"<td style='font-size:11px;text-align:left;color:#1E40AF;'>{ir['품목내_전체최저가_상세']}</td>"
            f"<td style='font-size:11px;text-align:left;color:#991B1B;'>{ir['품목내_전체최고가_상세']}</td>"
            f"</tr>"
        )

    # Table 2B: 전체 품목 × 등급별 상세 전수 비교표 (상위 100개 제한 해제 -> 100% 전수 출력!)
    tr_dash_html = []
    for r in dashboard_rows:
        gb_a = _fmt_price(r["근본한우_평균가(원/kg)"])
        cp_a = _fmt_price(r["타사통합_평균가(원/kg)"])
        mf_a = _fmt_price(r.get("미트프렌즈_평균가(원/kg)"))
        tm_a = _fmt_price(r.get("탑미트_평균가(원/kg)"))
        ek_a = _fmt_price(r.get("금천미트_평균가(원/kg)"))
        bd_a = _fmt_price(r.get("본대로_평균가(원/kg)"))
        dv = r["근본한우대비_타사평균_편차(%)"]
        dv_str = f"{dv:+.2f}%" if dv is not None else "-"
        dv_col = "#059669" if (dv is not None and dv > 0) else ("#DC2626" if (dv is not None and dv < 0) else "#64748B")
        tr_dash_html.append(
            f"<tr>"
            f"<td>{r['축종구분']}</td><td>{r['보관상태']}</td><td><b>{r['부위명']}</b></td><td><b>{r['등급']}</b></td>"
            f"<td style='text-align:right;background:#FEF2F2;font-weight:700;color:#B91C1C;'>{gb_a}</td>"
            f"<td style='text-align:right;font-weight:700;'>{cp_a}</td>"
            f"<td style='text-align:right;font-weight:700;color:{dv_col};'>{dv_str}</td>"
            f"<td style='font-size:11.5px;text-align:left;color:#1E40AF;'>{r['생산일자대비_최저가_상세']}</td>"
            f"<td style='font-size:11.5px;text-align:left;color:#991B1B;'>{r['생산일자대비_최고가_상세']}</td>"
            f"<td style='text-align:right;'>{mf_a}<br/><span style='font-size:10.5px;color:#64748B;'>{r.get('미트프렌즈_최저가_브랜드(생산일)', '-')}</span></td>"
            f"<td style='text-align:right;'>{tm_a}<br/><span style='font-size:10.5px;color:#64748B;'>{r.get('탑미트_최저가_브랜드(생산일)', '-')}</span></td>"
            f"<td style='text-align:right;'>{ek_a}<br/><span style='font-size:10.5px;color:#64748B;'>{r.get('금천미트_최저가_브랜드(생산일)', '-')}</span></td>"
            f"<td style='text-align:right;'>{bd_a}<br/><span style='font-size:10.5px;color:#64748B;'>{r.get('본대로_최저가_브랜드(생산일)', '-')}</span></td>"
            f"<td style='font-size:11.5px;text-align:left;'>{r['브랜드_가격라인업(최저~최고)']}</td>"
            f"</tr>"
        )

    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8"/>
<title>{today_str} 한우원물 전체품목 전수비교 시각화리포트 (미트박스 제외)</title>
<style>
  body {{ font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; background:#F1F5F9; color:#0F172A; margin:0; padding:24px; }}
  .container {{ max-width: 1720px; margin: 0 auto; }}
  .header {{ background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); color:#FFF; padding:28px 32px; border-radius:14px; margin-bottom:24px; }}
  .header h1 {{ margin:0 0 8px 0; font-size:24px; }}
  .header p {{ margin:0; color:#CBD5E1; font-size:14px; }}
  .card {{ background:#FFF; border-radius:12px; padding:24px; margin-bottom:24px; box-shadow:0 2px 10px rgba(15,23,42,0.06); overflow-x:auto; }}
  .card h2 {{ margin-top:0; font-size:18px; border-left:5px solid #DC2626; padding-left:12px; color:#0F172A; }}
  .chart-img {{ width:100%; height:auto; border-radius:8px; border:1px solid #E2E8F0; }}
  table {{ width:100%; border-collapse:collapse; font-size:12px; margin-top:12px; }}
  th {{ background:#1E293B; color:#FFF; padding:10px 6px; text-align:center; position:sticky; top:0; }}
  td {{ padding:7px 6px; border-bottom:1px solid #E2E8F0; text-align:center; }}
  tr:hover {{ background:#F8FAFC; }}
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>근본한우 온라인팀 MD | 한우 원물(박스육) 전체 품목 전수 비교 & 지육시세 분석 리포트 ({today_str}, 미트박스 제외)</h1>
    <p>비교 플랫폼: 근본한우(자사) vs 미트프렌즈 vs 탑미트 vs 금천미트 vs 본대로 | 총 수집 원물 옵션/박스: {analytics_meta['total_raw_rows']:,}건 | 전체 품목 그룹: {len(full_item_rows)}개 | 전체 품목×등급 그룹: {len(dashboard_rows)}개 | 고유 브랜드: {uniq_brands:,}개</p>
  </div>
  <div class="card">
    <h2>[Table 2A] 전체 품목 전수 비교표 (축종 × 보관상태 × 전체 {len(full_item_rows)}개 표준품목 전수 나열 · 요약/생략 없음)</h2>
    <table>
      <thead>
        <tr>
          <th>축종</th><th>보관</th><th>대분류</th><th>품목명(정밀분리)</th><th>취급등급</th><th>총수집건수</th>
          <th>근본한우 (최저~최고/평균)</th>
          <th>미트프렌즈 (최저~최고/평균)</th>
          <th>탑미트 (최저~최고/평균)</th>
          <th>금천미트 (최저~최고/평균)</th>
          <th>본대로 (최저~최고/평균)</th>
          <th>품목 내 최저가 상세 (브랜드/등급/생산일)</th>
          <th>품목 내 최고가 상세 (브랜드/등급/생산일)</th>
        </tr>
      </thead>
      <tbody>{"".join(tr_full_item_html)}</tbody>
    </table>
  </div>
  <div class="card">
    <h2>[Table 2B] 전체 품목 × 등급별 상세 전수 비교표 (총 {len(dashboard_rows)}개 품목×등급 조건 100% 전수 나열)</h2>
    <table>
      <thead>
        <tr>
          <th>축종</th><th>보관</th><th>품목명(정밀분리)</th><th>등급</th>
          <th>근본한우 평균</th><th>4사통합 평균</th><th>근본대비 편차</th>
          <th>생산일자대비 최저가 상세</th><th>생산일자대비 최고가 상세</th>
          <th>미트프렌즈(최저브랜드)</th><th>탑미트(최저브랜드)</th><th>금천미트(최저브랜드)</th><th>본대로(최저브랜드)</th>
          <th>동일조건 브랜드 라인업 (최저가~최고가)</th>
        </tr>
      </thead>
      <tbody>{"".join(tr_dash_html)}</tbody>
    </table>
  </div>
  <div class="card">
    <h2>[Chart 1] 주요 부위·등급·보관상태별 플랫폼 [최저가 - 평균가 - 최고가] 가격 분포 및 최저·최고 브랜드</h2>
    <img class="chart-img" src="data:image/png;base64,{b64_c1}" alt="Chart 1"/>
  </div>
  <div class="card">
    <h2>[Chart 2] 원물 생산일자 대비 최저가·최고가 가격 밴드 및 플랫폼·브랜드별 단가 스캐터/히트맵</h2>
    <img class="chart-img" src="data:image/png;base64,{b64_c2}" alt="Chart 2"/>
  </div>
  <div class="card">
    <h2>[Chart 3] 한우거세·암소 등급별 지육시세 변동률(전일/전주/전월/전년) 및 냉장 원물 연동성 분석</h2>
    <img class="chart-img" src="data:image/png;base64,{b64_c3}" alt="Chart 3"/>
  </div>
  <div class="card">
    <h2>[Table 1] 축산물품질평가원(EKAPE) 한우거세·한우암소 등급별 지육 경락가격 동향 (기준경매일: {ek_base_dt})</h2>
    <table>
      <thead>
        <tr>
          <th>축종구분</th><th>등급</th><th>기준일 경락가(원/kg)</th><th>경매두수</th>
          <th>전일 대비</th><th>전주 동요일 대비</th><th>전월 동기 대비</th><th>전년 동기(YoY) 대비</th>
        </tr>
      </thead>
      <tbody>{"".join(tr_ekape_html)}</tbody>
    </table>
  </div>
</div>
</body>
</html>"""
    html_path.write_text(html_content, encoding="utf-8")

    logger.info("시각화 결과물 생성 완료: PDF=%s, HTML=%s", pdf_path.name, html_path.name)
    return {
        "chart1": chart1_path,
        "chart2": chart2_path,
        "chart3": chart3_path,
        "pdf": pdf_path,
        "html": html_path,
    }


# ---------------------------------------------------------------------------
# [Task 5] 마스터 엑셀 누적 저장 (Append & Archive) 및 구글 드라이브 동기화
# ---------------------------------------------------------------------------
import shutil
GDRIVE_BIN = shutil.which("gdrive") or r"C:\Users\limkyungsoon\.gemini\antigravity\bin\gdrive.cmd"


def find_gdrive_file_id(folder_id: str, filename: str) -> Optional[str]:
    """지정된 구글 드라이브 폴더 내 특정 파일명의 file_id를 조회합니다."""
    try:
        cmd = [
            GDRIVE_BIN, "readonly", "search",
            "--parent-id", folder_id,
            "--name-exact", filename,
            "--json",
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=30)
        if res.returncode == 0 and res.stdout.strip():
            data = json.loads(res.stdout)
            files = data.get("files") if isinstance(data, dict) else data
            if isinstance(files, list) and files:
                return str(files[0].get("id") or "")
    except Exception as e:
        logger.warning("구글 드라이브 파일 검색 경고 (%s): %s", filename, e)
    return None


def download_existing_master_from_gdrive(folder_id: str, local_master_path: Path) -> Optional[str]:
    """
    구글 드라이브 폴더에 이미 마스터 엑셀 파일이 존재하면 로컬로 다운로드하여
    과거 일자 누적 데이터를 보존한 채 당일 데이터를 Append합니다.
    """
    file_id = find_gdrive_file_id(folder_id, MASTER_EXCEL_FILENAME)
    if file_id and not local_master_path.exists():
        logger.info("구글 드라이브에서 기존 마스터 파일(%s) 감지 -> 누적 병합을 위해 다운로드 중...", file_id)
        try:
            cmd = [
                GDRIVE_BIN, "readonly", "download", file_id,
                "--dest", str(local_master_path.parent),
                "--overwrite",
            ]
            subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=60)
        except Exception as e:
            logger.warning("기존 마스터 파일 다운로드 실패 (로컬 신규 생성으로 진행): %s", e)
    return file_id


def _read_existing_sheet_rows(wb: openpyxl.Workbook, sheet_name: str, date_col_name: str, today_str: str) -> List[Dict[str, Any]]:
    """
    기존 워크북의 특정 시트에서 오늘 날짜(today_str)가 아닌 과거 누적 행들을 읽어 반환합니다.
    (당일 재실행 시 중복 적재를 방지하면서 과거 일자 데이터는 100% 보존하는 Idempotent Append 방식)
    """
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows_iter = list(ws.iter_rows(values_only=True))
    if not rows_iter:
        return []

    # 헤더 행 탐색 (첫 5행 이내에서 date_col_name을 포함하는 행을 헤더로 인식)
    header_idx = None
    headers: List[str] = []
    for idx, r in enumerate(rows_iter[:5]):
        r_strs = [str(c).strip() if c is not None else "" for c in r]
        if date_col_name in r_strs:
            header_idx = idx
            headers = r_strs
            break

    if header_idx is None:
        return []

    past_rows: List[Dict[str, Any]] = []
    for r in rows_iter[header_idx + 1:]:
        if not any(r):
            continue
        row_dict = {headers[i]: (r[i] if i < len(r) else None) for i in range(len(headers)) if headers[i]}
        for k, v in list(row_dict.items()):
            if "(%)" in k and isinstance(v, (int, float)):
                if abs(float(v)) <= 3.5:
                    row_dict[k] = round(float(v) * 100.0, 2)
                else:
                    row_dict[k] = round(float(v), 2)
        dt_val = str(row_dict.get(date_col_name) or "").strip()
        if dt_val and dt_val != today_str:
            past_rows.append(row_dict)
    return past_rows


def _write_styled_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    title_banner: str,
    headers: List[str],
    rows_data: List[Dict[str, Any]],
    currency_cols: Set[str],
    pct_cols: Set[str],
    highlight_col: Optional[str] = None,
) -> None:
    """시트를 생성(또는 재작성)하고 임원 보고용 서식·열너비·틀고정·숫자포맷을 일괄 적용합니다."""
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    # 1행: 타이틀 배너
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    banner_cell = ws.cell(row=1, column=1, value=title_banner)
    banner_cell.font = Font(name="맑은 고딕", size=12, bold=True, color="FFFFFF")
    banner_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    banner_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28

    # 2행: 컬럼 헤더
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    gb_header_fill = PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    ws.row_dimensions[2].height = 25
    for col_idx, h_name in enumerate(headers, start=1):
        c = ws.cell(row=2, column=col_idx, value=h_name)
        c.font = header_font
        c.fill = gb_header_fill if "근본한우" in h_name else header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border

    # 3행~: 데이터 행
    body_font = Font(name="맑은 고딕", size=9.5, color="0F172A")
    bold_font = Font(name="맑은 고딕", size=9.5, bold=True, color="0F172A")
    gb_cell_fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
    pos_font = Font(name="맑은 고딕", size=9.5, bold=True, color="059669")
    neg_font = Font(name="맑은 고딕", size=9.5, bold=True, color="DC2626")
    align_right = Alignment(horizontal="right", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")

    # 대용량 시트(3,000행 초과: Sheet 3, Sheet 4, Sheet 5)는 고속 append 방식으로 기록하여 저장 시간을 수분 -> 수초로 단축
    if len(rows_data) > 3000:
        for r_dict in rows_data:
            row_vals = []
            for h_name in headers:
                val = r_dict.get(h_name)
                if h_name in pct_cols and isinstance(val, (int, float)):
                    val = round(float(val), 2)
                row_vals.append(val)
            ws.append(row_vals)
    else:
        for r_idx, r_dict in enumerate(rows_data, start=3):
            is_gb_row = highlight_col and str(r_dict.get(highlight_col) or "") == "근본한우"
            for c_idx, h_name in enumerate(headers, start=1):
                val = r_dict.get(h_name)
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                cell.font = bold_font if (is_gb_row or "근본한우" in h_name) else body_font

                if is_gb_row or "근본한우" in h_name:
                    cell.fill = gb_cell_fill

                if h_name in currency_cols and isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                    cell.alignment = align_right
                elif h_name in pct_cols and isinstance(val, (int, float)):
                    cell.number_format = "+0.00%;-0.00%;0.00%"
                    cell.value = float(val) / 100.0
                    cell.alignment = align_right
                    cell.font = pos_font if val > 0 else (neg_font if val < 0 else body_font)
                elif isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                    cell.alignment = align_right
                elif any(
                    k in h_name
                    for k in (
                        "상품URL", "브랜드_가격라인업", "최저가_플랫폼", "최고가_플랫폼", "원본상품명",
                        "상세", "최저~최고", "수행내역", "판정결과",
                    )
                ):
                    cell.alignment = align_left
                else:
                    cell.alignment = align_center

    # 틀 고정 및 오토필터
    ws.freeze_panes = "F3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{max(2, len(rows_data) + 2)}"

    # 열 너비 자동 최적화 (최대 500행 샘플링으로 속도 확보)
    sample_rows = rows_data[:400]
    for c_idx, h_name in enumerate(headers, start=1):
        max_len = len(h_name) * 1.65
        for r_d in sample_rows:
            v_str = str(r_d.get(h_name) or "")
            # 한글 폭 가중치 계산
            kor_cnt = sum(1 for ch in v_str if "\uac00" <= ch <= "\ud7a3")
            cur_w = len(v_str) + kor_cnt * 0.65
            if cur_w > max_len:
                max_len = cur_w
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max(max_len + 3.0, 11.0), 48.0)


def save_and_append_master_excel(
    local_master_path: Path,
    dashboard_rows: List[Dict[str, Any]],
    ekape_rows: List[Dict[str, Any]],
    sheet3_rows: List[Dict[str, Any]],
    all_raw_rows: List[HanwooProductRow],
    today_str: str,
    full_item_rows: Optional[List[Dict[str, Any]]] = None,
) -> Path:
    """
    [Task 5] `[누적_기초데이터]_한우원물_시세_및_플랫폼비교_마스터.xlsx` 파일을 생성/업데이트합니다.
    기존 파일이 있을 경우 과거 수집일자 데이터를 보존하고 당일 데이터를 하단에 누적(Append & Archive)합니다.
    (미트박스 제외 요청에 따라 미트박스 컬럼 및 타사 미트박스 행은 제외하고 5개 플랫폼 기준으로 구성)
    - Sheet 0: `전체품목_플랫폼별_전수비교` (요약/생략 없이 축종×보관상태×전체품목 전수 나열 비교표)
    - Sheet 1: `Dashboard_요약` (전체 품목×등급별 전수 비교표: 브랜드명 및 생산일자 대비 최저가·최고가 포함)
    - Sheet 2: `지육시세_추이`
    - Sheet 3: `생산일자_등급_보관별_통계` (동일 생산일자 구간 대비 최저가·최고가 브랜드 비교 포함)
    - Sheet 4: `Raw_Data_원물_누적` (2차검증 세부표준상품명 및 판정결과 포함)
    - Sheet 5: `상품명_2차검증_매핑표` (플랫폼별 상이한 상품명 1·2차 교차검증 결과표)
    """
    logger.info("=== [Task 5] 마스터 엑셀 6개 시트 누적 저장(Append & Archive, 미트박스 제외) 작업 시작 ===")

    if local_master_path.exists():
        try:
            wb_old = openpyxl.load_workbook(local_master_path, read_only=True, data_only=True)
        except Exception as e:
            logger.warning("기존 마스터 엑셀 로드 실패(신규 워크북으로 대체): %s", e)
            wb_old = openpyxl.Workbook()
    else:
        wb_old = openpyxl.Workbook()

    past_s0 = _read_existing_sheet_rows(wb_old, "전체품목_플랫폼별_전수비교", "수집일자", today_str)
    past_s1 = _read_existing_sheet_rows(wb_old, "Dashboard_요약", "수집일자", today_str)
    past_s2 = _read_existing_sheet_rows(wb_old, "지육시세_추이", "수집일자", today_str)
    past_s3 = _read_existing_sheet_rows(wb_old, "생산일자_등급_보관별_통계", "수집일자", today_str)
    past_s4 = _read_existing_sheet_rows(wb_old, "Raw_Data_원물_누적", "수집일자", today_str)
    past_s5 = _read_existing_sheet_rows(wb_old, "상품명_2차검증_매핑표", "수집일자", today_str)
    try:
        wb_old.close()
    except Exception:
        pass

    logger.info(
        "기존 누적 데이터 확인: Sheet1=%d건, Sheet2=%d건, Sheet3=%d건, Sheet4=%d건 -> 금일(%s) 데이터 Append 수행",
        len(past_s1), len(past_s2), len(past_s3), len(past_s4), today_str,
    )

    wb_new = openpyxl.Workbook()
    default_sheet = wb_new.active

    # -------------------------------------------------------------------------
    # Sheet 0: `전체품목_플랫폼별_전수비교` (전체 품목 요약 생략 없이 100% 전수 나열, 6대 플랫폼 비교)
    # -------------------------------------------------------------------------
    if full_item_rows:
        s0_headers = [
            "수집일자", "축종구분", "보관상태", "대분류부위(10대기준)", "품목명", "취급등급_분포",
            "전체_수집건수", "생산일자_범위",
            "근본한우_건수", "근본한우_최저가(원/kg)", "근본한우_평균가(원/kg)", "근본한우_최고가(원/kg)", "근본한우_생산일자", "근본한우_최저가_URL",
            "타사4곳통합_건수", "타사4곳통합_최저가(원/kg)", "타사4곳통합_평균가(원/kg)", "타사4곳통합_최고가(원/kg)", "근본한우대비_타사평균_편차(%)",
            "미트박스_건수", "미트박스_최저가(원/kg)", "미트박스_평균가(원/kg)", "미트박스_최고가(원/kg)", "미트박스_최저가_브랜드(생산일)", "미트박스_최고가_브랜드(생산일)", "미트박스_최저가_URL", "미트박스_최고가_URL",
            "미트프렌즈_건수", "미트프렌즈_최저가(원/kg)", "미트프렌즈_평균가(원/kg)", "미트프렌즈_최고가(원/kg)", "미트프렌즈_최저가_브랜드(생산일)", "미트프렌즈_최고가_브랜드(생산일)", "미트프렌즈_최저가_URL", "미트프렌즈_최고가_URL",
            "탑미트_건수", "탑미트_최저가(원/kg)", "탑미트_평균가(원/kg)", "탑미트_최고가(원/kg)", "탑미트_최저가_브랜드(생산일)", "탑미트_최고가_브랜드(생산일)", "탑미트_최저가_URL", "탑미트_최고가_URL",
            "금천미트_건수", "금천미트_최저가(원/kg)", "금천미트_평균가(원/kg)", "금천미트_최고가(원/kg)", "금천미트_최저가_브랜드(생산일)", "금천미트_최고가_브랜드(생산일)", "금천미트_최저가_URL", "금천미트_최고가_URL",
            "본대로_건수", "본대로_최저가(원/kg)", "본대로_평균가(원/kg)", "본대로_최고가(원/kg)", "본대로_최저가_브랜드(생산일)", "본대로_최고가_브랜드(생산일)", "본대로_최저가_URL", "본대로_최고가_URL",
            "품목내_전체최저가_상세", "품목내_전체최고가_상세", "최저가_상품URL", "최고가_상품URL",
        ]
        s0_all = full_item_rows + past_s0
        s0_curr_cols = {h for h in s0_headers if "(원/kg)" in h or "_건수" in h}
        s0_pct_cols = {h for h in s0_headers if "(%)" in h}
        _write_styled_sheet(
            wb_new,
            "전체품목_플랫폼별_전수비교",
            f"[전체 품목 전수 비교표] 축종 × 보관상태 × 전체 {len(full_item_rows)}개 표준품목 6대 플랫폼 전수 나열 (최신: {today_str})",
            s0_headers,
            s0_all,
            s0_curr_cols,
            s0_pct_cols,
        )

    # -------------------------------------------------------------------------
    # Sheet 1: `Dashboard_요약` (전체 품목 × 등급별 6대 플랫폼 전수 비교표)
    # -------------------------------------------------------------------------
    s1_headers = [
        "수집일자", "축종구분", "보관상태", "대분류부위(10대기준)", "부위명", "등급", "지육경락가(원/kg)",
        "생산일자_범위(최장기~최신)",
        "생산일자대비_최저가_상세", "생산일자대비_최고가_상세", "최저가_상품URL", "최고가_상품URL",
        "신선육(14일이내)_최저~최고(브랜드)", "숙성/장기(15일이상)_최저~최고(브랜드)",
        "근본한우_최저가(원/kg)", "근본한우_평균가(원/kg)", "근본한우_최고가(원/kg)", "근본한우_브랜드(생산일자)",
        "타사통합_최저가(원/kg)", "타사통합_평균가(원/kg)", "타사통합_최고가(원/kg)",
        "근본한우대비_타사평균_편차(%)", "근본한우대비_가격경쟁력판정",
        "미트박스_건수", "미트박스_최저가(원/kg)", "미트박스_최저가_브랜드(생산일)", "미트박스_평균가(원/kg)", "미트박스_최고가(원/kg)", "미트박스_최고가_브랜드(생산일)", "미트박스_근본대비편차(%)", "미트박스_최저가_URL", "미트박스_최고가_URL",
        "미트프렌즈_건수", "미트프렌즈_최저가(원/kg)", "미트프렌즈_최저가_브랜드(생산일)", "미트프렌즈_평균가(원/kg)", "미트프렌즈_최고가(원/kg)", "미트프렌즈_최고가_브랜드(생산일)", "미트프렌즈_근본대비편차(%)", "미트프렌즈_최저가_URL", "미트프렌즈_최고가_URL",
        "탑미트_건수", "탑미트_최저가(원/kg)", "탑미트_최저가_브랜드(생산일)", "탑미트_평균가(원/kg)", "탑미트_최고가(원/kg)", "탑미트_최고가_브랜드(생산일)", "탑미트_근본대비편차(%)", "탑미트_최저가_URL", "탑미트_최고가_URL",
        "금천미트_건수", "금천미트_최저가(원/kg)", "금천미트_최저가_브랜드(생산일)", "금천미트_평균가(원/kg)", "금천미트_최고가(원/kg)", "금천미트_최고가_브랜드(생산일)", "금천미트_근본대비편차(%)", "금천미트_최저가_URL", "금천미트_최고가_URL",
        "본대로_건수", "본대로_최저가(원/kg)", "본대로_최저가_브랜드(생산일)", "본대로_평균가(원/kg)", "본대로_최고가(원/kg)", "본대로_최고가_브랜드(생산일)", "본대로_근본대비편차(%)", "본대로_최저가_URL", "본대로_최고가_URL",
        "최저가_플랫폼(브랜드)", "최고가_플랫폼(브랜드)", "브랜드_가격라인업(최저~최고)",
    ]
    s1_all = dashboard_rows + past_s1
    s1_curr_cols = {h for h in s1_headers if "(원/kg)" in h or "_건수" in h}
    s1_pct_cols = {h for h in s1_headers if "(%)" in h}
    _write_styled_sheet(
        wb_new,
        "Dashboard_요약",
        f"[Sheet 1] 한우 원물(박스육) 전체 품목×등급별 6대 플랫폼 전수 비교 대시보드 (최신: {today_str} | 누적 {len(s1_all):,}행)",
        s1_headers,
        s1_all,
        s1_curr_cols,
        s1_pct_cols,
    )

    # -------------------------------------------------------------------------
    # Sheet 2: `지육시세_추이` (일자별 하단 누적 Append)
    # -------------------------------------------------------------------------
    s2_headers = [
        "수집일자", "기준경매일", "축종구분", "등급",
        "기준일_경락가(원/kg)", "기준일_경매두수(두)",
        "전일_경락가(원/kg)", "전일대비_증감액(원)", "전일대비_증감률(%)",
        "전주동요일_경락가(원/kg)", "전주대비_증감액(원)", "전주대비_증감률(%)",
        "전월동기_경락가(원/kg)", "전월동기_증감액(원)", "전월동기_증감률(%)",
        "전년동기_경락가(원/kg)", "전년동기_증감액(원)", "전년동기_증감률(%)",
    ]
    s2_all = past_s2 + ekape_rows
    s2_curr_cols = {h for h in s2_headers if "(원/kg)" in h or "(원)" in h or "(두)" in h}
    s2_pct_cols = {h for h in s2_headers if "(%)" in h}
    _write_styled_sheet(
        wb_new,
        "지육시세_추이",
        f"[Sheet 2] 축산물품질평가원(EKAPE) 한우거세·한우암소 등급별 지육 경락가격 시점별 비교 누적 (최신: {today_str} | 누적 {len(s2_all):,}행)",
        s2_headers,
        s2_all,
        s2_curr_cols,
        s2_pct_cols,
    )

    # -------------------------------------------------------------------------
    # Sheet 3: `생산일자_등급_보관별_통계` (일자별 하단 누적 Append)
    # -------------------------------------------------------------------------
    s3_headers = [
        "수집일자", "축종구분", "보관상태", "대분류부위(10대기준)", "부위명", "2차검증_세부표준상품명", "등급",
        "생산일자_주차구분", "상세생산일자_범위", "재고경과일수(일)", "생산경과일_구간",
        "플랫폼명", "브랜드명", "상품옵션수(재고)",
        "최저가(원/kg)", "평균가(원/kg)", "최고가(원/kg)",
        "최저가_상품URL", "최고가_상품URL",
        "동일생산일구간_전체최저가(원/kg)", "동일생산일구간_최저가_브랜드(플랫폼)", "동일생산일구간_최저가_URL",
        "동일생산일구간_전체최고가(원/kg)", "동일생산일구간_최고가_브랜드(플랫폼)", "동일생산일구간_최고가_URL",
        "생산일구간_최고최저_편차(원/kg)",
        "근본한우_동일조건_기준가(원/kg)", "근본한우대비_가격편차(%)",
        "동일조건_전체생산일_최저~최고(브랜드)",
    ]
    s3_all = past_s3 + sheet3_rows
    s3_curr_cols = {
        "최저가(원/kg)", "평균가(원/kg)", "최고가(원/kg)",
        "동일생산일구간_전체최저가(원/kg)", "동일생산일구간_전체최고가(원/kg)", "생산일구간_최고최저_편차(원/kg)",
        "근본한우_동일조건_기준가(원/kg)", "상품옵션수(재고)", "재고경과일수(일)",
    }
    s3_pct_cols = {"근본한우대비_가격편차(%)"}
    _write_styled_sheet(
        wb_new,
        "생산일자_등급_보관별_통계",
        f"[Sheet 3] 다차원 그룹핑 [축종 × 보관상태 × 부위 × 등급 × 생산일자 × 브랜드] 및 생산일자대비 최저·최고가 누적 통계 (누적 {len(s3_all):,}행)",
        s3_headers,
        s3_all,
        s3_curr_cols,
        s3_pct_cols,
        highlight_col="플랫폼명",
    )

    # -------------------------------------------------------------------------
    # Sheet 4: `Raw_Data_원물_누적` (크롤링된 원물 원본 데이터 일자별 하단 Append)
    # 필수 11개 컬럼 + 2차검증 세부표준상품명/판정결과/원본상품명
    # -------------------------------------------------------------------------
    s4_headers = [
        "수집일자",
        "플랫폼명",
        "축종구분",
        "보관상태",
        "가공상태",
        "대분류부위(10대기준)",
        "부위명",
        "2차검증_세부표준상품명",
        "등급",
        "브랜드명",
        "생산일자(가공일자/도축일자)",
        "단위단가",
        "상품URL",
        "옵션재고수량",
        "1차_플랫폼원본품목",
        "원본상품명",
        "2차검증_판정결과",
    ]
    today_s4_dicts = [r.to_excel_dict() for r in all_raw_rows]
    s4_all = past_s4 + today_s4_dicts
    s4_curr_cols = {"단위단가", "옵션재고수량"}
    _write_styled_sheet(
        wb_new,
        "Raw_Data_원물_누적",
        f"[Sheet 4] 6대 플랫폼 한우거세·한우암소 원물(박스육) 전수 크롤링 & 2차검증 Raw Data 누적 아카이브 (누적 {len(s4_all):,}행)",
        s4_headers,
        s4_all,
        s4_curr_cols,
        set(),
        highlight_col="플랫폼명",
    )

    # -------------------------------------------------------------------------
    # Sheet 5: `상품명_2차검증_매핑표` (플랫폼별 상이한 상품명 1·2차 교차검증 결과표)
    # -------------------------------------------------------------------------
    s5_headers = [
        "수집일자",
        "플랫폼명",
        "축종구분",
        "브랜드명",
        "1차_원본카테고리",
        "원본상품명",
        "1차_분류값",
        "2차검증_표준부위명",
        "2차검증_세부표준상품명",
        "2차검증_수행내역",
        "분석포함여부",
        "검증상품옵션수",
    ]
    audit_grouped: Dict[Tuple[str, str, str, str, str, str, str, str, str, str], int] = {}
    for alog in HanwooStandardizer.verification_audit_logs:
        akey = (
            str(alog.get("플랫폼명") or ""),
            str(alog.get("축종구분") or ""),
            str(alog.get("브랜드명") or ""),
            str(alog.get("1차_원본카테고리") or ""),
            str(alog.get("원본상품명") or ""),
            str(alog.get("1차_분류값") or ""),
            str(alog.get("2차검증_표준부위명") or ""),
            str(alog.get("2차검증_세부표준상품명") or ""),
            str(alog.get("2차검증_수행내역") or ""),
            str(alog.get("분석포함여부") or ""),
        )
        audit_grouped[akey] = audit_grouped.get(akey, 0) + 1

    # 2차보정 수행 건을 최상단에 배치하여 검증 내역을 한눈에 확인 가능하도록 정렬
    sorted_audit_items = sorted(
        audit_grouped.items(),
        key=lambda kv: (
            0 if "2차보정" in kv[0][8] else (1 if "제외" in kv[0][8] else 2),
            PLATFORMS_ORDER.index(kv[0][0]) if kv[0][0] in PLATFORMS_ORDER else 99,
            kv[0][6],
            -kv[1],
        ),
    )
    s5_rows = [
        {
            "수집일자": today_str,
            "플랫폼명": k[0],
            "축종구분": k[1],
            "브랜드명": k[2],
            "1차_원본카테고리": k[3],
            "원본상품명": k[4],
            "1차_분류값": k[5],
            "2차검증_표준부위명": k[6],
            "2차검증_세부표준상품명": k[7],
            "2차검증_수행내역": k[8],
            "분석포함여부": k[9],
            "검증상품옵션수": cnt,
        }
        for k, cnt in sorted_audit_items
    ]
    s5_all = s5_rows + past_s5
    _write_styled_sheet(
        wb_new,
        "상품명_2차검증_매핑표",
        f"[Sheet 5] 플랫폼별 상이한 상품명 1·2차 교차검증 및 원물 판정 매핑표 (기준일: {today_str} | 누적 {len(s5_all):,}종 상품 검증)",
        s5_headers,
        s5_all,
        {"검증상품옵션수"},
        set(),
        highlight_col="플랫폼명",
    )

    if default_sheet and default_sheet.title in wb_new.sheetnames and len(wb_new.sheetnames) > 4:
        wb_new.remove(default_sheet)

    local_master_path.parent.mkdir(parents=True, exist_ok=True)
    wb_new.save(local_master_path)
    logger.info("마스터 엑셀 파일(6개 시트) 저장 완료: %s", local_master_path)
    return local_master_path


def upload_or_update_gdrive_file(folder_id: str, local_path: Path) -> Optional[str]:
    """
    구글 드라이브 타겟 폴더에 파일을 업로드하거나, 이미 동일 파일명이 존재하면
    기존 파일 ID를 유지한 채 새 버전으로 업데이트(mutate update)합니다.
    """
    filename = local_path.name
    existing_id = find_gdrive_file_id(folder_id, filename)
    try:
        if existing_id:
            cmd = [GDRIVE_BIN, "mutate", "update", existing_id, str(local_path)]
            res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", timeout=120)
            if res.returncode == 0:
                logger.info("구글 드라이브 기존 파일 업데이트 완료: %s (ID: %s)", filename, existing_id)
                return existing_id
            logger.warning("기존 파일 업데이트 실패 -> 신규 업로드 재시도: %s", res.stderr)

        cmd_up = [
            GDRIVE_BIN, "mutate", "upload", str(local_path),
            "--parent", folder_id,
            "--json",
        ]
        res_up = subprocess.run(cmd_up, capture_output=True, text=True, encoding="utf-8", timeout=120)
        if res_up.returncode == 0:
            new_id = None
            if res_up.stdout.strip():
                try:
                    new_id = json.loads(res_up.stdout).get("id")
                except Exception:
                    pass
            if not new_id:
                new_id = find_gdrive_file_id(folder_id, filename)
            logger.info("구글 드라이브 신규 업로드 완료: %s (ID: %s)", filename, new_id)
            return new_id
        else:
            logger.error("구글 드라이브 업로드 오류 (%s): %s", filename, res_up.stderr)
    except Exception as e:
        logger.error("구글 드라이브 동기화 예외 (%s): %s", filename, e)
    return None


def sync_all_deliverables_to_gdrive(
    folder_id: str,
    master_excel_path: Path,
    visual_paths: Dict[str, Path],
) -> Dict[str, Optional[str]]:
    """마스터 엑셀, 통합 PDF 리포트, HTML 리포트 및 차트 PNG 3종을 지정된 구글 드라이브 폴더에 누적/업데이트합니다."""
    logger.info("=== [Task 5] 구글 드라이브 폴더(%s) 결과물 누적 업로드/동기화 시작 ===", folder_id)
    uploaded_ids: Dict[str, Optional[str]] = {}

    # 1. 마스터 엑셀 업로드/업데이트
    uploaded_ids[master_excel_path.name] = upload_or_update_gdrive_file(folder_id, master_excel_path)

    # 2. 일자별 시각화 PDF 리포트, HTML 리포트, 차트 PNG 및 누적 통합 대시보드 업로드
    for key in ["pdf", "html", "chart1", "chart2", "chart3", "cumulative_dashboard", "cumulative_archive"]:
        p = visual_paths.get(key)
        if p and p.exists():
            uploaded_ids[p.name] = upload_or_update_gdrive_file(folder_id, p)

    return uploaded_ids


# ---------------------------------------------------------------------------
# 메인 파이프라인 실행 엔트리포인트
# ---------------------------------------------------------------------------
def main() -> None:
    from cumulative_dashboard_builder import (
        CUMULATIVE_ARCHIVE_FILENAME,
        CUMULATIVE_DASHBOARD_FILENAME,
        update_and_build_cumulative_dashboard,
    )

    parser = argparse.ArgumentParser(description="근본한우 온라인팀 MD - 5대 B2B 플랫폼 한우원물 크롤링 및 시세 분석 자동화 파이프라인")
    parser.add_argument("--date", type=str, default=datetime.now().strftime("%Y-%m-%d"), help="기준 수집일자 (YYYY-MM-DD)")
    parser.add_argument("--folder-id", type=str, default=GDRIVE_FOLDER_ID, help="타겟 구글 드라이브 폴더 ID")
    parser.add_argument("--output-dir", type=str, default=str(Path(__file__).resolve().parent / "output"), help="로컬 산출물 저장 폴더")
    parser.add_argument("--skip-gdrive", action="store_true", help="구글 드라이브 업로드 생략 여부")
    parser.add_argument("--sync-only", action="store_true", help="이미 생성된 로컬 산출물을 구글 드라이브로만 동기화")
    parser.add_argument("--recompute-from-excel", action="store_true", help="기존 마스터 엑셀의 Raw_Data를 기반으로 정제/집계/시각화 재산출")
    args = parser.parse_args()

    today_str = args.date
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    local_master_path = output_dir / MASTER_EXCEL_FILENAME

    if args.sync_only:
        visual_paths = {
            "chart1": output_dir / f"{today_str}_차트1_등급별_보관상태별_6사_가격분포.png",
            "chart2": output_dir / f"{today_str}_차트2_생산일자별_단가스캐터_및_히트맵.png",
            "chart3": output_dir / f"{today_str}_차트3_지육시세_연동성_및_플랫폼비교.png",
            "pdf": output_dir / f"{today_str}_한우원물_6사비교_시각화리포트.pdf",
            "html": output_dir / f"{today_str}_한우원물_6사비교_시각화리포트.html",
            "cumulative_dashboard": output_dir / CUMULATIVE_DASHBOARD_FILENAME,
            "cumulative_archive": output_dir / CUMULATIVE_ARCHIVE_FILENAME,
        }
        uploaded_ids = sync_all_deliverables_to_gdrive(args.folder_id, local_master_path, visual_paths)
        summary_json_path = output_dir / f"{today_str}_pipeline_run_summary.json"
        if summary_json_path.exists():
            data = json.loads(summary_json_path.read_text(encoding="utf-8"))
            data["gdrive_uploaded_ids"] = uploaded_ids
            summary_json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info("구글 드라이브 단독 동기화 완료: %s", uploaded_ids)
        return

    logger.info("================================================================================")
    logger.info(" [근본한우 온라인팀 MD] 한우 원물(박스육) 6대 플랫폼 비교 & 지육시세 자동화 파이프라인 가동")
    logger.info(" - 기준일자: %s | 로컬 저장소: %s | 구글 드라이브 폴더: %s", today_str, output_dir, args.folder_id)
    logger.info("================================================================================")

    if args.recompute_from_excel and local_master_path.exists():
        logger.info("기존 마스터 엑셀(%s) 기반 정제·집계·시각화 재산출 모드 실행", local_master_path)
        wb_ex = openpyxl.load_workbook(local_master_path, data_only=True)
        # 1. 지육시세 복원
        ws_ek = wb_ex["지육시세_추이"]
        ek_headers = [ws_ek.cell(2, c).value for c in range(1, ws_ek.max_column + 1)]
        ekape_rows = []
        for r_idx in range(3, ws_ek.max_row + 1):
            d_row = {ek_headers[c - 1]: ws_ek.cell(r_idx, c).value for c in range(1, ws_ek.max_column + 1)}
            if str(d_row.get("수집일자")) == today_str:
                ekape_rows.append(d_row)
        geose_map = {str(r["등급"]): int(r["기준일_경락가(원/kg)"] or 0) for r in ekape_rows if r.get("축종구분") == "한우거세"}
        amso_map = {str(r["등급"]): int(r["기준일_경락가(원/kg)"] or 0) for r in ekape_rows if r.get("축종구분") == "한우암소"}
        ekape_summary = {
            "base_date": str(ekape_rows[0]["기준경매일"]) if ekape_rows else today_str,
            "geose_avg": geose_map.get("전국평균", 24450),
            "amso_avg": amso_map.get("전국평균", 21497),
            "geose_by_grade": geose_map,
            "amso_by_grade": amso_map,
            "geose_prev_day_pct": next((float(r["전일대비_증감률(%)"]) for r in ekape_rows if r.get("축종구분") == "한우거세" and r.get("등급") == "전국평균"), 0.0),
            "amso_prev_day_pct": next((float(r["전일대비_증감률(%)"]) for r in ekape_rows if r.get("축종구분") == "한우암소" and r.get("등급") == "전국평균"), 0.0),
            "geose_yoy_pct": next((float(r["전년동기_증감률(%)"]) for r in ekape_rows if r.get("축종구분") == "한우거세" and r.get("등급") == "전국평균"), 0.0),
            "amso_yoy_pct": next((float(r["전년동기_증감률(%)"]) for r in ekape_rows if r.get("축종구분") == "한우암소" and r.get("등급") == "전국평균"), 0.0),
        }
        # 2. Raw Data 복원 및 정밀 재분류 (미트박스 제외 & 할랄한우 제외)
        ws_raw = wb_ex["Raw_Data_원물_누적"]
        raw_headers = [ws_raw.cell(2, c).value for c in range(1, ws_raw.max_column + 1)]
        all_raw_rows = []
        for r_idx in range(3, ws_raw.max_row + 1):
            d_r = {raw_headers[c - 1]: ws_raw.cell(r_idx, c).value for c in range(1, ws_raw.max_column + 1)}
            if str(d_r.get("수집일자")) != today_str:
                continue
            plat = str(d_r.get("플랫폼명") or "").strip()
            if plat == "미트박스" or plat not in PLATFORMS_ORDER:
                continue
            cattle_t = str(d_r.get("축종구분") or "").strip()
            keep_t = str(d_r.get("보관상태") or "").strip()
            brand_t = str(d_r.get("브랜드명") or "").strip()
            raw_nm = str(d_r.get("원본상품명") or "").strip()
            prim_raw = str(d_r.get("1차_플랫폼원본품목") or d_r.get("부위명") or "").strip()
            u_prc = int(d_r.get("단위단가") or 0)
            m_part, p_std, sub_spec, v_stat = HanwooStandardizer.verify_and_normalize_part(prim_raw, raw_nm, u_prc)
            if "할랄" in brand_t or "할랄" in raw_nm:
                HanwooStandardizer.record_audit(
                    plat, cattle_t, prim_raw, raw_nm, brand_t,
                    prim_raw, p_std, sub_spec, "2차검증 제외: 할랄인증 특수선주문(실생산일 미존재·고정더미단가 시세왜곡 방지)", False,
                )
                continue
            HanwooStandardizer.record_audit(
                plat, cattle_t, prim_raw, raw_nm, brand_t,
                prim_raw, p_std, sub_spec, v_stat, True,
            )
            all_raw_rows.append(
                HanwooProductRow(
                    collected_date=today_str,
                    platform=plat,
                    cattle_type=cattle_t,
                    keeping=keep_t,
                    processing="원물",
                    part_name=p_std,
                    grade=str(d_r.get("등급") or "").strip(),
                    brand_name=brand_t,
                    production_date=str(d_r.get("생산일자(가공일자/도축일자)") or "-").strip(),
                    unit_price_kg=u_prc,
                    product_url=str(d_r.get("상품URL") or "").strip(),
                    raw_product_name=raw_nm,
                    stock_qty=int(d_r.get("옵션재고수량") or 1),
                    major_part=m_part,
                    primary_part_raw=prim_raw,
                    sub_spec_name=sub_spec,
                    verify_status=v_stat,
                )
            )
    else:
        session = build_http_session()

        # 0. 구글 드라이브에 기존 마스터 파일이 있으면 먼저 다운로드하여 누적 아카이브 기반 확보
        if not args.skip_gdrive:
            download_existing_master_from_gdrive(args.folder_id, local_master_path)

        # 1. [Task 3] 축산물품질평가원(EKAPE) 한우거세·한우암소 등급별 지육 경락가격 수집
        ekape_rows, ekape_summary = fetch_ekape_hanwoo_auction(session, today_str)

        # 2. [Task 1] 5개 타겟 B2B 플랫폼 + 근본한우(자사) 원물(박스육) 전수 크롤링
        mb_rows = fetch_meatbox_and_geunbon(session, today_str)
        mf_rows = fetch_meatfriends_hanwoo(session, today_str)
        tm_rows = fetch_topmeat_hanwoo(session, today_str)
        ek_rows = fetch_ekcm_hanwoo(session, today_str)
        bd_rows = fetch_bondaero_hanwoo(session, today_str)

        all_raw_rows = []
        for r_list in (mb_rows, mf_rows, tm_rows, ek_rows, bd_rows):
            all_raw_rows.extend(r_list)

    # 정렬: 플랫폼 순 -> 축종 -> 보관상태 -> 부위 -> 등급 -> 단가
    all_raw_rows.sort(
        key=lambda r: (
            PLATFORMS_ORDER.index(r.platform) if r.platform in PLATFORMS_ORDER else 99,
            0 if r.cattle_type == "한우거세" else 1,
            0 if r.keeping == "냉장" else 1,
            CORE_PARTS_ORDER.index(r.part_name) if r.part_name in CORE_PARTS_ORDER else 99,
            GRADES_ORDER.index(r.grade) if r.grade in GRADES_ORDER else 99,
            r.unit_price_kg,
        )
    )

    # 3. [Task 2] 다차원 그룹핑 및 통계 산출
    dashboard_rows, sheet3_rows, analytics_meta = build_multidimensional_analytics(
        all_raw_rows, ekape_summary, today_str
    )

    # 4. [Task 5] 마스터 엑셀 누적 저장 (6개 시트: 전체품목 전수비교표 포함) - 데이터 유실 방지를 위해 시각화보다 먼저 저장
    save_and_append_master_excel(
        local_master_path,
        dashboard_rows,
        ekape_rows,
        sheet3_rows,
        all_raw_rows,
        today_str,
        full_item_rows=analytics_meta.get("full_item_rows", []),
    )

    # 5. [누적 통합 대시보드] 크롤링 회차별 아카이브 JSON 업데이트 및 인터랙티브 대시보드 생성
    cum_paths = update_and_build_cumulative_dashboard(local_master_path, output_dir)

    # 6. [Task 4] 6대 브랜드/플랫폼 비교 시각화 차트(PNG 3종) + 통합 PDF + HTML 생성
    visual_paths = generate_visual_reports(
        all_raw_rows,
        dashboard_rows,
        sheet3_rows,
        ekape_rows,
        ekape_summary,
        analytics_meta,
        output_dir,
        today_str,
    )
    visual_paths["cumulative_dashboard"] = cum_paths["cumulative_dashboard_html"]
    visual_paths["cumulative_archive"] = cum_paths["archive_json"]

    uploaded_ids: Dict[str, Optional[str]] = {}
    if not args.skip_gdrive:
        uploaded_ids = sync_all_deliverables_to_gdrive(args.folder_id, local_master_path, visual_paths)

    # 실행 요약 JSON 저장 (검증 및 자동 리포팅용)
    summary_json_path = output_dir / f"{today_str}_pipeline_run_summary.json"
    summary_payload = {
        "collected_date": today_str,
        "ekape_base_date": ekape_summary.get("base_date"),
        "ekape_geose_avg": ekape_summary.get("geose_avg"),
        "ekape_amso_avg": ekape_summary.get("amso_avg"),
        "analytics_meta": analytics_meta,
        "master_excel": str(local_master_path),
        "visual_files": {k: str(v) for k, v in visual_paths.items()},
        "gdrive_uploaded_ids": uploaded_ids,
    }
    summary_json_path.write_text(json.dumps(summary_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    logger.info("================================================================================")
    logger.info(" 파이프라인 전 과정 완료! 총 수집 원물 옵션/박스: %d건", len(all_raw_rows))
    logger.info(" - 마스터 엑셀: %s", local_master_path)
    logger.info(" - 누적 통합 대시보드: %s", visual_paths["cumulative_dashboard"])
    logger.info(" - 시각화 PDF : %s", visual_paths["pdf"])
    logger.info("================================================================================")


if __name__ == "__main__":
    main()
