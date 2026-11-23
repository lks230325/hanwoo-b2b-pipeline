#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[근본한우 B2B 인텔리전스] GitHub Pages 정적 배포 빌드 및 자동 동기화 스크립트 (build_gh_pages.py)
- 로컬 실행 시: output/[누적_히스토리]_한우원물_크롤링_아카이브.json 및 누적 마스터 엑셀에서
  GitHub Pages용 경량화 데이터(static_data/slim_archive.json, status.json, hanwoo_master_cumulative.xlsx)를 추출하고 dist/ 생성
- GitHub Actions 실행 시 (--from-static-only): 저장소에 커밋된 static_data/ 와 hanwoo_ui_template.html을 결합하여
  https://lks230325.github.io/hanwoo-b2b-pipeline/ 하위 경로 완벽 호환 dist/ 정적 사이트 빌드
"""

import json
import os
import shutil
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Optional

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

WORK_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = WORK_DIR / "output"
STATIC_DATA_DIR = WORK_DIR / "static_data"
DIST_DIR = WORK_DIR / "dist"
ARCHIVE_JSON_PATH = OUTPUT_DIR / "[누적_히스토리]_한우원물_크롤링_아카이브.json"
MASTER_EXCEL_PATH = OUTPUT_DIR / "[누적_기초데이터]_한우원물_시세_및_플랫폼비교_마스터.xlsx"
UI_TEMPLATE_PATH = WORK_DIR / "hanwoo_ui_template.html"

import cumulative_dashboard_builder as cdb


def compute_next_weekday_run(now: Optional[datetime] = None) -> str:
    if now is None:
        now = datetime.now()
    for day_offset in range(0, 8):
        dt = now + timedelta(days=day_offset)
        if dt.weekday() >= 5:
            continue
        for hh, mm in ((9, 30), (16, 30)):
            cand = dt.replace(hour=hh, minute=mm, second=0, microsecond=0)
            if cand > now:
                weekday_kr = ["월", "화", "수", "목", "금", "토", "일"][cand.weekday()]
                return f"{cand.strftime('%Y-%m-%d')} ({weekday_kr}) {cand.strftime('%H:%M')}"
    return "평일 09:30 · 16:30 자동동기화"


def export_to_static_data(public_pages_url: str = "") -> Dict[str, Any]:
    """로컬 아카이브 JSON 및 마스터 엑셀에서 static_data/ 폴더로 정적 배포 데이터 추출"""
    STATIC_DATA_DIR.mkdir(parents=True, exist_ok=True)
    slim_path = STATIC_DATA_DIR / "slim_archive.json"

    if ARCHIVE_JSON_PATH.exists():
        raw_archive = json.loads(ARCHIVE_JSON_PATH.read_text(encoding="utf-8"))
        # _slim_archive_for_html 내부에서 URL 압축을 수행하므로 원본 복사 후 슬림화
        slim_payload = cdb._slim_archive_for_html(json.loads(json.dumps(raw_archive)))
        slim_path.write_text(
            json.dumps(slim_payload, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )
    elif slim_path.exists():
        slim_payload = json.loads(slim_path.read_text(encoding="utf-8"))
    else:
        raise FileNotFoundError(f"아카이브 JSON 및 static_data/slim_archive.json 모두 존재하지 않습니다: {ARCHIVE_JSON_PATH}")

    snaps = slim_payload.get("snapshots", [])
    latest_snap = snaps[-1] if snaps else {}
    latest_label = latest_snap.get("collected_date", "-")

    status_payload = {
        "is_crawling": False,
        "current_label": None,
        "last_completed_label": latest_label,
        "last_completed_at": slim_payload.get("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        "last_error": None,
        "schedule_rule": "평일(월~금) 오전 09:30 / 오후 16:30 자동 크롤링 (토·일 제외)",
        "next_scheduled_run": compute_next_weekday_run(),
        "total_snapshots": len(snaps),
        "github_pages_url": public_pages_url or "https://lks230325.github.io/hanwoo-b2b-pipeline/",
        "fixed_https_url": "https://cqwq4y4auzplvuyk.l.tunwg.com",
        "gdrive_folder_url": "https://drive.google.com/drive/folders/1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe",
    }
    (STATIC_DATA_DIR / "status.json").write_text(
        json.dumps(status_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    if MASTER_EXCEL_PATH.exists():
        shutil.copy2(MASTER_EXCEL_PATH, STATIC_DATA_DIR / "hanwoo_master_cumulative.xlsx")

    return slim_payload


def build_dist(public_pages_url: str = "") -> Path:
    """hanwoo_ui_template.html + static_data/ -> dist/ 정적 사이트 생성"""
    env_pages_url = os.environ.get("GH_PAGES_URL", "").strip()
    effective_url = public_pages_url or env_pages_url or "https://lks230325.github.io/hanwoo-b2b-pipeline/"

    if ARCHIVE_JSON_PATH.exists() and "--from-static-only" not in sys.argv:
        slim_payload = export_to_static_data(public_pages_url=effective_url)
    else:
        slim_path = STATIC_DATA_DIR / "slim_archive.json"
        if not slim_path.exists():
            raise FileNotFoundError(f"정적 데이터 파일이 없습니다: {slim_path}")
        slim_payload = json.loads(slim_path.read_text(encoding="utf-8"))

    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)
    DIST_DIR.mkdir(parents=True, exist_ok=True)

    dist_data_dir = DIST_DIR / "data"
    shutil.copytree(STATIC_DATA_DIR, dist_data_dir)

    pre_map = cdb._pre_render_initial_sections(slim_payload)
    embedded_json = json.dumps(slim_payload, ensure_ascii=False, separators=(",", ":")).replace("</script>", "<\\/script>")

    tpl_str = UI_TEMPLATE_PATH.read_text(encoding="utf-8")
    tpl_str = tpl_str.replace("{{EMBEDDED_JSON}}", embedded_json)
    for k, v in pre_map.items():
        tpl_str = tpl_str.replace(k, v)

    # GitHub Pages 정적 모드 플래그 주입 (하위 경로 https://lks230325.github.io/hanwoo-b2b-pipeline/ 100% 호환)
    static_flag_script = "  <script>window.__GH_PAGES_STATIC__ = true;</script>\n</head>"
    tpl_str = tpl_str.replace("</head>", static_flag_script, 1)

    (DIST_DIR / "index.html").write_text(tpl_str, encoding="utf-8")
    (DIST_DIR / ".nojekyll").write_text("", encoding="utf-8")

    snaps = slim_payload.get("snapshots", [])
    latest = snaps[-1] if snaps else {}
    print(
        f"[BUILD_OK] dist 생성 완료: index.html ({len(tpl_str):,} bytes), "
        f"총 {len(snaps)}개 회차 누적 (최신: {latest.get('collected_date', '-')}, "
        f"원물 {latest.get('total_raw_rows', 0):,}건)"
    )
    return DIST_DIR


def sync_static_data_to_github() -> bool:
    """정기 크롤링 완료 후 static_data/ 갱신 및 GitHub 원격 저장소(origin main) 자동 커밋·푸시"""
    try:
        export_to_static_data()
        git_dir = WORK_DIR / ".git"
        if not git_dir.exists():
            return False
        rem = subprocess.run(
            ["git", "remote", "get-url", "origin"],
            cwd=str(WORK_DIR),
            capture_output=True,
            text=True,
            timeout=10,
        )
        if rem.returncode != 0 or not rem.stdout.strip():
            return False

        subprocess.run(["git", "add", "static_data"], cwd=str(WORK_DIR), check=False, timeout=30)
        diff = subprocess.run(
            ["git", "diff", "--cached", "--quiet"],
            cwd=str(WORK_DIR),
            capture_output=True,
            timeout=10,
        )
        if diff.returncode != 0:
            st_path = STATIC_DATA_DIR / "status.json"
            label = "-"
            snap_cnt = 0
            if st_path.exists():
                st = json.loads(st_path.read_text(encoding="utf-8"))
                label = st.get("last_completed_label", "-")
                snap_cnt = st.get("total_snapshots", 0)
            msg = f"chore(data): sync Hanwoo crawl snapshot #{snap_cnt} ({label})"
            subprocess.run(["git", "commit", "-m", msg], cwd=str(WORK_DIR), check=False, timeout=30)
            subprocess.run(["git", "push", "origin", "main"], cwd=str(WORK_DIR), check=False, timeout=90)
            print(f"[GIT_PUSH_OK] GitHub Pages 원격 저장소 자동 푸시 완료: {msg}")
        return True
    except Exception as exc:
        print(f"[WARN] GitHub Pages 데이터 자동 푸시 건너뜀: {exc}")
        return False


if __name__ == "__main__":
    build_dist()
    if "--push" in sys.argv:
        sync_static_data_to_github()
