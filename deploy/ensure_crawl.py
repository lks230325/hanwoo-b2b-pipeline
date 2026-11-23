#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[근본한우 B2B 인텔리전스] 윈도우 작업 스케줄러 연동 평일(09:30 / 16:30) 안전 크롤링 & 구글드라이브 동기화 보장기
- 슈퍼바이저(server_supervisor.py) 미실행 시 백그라운드 자동 기동
- 이미 당일 해당 슬롯(09시30분 / 16시30분) 크롤링이 완료되었거나 진행 중이면 중복 실행 없이 즉시 종료
- 미실행 상태일 경우에만 안전하게 수집 -> 대시보드 빌드 -> 구글 드라이브(1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe) 및 바탕화면 동기화 실행
"""

import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
from datetime import datetime
from pathlib import Path

DEPLOY_DIR = Path(__file__).resolve().parent
WORK_DIR = DEPLOY_DIR.parent
OUTPUT_DIR = WORK_DIR / "output"
ARCHIVE_JSON_PATH = OUTPUT_DIR / "[누적_히스토리]_한우원물_크롤링_아카이브.json"
DASHBOARD_HTML_PATH = OUTPUT_DIR / "[누적_통합대시보드]_한우원물_5대플랫폼_시세비교_대시보드.html"
MASTER_EXCEL_PATH = OUTPUT_DIR / "[누적_기초데이터]_한우원물_시세_및_플랫폼비교_마스터.xlsx"
ARTIFACT_DASHBOARD_PATH = Path(
    r"C:\Users\limkyungsoon\.gemini\antigravity\brain\da933e6f-57ef-4289-864b-121c7d9cbc4f\hanwoo_cumulative_dashboard.html"
)
DESKTOP_DIR = Path(r"C:\Users\limkyungsoon\Desktop\온라인팀\근본한우")
VENV_PYTHON = Path(r"C:\Users\limkyungsoon\.gemini\antigravity\scratch\meatbox-scraper\.venv\Scripts\python.exe")

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def ensure_background_supervisor() -> None:
    sup_script = DEPLOY_DIR / "server_supervisor.py"
    pyw = Path(r"C:\Users\limkyungsoon\AppData\Local\Programs\Python\Python312\pythonw.exe")
    exe = str(pyw if pyw.exists() else (VENV_PYTHON if VENV_PYTHON.exists() else Path(sys.executable)))
    subprocess.Popen(
        [exe, str(sup_script)],
        cwd=str(WORK_DIR),
        creationflags=CREATE_NO_WINDOW,
    )


def has_slot_run(date_str: str, slot_kr: str) -> bool:
    if not ARCHIVE_JSON_PATH.exists():
        return False
    try:
        arc = json.loads(ARCHIVE_JSON_PATH.read_text(encoding="utf-8"))
        for snap in arc.get("snapshots", []):
            cdate = str(snap.get("collected_date", ""))
            if cdate.startswith(date_str) and slot_kr in cdate:
                return True
    except Exception:
        pass
    return False


def main() -> int:
    ensure_background_supervisor()
    now = datetime.now()
    if now.weekday() >= 5:  # 토·일 제외
        return 0

    slot_kr = "09시30분" if now.hour < 14 else "16시30분"
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if "9" in arg:
            slot_kr = "09시30분"
        elif "16" in arg:
            slot_kr = "16시30분"

    date_str = now.strftime("%Y-%m-%d")
    label = f"{date_str} {slot_kr}"

    # web_dashboard_server.py 내장 스케줄러가 먼저 시작할 수 있도록 15초 대기 후 확인
    time.sleep(15)
    if has_slot_run(date_str, slot_kr):
        return 0

    try:
        with urllib.request.urlopen("http://127.0.0.1:8080/api/status", timeout=5) as resp:
            st = json.loads(resp.read().decode("utf-8"))
            if st.get("is_crawling"):
                return 0
            # 웹 서버가 살아있으면 웹 서버 단일 락(STATE_LOCK)을 통해 실행하여 중복 프로세스 충돌 방지
            req = urllib.request.Request("http://127.0.0.1:8080/api/run-now", data=b"{}", method="POST")
            with urllib.request.urlopen(req, timeout=10) as _:
                return 0
    except Exception:
        pass

    if not has_slot_run(date_str, slot_kr):
        py_bin = str(VENV_PYTHON if VENV_PYTHON.exists() else sys.executable)
        pipeline_script = str(WORK_DIR / "hanwoo_pipeline.py")
        builder_script = str(WORK_DIR / "cumulative_dashboard_builder.py")

        subprocess.run(
            [py_bin, "-X", "utf8", pipeline_script, "--date", label, "--skip-gdrive"],
            cwd=str(WORK_DIR),
            creationflags=CREATE_NO_WINDOW,
        )
        subprocess.run(
            [py_bin, "-X", "utf8", builder_script, "--fast"],
            cwd=str(WORK_DIR),
            creationflags=CREATE_NO_WINDOW,
        )
        subprocess.run(
            [py_bin, "-X", "utf8", pipeline_script, "--date", label, "--sync-only"],
            cwd=str(WORK_DIR),
            creationflags=CREATE_NO_WINDOW,
        )
        gh_builder = WORK_DIR / "build_gh_pages.py"
        if gh_builder.exists():
            subprocess.run(
                [py_bin, "-X", "utf8", str(gh_builder), "--push"],
                cwd=str(WORK_DIR),
                creationflags=CREATE_NO_WINDOW,
            )
        try:
            if DASHBOARD_HTML_PATH.exists() and ARTIFACT_DASHBOARD_PATH.parent.exists():
                shutil.copyfile(str(DASHBOARD_HTML_PATH), str(ARTIFACT_DASHBOARD_PATH))
            if DESKTOP_DIR.exists():
                if DASHBOARD_HTML_PATH.exists():
                    shutil.copyfile(str(DASHBOARD_HTML_PATH), str(DESKTOP_DIR / DASHBOARD_HTML_PATH.name))
                if MASTER_EXCEL_PATH.exists():
                    shutil.copyfile(str(MASTER_EXCEL_PATH), str(DESKTOP_DIR / MASTER_EXCEL_PATH.name))
        except Exception:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
