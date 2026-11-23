#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[근본한우 온라인팀 MD] 6대 플랫폼 원물 비교 실시간 웹 대시보드 서버 & 평일 스케줄러 데몬
- 접속 URL: http://localhost:8080 (또는 http://127.0.0.1:8080)
- 정기 자동 크롤링 스케줄: 매주 월~금(토·일요일 제외) 오전 09:30, 오후 16:30 자동 실행
- 웹 UI에서 [지금 즉시 크롤링 실행], [누적 마스터 엑셀 다운로드], [실시간 상태 조회] 지원
"""

import json
import logging
import subprocess
import sys
import threading
import time
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import urlparse

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
DASHBOARD_HTML_PATH = OUTPUT_DIR / "[누적_통합대시보드]_한우원물_5대플랫폼_시세비교_대시보드.html"
ARCHIVE_JSON_PATH = OUTPUT_DIR / "[누적_히스토리]_한우원물_크롤링_아카이브.json"
MASTER_EXCEL_PATH = OUTPUT_DIR / "[누적_기초데이터]_한우원물_시세_및_플랫폼비교_마스터.xlsx"
ARTIFACT_DASHBOARD_PATH = Path(
    r"C:\Users\limkyungsoon\.gemini\antigravity\brain\da933e6f-57ef-4289-864b-121c7d9cbc4f\hanwoo_cumulative_dashboard.html"
)
PYTHON_EXE = Path(r"C:\Users\limkyungsoon\.gemini\antigravity\scratch\meatbox-scraper\.venv\Scripts\python.exe")
PIPELINE_SCRIPT = BASE_DIR / "hanwoo_pipeline.py"
BUILDER_SCRIPT = BASE_DIR / "cumulative_dashboard_builder.py"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] WebDashboard - %(message)s",
)
logger = logging.getLogger("HanwooWebServer")

STATE: Dict[str, Any] = {
    "is_crawling": False,
    "current_label": None,
    "last_completed_label": None,
    "last_completed_at": None,
    "last_error": None,
    "triggered_slots": set(),  # e.g. "2026-11-18_09:30"
    "arc_mtime": 0.0,
    "cached_snap_count": 0,
    "cached_latest_snap_date": "-",
}
STATE_LOCK = threading.Lock()


def get_cached_archive_meta() -> tuple[int, str]:
    if not ARCHIVE_JSON_PATH.exists():
        return 0, "-"
    try:
        mt = ARCHIVE_JSON_PATH.stat().st_mtime
        if mt == STATE.get("arc_mtime") and STATE.get("cached_snap_count", 0) > 0:
            return STATE["cached_snap_count"], STATE["cached_latest_snap_date"]
        arc = json.loads(ARCHIVE_JSON_PATH.read_text(encoding="utf-8"))
        snaps = arc.get("snapshots", [])
        cnt = len(snaps)
        latest = snaps[-1].get("collected_date", "-") if snaps else "-"
        STATE["arc_mtime"] = mt
        STATE["cached_snap_count"] = cnt
        STATE["cached_latest_snap_date"] = latest
        return cnt, latest
    except Exception:
        return STATE.get("cached_snap_count", 0), STATE.get("cached_latest_snap_date", "-")


def compute_next_weekday_run(now: Optional[datetime] = None) -> str:
    """월~금(토·일 제외) 09:30 / 16:30 기준 다음 예정 시각을 계산합니다."""
    if now is None:
        now = datetime.now()
    for day_offset in range(0, 8):
        dt = now + timedelta(days=day_offset)
        if dt.weekday() >= 5:  # 5=토요일, 6=일요일 제외
            continue
        for hh, mm in ((9, 30), (16, 30)):
            cand = dt.replace(hour=hh, minute=mm, second=0, microsecond=0)
            if cand > now:
                weekday_kr = ["월", "화", "수", "목", "금", "토", "일"][cand.weekday()]
                return f"{cand.strftime('%Y-%m-%d')} ({weekday_kr}) {cand.strftime('%H:%M')}"
    return "-"


def run_crawl_job(date_label: str) -> None:
    """백그라운드에서 6대 플랫폼 파이프라인 및 누적 대시보드 빌더를 실행합니다."""
    with STATE_LOCK:
        if STATE["is_crawling"]:
            logger.info("이미 크롤링이 진행 중이므로 중복 실행을 건너뜁니다: %s", date_label)
            return
        STATE["is_crawling"] = True
        STATE["current_label"] = date_label
        STATE["last_error"] = None

    logger.info("=== [웹 대시보드 자동/수동 크롤링 시작] 회차 라벨: %s ===", date_label)
    try:
        py_bin = str(PYTHON_EXE if PYTHON_EXE.exists() else sys.executable)
        cmd1 = [py_bin, "-X", "utf8", str(PIPELINE_SCRIPT), "--date", date_label, "--skip-gdrive"]
        res1 = subprocess.run(cmd1, cwd=str(BASE_DIR), capture_output=True, text=True, encoding="utf-8", errors="replace")
        if res1.returncode != 0:
            raise RuntimeError(f"hanwoo_pipeline.py 오류 (code {res1.returncode}): {res1.stderr[-500:]}")

        cmd2 = [py_bin, "-X", "utf8", str(BUILDER_SCRIPT), "--fast"]
        subprocess.run(cmd2, cwd=str(BASE_DIR), capture_output=True, text=True, encoding="utf-8", errors="replace")

        # 구글 드라이브 폴더(1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe) 자동 배포/동기화
        cmd3 = [py_bin, "-X", "utf8", str(PIPELINE_SCRIPT), "--date", date_label, "--sync-only"]
        subprocess.run(cmd3, cwd=str(BASE_DIR), capture_output=True, text=True, encoding="utf-8", errors="replace")

        # GitHub Pages 정적 사이트 빌드 및 원격 저장소(origin main) 자동 푸시
        gh_builder = BASE_DIR / "build_gh_pages.py"
        if gh_builder.exists():
            cmd4 = [py_bin, "-X", "utf8", str(gh_builder), "--push"]
            subprocess.run(cmd4, cwd=str(BASE_DIR), capture_output=True, text=True, encoding="utf-8", errors="replace")

        try:
            import shutil
            if DASHBOARD_HTML_PATH.exists() and ARTIFACT_DASHBOARD_PATH.parent.exists():
                shutil.copyfile(str(DASHBOARD_HTML_PATH), str(ARTIFACT_DASHBOARD_PATH))
            desktop_dir = Path(r"C:\Users\limkyungsoon\Desktop\온라인팀\근본한우")
            if desktop_dir.exists():
                if DASHBOARD_HTML_PATH.exists():
                    shutil.copyfile(str(DASHBOARD_HTML_PATH), str(desktop_dir / "[누적_통합대시보드]_한우원물_5대플랫폼_시세비교_대시보드.html"))
                if MASTER_EXCEL_PATH.exists():
                    shutil.copyfile(str(MASTER_EXCEL_PATH), str(desktop_dir / "[누적_기초데이터]_한우원물_시세_및_플랫폼비교_마스터.xlsx"))
        except Exception:
            pass

        with STATE_LOCK:
            STATE["last_completed_label"] = date_label
            STATE["last_completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logger.info("=== [웹 대시보드 크롤링 완료] 회차 라벨: %s ===", date_label)
    except Exception as e:
        logger.error("크롤링 실행 중 예외 발생: %s", e)
        with STATE_LOCK:
            STATE["last_error"] = str(e)
    finally:
        with STATE_LOCK:
            STATE["is_crawling"] = False
            STATE["current_label"] = None


def get_local_ip() -> str:
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def get_public_tunnel_url() -> str:
    fixed_file = OUTPUT_DIR / "fixed_tunnel_url.txt"
    if fixed_file.exists():
        try:
            u = fixed_file.read_text(encoding="utf-8").strip()
            if u:
                return u
        except Exception:
            pass
    tunnel_file = OUTPUT_DIR / "public_tunnel_url.txt"
    if tunnel_file.exists():
        try:
            return tunnel_file.read_text(encoding="utf-8").strip()
        except Exception:
            pass
    return ""


def get_fixed_tunnel_url() -> str:
    fixed_file = OUTPUT_DIR / "fixed_tunnel_url.txt"
    if fixed_file.exists():
        try:
            return fixed_file.read_text(encoding="utf-8").strip()
        except Exception:
            pass
    return ""


def weekday_scheduler_loop() -> None:
    """매주 월~금(토·일 제외) 09:30 및 16:30 정각에 자동으로 크롤링을 실행합니다."""
    logger.info("평일 스케줄러 가동 완료: 매주 월~금 09:30 / 16:30 자동 실행 (토·일 제외)")
    # 서버 재기동 시 금일 이미 수집된 슬롯을 STATE['triggered_slots']에 등록
    if ARCHIVE_JSON_PATH.exists():
        try:
            arc = json.loads(ARCHIVE_JSON_PATH.read_text(encoding="utf-8"))
            for s in arc.get("snapshots", []):
                cdate = str(s.get("collected_date", ""))
                if "09시30분" in cdate:
                    STATE["triggered_slots"].add(f"{cdate[:10]}_09:30")
                elif "16시30분" in cdate:
                    STATE["triggered_slots"].add(f"{cdate[:10]}_16:30")
        except Exception:
            pass
    while True:
        try:
            now = datetime.now()
            # 월(0)~금(4)만 실행, 토(5)·일(6)은 제외
            if now.weekday() < 5:
                hm = now.strftime("%H:%M")
                if hm in ("09:30", "16:30"):
                    slot_key = f"{now.strftime('%Y-%m-%d')}_{hm}"
                    if slot_key not in STATE["triggered_slots"]:
                        STATE["triggered_slots"].add(slot_key)
                        label = f"{now.strftime('%Y-%m-%d')} {'09시30분' if hm == '09:30' else '16시30분'}"
                        t = threading.Thread(target=run_crawl_job, args=(label,), daemon=True)
                        t.start()
        except Exception as e:
            logger.error("스케줄러 루프 오류: %s", e)
        time.sleep(15)


class DashboardRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        return

    def _send_json(self, payload: Dict[str, Any], status: int = 200) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html", "/dashboard"):
            if not DASHBOARD_HTML_PATH.exists():
                msg = "<h1>대시보드 파일을 생성 중입니다. 잠시 후 새로고침 해주세요.</h1>".encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(msg)
                return
            raw_html = DASHBOARD_HTML_PATH.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
            self.send_header("Content-Length", str(len(raw_html)))
            self.end_headers()
            self.wfile.write(raw_html)
            return

        if path in ("/api/status", "/data/status.json"):
            snap_count, latest_snap_date = get_cached_archive_meta()
            with STATE_LOCK:
                self._send_json({
                    "is_crawling": STATE["is_crawling"],
                    "current_label": STATE["current_label"],
                    "last_completed_label": STATE["last_completed_label"] or latest_snap_date,
                    "last_completed_at": STATE["last_completed_at"],
                    "last_error": STATE["last_error"],
                    "schedule_rule": "평일(월~금) 오전 09:30 / 오후 16:30 자동 크롤링 (토·일 제외)",
                    "next_scheduled_run": compute_next_weekday_run(),
                    "total_snapshots": snap_count,
                    "local_ip": get_local_ip(),
                    "local_url": "http://localhost:8080",
                    "lan_url": f"http://{get_local_ip()}:8080",
                    "fixed_portal_url": "https://telegra.ph/Geunbon-Hanwoo-B2B-Dashboard-11-19",
                    "github_pages_url": "https://lks230325.github.io/hanwoo-b2b-pipeline/",
                    "public_url": get_public_tunnel_url(),
                    "gdrive_folder_url": "https://drive.google.com/drive/folders/1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe",
                })
            return

        if path in ("/download/excel", "/data/hanwoo_master_cumulative.xlsx"):
            if not MASTER_EXCEL_PATH.exists():
                self.send_error(404, "Master Excel not found")
                return
            data = MASTER_EXCEL_PATH.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="hanwoo_master_cumulative.xlsx"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        if path == "/data/slim_archive.json":
            slim_file = BASE_DIR / "static_data" / "slim_archive.json"
            if not slim_file.exists():
                self.send_error(404, "slim_archive.json not found")
                return
            data = slim_file.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        self.send_error(404, "Not Found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/run-now":
            now = datetime.now()
            label = now.strftime("%Y-%m-%d %H시%M분")
            with STATE_LOCK:
                if STATE["is_crawling"]:
                    self._send_json({
                        "started": False,
                        "message": f"현재 '{STATE['current_label']}' 크롤링이 이미 진행 중입니다.",
                    })
                    return
            t = threading.Thread(target=run_crawl_job, args=(label,), daemon=True)
            t.start()
            self._send_json({
                "started": True,
                "label": label,
                "message": f"'{label}' 기준 6대 플랫폼 실시간 크롤링을 시작했습니다. (약 2분 소요)",
            })
            return
        self.send_error(404, "Not Found")


def main() -> None:
    port = 8080
    sched_thread = threading.Thread(target=weekday_scheduler_loop, daemon=True)
    sched_thread.start()

    server = ThreadingHTTPServer(("0.0.0.0", port), DashboardRequestHandler)
    logger.info("==================================================================")
    logger.info(" [근본한우 MD] 실시간 웹 대시보드 서버 오픈: http://localhost:%d", port)
    logger.info(" - 자동 스케줄: 매주 월~금 오전 09:30 / 오후 16:30 (토·일 제외)")
    logger.info(" - 다음 예정 시각: %s", compute_next_weekday_run())
    logger.info("==================================================================")
    server.serve_forever()


if __name__ == "__main__":
    main()
