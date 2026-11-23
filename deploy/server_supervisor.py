#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[근본한우 B2B 인텔리전스] 상시 서버 & 무료 영구 고정 HTTPS 배포 URL 슈퍼바이저 (Self-Healing Supervisor)
- 단일 인스턴스 락(127.0.0.1:48898)으로 중복 실행 100% 방지 (윈도우 스케줄러 워치독 연동)
- 웹 대시보드 서버(web_dashboard_server.py :8080) 상태 상시 감시 및 종료 시 즉시 자동 재기동
- 무료 영구 고정 HTTPS 배포 URL(tunwg + 로컬 WireGuard 고유키 + Let's Encrypt SSL) 상시 유지 (`TUNWG_PATH=deploy/tunwg_keys`)
  -> PC 재부팅·프로세스 재시작·와이파이 변경 시에도 동일한 고정 도메인(`https://*.l.tunwg.com`) 영구 유지
- 보조 백업 Cloudflare HTTPS 터널(`cloudflared.exe`) 및 고정 포털(`telegra.ph`) 동시 가동으로 다중화 보장
"""

import json
import logging
import os
import re
import socket
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

DEPLOY_DIR = Path(__file__).resolve().parent
WORK_DIR = DEPLOY_DIR.parent
OUTPUT_DIR = WORK_DIR / "output"
LOGS_DIR = DEPLOY_DIR / "logs"
TUNWG_KEYS_DIR = DEPLOY_DIR / "tunwg_keys"
LOGS_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TUNWG_KEYS_DIR.mkdir(parents=True, exist_ok=True)

STATUS_JSON_PATH = DEPLOY_DIR / "deploy_status.json"
TUNNEL_URL_FILE = OUTPUT_DIR / "public_tunnel_url.txt"
FIXED_TUNNEL_URL_FILE = OUTPUT_DIR / "fixed_tunnel_url.txt"
DESKTOP_DIR = Path(r"C:\Users\limkyungsoon\Desktop\온라인팀\근본한우")
URL_GUIDE_TXT_PATH = DESKTOP_DIR / "외부접속_공인URL.txt"
TUNWG_EXE = DEPLOY_DIR / "tunwg.exe"
CLOUDFLARED_EXE = WORK_DIR / "cloudflared.exe"
VENV_PYTHON = Path(r"C:\Users\limkyungsoon\.gemini\antigravity\scratch\meatbox-scraper\.venv\Scripts\python.exe")

SINGLE_INSTANCE_PORT = 48898
APP_PORT = 8080
CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

FIXED_PORTAL_URL = "https://telegra.ph/Geunbon-Hanwoo-B2B-Dashboard-11-19"
TELEGRAPH_TOKEN = "510c0a7176ad52351c2aada81646b3e9cd1c0160d22aa53a08ddd9316d48"
TELEGRAPH_PATH = "Geunbon-Hanwoo-B2B-Dashboard-11-19"
NTFY_TOPIC_URL = "https://ntfy.sh/geunbon_hanwoo_b2b_live_8080"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] HanwooSupervisor - %(message)s",
    handlers=[
        logging.FileHandler(LOGS_DIR / "supervisor.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("HanwooSupervisor")


def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def is_server_healthy() -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{APP_PORT}/api/status", timeout=4) as resp:
            return resp.status == 200
    except Exception:
        return False


def is_remote_url_healthy(url: str) -> bool:
    if not url:
        return False
    try:
        req = urllib.request.Request(
            f"{url.rstrip('/')}/api/status",
            headers={"User-Agent": "HanwooSupervisorHealthCheck/1.0"},
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            return resp.status == 200
    except Exception:
        return False


def get_python_exe() -> str:
    if VENV_PYTHON.exists():
        return str(VENV_PYTHON)
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        cand = exe.with_name("python.exe")
        if cand.exists():
            return str(cand)
    return str(exe)


class HanwooDeploymentSupervisor:
    def __init__(self) -> None:
        self.web_proc: Optional[subprocess.Popen] = None
        self.fixed_tunnel_proc: Optional[subprocess.Popen] = None
        self.cf_tunnel_proc: Optional[subprocess.Popen] = None
        self.fixed_public_url: str = ""
        self.fixed_started_at: str = ""
        self.backup_cf_url: str = ""
        self.cf_started_at: str = ""
        self.started_at: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._last_remote_probe_ts: float = 0.0
        self._remote_fail_streak: int = 0
        self._lock = threading.Lock()

        if STATUS_JSON_PATH.exists():
            try:
                prev = json.loads(STATUS_JSON_PATH.read_text(encoding="utf-8"))
                prev_fixed = str(prev.get("fixed_https_url") or "")
                if ".tunwg.com" in prev_fixed:
                    self.fixed_public_url = prev_fixed
            except Exception:
                pass

    def write_status_files(self) -> None:
        local_ip = get_local_ip()
        pc_url = f"http://localhost:{APP_PORT}"
        lan_url = f"http://{local_ip}:{APP_PORT}"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with self._lock:
            fixed_url = self.fixed_public_url
            f_start = self.fixed_started_at
            cf_url = self.backup_cf_url
            c_start = self.cf_started_at

        primary_url = fixed_url or cf_url
        if primary_url:
            try:
                TUNNEL_URL_FILE.write_text(primary_url, encoding="utf-8")
            except Exception:
                pass
        if fixed_url:
            try:
                FIXED_TUNNEL_URL_FILE.write_text(fixed_url, encoding="utf-8")
            except Exception:
                pass

        status_data: Dict[str, Any] = {
            "supervisor_pid": os.getpid(),
            "supervisor_started_at": self.started_at,
            "updated_at": now_str,
            "pc_url": pc_url,
            "lan_mobile_url": lan_url,
            "public_https_url": primary_url,
            "url_is_permanent_fixed": bool(fixed_url),
            "fixed_https_url": fixed_url,
            "fixed_tunnel_started_at": f_start,
            "backup_cloudflare_url": cf_url,
            "backup_tunnel_started_at": c_start,
            "fixed_portal_url": FIXED_PORTAL_URL,
            "server_healthy": is_server_healthy(),
            "gdrive_folder_url": "https://drive.google.com/drive/folders/1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe",
            "schedule": "매주 월~금 오전 09:30 / 오후 16:30 자동 크롤링 & 구글 드라이브 동기화 (토·일 제외)",
        }
        try:
            STATUS_JSON_PATH.write_text(json.dumps(status_data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            logger.warning("상태 JSON 저장 실패: %s", exc)

        guide_lines = [
            "==========================================================================",
            " [근본한우 B2B 인텔리전스] 한우 6대 플랫폼 대시보드 배포 접속 안내",
            "==========================================================================",
            f"- 상태 갱신 시각          : {now_str}",
            f"- [고정 배포 URL · 무료]  : {fixed_url or '(고정 터널 연결 중...)'}",
            "  * PC 재부팅/재기동/와이파이 변경 시에도 주소가 변경되지 않는 영구 고정 HTTPS 주소입니다.",
            "  * 즐겨찾기(북마크) 및 모바일 홈화면에 고정 등록하여 사용하세요.",
            f"- [보조 백업 HTTPS URL]   : {cf_url or '(보조 터널 연결 중...)'}",
            f"- [영구 고정 포털 URL]    : {FIXED_PORTAL_URL}",
            f"- 사내 동일 Wi-Fi 접속    : {lan_url}",
            f"- 현재 PC 로컬 접속       : {pc_url}",
            "- 구글 드라이브 폴더      : https://drive.google.com/drive/folders/1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe",
            "- 자동 크롤링 스케줄      : 매주 월~금 오전 09:30 / 오후 16:30 (자동 누적 & 드라이브 동기화)",
            "==========================================================================",
            "",
        ]
        try:
            DESKTOP_DIR.mkdir(parents=True, exist_ok=True)
            URL_GUIDE_TXT_PATH.write_text("\n".join(guide_lines), encoding="utf-8")
        except Exception as exc:
            logger.warning("안내 텍스트 저장 실패: %s", exc)

    def sync_fixed_portal_url(self, live_url: str) -> None:
        if not live_url:
            return
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._lock:
            fixed_u = self.fixed_public_url or live_url
            backup_u = self.backup_cf_url or live_url
        active_primary = fixed_u if (fixed_u and is_remote_url_healthy(fixed_u)) else (backup_u or fixed_u)
        try:
            req_n = urllib.request.Request(
                NTFY_TOPIC_URL,
                data=active_primary.encode("utf-8"),
                method="POST",
            )
            with urllib.request.urlopen(req_n, timeout=8) as _:
                pass
        except Exception as exc:
            logger.warning("ntfy 고정 채널 발행 실패: %s", exc)

        try:
            content = [
                {"tag": "h3", "children": ["🐂 [근본한우 B2B 인텔리전스] 6대 플랫폼 실시간 통합 대시보드 (고정 접속 포털)"]},
                {"tag": "p", "children": [{"tag": "strong", "children": ["👉 [지금 즉시 접속 · 모든 네트워크/핫스팟 호환]: "]}, {"tag": "a", "attrs": {"href": active_primary}, "children": [active_primary]}]},
                {"tag": "p", "children": [{"tag": "strong", "children": ["👉 [사내/일반 Wi-Fi 고정 도메인]: "]}, {"tag": "a", "attrs": {"href": fixed_u}, "children": [fixed_u]}]},
                {"tag": "p", "children": [{"tag": "strong", "children": ["👉 [Cloudflare TCP 백업 터널]: "]}, {"tag": "a", "attrs": {"href": backup_u}, "children": [backup_u]}]},
                {"tag": "p", "children": [f"• 최종 자동 갱신 시각: {now_str} (상시 슈퍼바이저 실시간 연동 중)"]},
                {"tag": "h4", "children": ["📊 정기 크롤링 및 구글 드라이브 아카이브"]},
                {"tag": "p", "children": ["• 정기 스케줄: 매주 월~금 오전 09:30 / 오후 16:30 자동 크롤링 및 동기화 (토·일 제외)"]},
                {"tag": "p", "children": ["• 구글 드라이브 마스터 폴더: ", {"tag": "a", "attrs": {"href": "https://drive.google.com/drive/folders/1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe"}, "children": ["구글 드라이브 바로가기"]}]},
            ]
            form_data = urllib.parse.urlencode({
                "access_token": TELEGRAPH_TOKEN,
                "path": TELEGRAPH_PATH,
                "title": "Geunbon-Hanwoo-B2B-Dashboard",
                "author_name": "근본한우 온라인팀 MD",
                "content": json.dumps(content, ensure_ascii=False),
                "return_content": "false",
            }).encode("utf-8")
            req_t = urllib.request.Request("https://api.telegra.ph/editPage", data=form_data, method="POST")
            with urllib.request.urlopen(req_t, timeout=8) as _:
                logger.info("영구 고정 포털 URL(%s) -> 고정 터널(%s) 자동 매핑 완료", FIXED_PORTAL_URL, fixed_u)
        except Exception as exc:
            logger.warning("영구 고정 포털 업데이트 실패: %s", exc)

    def ensure_web_server(self) -> None:
        if is_server_healthy():
            return
        if self.web_proc is not None and self.web_proc.poll() is None:
            time.sleep(2)
            if is_server_healthy():
                return

        py_exe = get_python_exe()
        srv_log = open(LOGS_DIR / "web_server.log", "a", encoding="utf-8")
        cmd = [py_exe, "-X", "utf8", str(WORK_DIR / "web_dashboard_server.py")]
        logger.info("한우 대시보드 서버 기동 시작: %s", " ".join(cmd))
        self.web_proc = subprocess.Popen(
            cmd,
            cwd=str(WORK_DIR),
            stdout=srv_log,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW,
        )
        for _ in range(20):
            time.sleep(0.5)
            if is_server_healthy():
                logger.info("한우 대시보드 서버 정상 응답 확인 (PID=%s)", self.web_proc.pid)
                break

    def _watch_fixed_tunnel_output(self, proc: subprocess.Popen) -> None:
        url_regex = re.compile(r"(https://[a-zA-Z0-9\-\.]+\.tunwg\.com)")
        tun_log_path = LOGS_DIR / "fixed_tunnel.log"
        with open(tun_log_path, "a", encoding="utf-8") as tlog:
            assert proc.stdout is not None
            for raw_line in iter(proc.stdout.readline, ""):
                if not raw_line:
                    break
                tlog.write(raw_line)
                tlog.flush()
                m = url_regex.search(raw_line)
                if m:
                    found_url = m.group(1).strip()
                    with self._lock:
                        changed = self.fixed_public_url != found_url
                        self.fixed_public_url = found_url
                        self.fixed_started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        self._remote_fail_streak = 0
                    if changed:
                        logger.info("무료 영구 고정 HTTPS 배포 터널(tunwg) 개통 완료: %s", found_url)
                    self.write_status_files()
                    self.sync_fixed_portal_url(found_url)

    def ensure_fixed_tunnel(self) -> None:
        if not TUNWG_EXE.exists():
            logger.error("tunwg.exe 파일이 존재하지 않습니다: %s", TUNWG_EXE)
            return

        if self.fixed_tunnel_proc is not None and self.fixed_tunnel_proc.poll() is None:
            now_ts = time.time()
            if self.fixed_public_url and (now_ts - self._last_remote_probe_ts >= 180):
                self._last_remote_probe_ts = now_ts
                if is_remote_url_healthy(self.fixed_public_url):
                    self._remote_fail_streak = 0
                    return
                self._remote_fail_streak += 1
                logger.warning(
                    "고정 HTTPS 터널 응답 지연 감지 (%d/2): %s",
                    self._remote_fail_streak,
                    self.fixed_public_url,
                )
                if self._remote_fail_streak >= 2:
                    logger.warning("고정 HTTPS 터널 재연결 실행 (동일 고정 URL 유지)")
                    try:
                        self.fixed_tunnel_proc.terminate()
                        self.fixed_tunnel_proc.wait(timeout=5)
                    except Exception:
                        try:
                            self.fixed_tunnel_proc.kill()
                        except Exception:
                            pass
                    self._remote_fail_streak = 0
                else:
                    return
            else:
                return

        env = os.environ.copy()
        env["TUNWG_PATH"] = str(TUNWG_KEYS_DIR)
        cmd = [str(TUNWG_EXE), f"--forward=http://127.0.0.1:{APP_PORT}"]
        logger.info("무료 영구 고정 HTTPS 터널(tunwg) 프로세스 시작 (키 경로: %s)", TUNWG_KEYS_DIR)
        self.fixed_tunnel_proc = subprocess.Popen(
            cmd,
            cwd=str(WORK_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            creationflags=CREATE_NO_WINDOW,
        )
        self._last_remote_probe_ts = time.time()
        t = threading.Thread(target=self._watch_fixed_tunnel_output, args=(self.fixed_tunnel_proc,), daemon=True)
        t.start()

    def _watch_cf_tunnel_output(self, proc: subprocess.Popen) -> None:
        url_regex = re.compile(r"(https://[a-zA-Z0-9\-]+\.trycloudflare\.com)")
        tun_log_path = LOGS_DIR / "tunnel.log"
        with open(tun_log_path, "a", encoding="utf-8") as tlog:
            assert proc.stderr is not None
            for raw_line in iter(proc.stderr.readline, ""):
                if not raw_line:
                    break
                tlog.write(raw_line)
                tlog.flush()
                m = url_regex.search(raw_line)
                if m:
                    found_url = m.group(1).strip()
                    with self._lock:
                        changed = self.backup_cf_url != found_url
                        self.backup_cf_url = found_url
                        self.cf_started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if changed:
                        logger.info("보조 Cloudflare HTTPS 터널 개통 완료: %s", found_url)
                        self.write_status_files()
                        self.sync_fixed_portal_url(self.fixed_public_url or found_url)

    def ensure_cf_backup_tunnel(self) -> None:
        if not CLOUDFLARED_EXE.exists():
            return
        if self.cf_tunnel_proc is not None and self.cf_tunnel_proc.poll() is None:
            return

        cmd = [
            str(CLOUDFLARED_EXE),
            "tunnel",
            "--url",
            f"http://127.0.0.1:{APP_PORT}",
            "--no-autoupdate",
        ]
        logger.info("보조 Cloudflare HTTPS 터널 프로세스 시작")
        self.cf_tunnel_proc = subprocess.Popen(
            cmd,
            cwd=str(WORK_DIR),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=CREATE_NO_WINDOW,
        )
        t = threading.Thread(target=self._watch_cf_tunnel_output, args=(self.cf_tunnel_proc,), daemon=True)
        t.start()

    def run_forever(self) -> None:
        logger.info("한우 슈퍼바이저 상시 가동 시작 (PID=%d)", os.getpid())
        self.ensure_web_server()
        self.ensure_fixed_tunnel()
        self.ensure_cf_backup_tunnel()
        self.write_status_files()

        while True:
            time.sleep(15)
            try:
                self.ensure_web_server()
                self.ensure_fixed_tunnel()
                self.ensure_cf_backup_tunnel()
                self.write_status_files()
            except Exception as exc:
                logger.error("슈퍼바이저 루프 오류: %s", exc)


def main() -> int:
    lock_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        lock_sock.bind(("127.0.0.1", SINGLE_INSTANCE_PORT))
        lock_sock.listen(1)
    except OSError:
        # 이미 슈퍼바이저가 정상 실행 중이므로 조용히 종료
        return 0

    sup = HanwooDeploymentSupervisor()
    sup.run_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
