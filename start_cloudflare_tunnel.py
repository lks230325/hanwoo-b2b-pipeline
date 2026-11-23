#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
[근본한우 B2B 대시보드] Cloudflare Quick Tunnel 데몬 래퍼
- http://localhost:8080 로컬 대시보드 서버를 외부 공인 HTTPS URL (*.trycloudflare.com)로 즉시 배포
- 발급된 공인 URL을 output/public_tunnel_url.txt 및 바탕화면 온라인팀\근본한우 폴더에 자동 기록
"""

import re
import subprocess
import sys
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CF_BIN = BASE_DIR / "cloudflared.exe"
OUTPUT_DIR = BASE_DIR / "output"
TUNNEL_URL_FILE = OUTPUT_DIR / "public_tunnel_url.txt"
DESKTOP_DIR = Path(r"C:\Users\limkyungsoon\Desktop\온라인팀\근본한우")
URL_REGEX = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com")


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if not CF_BIN.exists():
        print(f"ERROR: cloudflared.exe not found at {CF_BIN}", file=sys.stderr)
        sys.exit(1)

    cmd = [str(CF_BIN), "tunnel", "--url", "http://localhost:8080", "--no-autoupdate"]
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    found_url = None
    assert proc.stdout is not None
    for line in proc.stdout:
        sys.stdout.write(line)
        sys.stdout.flush()
        m = URL_REGEX.search(line)
        if m and not found_url:
            found_url = m.group(0)
            TUNNEL_URL_FILE.write_text(found_url, encoding="utf-8")
            print(f"\n[DEPLOYED_PUBLIC_HTTPS_URL] {found_url}\n", flush=True)
            try:
                if DESKTOP_DIR.exists():
                    (DESKTOP_DIR / "외부접속_공인URL.txt").write_text(
                        f"[근본한우 B2B 6대 플랫폼 통합 대시보드 접속 안내]\n"
                        f"1. 외부/모바일 공인 HTTPS URL (어디서나 접속): {found_url}\n"
                        f"2. 사내 Wi-Fi/LAN 접속 URL: http://192.168.0.58:8080\n"
                        f"3. 현재 PC 로컬 접속 URL: http://localhost:8080\n"
                        f"4. 구글 드라이브 자동 아카이브 폴더: https://drive.google.com/drive/folders/1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe\n",
                        encoding="utf-8",
                    )
            except Exception:
                pass

    proc.wait()


if __name__ == "__main__":
    main()
