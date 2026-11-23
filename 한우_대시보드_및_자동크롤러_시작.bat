@echo off
chcp 65001 >nul
title [근본한우 B2B] 한우 6대 플랫폼 대시보드 & 평일(09:30·16:30) 자동 크롤러/배포 서버
cd /d "C:\Users\limkyungsoon\.gemini\antigravity\scratch\hanwoo-b2b-pipeline"
echo ===========================================================================
echo  [근본한우 B2B 인텔리전스] 한우 6대 플랫폼 대시보드 서버 및 자동 크롤러 기동
echo  - 정기 크롤링: 매주 월~금 오전 09:30 / 오후 16:30 (토·일 제외)
echo  - 구글 드라이브 자동 배포: 1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe
echo  - PC 접속: http://localhost:8080
echo  - 사내 Wi-Fi/모바일 접속: http://192.168.0.58:8080
echo ===========================================================================
start "" "http://localhost:8080"
start "Hanwoo Cloudflare Tunnel" /min "C:\Users\limkyungsoon\.gemini\antigravity\scratch\meatbox-scraper\.venv\Scripts\python.exe" -X utf8 start_cloudflare_tunnel.py
"C:\Users\limkyungsoon\.gemini\antigravity\scratch\meatbox-scraper\.venv\Scripts\python.exe" -X utf8 web_dashboard_server.py
pause
