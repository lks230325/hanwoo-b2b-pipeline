@echo off
chcp 65001 > nul
echo ============================================================================
echo [근본한우 온라인팀 MD] 5대 축산 B2B 플랫폼 한우원물 크롤링 및 시세 분석 파이프라인 실행
echo ============================================================================
"C:\Users\limkyungsoon\.gemini\antigravity\scratch\meatbox-scraper\.venv\Scripts\python.exe" "%~dp0hanwoo_pipeline.py" %*
echo ============================================================================
echo 실행이 완료되었습니다. output 폴더 및 구글 드라이브(폴더 ID: 1up8hQdRmuvEFaVjWWtDEjbaVbgCciFSe)를 확인하세요.
pause
