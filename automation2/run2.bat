@echo off
REM ==============================================================
REM  Website verifier -- SYSTEM 2 (second copy, second login).
REM  Double-click this file to start a SYSTEM 2 run.
REM
REM  This runs website_verifier2.py, which uses:
REM      .env2      for credentials
REM      debug2\    for screenshots, form dumps and run_log.csv
REM
REM  SYSTEM 1 is run.bat / website_verifier.py and is unaffected.
REM ==============================================================
cd /d "%~dp0"
echo Starting website verifier SYSTEM 2 from %cd%
echo.
python "website_verifier2.py"
echo.
echo ==============================================================
echo  SYSTEM 2 run finished. Press any key to close this window.
echo ==============================================================
pause >nul
