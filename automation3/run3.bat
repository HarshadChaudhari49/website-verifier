@echo off
REM ==============================================================
REM  Website verifier -- SYSTEM 3  (CHROME).
REM  Double-click this file to start the portal + ChatGPT flow.
REM
REM  It uses:
REM      CHROME                    System 3 runs Chrome, not Firefox
REM      .env3                     portal credentials
REM      MASTER_RULES_v3.1.md      the rulebook (project root)
REM      debug3\                   screenshots, dumps, gpt_flow_log.csv
REM      chatgpt_profiles\default\ the saved ChatGPT sign-in
REM      portal_chrome_profile3\   the portal's own Chrome profile
REM
REM  FIRST TIME ONLY -- run this once and sign in by hand:
REM      run3.bat --chatgpt-login
REM  It opens a NORMAL Chrome window (not automated), which is why
REM  Google accepts the sign-in. Sign in, CLOSE that window, then
REM  press Enter in the terminal. Every run after that opens
REM  already logged in.
REM
REM  Any argument given here is passed straight through:
REM      run3.bat                        portal + ChatGPT  (default)
REM      run3.bat --chatgpt-login        one-time ChatGPT sign-in
REM      run3.bat --chatgpt-profile work use a second saved account
REM      run3.bat --login-only           portal login, then stop
REM      run3.bat --dump-form            read-only form dump
REM
REM  Systems 1 and 2 are unaffected: automation1\ and automation2\
REM  keep their own scripts, credentials and Firefox browser.
REM ==============================================================
cd /d "%~dp0"

set "MODE=%*"

echo Starting website verifier SYSTEM 3 from %cd%
if "%MODE%"=="" (echo Mode: portal + ChatGPT [default]) else (echo Mode: %MODE%)
echo.
python "website_verifier3.py" %MODE%
echo.
echo ==============================================================
echo  SYSTEM 3 run finished. Press any key to close this window.
echo ==============================================================
pause >nul
