@echo off
REM ==============================================================
REM  Website verifier -- SYSTEM 2.
REM  Double-click this file to start the portal + ChatGPT flow.
REM
REM  It uses:
REM      .env2               for credentials
REM      MASTER_RULES.md     the rulebook (project root)
REM      debug2\             screenshots, form dumps, gpt_flow_log.csv
REM
REM  The portal + ChatGPT flow is System 2's only engine. Any
REM  argument given here is passed straight through, so:
REM      run2.bat                  portal + ChatGPT   (default)
REM      run2.bat --login-only     log in and stop
REM      run2.bat --chatgpt-login  one-time ChatGPT sign-in
REM      run2.bat --dump-form      read-only form dump
REM
REM  The built-in crawler engine was removed from System 2 on
REM  2026-09-05. It lives on in SYSTEM 1: automation1\run.bat and
REM  automation1\website_verifier.py, which are unaffected.
REM ==============================================================
cd /d "%~dp0"

set "MODE=%*"
if "%MODE%"=="" set "MODE=--gpt-flow"

echo Starting website verifier SYSTEM 2 from %cd%
echo Mode: %MODE%
echo.
python "website_verifier2.py" %MODE%
echo.
echo ==============================================================
echo  SYSTEM 2 run finished. Press any key to close this window.
echo ==============================================================
pause >nul
