@echo off
setlocal

set "PROFILE=%~dp0.edge-debug-profile"
set "EDGE=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=C:\Program Files\Microsoft\Edge\Application\msedge.exe"

echo ============================================================
echo  Opening Edge for manual login
echo ============================================================
echo.
echo  STEP 1: Log in to pharmacylibrary.com
echo  STEP 2: Keep this Edge window OPEN
echo  STEP 3: In the terminal, run:  npm run check
echo.
echo ============================================================

start "" "%EDGE%" --remote-debugging-port=9222 --user-data-dir="%PROFILE%" https://pharmacylibrary.com/
