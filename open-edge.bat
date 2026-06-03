@echo off
setlocal

set "BROWSERS=%~1"
if "%BROWSERS%"=="" set "BROWSERS=1"

set "EDGE=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=C:\Program Files\Microsoft\Edge\Application\msedge.exe"

set "DIR=%~dp0"
set "BASE=%DIR%.edge-debug-profile-0"
set "OLD=%DIR%.edge-debug-profile"

echo ============================================================
echo  Opening %BROWSERS% Edge window(s) for DOI scan
echo ============================================================
echo.

:: Migrate old single-profile to profile-0 if needed
if exist "%OLD%" if not exist "%BASE%" rename "%OLD%" ".edge-debug-profile-0"

:: First-time setup: profile-0 does not exist yet
if exist "%BASE%" goto copy_extras
echo  [Setup] No base profile found.
echo  A fresh Edge window will open. Log in to pharmacylibrary.com.
echo  Then re-run:  .\open-edge.bat %BROWSERS%
echo.
start "" "%EDGE%" --remote-debugging-port=9222 --user-data-dir="%BASE%" https://pharmacylibrary.com/
goto done

:copy_extras
:: Clone profile-0 to any missing profiles before opening any window
set /a IDX=1
:copy_loop
if %IDX% GEQ %BROWSERS% goto open_browsers
set "DEST=%DIR%.edge-debug-profile-%IDX%"
if exist "%DEST%" goto skip_clone
echo  Cloning session to profile-%IDX% - one moment...
robocopy "%BASE%" "%DEST%" /E /NFL /NDL /NJH /NJS 1>nul 2>nul
echo  Done.
:skip_clone
set /a IDX+=1
goto copy_loop

:open_browsers
set /a PORT=9222
set /a IDX=0
:open_loop
if %IDX% GEQ %BROWSERS% goto done
set "PROFILE=%DIR%.edge-debug-profile-%IDX%"
echo  Starting Edge  port=%PORT%  profile-%IDX%
start "" "%EDGE%" --remote-debugging-port=%PORT% --user-data-dir="%PROFILE%" https://pharmacylibrary.com/
set /a PORT+=1
set /a IDX+=1
goto open_loop

:done
echo.
echo  Keep all windows open, then run:  npm run find-dois
echo ============================================================
