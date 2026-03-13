@echo off
setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"
set LOG_FILE=%SCRIPT_DIR%uninstall_launcher.log
set SHOULD_PAUSE=1

for %%A in (%*) do (
  if /I "%%~A"=="-NoPause" set SHOULD_PAUSE=0
)

echo === Subtitle Tool - Full Uninstall ===
echo Script directory: %SCRIPT_DIR%
echo.
echo This will remove the virtual environment, pip caches, and Whisper model
echo cache. It will NOT uninstall Python, ffmpeg, or VC++ from your system.
echo.

where powershell >nul 2>&1
if %ERRORLEVEL% neq 0 (
  echo ERROR: powershell executable not found in PATH.
  pause
  exit /b 1
)

echo [%DATE% %TIME%] Uninstall started>>"%LOG_FILE%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install_all_windows.ps1" -Uninstall
set EXITCODE=%ERRORLEVEL%
echo [%DATE% %TIME%] Uninstall exit code: %EXITCODE%>>"%LOG_FILE%"

if %EXITCODE% neq 0 (
  echo.
  echo Uninstall encountered errors. Check the messages above.
  if %SHOULD_PAUSE% neq 0 pause
  exit /b %EXITCODE%
)

if %SHOULD_PAUSE% neq 0 pause
exit /b 0
