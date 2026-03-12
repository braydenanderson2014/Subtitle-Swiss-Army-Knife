@echo off
setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"
set LOG_FILE=%SCRIPT_DIR%install_ffmpeg_windows_launcher.log

echo === FFmpeg Installer Launcher ===
echo Script directory: %SCRIPT_DIR%
echo Log file: %LOG_FILE%
echo.

where powershell >nul 2>&1
if %ERRORLEVEL% neq 0 (
  echo ERROR: powershell executable not found in PATH.
  echo This launcher cannot continue.
  pause
  exit /b 1
)

echo [%DATE% %TIME%] Launcher started>>"%LOG_FILE%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install_ffmpeg_windows.ps1" %*
set EXITCODE=%ERRORLEVEL%
echo [%DATE% %TIME%] Launcher exit code: %EXITCODE%>>"%LOG_FILE%"

if %EXITCODE% neq 0 (
  echo.
  echo ffmpeg installation failed. Check the error messages above.
  echo See launcher log: %LOG_FILE%
  exit /b %EXITCODE%
)

echo ffmpeg installation succeeded.
exit /b 0
