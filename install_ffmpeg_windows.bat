@echo off
setlocal

set SCRIPT_DIR=%~dp0
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install_ffmpeg_windows.ps1" %*
set EXITCODE=%ERRORLEVEL%

if %EXITCODE% neq 0 (
  echo ffmpeg installation failed.
  exit /b %EXITCODE%
)

echo ffmpeg installation succeeded.
exit /b 0
