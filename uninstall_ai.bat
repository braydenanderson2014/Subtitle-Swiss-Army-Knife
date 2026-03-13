@echo off
setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"
set LOG_FILE=%SCRIPT_DIR%uninstall_ai_launcher.log
set SHOULD_PAUSE=1

for %%A in (%*) do (
  if /I "%%~A"=="-NoPause" set SHOULD_PAUSE=0
)

echo === Subtitle Tool - Uninstall AI Libraries ===
echo Script directory: %SCRIPT_DIR%
echo.
echo This will remove AI packages (PyTorch, Whisper, pysubs2, cinemagoer)
echo from the virtual environment, and delete any downloaded Whisper model
echo files from disk. The core venv and app remain functional.
echo.

where powershell >nul 2>&1
if %ERRORLEVEL% neq 0 (
  echo ERROR: powershell executable not found in PATH.
  pause
  exit /b 1
)

echo [%DATE% %TIME%] AI uninstall started>>"%LOG_FILE%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install_all_windows.ps1" -UninstallAI
set EXITCODE=%ERRORLEVEL%
echo [%DATE% %TIME%] AI uninstall exit code: %EXITCODE%>>"%LOG_FILE%"

if %EXITCODE% neq 0 (
  echo.
  echo AI uninstall encountered errors. Check the messages above.
  if %SHOULD_PAUSE% neq 0 pause
  exit /b %EXITCODE%
)

if %SHOULD_PAUSE% neq 0 pause
exit /b 0
