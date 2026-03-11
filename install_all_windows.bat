@echo off
setlocal

set SCRIPT_DIR=%~dp0
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install_all_windows.ps1" %*
set EXITCODE=%ERRORLEVEL%

if %EXITCODE% neq 0 (
  echo.
  echo Full install automation failed. Check the error messages above.
  pause
  exit /b %EXITCODE%
)

echo Full install automation succeeded.
pause
exit /b 0
