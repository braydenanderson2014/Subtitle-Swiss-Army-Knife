@echo off
setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

set VENV_PYTHON=%SCRIPT_DIR%venv\Scripts\python.exe

if not exist "%VENV_PYTHON%" (
  echo Virtual environment not found at: %VENV_PYTHON%
  echo.
  echo Please run install_all_windows.bat first to set up the environment.
  pause
  exit /b 1
)

echo Launching Subtitle Tool GUI...
"%VENV_PYTHON%" subtitle_tool.py gui

if %ERRORLEVEL% neq 0 (
  echo.
  echo Failed to launch GUI. Exit code: %ERRORLEVEL%
  echo.
  echo If dependencies are missing, try running install_all_windows.bat again.
  pause
  exit /b %ERRORLEVEL%
)

exit /b 0
