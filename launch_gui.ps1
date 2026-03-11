$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPython = Join-Path $scriptDir "venv\Scripts\python.exe"
$subtitleTool = Join-Path $scriptDir "subtitle_tool.py"

Set-Location $scriptDir

if (-not (Test-Path $venvPython)) {
    Write-Host "Virtual environment not found at: $venvPython" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please run install_all_windows.bat or install_all_windows.ps1 first to set up the environment." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Launching Subtitle Tool GUI..." -ForegroundColor Cyan
Write-Host ""

& $venvPython $subtitleTool gui

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "Failed to launch GUI. Exit code: $LASTEXITCODE" -ForegroundColor Red
    Write-Host ""
    Write-Host "If dependencies are missing, try running install_all_windows.ps1 again." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit $LASTEXITCODE
}

exit 0
