param(
    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$PythonInstallMethod = "auto",

    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$FfmpegInstallMethod = "auto"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'

# Catch any terminating error, show it, and pause so the window stays open
trap {
    Write-Host ""
    Write-Host "=== INSTALLATION FAILED ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "The window will stay open so you can read the error above." -ForegroundColor Yellow
    Read-Host "Press Enter to close"
    exit 1
}

function Write-StatusOK {
    param([string]$Message)
    Write-Host "$Message..." -NoNewline -ForegroundColor White
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "OK" -NoNewline -ForegroundColor Green
    Write-Host " ]" -ForegroundColor White
}

function Write-StatusFail {
    param([string]$Message)
    Write-Host "$Message..." -NoNewline -ForegroundColor White
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "FAIL" -NoNewline -ForegroundColor Red
    Write-Host " ]" -ForegroundColor White
}

function Test-CommandAvailable {
    param([string]$CommandName)
    return [bool](Get-Command $CommandName -ErrorAction SilentlyContinue)
}

function Test-VCRedist {
    # Check for Visual C++ Redistributable 2015-2022 (required for PyTorch)
    # Check multiple registry paths for different VC++ versions
    $vcRedistKeys = @(
        # Visual Studio 2015+
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64",
        # Visual Studio 2017+
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\15.0\VC\Runtimes\x64",
        # Visual Studio 2019+
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\16.0\VC\Runtimes\x64",
        # Visual Studio 2022+
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\17.0\VC\Runtimes\x64",
        # WOW6432Node versions
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\15.0\VC\Runtimes\x64",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\16.0\VC\Runtimes\x64",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\17.0\VC\Runtimes\x64"
    )
    
    foreach ($key in $vcRedistKeys) {
        if (Test-Path $key) {
            $installed = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
            if ($installed -and $installed.Installed -eq 1) {
                return $true
            }
        }
    }
    return $false
}

function Install-VCRedistWithWinget {
    Write-Host "Installing Visual C++ Redistributable via winget..." -ForegroundColor Yellow
    try {
        $process = Start-Process winget -ArgumentList "install", "--id", "Microsoft.VCRedist.2015+.x64", "--silent", "--accept-package-agreements", "--accept-source-agreements" -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 996) {
            # Exit code 996 means already installed
            Write-Host "Visual C++ Redistributable ready via winget." -ForegroundColor Green
            return $true
        }
        return $false
    } catch {
        return $false
    }
}

function Install-VCRedistWithChoco {
    Write-Host "Installing Visual C++ Redistributable via chocolatey..." -ForegroundColor Yellow
    try {
        $process = Start-Process choco -ArgumentList "install", "vcredist-all", "-y" -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0) {
            Write-Host "Visual C++ Redistributable ready via chocolatey." -ForegroundColor Green
            return $true
        }
        return $false
    } catch {
        return $false
    }
}

function Install-VCRedist {
    # Try package managers first, then fall back to manual installation
    $installed = $false
    
    # Try winget
    if (Test-CommandAvailable "winget") {
        if (Install-VCRedistWithWinget) {
            $installed = $true
        }
    }
    
    # Try chocolatey if winget failed
    if (-not $installed -and (Test-CommandAvailable "choco")) {
        if (Install-VCRedistWithChoco) {
            $installed = $true
        }
    }
    
    # Fall back to manual installation
    if (-not $installed) {
        Write-Host "Downloading Visual C++ Redistributable manually..." -ForegroundColor Yellow
        $vcRedistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
        $vcRedistInstaller = Join-Path $env:TEMP "vc_redist.x64.exe"
        
        try {
            # Remove old installer if it exists
            if (Test-Path $vcRedistInstaller) {
                Remove-Item $vcRedistInstaller -Force -ErrorAction SilentlyContinue
            }
            
            Invoke-WebRequest -Uri $vcRedistUrl -OutFile $vcRedistInstaller -UseBasicParsing
            Write-Host "Installing Visual C++ Redistributable..." -ForegroundColor Yellow
            $process = Start-Process -FilePath $vcRedistInstaller -ArgumentList "/install", "/quiet", "/norestart" -Wait -PassThru
            
            # Wait a moment for registry to update
            Start-Sleep -Seconds 2
            
            Remove-Item $vcRedistInstaller -ErrorAction SilentlyContinue
            Write-Host "Visual C++ Redistributable installed successfully." -ForegroundColor Green
            $installed = $true
        } catch {
            Write-Host "Failed to install VC++ Redistributable: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }
    
    # Verify installation by checking registry again
    Start-Sleep -Seconds 1
    if (-not (Test-VCRedist)) {
        Write-Host "Warning: VC++ installation verification failed. May need manual restart." -ForegroundColor Yellow
        return $false
    }
    
    return $installed
}

function Install-Winget {
    Write-Host "Installing winget (App Installer)..."
    try {
        # Install App Installer from Microsoft Store
        $progressPreference = 'silentlyContinue'
        Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe
        Write-Host "Winget installed successfully."
    } catch {
        throw "Failed to install winget: $($_.Exception.Message)"
    }
}

function Install-Chocolatey {
    Write-Host "Installing Chocolatey..."
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        # Refresh PATH to include choco
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        Write-Host "Chocolatey installed successfully."
    } catch {
        throw "Failed to install Chocolatey: $($_.Exception.Message)"
    }
}

function Ensure-PackageManager {
    Write-Host "Checking for package managers" -NoNewline -ForegroundColor White
    
    if (Test-CommandAvailable "winget") {
        Write-Host "..." -NoNewline -ForegroundColor White
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White
        return
    }
    
    if (Test-CommandAvailable "choco") {
        Write-Host "..." -NoNewline -ForegroundColor White
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White
        return
    }
    
    Write-Host "" # New line
    Write-Host "No package manager found. Installing one..." -ForegroundColor Yellow
    
    # Try winget first (modern Windows 10/11)
    try {
        Install-Winget
        if (Test-CommandAvailable "winget") {
            return
        }
    } catch {
        Write-Host "Winget installation failed: $($_.Exception.Message)"
    }
    
    # Fall back to Chocolatey
    try {
        Install-Chocolatey
        if (Test-CommandAvailable "choco") {
            return
        }
    } catch {
        throw "Failed to install any package manager. Please install winget or Chocolatey manually."
    }
    
    throw "Package manager installation completed but commands are not available."
}

function Find-PythonCommand {
    if (Test-CommandAvailable "python") {
        return "python"
    }
    if (Test-CommandAvailable "py") {
        return "py -3"
    }
    return $null
}

function Install-PythonWithWinget {
    Write-Host "Installing Python using winget..."
    winget install --id Python.Python.3.12 --exact --silent --accept-source-agreements --accept-package-agreements
}

function Install-PythonWithChoco {
    Write-Host "Installing Python using Chocolatey..."
    choco install python --yes
}

function Install-PythonWithScoop {
    Write-Host "Installing Python using Scoop..."
    scoop install python
}

function Install-Python {
    param([string]$Method)

    $attempted = @()

    if ($Method -eq "winget") {
        if (-not (Test-CommandAvailable "winget")) { throw "winget not found." }
        Install-PythonWithWinget
        $attempted += "winget"
    } elseif ($Method -eq "choco") {
        if (-not (Test-CommandAvailable "choco")) { throw "choco not found." }
        Install-PythonWithChoco
        $attempted += "choco"
    } elseif ($Method -eq "scoop") {
        if (-not (Test-CommandAvailable "scoop")) { throw "scoop not found." }
        Install-PythonWithScoop
        $attempted += "scoop"
    } else {
        if (Test-CommandAvailable "winget") {
            try {
                Install-PythonWithWinget
                $attempted += "winget"
            } catch {
                Write-Host "winget Python install failed: $($_.Exception.Message)"
            }
        }
        if (-not (Find-PythonCommand) -and (Test-CommandAvailable "choco")) {
            try {
                Install-PythonWithChoco
                $attempted += "choco"
            } catch {
                Write-Host "choco Python install failed: $($_.Exception.Message)"
            }
        }
        if (-not (Find-PythonCommand) -and (Test-CommandAvailable "scoop")) {
            try {
                Install-PythonWithScoop
                $attempted += "scoop"
            } catch {
                Write-Host "scoop Python install failed: $($_.Exception.Message)"
            }
        }
    }

    if (-not (Find-PythonCommand)) {
        if ($attempted.Count -eq 0) {
            throw "Python not found and no supported package manager is available (winget/choco/scoop)."
        }
        throw "Python installation failed after trying: $($attempted -join ', ')."
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$requirementsPath = Join-Path $scriptDir "requirements.txt"
$requirementsAIPath = Join-Path $scriptDir "requirements_ai.txt"
$ffmpegInstallerPath = Join-Path $scriptDir "install_ffmpeg_windows.ps1"
$subtitleToolPath = Join-Path $scriptDir "subtitle_tool.py"
$venvPath = Join-Path $scriptDir "venv"

# When running from a non-system drive, redirect pip's cache and temp directories
# to that same drive so large downloads (e.g. PyTorch ~2GB) don't fill up C:\
$scriptDrive = Split-Path -Qualifier $scriptDir
$systemDrive = $env:SystemDrive
if ($scriptDrive -ne $systemDrive) {
    $pipCacheDir = Join-Path $scriptDir ".pip-cache"
    $pipTmpDir   = Join-Path $scriptDir ".pip-tmp"
    if (-not (Test-Path $pipCacheDir)) { New-Item -ItemType Directory -Path $pipCacheDir -Force | Out-Null }
    if (-not (Test-Path $pipTmpDir))   { New-Item -ItemType Directory -Path $pipTmpDir   -Force | Out-Null }
    $env:PIP_CACHE_DIR = $pipCacheDir
    $env:TMPDIR        = $pipTmpDir
    $env:TEMP          = $pipTmpDir
    $env:TMP           = $pipTmpDir
    Write-Host "Running on $scriptDrive - pip cache and temp redirected to script directory." -ForegroundColor Cyan
}

Write-Host "=== Subtitle Tool Installation ==="
Write-Host ""

# Ensure we have a package manager
Ensure-PackageManager

Write-Host "Checking Python installation" -NoNewline -ForegroundColor White
$pythonCmd = Find-PythonCommand
if (-not $pythonCmd) {
    Write-Host "" # New line
    Write-Host "Python not found. Installing..." -ForegroundColor Yellow
    Install-Python -Method $PythonInstallMethod
    $pythonCmd = Find-PythonCommand
}

if (-not $pythonCmd) {
    Write-StatusFail "Python installation"
    throw "Python is still unavailable after installation attempts."
}

Write-Host "..." -NoNewline -ForegroundColor White
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White

# Refresh PATH in current session in case installer changed machine/user PATH.
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("Path", "User")

# Check if virtual environment exists
if (Test-Path $venvPath) {
    Write-StatusOK "Virtual environment exists"
} else {
    Write-Host "Creating virtual environment" -NoNewline -ForegroundColor White
    Invoke-Expression "$pythonCmd -m venv `"$venvPath`"" 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-StatusFail "Virtual environment creation"
        throw "Failed to create virtual environment."
    }
    Write-Host "..." -NoNewline -ForegroundColor White
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "OK" -NoNewline -ForegroundColor Green
    Write-Host " ]" -ForegroundColor White
}

# Activate virtual environment and use its Python
$venvPythonCmd = Join-Path $venvPath "Scripts\python.exe"
if (-not (Test-Path $venvPythonCmd)) {
    throw "Virtual environment Python not found at $venvPythonCmd"
}

Write-Host "Upgrading pip/setuptools/wheel..." -ForegroundColor White
& $venvPythonCmd -m pip install --upgrade pip setuptools wheel 2>&1 | ForEach-Object { 
    if ($_ -match "Successfully installed|Requirement already satisfied|Collecting") {
        Write-Host "  $_" -ForegroundColor Gray
    }
}
if ($LASTEXITCODE -ne 0) {
    Write-StatusFail "pip upgrade"
    throw "Failed to upgrade pip"
}
Write-StatusOK "pip/setuptools/wheel upgraded"

if (-not (Test-Path $requirementsPath)) {
    throw "requirements.txt not found at $requirementsPath"
}

Write-Host "Installing Python dependencies..." -ForegroundColor White
& $venvPythonCmd -m pip install -r "$requirementsPath" 2>&1 | ForEach-Object { 
    if ($_ -match "Successfully installed|Requirement already satisfied|Collecting") {
        Write-Host "  $_" -ForegroundColor Gray
    }
}
if ($LASTEXITCODE -ne 0) {
    Write-StatusFail "Python dependencies"
    throw "Failed to install Python dependencies"
}
Write-StatusOK "Core dependencies installed"

$installAI = $null
$aiSettingOverride = $null
$showSkipAiMessage = $true

# Check if AI libraries are already installed AND working
Write-Host ""
Write-Host "Checking for AI libraries..." -NoNewline -ForegroundColor White
$previousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "SilentlyContinue"

# First try to import torch to verify it works (torch is required for whisper)
$torchCheckResult = & $venvPythonCmd -c "import torch; print('ok')" 2>$null
$torchWorks = ($torchCheckResult -match "ok")

# Only mark AI as installed if torch actually works
if ($torchWorks) {
    $whisperCheckResult = & $venvPythonCmd -c "import whisper; print('installed')" 2>$null
    $aiAlreadyInstalled = ($whisperCheckResult -match "installed")
} else {
    $aiAlreadyInstalled = $false
}
$ErrorActionPreference = $previousErrorActionPreference

if ($aiAlreadyInstalled) {
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "ALREADY INSTALLED" -NoNewline -ForegroundColor Green
    Write-Host " ]" -ForegroundColor White
    $showSkipAiMessage = $false
} else {
    Write-Host " NOT FOUND" -ForegroundColor Yellow
    
    # Check and install VC++ if needed (before prompting about AI)
    Write-Host "Checking Visual C++ Redistributable..." -NoNewline -ForegroundColor White
    if (-not (Test-VCRedist)) {
        Write-Host " NOT FOUND" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "PyTorch (required for AI features) needs Visual C++ Redistributable." -ForegroundColor Yellow
        Write-Host "Installing VC++ Redistributable..." -ForegroundColor White
        $vcInstalled = Install-VCRedist
        if (-not $vcInstalled) {
            Write-Host ""
            Write-Host "WARNING: VC++ Redistributable installation failed." -ForegroundColor Red
            Write-Host "AI features will not be available. Install manually from:" -ForegroundColor Yellow
            Write-Host "https://aka.ms/vs/17/release/vc_redist.x64.exe" -ForegroundColor Cyan
            Write-Host ""
            $aiSettingOverride = $false
            $installAI = "N"  # Skip AI installation prompt since VC++ failed
        } else {
            Write-StatusOK "Visual C++ Redistributable installed"
        }
    } else {
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White
    }
    
    # Only prompt for AI installation if whisper not installed and VC++ is available
    if ($installAI -ne "N") {
        Write-Host ""
        Write-Host "=== AI Subtitle Generation (Optional) ===" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Whisper AI can automatically generate subtitles from video audio." -ForegroundColor White
        Write-Host "This feature is completely optional and requires significant disk space:" -ForegroundColor White
        Write-Host "  - Base installation: ~3-4GB (PyTorch + dependencies)" -ForegroundColor White
        Write-Host "  - Models download on first use: 72MB (tiny) to 2.9GB (large-v3)" -ForegroundColor White
        Write-Host "  - Total disk space needed: Up to ~10GB with all models" -ForegroundColor White
        Write-Host ""
        Write-Host "Install AI libraries? (openai-whisper, pysubs2, PyTorch)" -ForegroundColor Yellow
        Write-Host "  [Y] Yes (enables AI subtitle generation)" -ForegroundColor White
        Write-Host "  [N] No  (skip AI features, saves disk space)" -ForegroundColor White
        Write-Host ""
        $installAI = Read-Host "Your choice [Y/N]"
        if ($installAI -notmatch "^[Yy]") {
            $aiSettingOverride = $false
        }
    }
}

if ($installAI -match "^[Yy]") {
    if (-not (Test-Path $requirementsAIPath)) {
        Write-Host "Warning: requirements_ai.txt not found. Skipping AI installation." -ForegroundColor Yellow
        $aiSettingOverride = $false
    } else {
        Write-Host "Installing AI libraries (PyTorch, Whisper, pysubs2)..." -ForegroundColor White
        Write-Host "This will take several minutes. Package installation progress:" -ForegroundColor Yellow
        Write-Host ""
        
        & $venvPythonCmd -m pip install -r "$requirementsAIPath" 2>&1 | ForEach-Object { 
            if ($_ -match "Successfully installed|Requirement already satisfied|Collecting|Downloading") {
                Write-Host "  $_" -ForegroundColor Gray
            }
        }
        
        if ($LASTEXITCODE -ne 0) {
            Write-StatusFail "AI libraries installation"
            Write-Host ""
            Write-Host "Warning: AI libraries failed to install. The tool will work without AI features." -ForegroundColor Yellow
            Write-Host "Common issues:" -ForegroundColor Yellow
            Write-Host "  - Missing Visual C++ Redistributable" -ForegroundColor Yellow
            Write-Host "  - Insufficient disk space (~10GB needed)" -ForegroundColor Yellow
            Write-Host "  - Network/download issues" -ForegroundColor Yellow
            $aiSettingOverride = $false
        } else {
            Write-Host ""
            Write-StatusOK "AI libraries installed"
            Write-Host ""
            
            # Verify PyTorch installation
            Write-Host "Verifying PyTorch (AI backend)..." -NoNewline -ForegroundColor White
            $previousErrorActionPreference = $ErrorActionPreference
            $ErrorActionPreference = "SilentlyContinue"
            $torchTest = & $venvPythonCmd -c "import torch; print(torch.__version__)" 2>&1
            $ErrorActionPreference = $previousErrorActionPreference
            $torchExitCode = $LASTEXITCODE
            
            if ($torchExitCode -ne 0) {
                Write-StatusFail "PyTorch"
                Write-Host ""
                Write-Host "ERROR: PyTorch import failed." -ForegroundColor Red
                Write-Host "Error details:" -ForegroundColor Yellow
                Write-Host $torchTest -ForegroundColor Red
                Write-Host ""
                
                # Check if it's a DLL error
                if ($torchTest -match "c10\.dll|torch\.dll|DLL initialization") {
                    Write-Host "This is a Visual C++ Redistributable issue." -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "Attempting to fix by reinstalling VC++ Redistributable..." -ForegroundColor Yellow
                    $vcFixed = Install-VCRedist
                    
                    if ($vcFixed) {
                        Write-Host ""
                        Write-Host "VC++ Redistributable reinstalled. Testing PyTorch again..." -NoNewline -ForegroundColor Yellow
                        Start-Sleep -Seconds 2
                        $torchRetry = & $venvPythonCmd -c "import torch; print(torch.__version__)" 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            Write-Host " SUCCESS!" -ForegroundColor Green
                            Write-StatusOK "PyTorch verified ($torchRetry)"
                        } else {
                            Write-Host " FAILED" -ForegroundColor Red
                            Write-Host "You may need to restart your computer for VC++ changes to take effect." -ForegroundColor Yellow
                            $aiSettingOverride = $false
                        }
                    } else {
                        Write-Host "Failed to install VC++ Redistributable." -ForegroundColor Red
                        Write-Host "You may need to restart your computer or install manually." -ForegroundColor Yellow
                        $aiSettingOverride = $false
                    }
                } else {
                    Write-Host "Common fixes:" -ForegroundColor Yellow
                    Write-Host "  1. Delete the venv folder and re-run this installer" -ForegroundColor White
                    Write-Host "  2. Ensure you have ~10GB of free disk space" -ForegroundColor White
                    Write-Host "  3. Check your internet connection" -ForegroundColor White
                    $aiSettingOverride = $false
                }
            } else {
                Write-Host " [ " -NoNewline -ForegroundColor White
                Write-Host "OK" -NoNewline -ForegroundColor Green
                Write-Host " ] v$torchTest" -ForegroundColor White
                
                # Verify Whisper installation
                Write-Host "Verifying Whisper AI..." -NoNewline -ForegroundColor White
                $whisperTest = & $venvPythonCmd -c "import whisper; print(whisper.__version__ if hasattr(whisper, '__version__') else 'installed')" 2>&1
                if ($LASTEXITCODE -ne 0) {
                    Write-StatusFail "Whisper"
                    Write-Host ""
                    Write-Host "Warning: Whisper import failed. AI features will be disabled." -ForegroundColor Yellow
                    Write-Host "Error: $whisperTest" -ForegroundColor Red
                    $aiSettingOverride = $false
                } else {
                    Write-Host " [ " -NoNewline -ForegroundColor White
                    Write-Host "OK" -NoNewline -ForegroundColor Green
                    Write-Host " ] $whisperTest" -ForegroundColor White
                    
                    # Verify pysubs2
                    Write-Host "Verifying pysubs2..." -NoNewline -ForegroundColor White
                    $pysubs2Test = & $venvPythonCmd -c "import pysubs2; print(pysubs2.VERSION if hasattr(pysubs2, 'VERSION') else 'installed')" 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-StatusFail "pysubs2"
                        Write-Host ""
                        Write-Host "Warning: pysubs2 import failed." -ForegroundColor Yellow
                    } else {
                        Write-Host " [ " -NoNewline -ForegroundColor White
                        Write-Host "OK" -NoNewline -ForegroundColor Green
                        Write-Host " ] $pysubs2Test" -ForegroundColor White
                    }

                    # AI install and verification succeeded.
                    $aiSettingOverride = $true
                }
            }
        }
    }
} else {
    if ($showSkipAiMessage) {
        Write-Host ""
        Write-Host "Skipping AI libraries installation." -ForegroundColor Yellow
        Write-Host "You can install them later with:" -ForegroundColor White
        Write-Host "  pip install -r requirements_ai.txt" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Note: AI installation requires:" -ForegroundColor White
        Write-Host "  - Visual C++ Redistributable 2015-2022" -ForegroundColor White
        Write-Host "  - ~10GB disk space (including models)" -ForegroundColor White
        Write-Host "  - Stable internet connection" -ForegroundColor White
    }
}

Write-Host ""

if (-not (Test-Path $ffmpegInstallerPath)) {
    throw "ffmpeg installer script not found at $ffmpegInstallerPath"
}

Write-Host "Installing ffmpeg/ffprobe" -NoNewline -ForegroundColor White
powershell -ExecutionPolicy Bypass -File "$ffmpegInstallerPath" -Method $FfmpegInstallMethod 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-StatusFail "ffmpeg installation"
    throw "ffmpeg installation failed."
}
Write-Host "..." -NoNewline -ForegroundColor White
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White

Write-Host ""
Write-Host "=== Package Verification ==="

# Test PyQt6
Write-Host "PyQt6" -NoNewline -ForegroundColor White
& $venvPythonCmd -c "from PyQt6.QtCore import PYQT_VERSION_STR" 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-StatusFail "PyQt6"
    throw "PyQt6 verification failed"
}
Write-Host "..." -NoNewline -ForegroundColor White
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White

# Test fastapi
Write-Host "fastapi" -NoNewline -ForegroundColor White
& $venvPythonCmd -c "import fastapi" 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-StatusFail "fastapi"
    throw "fastapi verification failed"
}
Write-Host "..." -NoNewline -ForegroundColor White
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White

# Test uvicorn
Write-Host "uvicorn" -NoNewline -ForegroundColor White
& $venvPythonCmd -c "import uvicorn" 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-StatusFail "uvicorn"
    throw "uvicorn verification failed"
}
Write-Host "..." -NoNewline -ForegroundColor White
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White

Write-Host ""
Write-Host "=== Installation Complete ==="
Write-StatusOK "All components installed and verified"
Write-Host ""
Write-Host "Launching Subtitle Tool GUI..." -ForegroundColor Cyan
Write-Host ""

# Launch the GUI using the virtual environment Python
$guiArgs = @($subtitleToolPath, "gui")
if ($null -ne $aiSettingOverride) {
    if ($aiSettingOverride) {
        $guiArgs += "--use-ai"
        Write-Host "Launching GUI with --use-ai (app will persist this setting)." -ForegroundColor Gray
    } else {
        $guiArgs += "--no-ai"
        Write-Host "Launching GUI with --no-ai (app will persist this setting)." -ForegroundColor Gray
    }
}

& $venvPythonCmd @guiArgs
if ($LASTEXITCODE -ne 0) {
    throw "Subtitle Tool failed to launch correctly (exit code: $LASTEXITCODE)."
}

exit 0
