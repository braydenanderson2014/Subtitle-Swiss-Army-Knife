param(
    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$PythonInstallMethod = "auto",

    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$FfmpegInstallMethod = "auto",

    [switch]$NoPause,

    [switch]$KeepInstallArtifacts,

    # Uninstall ALL dependencies: deletes venv, pip cache dirs, and reports what
    # was removed. Does NOT uninstall system-wide Python, ffmpeg, or VC++.
    [switch]$Uninstall,

    # Uninstall only the AI/IMDB extras: torch, openai-whisper, pysubs2,
    # cinemagoer, plus the downloaded Whisper model cache on disk.
    [switch]$UninstallAI
)

$ErrorActionPreference = "Stop"
$ProgressPreference = 'SilentlyContinue'
$script:CleanupTargets = @()

function Cleanup-InstallerArtifacts {
    if ($KeepInstallArtifacts) {
        return
    }

    foreach ($path in $script:CleanupTargets) {
        try {
            if ($path -and (Test-Path $path)) {
                Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch {
            # Best-effort cleanup only.
        }
    }
}

function Remove-PathRobust {
    param([string]$Path)

    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        return $true
    }

    # First pass: clear common attributes and try PowerShell deletion a few times.
    try {
        attrib -r -s -h "$Path" /s /d 2>$null | Out-Null
    } catch {
        # Best effort only.
    }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
        } catch {
            # Keep trying with fallback paths below.
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            return $true
        }

        Start-Sleep -Milliseconds 250
    }

    # Fallback: use cmd rmdir/del with long-path prefix support.
    try {
        if (Test-Path -LiteralPath $Path) {
            $fullPath = [System.IO.Path]::GetFullPath($Path)
            $longPath = if ($fullPath.StartsWith("\\?\")) { $fullPath } else { "\\?\$fullPath" }
            $isDir = (Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue).PSIsContainer
            if ($isDir) {
                cmd /c "rmdir /s /q \"$longPath\"" 2>$null | Out-Null
            } else {
                cmd /c "del /f /q \"$longPath\"" 2>$null | Out-Null
            }
        }
    } catch {
        # Best effort only.
    }

    return (-not (Test-Path -LiteralPath $Path))
}

# Catch any terminating error, show it, and pause so the window stays open
trap {
    Write-Host ""
    Write-Host "=== INSTALLATION FAILED ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Cleanup-InstallerArtifacts
    Write-Host "The window will stay open so you can read the error above." -ForegroundColor Yellow
    if (-not $NoPause) {
        Read-Host "Press Enter to close"
    }
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

function Test-IsAdministrator {
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Test-PythonCommandVersionSupported {
    param([string]$Command)

    try {
        $version = Invoke-Expression "$Command -c \"import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')\"" 2>$null
        if (-not $version) { return $false }
        return ($version -match '^3\.(10|11|12)$')
    } catch {
        return $false
    }
}

function Get-PythonCommandVersion {
    param([string]$Command)

    try {
        $version = Invoke-Expression "$Command -c \"import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')\"" 2>$null
        return "$version".Trim()
    } catch {
        return ""
    }
}

function Get-PythonCommandExecutablePath {
    param([string]$Command)

    try {
        $path = Invoke-Expression "$Command -c \"import sys; print(sys.executable)\"" 2>$null
        return "$path".Trim()
    } catch {
        return ""
    }
}

function Test-PythonCommandSuitable {
    param([string]$Command)

    if (-not (Test-PythonCommandVersionSupported $Command)) {
        return $false
    }

    $exePath = Get-PythonCommandExecutablePath $Command
    if (-not $exePath) {
        return $false
    }

    # Reject embedded/bundled Python distributions like PlatformIO's private runtime.
    if ($exePath -match '(?i)platformio') {
        return $false
    }

    return $true
}

function Find-Python312Path {
    $candidates = @(
        (Join-Path $env:LocalAppData "Programs\Python\Python312\python.exe"),
        (Join-Path $env:ProgramFiles "Python312\python.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Python312\python.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Find-Python311Path {
    $candidates = @(
        (Join-Path $env:LocalAppData "Programs\Python\Python311\python.exe"),
        (Join-Path $env:ProgramFiles "Python311\python.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Python311\python.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Find-PyLauncherPath {
    $candidates = @(
        (Join-Path $env:LocalAppData "Programs\Python\Launcher\py.exe"),
        (Join-Path $env:WINDIR "py.exe")
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Find-PythonFromPyList {
    try {
        $pyListOutput = & py -0p 2>$null
        if (-not $pyListOutput) {
            return $null
        }

        foreach ($line in $pyListOutput) {
            $trimmed = "$line".Trim()
            if ($trimmed -match '^\-V:3\.(11|10|12).*\s+(.+python\.exe)$') {
                $candidatePath = $matches[2].Trim()
                if ((Test-Path $candidatePath) -and ($candidatePath -notmatch '(?i)platformio')) {
                    return "& '$candidatePath'"
                }
            }
        }
    } catch {
        return $null
    }

    return $null
}

function Find-PythonCommand {
    # Prefer versions known to work well with PyTorch on Windows.
    $pyLauncherPath = Find-PyLauncherPath
    if ($pyLauncherPath) {
        if (Test-PythonCommandSuitable "& '$pyLauncherPath' -3.11") { return "& '$pyLauncherPath' -3.11" }
        if (Test-PythonCommandSuitable "& '$pyLauncherPath' -3.10") { return "& '$pyLauncherPath' -3.10" }
        if (Test-PythonCommandSuitable "& '$pyLauncherPath' -3.12") { return "& '$pyLauncherPath' -3.12" }
        if (Test-PythonCommandSuitable "& '$pyLauncherPath' -3") { return "& '$pyLauncherPath' -3" }
    }

    if (Test-CommandAvailable "py") {
        if (Test-PythonCommandSuitable "py -3.11") { return "py -3.11" }
        if (Test-PythonCommandSuitable "py -3.10") { return "py -3.10" }
        if (Test-PythonCommandSuitable "py -3.12") { return "py -3.12" }
        if (Test-PythonCommandSuitable "py -3") { return "py -3" }
    }

    $fromPyList = Find-PythonFromPyList
    if ($fromPyList) {
        return $fromPyList
    }

    $python311Path = Find-Python311Path
    if ($python311Path -and (Test-PythonCommandSuitable "& '$python311Path'")) {
        return "& '$python311Path'"
    }

    $python312Path = Find-Python312Path
    if ($python312Path -and (Test-PythonCommandSuitable "& '$python312Path'")) {
        return "& '$python312Path'"
    }

    if (Test-CommandAvailable "python") {
        if (Test-PythonCommandSuitable "python") {
            return "python"
        }
    }
    return $null
}

function Install-PythonWithWinget {
    Write-Host "Installing Python 3.11 using winget..."
    winget install --id Python.Python.3.11 --exact --scope user --silent --accept-source-agreements --accept-package-agreements
}

function Install-PythonWithChoco {
    Write-Host "Installing Python 3.11 using Chocolatey..."
    choco install python311 --yes
}

function Install-PythonWithScoop {
    throw "Scoop automatic Python install is not configured for the required 3.12 version. Use winget or Chocolatey."
}

function Install-PythonManually {
    Write-Host "Installing Python 3.11 with per-user installer..." -ForegroundColor Yellow

    $pythonInstallerUrls = @(
        "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe",
        "https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe"
    )
    $pythonInstallerPath = Join-Path $env:TEMP "python-3.11-amd64.exe"

    try {
        if (Test-Path $pythonInstallerPath) {
            Remove-Item $pythonInstallerPath -Force -ErrorAction SilentlyContinue
        }

        $downloaded = $false
        foreach ($url in $pythonInstallerUrls) {
            try {
                Invoke-WebRequest -Uri $url -OutFile $pythonInstallerPath -UseBasicParsing
                $downloaded = $true
                break
            } catch {
                Write-Host "Manual Python installer URL failed: $url" -ForegroundColor DarkYellow
            }
        }

        if (-not $downloaded) {
            throw "Could not download a Python 3.11 installer from known URLs."
        }

        $process = Start-Process -FilePath $pythonInstallerPath -ArgumentList @(
            "/quiet",
            "InstallAllUsers=0",
            "PrependPath=1",
            "Include_launcher=1",
            "InstallLauncherAllUsers=0",
            "Include_test=0",
            "Shortcuts=0"
        ) -Wait -PassThru

        Remove-Item $pythonInstallerPath -Force -ErrorAction SilentlyContinue

        if ($process.ExitCode -ne 0) {
            throw "Python installer exited with code $($process.ExitCode)."
        }

        $userPythonDir = Join-Path $env:LocalAppData "Programs\Python\Python311"
        $userPythonScriptsDir = Join-Path $userPythonDir "Scripts"
        $env:Path = "$userPythonDir;$userPythonScriptsDir;" +
                    [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("Path", "User")
    } catch {
        throw "Failed to install Python 3.11 manually: $($_.Exception.Message)"
    }
}

function Wait-ForPythonCommand {
    param([int]$Seconds = 25)

    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        $candidate = Find-PythonCommand
        if ($candidate) {
            return $candidate
        }

        # Refresh PATH between retries because installers often update User PATH.
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("Path", "User")
        Start-Sleep -Milliseconds 1200
    }

    return $null
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

                $resolvedAfterWinget = Wait-ForPythonCommand -Seconds 30
                if ($resolvedAfterWinget) {
                    return
                }

                Write-Host "winget reported success but Python was not discoverable yet." -ForegroundColor Yellow
            } catch {
                Write-Host "winget Python install failed: $($_.Exception.Message)"
            }
        }

        if (-not (Find-PythonCommand)) {
            try {
                Install-PythonManually
                $attempted += "manual"

                $resolvedAfterManual = Wait-ForPythonCommand -Seconds 30
                if ($resolvedAfterManual) {
                    return
                }

                Write-Host "manual Python install completed but Python was not discoverable yet." -ForegroundColor Yellow
            } catch {
                Write-Host "manual Python install failed: $($_.Exception.Message)"
            }
        }

        if (-not (Find-PythonCommand) -and (Test-CommandAvailable "choco") -and (Test-IsAdministrator)) {
            try {
                Install-PythonWithChoco
                $attempted += "choco"
            } catch {
                Write-Host "choco Python install failed: $($_.Exception.Message)"
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

function Repair-TorchInVenv {
    param([string]$PythonExe)

    Write-Host "Attempting PyTorch repair (CPU build) in current venv..." -ForegroundColor Yellow

    # Remove potentially broken Torch packages first.
    & $PythonExe -m pip uninstall -y torch torchvision torchaudio 2>&1 | Out-Null

    # Clear pip cache to avoid reusing a corrupted wheel.
    & $PythonExe -m pip cache purge 2>&1 | Out-Null

    # Install a fresh CPU wheel directly from the official PyTorch CPU index.
    & $PythonExe -m pip install --no-cache-dir --force-reinstall --index-url https://download.pytorch.org/whl/cpu torch 2>&1 | ForEach-Object {
        if ($_ -match "Successfully installed|Requirement already satisfied|Collecting|Downloading") {
            Write-Host "  $_" -ForegroundColor Gray
        }
    }

    return ($LASTEXITCODE -eq 0)
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$requirementsPath = Join-Path $scriptDir "requirements.txt"
$requirementsAIPath = Join-Path $scriptDir "requirements_ai.txt"
$ffmpegInstallerPath = Join-Path $scriptDir "install_ffmpeg_windows.ps1"
$subtitleToolPath = Join-Path $scriptDir "subtitle_tool.py"
$venvPath = Join-Path $scriptDir "venv"
$manifestPath = Join-Path $scriptDir ".install_manifest.json"

Set-Location $scriptDir

# ---------------------------------------------------------------------------
# Manifest helpers - track what THIS script installed vs. what pre-existed.
# The manifest is written at the end of a successful install and read during
# uninstall so that we only remove components that we actually put there.
# ---------------------------------------------------------------------------
function Read-InstallManifest {
    if (-not (Test-Path $manifestPath)) { return @{} }
    try {
        $raw = Get-Content $manifestPath -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json
        # Convert PSCustomObject to hashtable so callers can use .ContainsKey etc.
        $ht = @{}
        $obj.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
        return $ht
    } catch {
        return @{}
    }
}

function Save-InstallManifest {
    param([hashtable]$Manifest)
    try {
        $Manifest | ConvertTo-Json -Depth 5 | Set-Content $manifestPath -Encoding UTF8
    } catch {
        Write-Host "Warning: could not save install manifest: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }
}

# ===========================================================================
# UNINSTALL AI  (-UninstallAI)
# Removes AI / IMDB-lookup packages from the venv and deletes Whisper models.
# Leaves the core venv (PyQt6, fastapi, etc.) intact.
# ===========================================================================
if ($UninstallAI) {
    Write-Host ""
    Write-Host "=== Uninstall AI Libraries ===" -ForegroundColor Cyan
    Write-Host ""

    $venvPy = Join-Path $venvPath "Scripts\python.exe"

    if (-not (Test-Path $venvPy)) {
        Write-Host "No virtual environment found at $venvPath. Nothing to uninstall." -ForegroundColor Yellow
    } else {
        $aiPackages = @("torch", "torchvision", "torchaudio", "openai-whisper", "whisper", "pysubs2", "cinemagoer", "imdbpy")

        Write-Host "Uninstalling AI packages from venv..." -ForegroundColor White
        foreach ($pkg in $aiPackages) {
            Write-Host "  Removing $pkg..." -NoNewline -ForegroundColor Gray
            $result = & $venvPy -m pip uninstall -y $pkg 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host " done" -ForegroundColor Green
            } else {
                # pip exits non-zero when the package isn't installed - that's fine.
                Write-Host " (not installed, skipping)" -ForegroundColor DarkGray
            }
        }

        # Also remove stale torch-related packages that may linger after uninstall
        Write-Host "  Cleaning up torch dependencies..." -NoNewline -ForegroundColor Gray
        & $venvPy -m pip uninstall -y filelock sympy networkx jinja2-markupsafe 2>&1 | Out-Null
        Write-Host " done" -ForegroundColor Green

        Write-StatusOK "AI packages removed from venv"
    }

    # Delete downloaded Whisper model cache
    # Whisper stores models in %USERPROFILE%\.cache\whisper on Windows
    $whisperModelDirs = @(
        (Join-Path $env:USERPROFILE ".cache\whisper"),
        (Join-Path $env:LOCALAPPDATA "whisper"),
        (Join-Path $env:APPDATA "whisper")
    )
    Write-Host ""
    Write-Host "Searching for Whisper model cache..." -ForegroundColor White
    $foundAny = $false
    foreach ($dir in $whisperModelDirs) {
        if (Test-Path $dir) {
            $foundAny = $true
            $size = (Get-ChildItem $dir -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 1)
            Write-Host "  Found: $dir  ($($sizeMB)MB)" -ForegroundColor Yellow
            Write-Host "  Deleting..." -NoNewline -ForegroundColor Gray
            Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
            if (-not (Test-Path $dir)) {
                Write-Host " deleted" -ForegroundColor Green
            } else {
                Write-Host " some files could not be deleted (may be in use)" -ForegroundColor Yellow
            }
        }
    }
    if (-not $foundAny) {
        Write-Host "  No Whisper model cache found." -ForegroundColor Gray
    }

    # Offer to remove VC++ if the manifest shows this script installed it for AI
    $mAI = Read-InstallManifest
    if ($mAI.Count -gt 0) {
        $mVC = $mAI["vcredist"]
        if ($null -ne $mVC) {
            $vcWasOurs = -not [bool]($mVC.PSObject.Properties["pre_existed"] | Select-Object -ExpandProperty Value)
            $vcMeth    = [string]($mVC.PSObject.Properties["install_method"] | Select-Object -ExpandProperty Value)
            if ($vcWasOurs -and $vcMeth) {
                Write-Host ""
                Write-Host "The install manifest shows that Visual C++ Redistributable was installed" -ForegroundColor Yellow
                Write-Host "by this script (required for PyTorch). Since you are removing AI features," -ForegroundColor Yellow
                Write-Host "you may want to remove it too (install method: $vcMeth)." -ForegroundColor Yellow
                if (-not $NoPause) {
                    $vcAnswer = Read-Host "Remove Visual C++ Redistributable as well? [Y/N]"
                } else {
                    $vcAnswer = "N"
                }
                if ($vcAnswer -match "^[Yy]") {
                    Write-Host "Uninstalling Visual C++ Redistributable via $vcMeth..." -ForegroundColor White
                    switch ($vcMeth) {
                        "winget" {
                            winget uninstall --id Microsoft.VCRedist.2015+.x64 --silent 2>&1 | Out-Null
                            Write-Host "  VC++ Redistributable uninstalled." -ForegroundColor Green
                        }
                        "choco" {
                            choco uninstall vcredist-all -y 2>&1 | Out-Null
                            Write-Host "  VC++ Redistributable uninstalled via Chocolatey." -ForegroundColor Green
                        }
                        default {
                            Write-Host "  Please remove VC++ manually via Apps & Features." -ForegroundColor Yellow
                        }
                    }
                }
            }
        }
    }

    Write-Host ""
    Write-Host "=== AI Uninstall Complete ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "The core venv (PyQt6, fastapi, etc.) remains intact." -ForegroundColor White
    Write-Host "Re-run the installer and choose 'Y' at the AI prompt to reinstall AI features." -ForegroundColor White
    Write-Host ""
    if (-not $NoPause) { Read-Host "Press Enter to close" }
    exit 0
}

# ===========================================================================
# FULL UNINSTALL  (-Uninstall)
# Reads .install_manifest.json to know which system components were installed
# by this script, then removes only those along with the venv and caches.
# ===========================================================================
if ($Uninstall) {
    Write-Host ""
    Write-Host "=== Full Uninstall ===" -ForegroundColor Cyan
    Write-Host ""

    # Read manifest - tells us what this script originally installed
    $m = Read-InstallManifest
    $hasMfst = $m.Count -gt 0

    # Helper: safely read nested PSCustomObject or hashtable property
    function Get-MfstProp {
        param($obj, [string]$key, $default = $null)
        if ($null -eq $obj) { return $default }
        if ($obj -is [hashtable]) {
            if ($obj.ContainsKey($key)) { return $obj[$key] }
            return $default
        }
        $v = $obj.PSObject.Properties[$key]
        if ($null -eq $v) { return $default }
        return $v.Value
    }

    $mPython   = if ($hasMfst) { Get-MfstProp $m "python"   } else { $null }
    $mVCRedist = if ($hasMfst) { Get-MfstProp $m "vcredist" } else { $null }
    $mFfmpeg   = if ($hasMfst) { Get-MfstProp $m "ffmpeg"   } else { $null }

    $pyPreExisted  = if ($mPython)   { [bool](Get-MfstProp $mPython   "pre_existed" $true) } else { $true }
    $pyMethod      = if ($mPython)   { [string](Get-MfstProp $mPython   "install_method" "") } else { "" }
    $vcPreExisted  = if ($mVCRedist) { [bool](Get-MfstProp $mVCRedist "pre_existed" $true) } else { $true }
    $vcMethod      = if ($mVCRedist) { [string](Get-MfstProp $mVCRedist "install_method" "") } else { "" }
    $ffPreExisted  = if ($mFfmpeg)   { [bool](Get-MfstProp $mFfmpeg   "pre_existed" $true) } else { $true }
    $ffMethod      = if ($mFfmpeg)   { [string](Get-MfstProp $mFfmpeg   "install_method" "") } else { "" }

    # Show summary of what will happen
    Write-Host "The following will always be removed:" -ForegroundColor White
    Write-Host "  - Virtual environment:  $venvPath" -ForegroundColor White
    Write-Host "  - pip cache dirs (.pip-cache, .pip-tmp)" -ForegroundColor White
    Write-Host "  - __pycache__ folders in the script directory" -ForegroundColor White
    Write-Host "  - Downloaded Whisper AI model cache" -ForegroundColor White
        Write-Host "  - Application settings (.subtitle_tool_settings.json)" -ForegroundColor White
    Write-Host ""

    if (-not $hasMfst) {
        Write-Host "No install manifest found (.install_manifest.json)." -ForegroundColor Yellow
        Write-Host "System packages (Python, ffmpeg, VC++) will NOT be touched." -ForegroundColor Yellow
        Write-Host "If you want to remove them, delete them manually via Apps & Features." -ForegroundColor Yellow
    } else {
        Write-Host "System packages (based on install manifest):" -ForegroundColor White
        if (-not $pyPreExisted -and $pyMethod) {
            Write-Host "  - Python 3.11    will be uninstalled via $pyMethod" -ForegroundColor Cyan
        } else {
            Write-Host "  - Python         was already installed before this script - will NOT be touched" -ForegroundColor Gray
        }
        if (-not $ffPreExisted -and $ffMethod) {
            Write-Host "  - ffmpeg         will be uninstalled via $ffMethod" -ForegroundColor Cyan
        } else {
            Write-Host "  - ffmpeg         was already installed before this script - will NOT be touched" -ForegroundColor Gray
        }
        if (-not $vcPreExisted -and $vcMethod) {
            Write-Host "  - VC++ Redist    will be uninstalled via $vcMethod" -ForegroundColor Cyan
        } else {
            Write-Host "  - VC++ Redist    was already installed before this script - will NOT be touched" -ForegroundColor Gray
        }
    }

    Write-Host ""
    if (-not $NoPause) {
        $confirm = Read-Host "Continue with uninstall? [Y/N]"
        if ($confirm -notmatch "^[Yy]") {
            Write-Host "Uninstall cancelled." -ForegroundColor Yellow
            exit 0
        }
    }

    Write-Host ""

    # 1. Delete venv
    if (Test-Path $venvPath) {
        Write-Host "Deleting virtual environment..." -NoNewline -ForegroundColor White
        $venvRemoved = Remove-PathRobust -Path $venvPath
        if ($venvRemoved) {
            Write-Host " [ " -NoNewline -ForegroundColor White
            Write-Host "OK" -NoNewline -ForegroundColor Green
            Write-Host " ]" -ForegroundColor White
        } else {
            Write-Host " [ some files in use, partial delete ]" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Virtual environment not found - skipping." -ForegroundColor Gray
    }

    # 2. Delete installer-created cache dirs
    foreach ($dir in @((Join-Path $scriptDir ".pip-cache"), (Join-Path $scriptDir ".pip-tmp"))) {
        if (Test-Path $dir) {
            Write-Host "Deleting $dir..." -NoNewline -ForegroundColor White
            if (Remove-PathRobust -Path $dir) {
                Write-Host " [ " -NoNewline -ForegroundColor White
                Write-Host "OK" -NoNewline -ForegroundColor Green
                Write-Host " ]" -ForegroundColor White
            } else {
                Write-Host " [ some files in use, partial delete ]" -ForegroundColor Yellow
            }
        }
    }

    # 3. Delete __pycache__ under script dir
    Write-Host "Cleaning __pycache__..." -NoNewline -ForegroundColor White
    $pycacheDirs = Get-ChildItem -Path $scriptDir -Filter "__pycache__" -Recurse -Directory -ErrorAction SilentlyContinue
    $pycacheFailures = 0
    foreach ($pc in $pycacheDirs) {
        if (-not (Remove-PathRobust -Path $pc.FullName)) {
            $pycacheFailures++
        }
    }
    if ($pycacheFailures -eq 0) {
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White
    } else {
        Write-Host " [ $pycacheFailures folder(s) could not be fully removed ]" -ForegroundColor Yellow
    }

    # 3b. Delete application settings file
    $settingsFilePath = Join-Path $scriptDir ".subtitle_tool_settings.json"
    if (Test-Path $settingsFilePath) {
        Write-Host "Deleting settings file (.subtitle_tool_settings.json)..." -NoNewline -ForegroundColor White
        if (Remove-PathRobust -Path $settingsFilePath) {
            Write-Host " [ " -NoNewline -ForegroundColor White
            Write-Host "OK" -NoNewline -ForegroundColor Green
            Write-Host " ]" -ForegroundColor White
        } else {
            Write-Host " [ failed ]" -ForegroundColor Yellow
        }
    }

    # 4. Delete Whisper model cache

    foreach ($dir in @(
        (Join-Path $env:USERPROFILE ".cache\whisper"),
        (Join-Path $env:LOCALAPPDATA "whisper"),
        (Join-Path $env:APPDATA "whisper")
    )) {
        if (Test-Path $dir) {
            $sizeMB = [math]::Round(
                (Get-ChildItem $dir -Recurse -ErrorAction SilentlyContinue |
                 Measure-Object -Property Length -Sum).Sum / 1MB, 1)
            Write-Host "Deleting Whisper cache ($($sizeMB)MB) $dir..." -NoNewline -ForegroundColor White
            if (Remove-PathRobust -Path $dir) {
                Write-Host " [ " -NoNewline -ForegroundColor White
                Write-Host "OK" -NoNewline -ForegroundColor Green
                Write-Host " ]" -ForegroundColor White
            } else {
                Write-Host " [ some files in use, partial delete ]" -ForegroundColor Yellow
            }
        }
    }

    # 5. Conditionally remove system packages (only what the script installed)
    if ($hasMfst) {
        # --- Python ---
        if (-not $pyPreExisted -and $pyMethod) {
            Write-Host ""
            Write-Host "Uninstalling Python 3.11 via $pyMethod..." -ForegroundColor White
            switch ($pyMethod) {
                "winget" {
                    winget uninstall --id Python.Python.3.11 --silent --accept-source-agreements 2>&1 | Out-Null
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "  Python uninstalled." -ForegroundColor Green
                    } else {
                        # Also try 3.11.* catch-all
                        winget uninstall --name "Python 3.11" --silent 2>&1 | Out-Null
                        Write-Host "  Python uninstall attempted (check Apps & Features if it persists)." -ForegroundColor Yellow
                    }
                }
                "choco" {
                    choco uninstall python311 -y --remove-dependencies 2>&1 | Out-Null
                    Write-Host "  Python uninstalled via Chocolatey." -ForegroundColor Green
                }
                default {
                    Write-Host "  Install method '$pyMethod' - please remove Python manually via Apps & Features." -ForegroundColor Yellow
                }
            }
        }

        # --- ffmpeg ---
        if (-not $ffPreExisted -and $ffMethod) {
            Write-Host ""
            Write-Host "Uninstalling ffmpeg via $ffMethod..." -ForegroundColor White
            switch ($ffMethod) {
                "winget" {
                    winget uninstall --id Gyan.FFmpeg --silent 2>&1 | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        # Some bundles register under a slightly different id
                        winget uninstall --name "ffmpeg" --silent 2>&1 | Out-Null
                    }
                    Write-Host "  ffmpeg uninstalled." -ForegroundColor Green
                }
                "choco" {
                    choco uninstall ffmpeg -y 2>&1 | Out-Null
                    Write-Host "  ffmpeg uninstalled via Chocolatey." -ForegroundColor Green
                }
                "scoop" {
                    scoop uninstall ffmpeg 2>&1 | Out-Null
                    Write-Host "  ffmpeg uninstalled via Scoop." -ForegroundColor Green
                }
                default {
                    Write-Host "  Install method '$ffMethod' - please remove ffmpeg manually." -ForegroundColor Yellow
                }
            }
        }

        # --- Visual C++ Redistributable ---
        if (-not $vcPreExisted -and $vcMethod) {
            Write-Host ""
            Write-Host "Uninstalling Visual C++ Redistributable via $vcMethod..." -ForegroundColor White
            switch ($vcMethod) {
                "winget" {
                    winget uninstall --id Microsoft.VCRedist.2015+.x64 --silent 2>&1 | Out-Null
                    Write-Host "  VC++ Redistributable uninstalled." -ForegroundColor Green
                }
                "choco" {
                    choco uninstall vcredist-all -y 2>&1 | Out-Null
                    Write-Host "  VC++ Redistributable uninstalled via Chocolatey." -ForegroundColor Green
                }
                default {
                    Write-Host "  Install method '$vcMethod' - please remove VC++ manually via Apps & Features." -ForegroundColor Yellow
                }
            }
        }

        # 6. Delete the manifest itself (clean slate for future installs)
        if (Test-Path $manifestPath) {
            Remove-Item $manifestPath -Force -ErrorAction SilentlyContinue
            Write-Host ""
            Write-Host "Install manifest removed." -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "=== Uninstall Complete ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "To reinstall, run:  install_all_windows.bat" -ForegroundColor Cyan
    Write-Host ""
    if (-not $NoPause) { Read-Host "Press Enter to close" }
    exit 0
}

Write-Host "=== Installer Bootstrap ===" -ForegroundColor Cyan
Write-Host "Script path: $($MyInvocation.MyCommand.Path)"
Write-Host "Working directory: $scriptDir"
Write-Host "PowerShell version: $($PSVersionTable.PSVersion)"
Write-Host ""

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
    $script:CleanupTargets += $pipCacheDir
    $script:CleanupTargets += $pipTmpDir
    Write-Host "Running on $scriptDrive - pip cache and temp redirected to script directory." -ForegroundColor Cyan
}

Write-Host "=== Subtitle Tool Installation ==="
Write-Host ""

# Ensure we have a package manager
Ensure-PackageManager

# Record whether each system component pre-existed before this install run.
# This drives the uninstall logic - we only remove what we put there.
$manifest = Read-InstallManifest
# Initialise tracking vars (may be overwritten below)
$script:PythonPreExisted  = $true
$script:PythonInstallUsed = ""       # winget | choco | manual
$script:VCRedistPreExisted   = $true
$script:VCRedistInstallUsed  = ""
$script:FfmpegPreExisted  = $true
$script:FfmpegInstallUsed = ""

Write-Host "Checking Python installation" -NoNewline -ForegroundColor White
$pythonCmd = Find-PythonCommand
if (-not $pythonCmd) {
    $script:PythonPreExisted = $false
    Write-Host "" # New line
    $detectedVersion = ""
    if (Test-CommandAvailable "python") {
        $detectedVersion = Get-PythonCommandVersion "python"
    }
    if (-not $detectedVersion -and (Test-CommandAvailable "py")) {
        $detectedVersion = Get-PythonCommandVersion "py -3"
    }

    if ($detectedVersion) {
        Write-Host "Detected unsupported Python version $detectedVersion. Installing Python 3.11..." -ForegroundColor Yellow
    } else {
        Write-Host "Python not found. Installing Python 3.11..." -ForegroundColor Yellow
    }
    # Determine which method actually installs Python so we can uninstall later
    if ($PythonInstallMethod -eq "winget" -or ($PythonInstallMethod -eq "auto" -and (Test-CommandAvailable "winget"))) {
        $script:PythonInstallUsed = "winget"
    } elseif ($PythonInstallMethod -eq "choco" -or ($PythonInstallMethod -eq "auto" -and (Test-CommandAvailable "choco"))) {
        $script:PythonInstallUsed = "choco"
    } else {
        $script:PythonInstallUsed = "manual"
    }
    Install-Python -Method $PythonInstallMethod
    $pythonCmd = Find-PythonCommand
}

if (-not $pythonCmd) {
    Write-StatusFail "Python installation"
    if (Test-CommandAvailable "py") {
        Write-Host "Detected by py launcher:" -ForegroundColor Yellow
        & py -0p 2>$null | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    }
    $py311 = Find-Python311Path
    $py312 = Find-Python312Path
    Write-Host "Checked Python311 path: $py311" -ForegroundColor Yellow
    Write-Host "Checked Python312 path: $py312" -ForegroundColor Yellow
    throw "Python 3.10-3.12 is required for stable AI support. Install Python 3.11 or 3.12 and re-run installer."
}

Write-Host "..." -NoNewline -ForegroundColor White
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White

# Refresh PATH in current session in case installer changed machine/user PATH.
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("Path", "User")

# Check if virtual environment exists and uses a supported Python version.
$createVenv = $true
if (Test-Path $venvPath) {
    $existingVenvPython = Join-Path $venvPath "Scripts\python.exe"
    if (Test-Path $existingVenvPython) {
        $existingVenvVersion = & $existingVenvPython -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>$null
        if ($LASTEXITCODE -eq 0 -and $existingVenvVersion -match '^3\.(10|11|12)$') {
            Write-StatusOK "Virtual environment exists"
            $createVenv = $false
        } else {
            Write-Host "Existing venv uses unsupported Python version ($existingVenvVersion). Recreating..." -ForegroundColor Yellow
            Remove-Item -Path $venvPath -Recurse -Force
            $createVenv = $true
        }
    } else {
        Write-Host "Existing venv is invalid (missing python.exe). Recreating..." -ForegroundColor Yellow
        Remove-Item -Path $venvPath -Recurse -Force
        $createVenv = $true
    }
}

if ($createVenv) {
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

$venvVersion = & $venvPythonCmd -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
if ($LASTEXITCODE -ne 0 -or $venvVersion -notmatch '^3\.(10|11|12)$') {
    throw "Unsupported venv Python version ($venvVersion). Use Python 3.10-3.12 for AI features."
}

Write-Host "Using virtual environment Python: $venvPythonCmd" -ForegroundColor Cyan
Write-Host "Using Python version: $venvVersion" -ForegroundColor Cyan

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
        $script:VCRedistPreExisted = $false
        # Record which package manager will be used for VC++ so uninstall can reverse it
        if (Test-CommandAvailable "winget") {
            $script:VCRedistInstallUsed = "winget"
        } elseif (Test-CommandAvailable "choco") {
            $script:VCRedistInstallUsed = "choco"
        } else {
            $script:VCRedistInstallUsed = "manual"
        }
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
                            Write-Host "VC++ reinstall alone did not fix PyTorch. Trying fresh CPU PyTorch install..." -ForegroundColor Yellow

                            $torchRepair = Repair-TorchInVenv -PythonExe $venvPythonCmd
                            if ($torchRepair) {
                                $torchRetry2 = & $venvPythonCmd -c "import torch; print(torch.__version__)" 2>&1
                                if ($LASTEXITCODE -eq 0) {
                                    Write-StatusOK "PyTorch repaired and verified ($torchRetry2)"
                                } else {
                                    Write-Host "PyTorch still fails after repair. A reboot may be required for VC++ runtime changes." -ForegroundColor Yellow
                                    $aiSettingOverride = $false
                                }
                            } else {
                                Write-Host "PyTorch CPU reinstall failed." -ForegroundColor Red
                                Write-Host "You may need to restart your computer for VC++ changes to take effect." -ForegroundColor Yellow
                                $aiSettingOverride = $false
                            }
                        }
                    } else {
                        Write-Host "Failed to install VC++ Redistributable." -ForegroundColor Red
                        Write-Host "Trying fresh CPU PyTorch install anyway..." -ForegroundColor Yellow

                        $torchRepair = Repair-TorchInVenv -PythonExe $venvPythonCmd
                        if ($torchRepair) {
                            $torchRetry3 = & $venvPythonCmd -c "import torch; print(torch.__version__)" 2>&1
                            if ($LASTEXITCODE -eq 0) {
                                Write-StatusOK "PyTorch repaired and verified ($torchRetry3)"
                            } else {
                                Write-Host "PyTorch still fails after repair." -ForegroundColor Red
                                Write-Host "You may need to restart your computer or install VC++ manually." -ForegroundColor Yellow
                                $aiSettingOverride = $false
                            }
                        } else {
                            Write-Host "PyTorch CPU reinstall failed." -ForegroundColor Red
                            Write-Host "You may need to restart your computer or install manually." -ForegroundColor Yellow
                            $aiSettingOverride = $false
                        }
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

# Track whether ffmpeg pre-existed before this run
if (-not (Test-CommandAvailable "ffmpeg") -or -not (Test-CommandAvailable "ffprobe")) {
    $script:FfmpegPreExisted = $false
    if ($FfmpegInstallMethod -eq "winget" -or ($FfmpegInstallMethod -eq "auto" -and (Test-CommandAvailable "winget"))) {
        $script:FfmpegInstallUsed = "winget"
    } elseif ($FfmpegInstallMethod -eq "choco" -or ($FfmpegInstallMethod -eq "auto" -and (Test-CommandAvailable "choco"))) {
        $script:FfmpegInstallUsed = "choco"
    } elseif ($FfmpegInstallMethod -eq "scoop" -or ($FfmpegInstallMethod -eq "auto" -and (Test-CommandAvailable "scoop"))) {
        $script:FfmpegInstallUsed = "scoop"
    } else {
        $script:FfmpegInstallUsed = "unknown"
    }
}

Write-Host "Installing ffmpeg/ffprobe" -NoNewline -ForegroundColor White
powershell -NoProfile -ExecutionPolicy Bypass -File "$ffmpegInstallerPath" -Method $FfmpegInstallMethod 2>&1 | Out-Null
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

# Test cinemagoer (optional - IMDB episode name lookup)
Write-Host "cinemagoer (IMDB lookup)" -NoNewline -ForegroundColor White
$previousErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = "SilentlyContinue"
$cinemagoerTest = & $venvPythonCmd -c "from imdb import Cinemagoer; print('ok')" 2>&1
$ErrorActionPreference = $previousErrorActionPreference
if ($LASTEXITCODE -ne 0) {
    Write-Host "..." -NoNewline -ForegroundColor White
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "SKIP" -NoNewline -ForegroundColor Yellow
    Write-Host " ]" -ForegroundColor White
    Write-Host "  cinemagoer not available - IMDB episode lookup will be disabled." -ForegroundColor DarkYellow
    Write-Host "  Install with: pip install cinemagoer" -ForegroundColor DarkYellow
} else {
    Write-Host "..." -NoNewline -ForegroundColor White
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "OK" -NoNewline -ForegroundColor Green
    Write-Host " ]" -ForegroundColor White
}
Write-StatusOK "All components installed and verified"

# ---------------------------------------------------------------------------
# Write install manifest so the uninstaller knows what we set up.
# ---------------------------------------------------------------------------
$newManifest = @{
    install_date        = (Get-Date -Format "yyyy-MM-dd HH:mm")
    script_version      = "1.0"
    python = @{
        pre_existed     = $script:PythonPreExisted
        install_method  = if ($script:PythonPreExisted) { "" } else { $script:PythonInstallUsed }
        version         = if ($venvVersion) { $venvVersion } else { "" }
    }
    vcredist = @{
        pre_existed     = $script:VCRedistPreExisted
        install_method  = if ($script:VCRedistPreExisted) { "" } else { $script:VCRedistInstallUsed }
    }
    ffmpeg = @{
        pre_existed     = $script:FfmpegPreExisted
        install_method  = if ($script:FfmpegPreExisted) { "" } else { $script:FfmpegInstallUsed }
    }
    venv = @{
        path            = $venvPath
        created_by_script = $true
    }
    ai = @{
        installed       = ($aiSettingOverride -eq $true)
    }
}
Save-InstallManifest -Manifest $newManifest
Write-Host "Install manifest saved to: $manifestPath" -ForegroundColor Gray

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

Cleanup-InstallerArtifacts

if (-not $NoPause) {
    Write-Host ""
    Read-Host "Installation finished. Press Enter to close"
}

exit 0
