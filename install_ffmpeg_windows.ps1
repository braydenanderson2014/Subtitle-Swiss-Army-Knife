param(
    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$Method = "auto",
    [string]$MethodOutFile = ""
)

$ErrorActionPreference = "Stop"

function Test-CommandAvailable {
    param([string]$CommandName)
    return [bool](Get-Command $CommandName -ErrorAction SilentlyContinue)
}

function Test-FfmpegInstalled {
    return (Test-CommandAvailable "ffmpeg") -and (Test-CommandAvailable "ffprobe")
}

function Save-MethodOut {
    param([string]$UsedMethod)

    if (-not $MethodOutFile) { return }
    try {
        $dir = Split-Path -Parent $MethodOutFile
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        Set-Content -LiteralPath $MethodOutFile -Value ([string]$UsedMethod) -Encoding ASCII
    } catch {
        # Best-effort only.
    }
}

function Install-WithWinget {
    Write-Host "Installing ffmpeg using winget..."
    winget install --id Gyan.FFmpeg --exact --silent --accept-source-agreements --accept-package-agreements
}

function Install-WithChoco {
    Write-Host "Installing ffmpeg using Chocolatey..."
    choco install ffmpeg -y
}

function Install-WithScoop {
    Write-Host "Installing ffmpeg using Scoop..."
    scoop install ffmpeg
}

if (Test-FfmpegInstalled) {
    Write-Host "ffmpeg and ffprobe are already installed."
    Save-MethodOut -UsedMethod ""
    exit 0
}

$attempted = @()
$usedMethod = ""

try {
    if ($Method -eq "winget") {
        if (-not (Test-CommandAvailable "winget")) { throw "winget not found." }
        Install-WithWinget
        $attempted += "winget"
        if (Test-FfmpegInstalled) { $usedMethod = "winget" }
    } elseif ($Method -eq "choco") {
        if (-not (Test-CommandAvailable "choco")) { throw "choco not found." }
        Install-WithChoco
        $attempted += "choco"
        if (Test-FfmpegInstalled) { $usedMethod = "choco" }
    } elseif ($Method -eq "scoop") {
        if (-not (Test-CommandAvailable "scoop")) { throw "scoop not found." }
        Install-WithScoop
        $attempted += "scoop"
        if (Test-FfmpegInstalled) { $usedMethod = "scoop" }
    } else {
        if (Test-CommandAvailable "winget") {
            Install-WithWinget
            $attempted += "winget"
            if (Test-FfmpegInstalled) { $usedMethod = "winget" }
        } elseif (Test-CommandAvailable "choco") {
            Install-WithChoco
            $attempted += "choco"
            if (Test-FfmpegInstalled) { $usedMethod = "choco" }
        } elseif (Test-CommandAvailable "scoop") {
            Install-WithScoop
            $attempted += "scoop"
            if (Test-FfmpegInstalled) { $usedMethod = "scoop" }
        } else {
            throw "No supported package manager found. Install winget, choco, or scoop."
        }
    }
} catch {
    Write-Host "Installation attempt failed: $($_.Exception.Message)"
    if ($Method -eq "auto") {
        if ($attempted -notcontains "winget" -and (Test-CommandAvailable "winget")) {
            try {
                Install-WithWinget
                $attempted += "winget"
                if (Test-FfmpegInstalled) { $usedMethod = "winget" }
            } catch {}
        }
        if ($attempted -notcontains "choco" -and (Test-CommandAvailable "choco")) {
            try {
                Install-WithChoco
                $attempted += "choco"
                if (Test-FfmpegInstalled) { $usedMethod = "choco" }
            } catch {}
        }
        if ($attempted -notcontains "scoop" -and (Test-CommandAvailable "scoop")) {
            try {
                Install-WithScoop
                $attempted += "scoop"
                if (Test-FfmpegInstalled) { $usedMethod = "scoop" }
            } catch {}
        }
    }
}

if (Test-FfmpegInstalled) {
    Write-Host "ffmpeg installation complete."
    Save-MethodOut -UsedMethod $usedMethod
    exit 0
}

Write-Error "ffmpeg installation was not successful. Tried: $($attempted -join ', ')."
Save-MethodOut -UsedMethod ""
exit 1
