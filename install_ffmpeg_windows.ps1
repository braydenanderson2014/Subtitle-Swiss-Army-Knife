param(
    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$Method = "auto"
)

$ErrorActionPreference = "Stop"

function Test-CommandAvailable {
    param([string]$CommandName)
    return [bool](Get-Command $CommandName -ErrorAction SilentlyContinue)
}

function Test-FfmpegInstalled {
    return (Test-CommandAvailable "ffmpeg") -and (Test-CommandAvailable "ffprobe")
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
    exit 0
}

$attempted = @()

try {
    if ($Method -eq "winget") {
        if (-not (Test-CommandAvailable "winget")) { throw "winget not found." }
        Install-WithWinget
        $attempted += "winget"
    } elseif ($Method -eq "choco") {
        if (-not (Test-CommandAvailable "choco")) { throw "choco not found." }
        Install-WithChoco
        $attempted += "choco"
    } elseif ($Method -eq "scoop") {
        if (-not (Test-CommandAvailable "scoop")) { throw "scoop not found." }
        Install-WithScoop
        $attempted += "scoop"
    } else {
        if (Test-CommandAvailable "winget") {
            Install-WithWinget
            $attempted += "winget"
        } elseif (Test-CommandAvailable "choco") {
            Install-WithChoco
            $attempted += "choco"
        } elseif (Test-CommandAvailable "scoop") {
            Install-WithScoop
            $attempted += "scoop"
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
            } catch {}
        }
        if ($attempted -notcontains "choco" -and (Test-CommandAvailable "choco")) {
            try {
                Install-WithChoco
                $attempted += "choco"
            } catch {}
        }
        if ($attempted -notcontains "scoop" -and (Test-CommandAvailable "scoop")) {
            try {
                Install-WithScoop
                $attempted += "scoop"
            } catch {}
        }
    }
}

if (Test-FfmpegInstalled) {
    Write-Host "ffmpeg installation complete."
    exit 0
}

Write-Error "ffmpeg installation was not successful. Tried: $($attempted -join ', ')."
exit 1
