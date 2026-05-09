param(
    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$PythonInstallMethod = "auto",

    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$FfmpegInstallMethod = "auto",

    [ValidateSet("auto", "winget", "choco", "scoop")]
    [string]$ToolInstallMethod = "auto",

    [ValidateSet("auto", "uv", "pip")]
    [string]$PythonPackageInstallBackend = "auto",

    [ValidateSet("default", "user", "machine")]
    [string]$WingetInstallScope = "default",

    [string]$WingetInstallLocation = "",

    [string]$ChocoCacheLocation = "",

    [switch]$DisableAutoElevation,

    [switch]$InstallAiAll,
    [switch]$InstallAiOpenAIWhisper,
    [switch]$InstallAiFasterWhisper,
    [switch]$InstallAiWhisperX,
    [switch]$InstallAiStableTs,
    [switch]$InstallAiWhisperTimestamped,
    [switch]$InstallAiSpeechBrain,
    [switch]$InstallAiVosk,
    [switch]$InstallAiAeneas,
    [switch]$SkipAiSelectionPrompt,

    [switch]$InteractiveMenu,
    [switch]$NoMenu,
    [switch]$DisableAutoPathBridge,
    [switch]$DisableClickSelection,

    [switch]$NoPause,

    [switch]$KeepInstallArtifacts,

    # Suppress per-line pip/uv output (Collecting..., Downloading..., Resolved...).
    # Errors and status lines are always shown regardless of this setting.
    [switch]$QuietInstallOutput,

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
$script:AutoPathBridgeEnabled = (-not $DisableAutoPathBridge)
$script:ClickSelectionEnabled = (-not $DisableClickSelection)
$script:OptionalToolInstallMode = "prompt"
$script:MkvtoolnixInstallMethod = $ToolInstallMethod
$script:HandBrakeInstallMethod = $ToolInstallMethod
$script:MakeMKVInstallMethod = $ToolInstallMethod
$script:PythonPackageInstallBackend = $PythonPackageInstallBackend
$script:WingetInstallScope = $WingetInstallScope
$script:WingetInstallLocation = [string]$WingetInstallLocation
$script:ChocoCacheLocation = [string]$ChocoCacheLocation
$script:AutoElevationEnabled = (-not $DisableAutoElevation)
$script:OptionalToolAutoInstallKeys = @()
$script:UvExe = ""
$script:QuietOutput = [bool]$QuietInstallOutput
$script:LastPackageInstallBackendUsed = ""
$script:LastPackageInstallError = ""
$script:verboseLogPath = ""          # Set after $scriptDir is resolved
$script:PipBackendForRequirements = ""  # Backend used for requirements.txt
$script:PipBackendForAI = ""            # Backend used for AI packages

# On PowerShell 7+, do not treat native command stderr text (warnings) as a
# terminating error. The script already validates native command exit codes.
if ($null -ne (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue)) {
    $PSNativeCommandUseErrorActionPreference = $false
}

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
    if ($script:LastPackageInstallBackendUsed) {
        Write-Host "Package backend used: $($script:LastPackageInstallBackendUsed)" -ForegroundColor Yellow
    }
    if ($script:LastPackageInstallError) {
        Write-Host "Package install diagnostics:" -ForegroundColor Yellow
        Write-Host $script:LastPackageInstallError -ForegroundColor DarkYellow
    }
    $fullErr = ($_ | Out-String).Trim()
    if ($fullErr) {
        Write-Host "Full error record:" -ForegroundColor Yellow
        Write-Host $fullErr -ForegroundColor DarkYellow
    }
    Write-Host ""
    if ($script:verboseLogPath -and (Test-Path $script:verboseLogPath)) {
        Write-Host "Verbose log with full details: $($script:verboseLogPath)" -ForegroundColor Cyan
    }
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

# ---------------------------------------------------------------------------
# Verbose / detail logging helpers.
# Write-VerboseLog appends lines to install_verbose.log (never to transcript).
# Write-PhaseLog writes a timestamped phase banner to both console and log.
# ---------------------------------------------------------------------------
function Write-VerboseLog {
    param(
        [string[]]$Lines,
        [string]$Prefix = "  "
    )
    if (-not $script:verboseLogPath -or -not $Lines) { return }
    $ts = Get-Date -Format "HH:mm:ss"
    try {
        foreach ($line in $Lines) {
            if ($null -ne $line) {
                Add-Content -Path $script:verboseLogPath -Value "[$ts]$Prefix$([string]$line)" -Encoding UTF8 -ErrorAction SilentlyContinue
            }
        }
    } catch { }
}

function Write-VerboseLogBanner {
    param([string]$Message)
    if (-not $script:verboseLogPath) { return }
    $ts = Get-Date -Format "HH:mm:ss"
    $banner = "[$ts] === $Message ==="
    try {
        Add-Content -Path $script:verboseLogPath -Value "" -Encoding UTF8 -ErrorAction SilentlyContinue
        Add-Content -Path $script:verboseLogPath -Value $banner -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch { }
    # Also emit to host so transcript captures the timestamp
    Write-Host "[$ts] $Message" -ForegroundColor DarkGray
}

function Test-CommandAvailable {
    param([string]$CommandName)
    return [bool](Get-Command $CommandName -ErrorAction SilentlyContinue)
}

function Build-ScriptRelaunchArgs {
    $args = @(
        "-PythonInstallMethod", [string]$PythonInstallMethod,
        "-FfmpegInstallMethod", [string]$FfmpegInstallMethod,
        "-ToolInstallMethod", [string]$ToolInstallMethod,
        "-PythonPackageInstallBackend", [string]$script:PythonPackageInstallBackend,
        "-WingetInstallScope", [string]$script:WingetInstallScope
    )

    if ($script:WingetInstallLocation) {
        $args += "-WingetInstallLocation", ([string]$script:WingetInstallLocation)
    }
    if ($script:ChocoCacheLocation) {
        $args += "-ChocoCacheLocation", ([string]$script:ChocoCacheLocation)
    }

    if ($InstallAiAll) { $args += "-InstallAiAll" }
    if ($InstallAiOpenAIWhisper) { $args += "-InstallAiOpenAIWhisper" }
    if ($InstallAiFasterWhisper) { $args += "-InstallAiFasterWhisper" }
    if ($InstallAiWhisperX) { $args += "-InstallAiWhisperX" }
    if ($InstallAiStableTs) { $args += "-InstallAiStableTs" }
    if ($InstallAiWhisperTimestamped) { $args += "-InstallAiWhisperTimestamped" }
    if ($InstallAiSpeechBrain) { $args += "-InstallAiSpeechBrain" }
    if ($InstallAiVosk) { $args += "-InstallAiVosk" }
    if ($InstallAiAeneas) { $args += "-InstallAiAeneas" }

    if ($SkipAiSelectionPrompt) { $args += "-SkipAiSelectionPrompt" }
    if ($InteractiveMenu) { $args += "-InteractiveMenu" }
    if ($NoMenu) { $args += "-NoMenu" }
    if ($DisableAutoPathBridge) { $args += "-DisableAutoPathBridge" }
    if ($DisableClickSelection) { $args += "-DisableClickSelection" }
    if ($NoPause) { $args += "-NoPause" }
    if ($KeepInstallArtifacts) { $args += "-KeepInstallArtifacts" }
    if ($QuietInstallOutput) { $args += "-QuietInstallOutput" }
    if ($Uninstall) { $args += "-Uninstall" }
    if ($UninstallAI) { $args += "-UninstallAI" }
    if ($DisableAutoElevation) { $args += "-DisableAutoElevation" }

    return @($args)
}

function Get-WingetInstallBaseArgs {
    $args = @("--silent", "--accept-package-agreements", "--accept-source-agreements")
    if ($script:WingetInstallScope -eq "user" -or $script:WingetInstallScope -eq "machine") {
        $args += "--scope", $script:WingetInstallScope
    }
    return @($args)
}

function Invoke-WingetInstallCommand {
    param(
        [string[]]$CommandArgs,
        [switch]$AllowRetryWithoutLocation,
        [switch]$DisableCustomLocation
    )

    $baseArgs = @($CommandArgs)
    $hadLocation = $false
    if (-not $DisableCustomLocation -and $script:WingetInstallLocation) {
        $location = [string]$script:WingetInstallLocation
        $location = $location.Trim()
        if ($location) {
            if (-not (Test-Path -LiteralPath $location)) {
                New-Item -ItemType Directory -Path $location -Force | Out-Null
            }
            $baseArgs += "--location", $location
            $hadLocation = $true
        }
    }

    try {
        $process = Start-Process winget -ArgumentList $baseArgs -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 996) {
            return $true
        }
        if ($AllowRetryWithoutLocation -and $hadLocation) {
            Write-Host "Winget location override failed for this package; retrying with default location..." -ForegroundColor DarkYellow
            $retryProcess = Start-Process winget -ArgumentList $CommandArgs -Wait -NoNewWindow -PassThru
            return ($retryProcess.ExitCode -eq 0 -or $retryProcess.ExitCode -eq 996)
        }
        return $false
    } catch {
        if ($AllowRetryWithoutLocation -and $hadLocation) {
            try {
                Write-Host "Winget location override threw an exception; retrying with default location..." -ForegroundColor DarkYellow
                $retryProcess = Start-Process winget -ArgumentList $CommandArgs -Wait -NoNewWindow -PassThru
                return ($retryProcess.ExitCode -eq 0 -or $retryProcess.ExitCode -eq 996)
            } catch {
                return $false
            }
        }
        return $false
    }
}

function Get-ChocoCommonArgs {
    param(
        [switch]$DisableCustomCache
    )

    $args = @("-y")
    if (-not $DisableCustomCache -and $script:ChocoCacheLocation) {
        $cache = [string]$script:ChocoCacheLocation
        $cache = $cache.Trim()
        if ($cache) {
            if (-not (Test-Path -LiteralPath $cache)) {
                New-Item -ItemType Directory -Path $cache -Force | Out-Null
            }
            $args += "--cache-location", $cache
        }
    }
    return @($args)
}

function Refresh-ProcessPathFromRegistry {
    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machinePath;$userPath"
}

function Get-UserPathEntries {
    $rawPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    if (-not $rawPath) { return @() }
    return @($rawPath -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Set-UserPathEntries {
    param([string[]]$Entries)

    $normalized = @()
    foreach ($entry in $Entries) {
        if (-not $entry) { continue }
        $trimmed = [string]$entry
        $trimmed = $trimmed.Trim()
        if (-not $trimmed) { continue }
        $trimmed = $trimmed.TrimEnd("\\")
        if (-not $normalized.Contains($trimmed)) {
            $normalized += $trimmed
        }
    }

    [System.Environment]::SetEnvironmentVariable("Path", ($normalized -join ";"), "User")
    Refresh-ProcessPathFromRegistry
}

function Add-UserPathEntry {
    param([string]$PathEntry)

    if (-not $PathEntry) { return $false }
    if (-not (Test-Path -LiteralPath $PathEntry)) { return $false }

    $entries = Get-UserPathEntries
    $candidate = [string]$PathEntry
    $candidate = $candidate.Trim().TrimEnd("\\")
    foreach ($existing in $entries) {
        if ([string]::Equals($existing.Trim().TrimEnd("\\"), $candidate, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }
    }

    $entries += $candidate
    Set-UserPathEntries -Entries $entries
    return $true
}

function Set-UserEnvironmentVariable {
    param(
        [string]$Name,
        [string]$Value
    )

    if (-not $Name) { return $false }
    try {
        [System.Environment]::SetEnvironmentVariable($Name, $Value, "User")
        return $true
    } catch {
        return $false
    }
}

function New-MakeMKVCommandShim {
    param([string]$MakeMKVBinaryPath)

    if (-not $MakeMKVBinaryPath -or -not (Test-Path -LiteralPath $MakeMKVBinaryPath)) {
        return ""
    }

    $shimDir = Join-Path $env:LOCALAPPDATA "SubtitleTool\bin"
    New-Item -ItemType Directory -Path $shimDir -Force | Out-Null

    $shimPath = Join-Path $shimDir "makemkvcon.cmd"
    $shimContent = "@echo off`r`n`"$MakeMKVBinaryPath`" %*`r`n"
    Set-Content -Path $shimPath -Value $shimContent -Encoding ASCII -Force

    Add-UserPathEntry -PathEntry $shimDir | Out-Null
    return $shimPath
}

function Ensure-ToolCliReachable {
    param(
        [string]$ToolCommand,
        [string]$DetectedBinaryPath,
        [string]$ToolDisplayName
    )

    $envVarName = ""
    switch ($ToolCommand) {
        "mkvmerge" { $envVarName = "SUBTITLE_TOOL_MKVMERGE_BIN" }
        "HandBrakeCLI" { $envVarName = "SUBTITLE_TOOL_HANDBRAKE_BIN" }
        "makemkvcon" { $envVarName = "SUBTITLE_TOOL_MAKEMKVCON_BIN" }
    }

    if (-not $script:AutoPathBridgeEnabled) {
        return @{ changed = $false; reachable = [bool](Test-CommandAvailable $ToolCommand); reason = "disabled" }
    }

    if ($envVarName -and $DetectedBinaryPath) {
        Set-UserEnvironmentVariable -Name $envVarName -Value $DetectedBinaryPath | Out-Null
    }
    if (-not $DetectedBinaryPath) {
        return @{ changed = $false; reachable = $false; reason = "binary_not_detected" }
    }
    if (Test-CommandAvailable $ToolCommand) {
        return @{ changed = $false; reachable = $true; reason = "already_reachable" }
    }

    $changed = $false
    $reason = ""
    $fileName = [string](Split-Path -Leaf $DetectedBinaryPath)

    if ($ToolCommand -eq "makemkvcon" -and $fileName -ieq "makemkvcon64.exe") {
        $shimPath = New-MakeMKVCommandShim -MakeMKVBinaryPath $DetectedBinaryPath
        if ($shimPath) {
            $changed = $true
            $reason = "created_makemkv_shim"
        }
    } else {
        $binaryDir = Split-Path -Parent $DetectedBinaryPath
        if ($binaryDir) {
            if (Add-UserPathEntry -PathEntry $binaryDir) {
                $changed = $true
                $reason = "added_tool_dir_to_user_path"
            }
        }
    }

    Refresh-ProcessPathFromRegistry
    $reachable = [bool](Test-CommandAvailable $ToolCommand)
    if (-not $reason) {
        $reason = if ($reachable) { "reachable_after_refresh" } else { "not_reachable" }
    }

    if ($changed -and $reachable) {
        Write-Host "Auto-bridged CLI for $ToolDisplayName into PATH." -ForegroundColor Green
    } elseif ($changed -and -not $reachable) {
        Write-Host "Tried to auto-bridge CLI for $ToolDisplayName, but command is still not callable." -ForegroundColor Yellow
    }

    return @{ changed = $changed; reachable = $reachable; reason = $reason }
}

function Set-AiSelectionFlagsFromKeys {
    param([string[]]$SelectedKeys)

    $script:InstallAiAll = $false
    $script:InstallAiOpenAIWhisper = $false
    $script:InstallAiFasterWhisper = $false
    $script:InstallAiWhisperX = $false
    $script:InstallAiStableTs = $false
    $script:InstallAiWhisperTimestamped = $false
    $script:InstallAiSpeechBrain = $false
    $script:InstallAiVosk = $false
    $script:InstallAiAeneas = $false

    foreach ($key in $SelectedKeys) {
        switch ($key) {
            "openai-whisper" { $script:InstallAiOpenAIWhisper = $true }
            "faster-whisper" { $script:InstallAiFasterWhisper = $true }
            "whisperx" { $script:InstallAiWhisperX = $true }
            "stable-ts" { $script:InstallAiStableTs = $true }
            "whisper-timestamped" { $script:InstallAiWhisperTimestamped = $true }
            "speechbrain" { $script:InstallAiSpeechBrain = $true }
            "vosk" { $script:InstallAiVosk = $true }
            "aeneas" { $script:InstallAiAeneas = $true }
        }
    }

    $defs = Get-AiBackendDefinitions
    if ($SelectedKeys.Count -ge $defs.Count -and $defs.Count -gt 0) {
        $script:InstallAiAll = $true
    }
}

function Test-ClickSelectionAvailable {
    if (-not $script:ClickSelectionEnabled) {
        return $false
    }

    return [bool](Get-Command Out-GridView -ErrorAction SilentlyContinue)
}

function Test-WinFormsAvailable {
    if (-not $script:ClickSelectionEnabled) {
        return $false
    }

    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
        return $false
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

function Select-DefinitionKeysWithMouse {
    param(
        [array]$Definitions,
        [string[]]$PreselectedKeys = @(),
        [string]$Title = "Select items"
    )

    if (-not (Test-ClickSelectionAvailable)) {
        return @{
            used_gui = $false
            cancelled = $false
            selected = @()
        }
    }

    $selectedSet = @{}
    foreach ($key in $PreselectedKeys) {
        $selectedSet[[string]$key] = $true
    }

    $rows = @()
    foreach ($def in $Definitions) {
        $key = [string]$def.key
        $rows += [PSCustomObject]@{
            Selected = [bool]$selectedSet.ContainsKey($key)
            Name = [string]$def.display
            Key = $key
        }
    }

    try {
        $selectedRows = @(
            $rows |
                Sort-Object -Property @{ Expression = "Selected"; Descending = $true }, Name |
                Out-GridView -Title "$Title (multi-select with Ctrl/Shift, then click OK)" -PassThru
        )

        $isEmptySelection = ($selectedRows.Count -eq 0) -or (($selectedRows.Count -eq 1) -and $null -eq $selectedRows[0])
        if ($isEmptySelection) {
            return @{
                used_gui = $true
                cancelled = $true
                selected = @($PreselectedKeys)
            }
        }

        return @{
            used_gui = $true
            cancelled = $false
            selected = @($selectedRows | ForEach-Object { [string]$_.Key } | Select-Object -Unique)
        }
    } catch {
        Write-Host "Mouse click selection is unavailable in this host. Falling back to keyboard prompts." -ForegroundColor DarkYellow
        return @{
            used_gui = $false
            cancelled = $false
            selected = @()
        }
    }
}

function Show-InteractiveInstallerControlPanel {
    param(
        [array]$Definitions,
        [string[]]$InitialSelectedAiKeys = @()
    )

    if (-not (Test-WinFormsAvailable)) {
        return @{ used_gui = $false; submitted = $false }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Subtitle Tool Installer Control Panel"
    $form.Size = New-Object System.Drawing.Size(980, 880)
    $form.MinimumSize = New-Object System.Drawing.Size(940, 820)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.KeyPreview = $true
    $form.AutoScroll = $true

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Configure install options, then click Continue"
    $titleLabel.AutoSize = $false
    $titleLabel.Size = New-Object System.Drawing.Size(900, 28)
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $form.Controls.Add($titleLabel)

    $shortcutLabel = New-Object System.Windows.Forms.Label
    $shortcutLabel.Text = "Shortcuts: Ctrl+R Recommended, Ctrl+F Full AI, Ctrl+M Minimal, Enter Continue, Esc Cancel"
    $shortcutLabel.AutoSize = $false
    $shortcutLabel.Size = New-Object System.Drawing.Size(920, 22)
    $shortcutLabel.Location = New-Object System.Drawing.Point(20, 46)
    $form.Controls.Add($shortcutLabel)

    $methodValues = @("auto", "winget", "choco", "scoop")
    $wingetScopeValues = @("default", "user", "machine")
    $packageBackendValues = @("auto", "uv", "pip")
    $toolDefs = @(
        @{ key = "mkvtoolnix"; display = "MKVToolNix (mkvmerge)" },
        @{ key = "handbrake"; display = "HandBrakeCLI" },
        @{ key = "makemkv"; display = "MakeMKV (makemkvcon)" }
    )

    $lblPython = New-Object System.Windows.Forms.Label
    $lblPython.Text = "Python install method"
    $lblPython.AutoSize = $false
    $lblPython.Size = New-Object System.Drawing.Size(220, 22)
    $lblPython.Location = New-Object System.Drawing.Point(20, 86)
    $form.Controls.Add($lblPython)

    $comboPython = New-Object System.Windows.Forms.ComboBox
    $comboPython.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboPython.Items.AddRange($methodValues)
    $comboPython.SelectedItem = $script:PythonInstallMethod
    $comboPython.Location = New-Object System.Drawing.Point(250, 82)
    $comboPython.Width = 170
    $form.Controls.Add($comboPython)

    $lblFfmpeg = New-Object System.Windows.Forms.Label
    $lblFfmpeg.Text = "FFmpeg install method"
    $lblFfmpeg.AutoSize = $false
    $lblFfmpeg.Size = New-Object System.Drawing.Size(220, 22)
    $lblFfmpeg.Location = New-Object System.Drawing.Point(20, 120)
    $form.Controls.Add($lblFfmpeg)

    $comboFfmpeg = New-Object System.Windows.Forms.ComboBox
    $comboFfmpeg.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboFfmpeg.Items.AddRange($methodValues)
    $comboFfmpeg.SelectedItem = $script:FfmpegInstallMethod
    $comboFfmpeg.Location = New-Object System.Drawing.Point(250, 116)
    $comboFfmpeg.Width = 170
    $form.Controls.Add($comboFfmpeg)

    $lblPkgBackend = New-Object System.Windows.Forms.Label
    $lblPkgBackend.Text = "Python package backend"
    $lblPkgBackend.AutoSize = $false
    $lblPkgBackend.Size = New-Object System.Drawing.Size(220, 22)
    $lblPkgBackend.Location = New-Object System.Drawing.Point(450, 120)
    $form.Controls.Add($lblPkgBackend)

    $comboPkgBackend = New-Object System.Windows.Forms.ComboBox
    $comboPkgBackend.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboPkgBackend.Items.AddRange($packageBackendValues)
    $comboPkgBackend.SelectedItem = $script:PythonPackageInstallBackend
    $comboPkgBackend.Location = New-Object System.Drawing.Point(680, 116)
    $comboPkgBackend.Width = 210
    $form.Controls.Add($comboPkgBackend)

    $toolsGroup = New-Object System.Windows.Forms.GroupBox
    $toolsGroup.Text = "External Tools (Optional)"
    $toolsGroup.Location = New-Object System.Drawing.Point(20, 156)
    $toolsGroup.Size = New-Object System.Drawing.Size(920, 320)
    $form.Controls.Add($toolsGroup)

    $lblMkvMethod = New-Object System.Windows.Forms.Label
    $lblMkvMethod.Text = "MKVToolNix method"
    $lblMkvMethod.AutoSize = $false
    $lblMkvMethod.Size = New-Object System.Drawing.Size(220, 22)
    $lblMkvMethod.Location = New-Object System.Drawing.Point(16, 32)
    $toolsGroup.Controls.Add($lblMkvMethod)

    $comboMkvMethod = New-Object System.Windows.Forms.ComboBox
    $comboMkvMethod.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboMkvMethod.Items.AddRange($methodValues)
    $comboMkvMethod.SelectedItem = $script:MkvtoolnixInstallMethod
    $comboMkvMethod.Location = New-Object System.Drawing.Point(250, 28)
    $comboMkvMethod.Width = 170
    $toolsGroup.Controls.Add($comboMkvMethod)

    $lblHandBrakeMethod = New-Object System.Windows.Forms.Label
    $lblHandBrakeMethod.Text = "HandBrakeCLI method"
    $lblHandBrakeMethod.AutoSize = $false
    $lblHandBrakeMethod.Size = New-Object System.Drawing.Size(220, 22)
    $lblHandBrakeMethod.Location = New-Object System.Drawing.Point(16, 66)
    $toolsGroup.Controls.Add($lblHandBrakeMethod)

    $comboHandBrakeMethod = New-Object System.Windows.Forms.ComboBox
    $comboHandBrakeMethod.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboHandBrakeMethod.Items.AddRange($methodValues)
    $comboHandBrakeMethod.SelectedItem = $script:HandBrakeInstallMethod
    $comboHandBrakeMethod.Location = New-Object System.Drawing.Point(250, 62)
    $comboHandBrakeMethod.Width = 170
    $toolsGroup.Controls.Add($comboHandBrakeMethod)

    $lblMakeMKVMethod = New-Object System.Windows.Forms.Label
    $lblMakeMKVMethod.Text = "MakeMKV method"
    $lblMakeMKVMethod.AutoSize = $false
    $lblMakeMKVMethod.Size = New-Object System.Drawing.Size(220, 22)
    $lblMakeMKVMethod.Location = New-Object System.Drawing.Point(16, 100)
    $toolsGroup.Controls.Add($lblMakeMKVMethod)

    $comboMakeMKVMethod = New-Object System.Windows.Forms.ComboBox
    $comboMakeMKVMethod.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboMakeMKVMethod.Items.AddRange($methodValues)
    $comboMakeMKVMethod.SelectedItem = $script:MakeMKVInstallMethod
    $comboMakeMKVMethod.Location = New-Object System.Drawing.Point(250, 96)
    $comboMakeMKVMethod.Width = 170
    $toolsGroup.Controls.Add($comboMakeMKVMethod)

    $lblAutoInstallTools = New-Object System.Windows.Forms.Label
    $lblAutoInstallTools.Text = "Auto-install if missing"
    $lblAutoInstallTools.AutoSize = $false
    $lblAutoInstallTools.Size = New-Object System.Drawing.Size(220, 22)
    $lblAutoInstallTools.Location = New-Object System.Drawing.Point(450, 30)
    $toolsGroup.Controls.Add($lblAutoInstallTools)

    $toolAutoInstallList = New-Object System.Windows.Forms.CheckedListBox
    $toolAutoInstallList.CheckOnClick = $true
    $toolAutoInstallList.Location = New-Object System.Drawing.Point(450, 54)
    $toolAutoInstallList.Size = New-Object System.Drawing.Size(440, 76)
    foreach ($tool in $toolDefs) {
        [void]$toolAutoInstallList.Items.Add([string]$tool.display)
    }
    $toolsGroup.Controls.Add($toolAutoInstallList)

    $lblRemainingTools = New-Object System.Windows.Forms.Label
    $lblRemainingTools.Text = "For tools not auto-selected"
    $lblRemainingTools.AutoSize = $false
    $lblRemainingTools.Size = New-Object System.Drawing.Size(220, 22)
    $lblRemainingTools.Location = New-Object System.Drawing.Point(450, 140)
    $toolsGroup.Controls.Add($lblRemainingTools)

    $comboOptionalToolMode = New-Object System.Windows.Forms.ComboBox
    $comboOptionalToolMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboOptionalToolMode.Items.AddRange(@("prompt", "none"))
    $modeDefault = if ($script:OptionalToolInstallMode -eq "none") { "none" } else { "prompt" }
    $comboOptionalToolMode.SelectedItem = $modeDefault
    $comboOptionalToolMode.Location = New-Object System.Drawing.Point(680, 136)
    $comboOptionalToolMode.Width = 210
    $toolsGroup.Controls.Add($comboOptionalToolMode)

    $lblWingetScope = New-Object System.Windows.Forms.Label
    $lblWingetScope.Text = "Winget scope"
    $lblWingetScope.AutoSize = $false
    $lblWingetScope.Size = New-Object System.Drawing.Size(220, 22)
    $lblWingetScope.Location = New-Object System.Drawing.Point(16, 176)
    $toolsGroup.Controls.Add($lblWingetScope)

    $comboWingetScope = New-Object System.Windows.Forms.ComboBox
    $comboWingetScope.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$comboWingetScope.Items.AddRange($wingetScopeValues)
    $comboWingetScope.SelectedItem = [string]$script:WingetInstallScope
    $comboWingetScope.Location = New-Object System.Drawing.Point(250, 172)
    $comboWingetScope.Width = 170
    $toolsGroup.Controls.Add($comboWingetScope)

    $lblWingetLocation = New-Object System.Windows.Forms.Label
    $lblWingetLocation.Text = "Winget install location"
    $lblWingetLocation.AutoSize = $false
    $lblWingetLocation.Size = New-Object System.Drawing.Size(220, 22)
    $lblWingetLocation.Location = New-Object System.Drawing.Point(16, 210)
    $toolsGroup.Controls.Add($lblWingetLocation)

    $txtWingetLocation = New-Object System.Windows.Forms.TextBox
    $txtWingetLocation.Location = New-Object System.Drawing.Point(250, 206)
    $txtWingetLocation.Size = New-Object System.Drawing.Size(640, 24)
    $txtWingetLocation.Text = [string]$script:WingetInstallLocation
    $toolsGroup.Controls.Add($txtWingetLocation)

    $lblChocoCache = New-Object System.Windows.Forms.Label
    $lblChocoCache.Text = "Chocolatey cache location"
    $lblChocoCache.AutoSize = $false
    $lblChocoCache.Size = New-Object System.Drawing.Size(220, 22)
    $lblChocoCache.Location = New-Object System.Drawing.Point(16, 244)
    $toolsGroup.Controls.Add($lblChocoCache)

    $txtChocoCache = New-Object System.Windows.Forms.TextBox
    $txtChocoCache.Location = New-Object System.Drawing.Point(250, 240)
    $txtChocoCache.Size = New-Object System.Drawing.Size(640, 24)
    $txtChocoCache.Text = [string]$script:ChocoCacheLocation
    $toolsGroup.Controls.Add($txtChocoCache)

    $lblPathWarning = New-Object System.Windows.Forms.Label
    $lblPathWarning.Text = "Warning: custom package locations are best effort only. Some winget/choco packages may ignore these settings."
    $lblPathWarning.AutoSize = $false
    $lblPathWarning.Size = New-Object System.Drawing.Size(890, 38)
    $lblPathWarning.Location = New-Object System.Drawing.Point(16, 274)
    $lblPathWarning.ForeColor = [System.Drawing.Color]::DarkGoldenrod
    $toolsGroup.Controls.Add($lblPathWarning)

    $chkAutoPathBridge = New-Object System.Windows.Forms.CheckBox
    $chkAutoPathBridge.Text = "Enable auto PATH bridge for detected GUI-installed tools"
    $chkAutoPathBridge.Checked = [bool]$script:AutoPathBridgeEnabled
    $chkAutoPathBridge.AutoSize = $false
    $chkAutoPathBridge.Size = New-Object System.Drawing.Size(900, 24)
    $chkAutoPathBridge.Location = New-Object System.Drawing.Point(20, 488)
    $form.Controls.Add($chkAutoPathBridge)

    $chkSkipAiPrompt = New-Object System.Windows.Forms.CheckBox
    $chkSkipAiPrompt.Text = "Skip AI selection prompt if nothing selected"
    $chkSkipAiPrompt.Checked = [bool]$script:SkipAiSelectionPrompt
    $chkSkipAiPrompt.AutoSize = $false
    $chkSkipAiPrompt.Size = New-Object System.Drawing.Size(900, 24)
    $chkSkipAiPrompt.Location = New-Object System.Drawing.Point(20, 516)
    $form.Controls.Add($chkSkipAiPrompt)

    $chkRunUninstall = New-Object System.Windows.Forms.CheckBox
    $chkRunUninstall.Text = "Run full uninstall instead of install (removes venv + script-installed components)"
    $chkRunUninstall.Checked = [bool]$Uninstall
    $chkRunUninstall.AutoSize = $false
    $chkRunUninstall.Size = New-Object System.Drawing.Size(900, 24)
    $chkRunUninstall.Location = New-Object System.Drawing.Point(20, 540)
    $form.Controls.Add($chkRunUninstall)

    $chkRunUninstallAI = New-Object System.Windows.Forms.CheckBox
    $chkRunUninstallAI.Text = "Run AI-only uninstall instead of install (keeps core app dependencies)"
    $chkRunUninstallAI.Checked = [bool]$UninstallAI
    $chkRunUninstallAI.AutoSize = $false
    $chkRunUninstallAI.Size = New-Object System.Drawing.Size(900, 24)
    $chkRunUninstallAI.Location = New-Object System.Drawing.Point(20, 564)
    $form.Controls.Add($chkRunUninstallAI)

    $chkRunUninstall.Add_CheckedChanged({
        if ($chkRunUninstall.Checked) {
            $chkRunUninstallAI.Checked = $false
        }
    })
    $chkRunUninstallAI.Add_CheckedChanged({
        if ($chkRunUninstallAI.Checked) {
            $chkRunUninstall.Checked = $false
        }
    })

    $aiLabel = New-Object System.Windows.Forms.Label
    $aiLabel.Text = "AI backends to install"
    $aiLabel.AutoSize = $false
    $aiLabel.Size = New-Object System.Drawing.Size(300, 22)
    $aiLabel.Location = New-Object System.Drawing.Point(20, 590)
    $form.Controls.Add($aiLabel)

    $aiCheckedList = New-Object System.Windows.Forms.CheckedListBox
    $aiCheckedList.Location = New-Object System.Drawing.Point(20, 614)
    $aiCheckedList.Size = New-Object System.Drawing.Size(920, 120)
    $aiCheckedList.CheckOnClick = $true
    foreach ($def in $Definitions) {
        $label = "{0} ({1})" -f ([string]$def.display), ([string]$def.key)
        [void]$aiCheckedList.Items.Add($label)
    }
    $form.Controls.Add($aiCheckedList)

    $setAiCheckedState = {
        param([string[]]$Keys)
        $keySet = @{}
        foreach ($k in $Keys) { $keySet[[string]$k] = $true }
        for ($i = 0; $i -lt $Definitions.Count; $i++) {
            $k = [string]$Definitions[$i].key
            $aiCheckedList.SetItemChecked($i, [bool]$keySet.ContainsKey($k))
        }
    }

    $setToolAutoInstallCheckedState = {
        param([string[]]$Keys)
        $keySet = @{}
        foreach ($k in $Keys) { $keySet[[string]$k] = $true }
        for ($i = 0; $i -lt $toolDefs.Count; $i++) {
            $key = [string]$toolDefs[$i].key
            $toolAutoInstallList.SetItemChecked($i, [bool]$keySet.ContainsKey($key))
        }
    }

    $initialToolAutoInstallKeys = @($script:OptionalToolAutoInstallKeys | Select-Object -Unique)
    if ($initialToolAutoInstallKeys.Count -eq 0 -and $script:OptionalToolInstallMode -eq "all") {
        $initialToolAutoInstallKeys = @($toolDefs | ForEach-Object { [string]$_.key })
    }
    & $setToolAutoInstallCheckedState $initialToolAutoInstallKeys

    $selectedFromFlags = @($InitialSelectedAiKeys | Select-Object -Unique)
    if ($selectedFromFlags.Count -eq 0) {
        $selectedFromFlags = @("openai-whisper")
    }
    & $setAiCheckedState $selectedFromFlags

    $applyRecommended = {
        $comboPython.SelectedItem = "auto"
        $comboFfmpeg.SelectedItem = "auto"
        $comboPkgBackend.SelectedItem = "auto"
        $comboWingetScope.SelectedItem = "default"
        $txtWingetLocation.Text = ""
        $txtChocoCache.Text = ""
        $comboMkvMethod.SelectedItem = "auto"
        $comboHandBrakeMethod.SelectedItem = "auto"
        $comboMakeMKVMethod.SelectedItem = "auto"
        $chkAutoPathBridge.Checked = $true
        $chkSkipAiPrompt.Checked = $false
        $chkRunUninstall.Checked = $false
        $chkRunUninstallAI.Checked = $false
        $comboOptionalToolMode.SelectedItem = "prompt"
        & $setToolAutoInstallCheckedState @()
        & $setAiCheckedState @("openai-whisper")
    }

    $applyFull = {
        $comboPython.SelectedItem = "auto"
        $comboFfmpeg.SelectedItem = "auto"
        $comboPkgBackend.SelectedItem = "auto"
        $comboWingetScope.SelectedItem = "default"
        $txtWingetLocation.Text = ""
        $txtChocoCache.Text = ""
        $comboMkvMethod.SelectedItem = "auto"
        $comboHandBrakeMethod.SelectedItem = "auto"
        $comboMakeMKVMethod.SelectedItem = "auto"
        $chkAutoPathBridge.Checked = $true
        $chkSkipAiPrompt.Checked = $true
        $chkRunUninstall.Checked = $false
        $chkRunUninstallAI.Checked = $false
        $comboOptionalToolMode.SelectedItem = "none"
        & $setToolAutoInstallCheckedState @($toolDefs | ForEach-Object { [string]$_.key })
        & $setAiCheckedState @($Definitions | ForEach-Object { [string]$_.key })
    }

    $applyMinimal = {
        $comboPython.SelectedItem = "auto"
        $comboFfmpeg.SelectedItem = "auto"
        $comboPkgBackend.SelectedItem = "pip"
        $comboWingetScope.SelectedItem = "default"
        $txtWingetLocation.Text = ""
        $txtChocoCache.Text = ""
        $comboMkvMethod.SelectedItem = "auto"
        $comboHandBrakeMethod.SelectedItem = "auto"
        $comboMakeMKVMethod.SelectedItem = "auto"
        $chkAutoPathBridge.Checked = $false
        $chkSkipAiPrompt.Checked = $true
        $chkRunUninstall.Checked = $false
        $chkRunUninstallAI.Checked = $false
        $comboOptionalToolMode.SelectedItem = "none"
        & $setToolAutoInstallCheckedState @()
        & $setAiCheckedState @()
    }

    $btnRecommended = New-Object System.Windows.Forms.Button
    $btnRecommended.Text = "Recommended"
    $btnRecommended.Location = New-Object System.Drawing.Point(20, 754)
    $btnRecommended.Size = New-Object System.Drawing.Size(120, 32)
    $btnRecommended.Add_Click({ & $applyRecommended })
    $form.Controls.Add($btnRecommended)

    $btnFull = New-Object System.Windows.Forms.Button
    $btnFull.Text = "Full AI"
    $btnFull.Location = New-Object System.Drawing.Point(150, 754)
    $btnFull.Size = New-Object System.Drawing.Size(120, 32)
    $btnFull.Add_Click({ & $applyFull })
    $form.Controls.Add($btnFull)

    $btnMinimal = New-Object System.Windows.Forms.Button
    $btnMinimal.Text = "Minimal"
    $btnMinimal.Location = New-Object System.Drawing.Point(280, 754)
    $btnMinimal.Size = New-Object System.Drawing.Size(120, 32)
    $btnMinimal.Add_Click({ & $applyMinimal })
    $form.Controls.Add($btnMinimal)

    $script:__InstallerControlPanelSubmitted = $false

    $btnContinue = New-Object System.Windows.Forms.Button
    $btnContinue.Text = "Continue Install"
    $btnContinue.Location = New-Object System.Drawing.Point(710, 754)
    $btnContinue.Size = New-Object System.Drawing.Size(120, 32)
    $btnContinue.Add_Click({
        $script:PythonInstallMethod = Normalize-InstallMethod -Value $comboPython.SelectedItem
        $script:FfmpegInstallMethod = Normalize-InstallMethod -Value $comboFfmpeg.SelectedItem
        $script:PythonPackageInstallBackend = [string]$comboPkgBackend.SelectedItem
        $script:WingetInstallScope = [string]$comboWingetScope.SelectedItem
        $script:WingetInstallLocation = [string]$txtWingetLocation.Text
        $script:ChocoCacheLocation = [string]$txtChocoCache.Text
        if ($script:WingetInstallLocation) { $script:WingetInstallLocation = $script:WingetInstallLocation.Trim() }
        if ($script:ChocoCacheLocation) { $script:ChocoCacheLocation = $script:ChocoCacheLocation.Trim() }
        $script:MkvtoolnixInstallMethod = Normalize-InstallMethod -Value $comboMkvMethod.SelectedItem
        $script:HandBrakeInstallMethod = Normalize-InstallMethod -Value $comboHandBrakeMethod.SelectedItem
        $script:MakeMKVInstallMethod = Normalize-InstallMethod -Value $comboMakeMKVMethod.SelectedItem
        $script:AutoPathBridgeEnabled = [bool]$chkAutoPathBridge.Checked
        $script:SkipAiSelectionPrompt = [bool]$chkSkipAiPrompt.Checked
        $script:OptionalToolInstallMode = [string]$comboOptionalToolMode.SelectedItem

        # Installer action mode (mutually exclusive).
        if ([bool]$chkRunUninstall.Checked) {
            $Uninstall = $true
            $UninstallAI = $false
        } elseif ([bool]$chkRunUninstallAI.Checked) {
            $Uninstall = $false
            $UninstallAI = $true
        } else {
            $Uninstall = $false
            $UninstallAI = $false
        }

        $methodSet = @(
            @(
                $script:MkvtoolnixInstallMethod,
                $script:HandBrakeInstallMethod,
                $script:MakeMKVInstallMethod
            ) | ForEach-Object { Normalize-InstallMethod -Value $_ } | Select-Object -Unique
        )
        if ($methodSet.Count -eq 1) {
            $script:ToolInstallMethod = Normalize-InstallMethod -Value $methodSet[0]
        } else {
            $script:ToolInstallMethod = "auto"
        }

        $selectedToolKeys = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $toolDefs.Count; $i++) {
            if ($toolAutoInstallList.GetItemChecked($i)) {
                $selectedToolKeys.Add([string]$toolDefs[$i].key)
            }
        }
        $script:OptionalToolAutoInstallKeys = @($selectedToolKeys.ToArray() | Select-Object -Unique)

        $selectedKeys = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $Definitions.Count; $i++) {
            if ($aiCheckedList.GetItemChecked($i)) {
                $selectedKeys.Add([string]$Definitions[$i].key)
            }
        }

        Set-AiSelectionFlagsFromKeys -SelectedKeys @($selectedKeys.ToArray())
        if ($selectedKeys.Count -gt 0) {
            $script:SkipAiSelectionPrompt = $true
        }

        $script:__InstallerControlPanelSubmitted = $true
        $form.Close()
    })
    $form.Controls.Add($btnContinue)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(840, 754)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
    $btnCancel.Add_Click({
        $script:__InstallerControlPanelSubmitted = $false
        $form.Close()
    })
    $form.Controls.Add($btnCancel)

    $form.AcceptButton = $btnContinue
    $form.CancelButton = $btnCancel
    $form.Add_KeyDown({
        param($sender, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::R) {
            & $applyRecommended
            $e.SuppressKeyPress = $true
        } elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            & $applyFull
            $e.SuppressKeyPress = $true
        } elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::M) {
            & $applyMinimal
            $e.SuppressKeyPress = $true
        }
    })

    [void]$form.ShowDialog()

    return @{
        used_gui = $true
        submitted = [bool]$script:__InstallerControlPanelSubmitted
    }
}

function Read-InstallMethodChoice {
    param(
        [string]$Title,
        [string]$CurrentMethod
    )

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    Write-Host "  1) auto"
    Write-Host "  2) winget"
    Write-Host "  3) choco"
    Write-Host "  4) scoop"
    $defaultChoice = switch ($CurrentMethod) {
        "winget" { "2" }
        "choco" { "3" }
        "scoop" { "4" }
        default { "1" }
    }

    while ($true) {
        $raw = Read-Host "Choose [1-4] (default: $defaultChoice)"
        if ([string]::IsNullOrWhiteSpace("$raw")) { $raw = $defaultChoice }
        switch ($raw.Trim()) {
            "1" { return "auto" }
            "2" { return "winget" }
            "3" { return "choco" }
            "4" { return "scoop" }
        }
        Write-Host "Invalid choice. Try again." -ForegroundColor Yellow
    }
}

function Read-PackageBackendChoice {
    param([string]$CurrentBackend)

    Write-Host ""
    Write-Host "Python package backend" -ForegroundColor Cyan
    Write-Host "  1) auto (prefer uv, fallback to pip)"
    Write-Host "  2) uv (install if missing, fail if unavailable)"
    Write-Host "  3) pip"
    $defaultChoice = switch ($CurrentBackend) {
        "uv" { "2" }
        "pip" { "3" }
        default { "1" }
    }

    while ($true) {
        $raw = Read-Host "Choose [1-3] (default: $defaultChoice)"
        if ([string]::IsNullOrWhiteSpace("$raw")) { $raw = $defaultChoice }
        switch ($raw.Trim()) {
            "1" { return "auto" }
            "2" { return "uv" }
            "3" { return "pip" }
        }
        Write-Host "Invalid choice. Try again." -ForegroundColor Yellow
    }
}

function Read-WingetScopeChoice {
    param([string]$CurrentScope)

    Write-Host ""
    Write-Host "Winget install scope" -ForegroundColor Cyan
    Write-Host "  1) default"
    Write-Host "  2) user"
    Write-Host "  3) machine"
    $defaultChoice = switch ($CurrentScope) {
        "user" { "2" }
        "machine" { "3" }
        default { "1" }
    }

    while ($true) {
        $raw = Read-Host "Choose [1-3] (default: $defaultChoice)"
        if ([string]::IsNullOrWhiteSpace("$raw")) { $raw = $defaultChoice }
        switch ($raw.Trim()) {
            "1" { return "default" }
            "2" { return "user" }
            "3" { return "machine" }
        }
        Write-Host "Invalid choice. Try again." -ForegroundColor Yellow
    }
}

function Read-OptionalPathValue {
    param(
        [string]$Title,
        [string]$CurrentValue
    )

    Write-Host ""
    Write-Host $Title -ForegroundColor Cyan
    if ($CurrentValue) {
        Write-Host "Current: $CurrentValue" -ForegroundColor DarkGray
    } else {
        Write-Host "Current: (empty/default)" -ForegroundColor DarkGray
    }
    Write-Host "Enter a full path, or leave blank to use default behavior."
    $raw = Read-Host "Path"
    if ([string]::IsNullOrWhiteSpace("$raw")) {
        return ""
    }
    return $raw.Trim()
}

function Normalize-InstallMethod {
    param([object]$Value)

    $candidate = [string]$Value
    if ($candidate -in @("auto", "winget", "choco", "scoop")) {
        return $candidate
    }
    return "auto"
}

function Show-InteractiveInstallerMenu {
    if ($NoMenu -or $Uninstall -or $UninstallAI) {
        return
    }

    $shouldShow = $InteractiveMenu
    if (-not $shouldShow) {
        $hasAnyAiFlag = $InstallAiAll -or $InstallAiOpenAIWhisper -or $InstallAiFasterWhisper -or $InstallAiWhisperX -or $InstallAiStableTs -or $InstallAiWhisperTimestamped -or $InstallAiSpeechBrain -or $InstallAiVosk -or $InstallAiAeneas
        $usingDefaultMethods = ($PythonInstallMethod -eq "auto" -and $FfmpegInstallMethod -eq "auto" -and $ToolInstallMethod -eq "auto")
        if ($usingDefaultMethods -and -not $hasAnyAiFlag -and -not $SkipAiSelectionPrompt) {
            $shouldShow = $true
        }
    }

    if (-not $shouldShow) {
        return
    }

    if (-not $script:PythonPackageInstallBackend) {
        $script:PythonPackageInstallBackend = "auto"
    }

    $defs = Get-AiBackendDefinitions
    $selectedFromFlags = Get-AiBackendSelectionFromFlags -Definitions $defs
    $selectedAiKeys = @($selectedFromFlags.selected)

    $controlPanelResult = Show-InteractiveInstallerControlPanel -Definitions $defs -InitialSelectedAiKeys $selectedAiKeys
    if ([bool]$controlPanelResult.used_gui) {
        if ([bool]$controlPanelResult.submitted) {
            return
        }
        exit 0
    }

    while ($true) {
        Write-Host ""
        Write-Host "=== Interactive Installer Menu ===" -ForegroundColor Cyan
        Write-Host "1) Python install method      : $script:PythonInstallMethod"
        Write-Host "2) FFmpeg install method      : $script:FfmpegInstallMethod"
        Write-Host "3) Python package backend    : $script:PythonPackageInstallBackend"
        Write-Host "4) Tool install method        : $script:ToolInstallMethod"
        Write-Host "5) Auto PATH bridge           : $(if ($script:AutoPathBridgeEnabled) { 'enabled' } else { 'disabled' })"
        Write-Host "6) AI backend selection       : $(if ($selectedAiKeys.Count -gt 0) { ($selectedAiKeys -join ', ') } else { 'prompt later' })"
        Write-Host "7) Click selection UI         : $(if ($script:ClickSelectionEnabled) { 'enabled' } else { 'disabled' })"
        Write-Host "8) Skip AI prompt             : $(if ($script:SkipAiSelectionPrompt) { 'yes' } else { 'no' })"
        Write-Host "9) Winget scope               : $script:WingetInstallScope"
        Write-Host "10) Winget install location   : $(if ($script:WingetInstallLocation) { $script:WingetInstallLocation } else { 'default' })"
        Write-Host "11) Chocolatey cache location : $(if ($script:ChocoCacheLocation) { $script:ChocoCacheLocation } else { 'default' })"
        Write-Host "12) Continue install"
        Write-Host "13) Continue full uninstall"
        Write-Host "14) Continue AI-only uninstall"
        Write-Host "15) Exit"

        $choice = Read-Host "Choose an option [1-15]"
        switch ($choice.Trim()) {
            "1" { $script:PythonInstallMethod = Read-InstallMethodChoice -Title "Python Install Method" -CurrentMethod $script:PythonInstallMethod }
            "2" { $script:FfmpegInstallMethod = Read-InstallMethodChoice -Title "FFmpeg Install Method" -CurrentMethod $script:FfmpegInstallMethod }
            "3" { $script:PythonPackageInstallBackend = Read-PackageBackendChoice -CurrentBackend $script:PythonPackageInstallBackend }
            "4" {
                $selectedToolMethod = Read-InstallMethodChoice -Title "Tool Install Method" -CurrentMethod $script:ToolInstallMethod
                $script:ToolInstallMethod = $selectedToolMethod
                $script:MkvtoolnixInstallMethod = $selectedToolMethod
                $script:HandBrakeInstallMethod = $selectedToolMethod
                $script:MakeMKVInstallMethod = $selectedToolMethod
            }
            "5" { $script:AutoPathBridgeEnabled = -not $script:AutoPathBridgeEnabled }
            "6" {
                if (Test-ClickSelectionAvailable) {
                    $mouseSelected = Select-DefinitionKeysWithMouse `
                        -Definitions $defs `
                        -PreselectedKeys $selectedAiKeys `
                        -Title "Select AI backends to install"

                    if ([bool]$mouseSelected.used_gui -and -not [bool]$mouseSelected.cancelled) {
                        $selectedAiKeys = @($mouseSelected.selected)
                        Set-AiSelectionFlagsFromKeys -SelectedKeys $selectedAiKeys
                        if ($selectedAiKeys.Count -gt 0) {
                            $script:SkipAiSelectionPrompt = $true
                        }
                        continue
                    }

                    if ([bool]$mouseSelected.cancelled) {
                        Write-Host "Selection window cancelled. Keeping existing AI selections." -ForegroundColor Yellow
                        continue
                    }
                }

                while ($true) {
                    Write-Host ""
                    Write-Host "AI Backends (toggle by number, A=all, N=none, D=done)" -ForegroundColor Cyan
                    for ($i = 0; $i -lt $defs.Count; $i++) {
                        $key = [string]$defs[$i].key
                        $label = [string]$defs[$i].display
                        $mark = if ($selectedAiKeys -contains $key) { "x" } else { " " }
                        Write-Host "  $($i + 1)) [$mark] $label ($key)"
                    }
                    $cmd = Read-Host "Enter choice"
                    if ([string]::IsNullOrWhiteSpace("$cmd")) { continue }
                    $cmdNorm = $cmd.Trim().ToUpperInvariant()
                    if ($cmdNorm -eq "D") { break }
                    if ($cmdNorm -eq "A") {
                        $selectedAiKeys = @($defs | ForEach-Object { [string]$_.key })
                        continue
                    }
                    if ($cmdNorm -eq "N") {
                        $selectedAiKeys = @()
                        continue
                    }

                    $tokens = $cmd -split ","
                    foreach ($token in $tokens) {
                        $numRaw = $token.Trim()
                        $num = 0
                        if ([int]::TryParse($numRaw, [ref]$num)) {
                            if ($num -ge 1 -and $num -le $defs.Count) {
                                $key = [string]$defs[$num - 1].key
                                if ($selectedAiKeys -contains $key) {
                                    $selectedAiKeys = @($selectedAiKeys | Where-Object { $_ -ne $key })
                                } else {
                                    $selectedAiKeys += $key
                                }
                            }
                        }
                    }
                    $selectedAiKeys = @($selectedAiKeys | Select-Object -Unique)
                }

                Set-AiSelectionFlagsFromKeys -SelectedKeys $selectedAiKeys
                if ($selectedAiKeys.Count -gt 0) {
                    $script:SkipAiSelectionPrompt = $true
                }
            }
            "7" { $script:ClickSelectionEnabled = -not $script:ClickSelectionEnabled }
            "8" { $script:SkipAiSelectionPrompt = -not $script:SkipAiSelectionPrompt }
            "9" { $script:WingetInstallScope = Read-WingetScopeChoice -CurrentScope $script:WingetInstallScope }
            "10" { $script:WingetInstallLocation = Read-OptionalPathValue -Title "Winget install location" -CurrentValue $script:WingetInstallLocation }
            "11" { $script:ChocoCacheLocation = Read-OptionalPathValue -Title "Chocolatey cache location" -CurrentValue $script:ChocoCacheLocation }
            "12" { return }
            "13" {
                $Uninstall = $true
                $UninstallAI = $false
                return
            }
            "14" {
                $Uninstall = $false
                $UninstallAI = $true
                return
            }
            "15" { exit 0 }
            default { Write-Host "Invalid choice. Try again." -ForegroundColor Yellow }
        }
    }
}

function Resolve-ExecutablePath {
    param(
        [string]$CommandName,
        [string[]]$FallbackPaths = @()
    )

    try {
        $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cmd) {
            if ($cmd.Path) {
                return [string]$cmd.Path
            }
            if ($cmd.Source) {
                return [string]$cmd.Source
            }
        }
    } catch {
        # Continue to fallback path checks.
    }

    foreach ($candidate in $FallbackPaths) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return [string]$candidate
        }
    }

    return ""
}

function Get-ToolInstallDirFromRegistry {
    # Searches Windows uninstall registry entries for a tool by display name pattern.
    # Returns an install directory if found and on disk, or "" otherwise.
    # Falls back to parsing DisplayIcon / UninstallString when InstallLocation is blank
    # (common for NSIS installers and winget-managed packages).
    param([string]$DisplayNamePattern)
    $regRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($root in $regRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        try {
            foreach ($key in (Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
                $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                if ([string]$props.DisplayName -match $DisplayNamePattern) {
                    $loc = [string]$props.InstallLocation
                    if ($loc -and (Test-Path -LiteralPath $loc)) { return $loc }
                    # InstallLocation absent — derive directory from DisplayIcon or UninstallString
                    foreach ($rawVal in @([string]$props.DisplayIcon, [string]$props.UninstallString)) {
                        if (-not $rawVal) { continue }
                        # Strip leading/trailing quotes and any trailing arguments
                        $exePath = $rawVal.Trim('"').Split('"')[0].Trim()
                        $dir = Split-Path -Parent $exePath -ErrorAction SilentlyContinue
                        if ($dir -and (Test-Path -LiteralPath $dir)) { return $dir }
                    }
                }
            }
        } catch { }
    }
    return ""
}

function Get-OptionalVideoToolStatus {
    # Build extended fallback lists augmented with registry-detected install dirs.
    # HandBrake's display name includes the version (e.g. "HandBrake 1.8.0"), so don't anchor with $.
    $hbRegDir  = Get-ToolInstallDirFromRegistry -DisplayNamePattern "(?i)^HandBrake(\s|$)"
    $mkvRegDir = Get-ToolInstallDirFromRegistry -DisplayNamePattern "(?i)MKVToolNix"
    $mmvRegDir = Get-ToolInstallDirFromRegistry -DisplayNamePattern "(?i)^MakeMKV"

    $hbFallbacks = @(
        (Join-Path $env:ProgramFiles "HandBrake\HandBrakeCLI.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "HandBrake\HandBrakeCLI.exe"),
        (Join-Path $env:LOCALAPPDATA "HandBrake\HandBrakeCLI.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\HandBrake\HandBrakeCLI.exe")
    )
    if ($hbRegDir) { $hbFallbacks += (Join-Path $hbRegDir "HandBrakeCLI.exe") }
    # Also search the WinGet packages directory (user-scope winget installs land here for portable packages).
    $wingetPkgBase = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages"
    if (Test-Path -LiteralPath $wingetPkgBase) {
        $hbPkgDirs = Get-ChildItem -Path $wingetPkgBase -Directory -Filter "HandBrake*" -ErrorAction SilentlyContinue
        foreach ($d in $hbPkgDirs) {
            $hbFallbacks += (Join-Path $d.FullName "HandBrakeCLI.exe")
        }
    }

    $mkvFallbacks = @(
        (Join-Path $env:ProgramFiles "MKVToolNix\mkvmerge.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "MKVToolNix\mkvmerge.exe")
    )
    if ($mkvRegDir) { $mkvFallbacks += (Join-Path $mkvRegDir "mkvmerge.exe") }

    $mmvFallbacks = @(
        (Join-Path $env:ProgramFiles "MakeMKV\makemkvcon64.exe"),
        (Join-Path $env:ProgramFiles "MakeMKV\makemkvcon.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "MakeMKV\makemkvcon64.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "MakeMKV\makemkvcon.exe")
    )
    if ($mmvRegDir) {
        $mmvFallbacks += (Join-Path $mmvRegDir "makemkvcon64.exe")
        $mmvFallbacks += (Join-Path $mmvRegDir "makemkvcon.exe")
    }

    $mkvmergePath  = Resolve-ExecutablePath -CommandName "mkvmerge"    -FallbackPaths $mkvFallbacks
    $handbrakePath = Resolve-ExecutablePath -CommandName "HandBrakeCLI" -FallbackPaths $hbFallbacks
    $makemkvPath   = Resolve-ExecutablePath -CommandName "makemkvcon"   -FallbackPaths $mmvFallbacks

    return @{
        mkvtoolnix = @{
            found = [bool]$mkvmergePath
            path = $mkvmergePath
            command = "mkvmerge"
            command_on_path = [bool](Test-CommandAvailable "mkvmerge")
            display = "MKVToolNix (mkvmerge)"
            winget_id = "MoritzBunkus.MKVToolNix"
            choco_id = "mkvtoolnix"
            scoop_id = "mkvtoolnix"
        }
        handbrake = @{
            found = [bool]$handbrakePath
            path = $handbrakePath
            command = "HandBrakeCLI"
            command_on_path = [bool](Test-CommandAvailable "HandBrakeCLI")
            display = "HandBrakeCLI"
            winget_id = "HandBrake.HandBrakeCLI"
            choco_id = "handbrake-cli"
            scoop_id = "handbrake-cli"
        }
        makemkv = @{
            found = [bool]$makemkvPath
            path = $makemkvPath
            command = "makemkvcon"
            command_on_path = [bool](Test-CommandAvailable "makemkvcon")
            display = "MakeMKV (makemkvcon)"
            winget_id = "GuinpinSoft.MakeMKV"
            choco_id = "makemkv"
            scoop_id = "makemkv"
        }
    }
}

function Install-WithWingetId {
    param([string]$PackageId)
    if (-not (Test-CommandAvailable "winget")) { return $false }
    $installArgs = @("install", "--id", $PackageId, "--exact") + (Get-WingetInstallBaseArgs)
    if (Invoke-WingetInstallCommand -CommandArgs $installArgs -AllowRetryWithoutLocation) { return $true }
    # Winget exits with a non-zero code when a package is already installed and no
    # upgrade is available. Treat that as success via a list verification.
    try {
        $listOut = & winget list --id $PackageId --exact --accept-source-agreements 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0 -and $listOut -match [regex]::Escape($PackageId)) { return $true }
    } catch { }
    return $false
}

function Install-WithChocoPackage {
    param([string]$PackageName)
    if (-not (Test-CommandAvailable "choco")) { return $false }
    try {
        $chocoArgs = @("install", $PackageName) + (Get-ChocoCommonArgs)
        $process = Start-Process choco -ArgumentList $chocoArgs -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -eq 0) { return $true }
        # Choco may exit non-zero when the package is already installed.
        # Verify via a local list check.
        $listOut = & choco list --local-only $PackageName 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0 -and $listOut -match ("(?i)\b" + [regex]::Escape($PackageName) + "\b.*\d")) { return $true }
        return $false
    } catch {
        return $false
    }
}

function Install-WithScoopPackage {
    param([string]$PackageName)
    if (-not (Test-CommandAvailable "scoop")) { return $false }
    try {
        $process = Start-Process scoop -ArgumentList "install", $PackageName -Wait -NoNewWindow -PassThru
        return ($process.ExitCode -eq 0)
    } catch {
        return $false
    }
}

function Install-OptionalSystemTool {
    param(
        [string]$DisplayName,
        [string]$WingetId,
        [string]$ChocoId,
        [string]$ScoopId,
        [string]$Method = "auto"
    )

    $attempted = @()
    $usedMethod = ""
    $success = $false
    $methodsToTry = if ($Method -eq "auto") { @("winget", "choco", "scoop") } else { @($Method) }

    foreach ($candidateMethod in $methodsToTry) {
        if ($candidateMethod -eq "winget") {
            if (-not $WingetId -or -not (Test-CommandAvailable "winget")) { continue }
            $attempted += "winget"
            Write-Host "Installing $DisplayName via winget..." -ForegroundColor Yellow
            if (Install-WithWingetId -PackageId $WingetId) {
                $success = $true
                $usedMethod = "winget"
                break
            }
        } elseif ($candidateMethod -eq "choco") {
            if (-not $ChocoId -or -not (Test-CommandAvailable "choco")) { continue }
            $attempted += "choco"
            Write-Host "Installing $DisplayName via Chocolatey..." -ForegroundColor Yellow
            if (Install-WithChocoPackage -PackageName $ChocoId) {
                $success = $true
                $usedMethod = "choco"
                break
            }
        } elseif ($candidateMethod -eq "scoop") {
            if (-not $ScoopId -or -not (Test-CommandAvailable "scoop")) { continue }
            $attempted += "scoop"
            Write-Host "Installing $DisplayName via Scoop..." -ForegroundColor Yellow
            if (Install-WithScoopPackage -PackageName $ScoopId) {
                $success = $true
                $usedMethod = "scoop"
                break
            }
        }
    }

    return @{
        success = $success
        install_method = $usedMethod
        attempted = $attempted
    }
}

function Get-AiBackendDefinitions {
    return @(
        @{ key = "openai-whisper"; display = "OpenAI Whisper (original)"; import = "whisper"; packages = @("torch", "openai-whisper", "pysubs2"); needs_vcredist = $true }
        @{ key = "faster-whisper"; display = "faster-whisper"; import = "faster_whisper"; packages = @("faster-whisper", "pysubs2"); needs_vcredist = $true }
        @{ key = "whisperx"; display = "WhisperX"; import = "whisperx"; packages = @("whisperx", "pysubs2"); needs_vcredist = $true }
        @{ key = "stable-ts"; display = "stable-ts"; import = "stable_whisper"; packages = @("stable-ts", "pysubs2"); needs_vcredist = $true }
        @{ key = "whisper-timestamped"; display = "whisper-timestamped"; import = "whisper_timestamped"; packages = @("whisper-timestamped", "pysubs2"); needs_vcredist = $true }
        @{ key = "speechbrain"; display = "SpeechBrain"; import = "speechbrain"; packages = @("speechbrain", "soundfile", "pysubs2"); needs_vcredist = $true }
        @{ key = "vosk"; display = "Vosk"; import = "vosk"; packages = @("vosk", "pysubs2"); needs_vcredist = $false }
        @{ key = "aeneas"; display = "Aeneas"; import = "aeneas"; packages = @("aeneas", "pysubs2"); needs_vcredist = $false }
    )
}

function Get-AiBackendSelectionFromFlags {
    param([array]$Definitions)

    $selected = New-Object System.Collections.Generic.List[string]
    $hasFlags = $false

    if ($InstallAiAll) {
        foreach ($def in $Definitions) {
            $selected.Add([string]$def.key)
        }
        return @{ has_flags = $true; selected = @($selected.ToArray() | Select-Object -Unique) }
    }

    $flagMap = @{
        "openai-whisper" = [bool]$InstallAiOpenAIWhisper
        "faster-whisper" = [bool]$InstallAiFasterWhisper
        "whisperx" = [bool]$InstallAiWhisperX
        "stable-ts" = [bool]$InstallAiStableTs
        "whisper-timestamped" = [bool]$InstallAiWhisperTimestamped
        "speechbrain" = [bool]$InstallAiSpeechBrain
        "vosk" = [bool]$InstallAiVosk
        "aeneas" = [bool]$InstallAiAeneas
    }

    foreach ($entry in $flagMap.GetEnumerator()) {
        if ([bool]$entry.Value) {
            $hasFlags = $true
            $selected.Add([string]$entry.Key)
        }
    }

    return @{ has_flags = $hasFlags; selected = @($selected.ToArray() | Select-Object -Unique) }
}

function Prompt-AiBackendSelections {
    param([array]$Definitions)

    if (Test-ClickSelectionAvailable) {
        $defaultSelection = @("openai-whisper")
        $mouseSelected = Select-DefinitionKeysWithMouse `
            -Definitions $Definitions `
            -PreselectedKeys $defaultSelection `
            -Title "Select AI backends to install"

        if ([bool]$mouseSelected.used_gui -and -not [bool]$mouseSelected.cancelled) {
            return @($mouseSelected.selected)
        }

        if ([bool]$mouseSelected.cancelled) {
            Write-Host "Selection window cancelled. Falling back to keyboard prompts." -ForegroundColor Yellow
        }
    }

    $selected = New-Object System.Collections.Generic.List[string]
    Write-Host ""
    Write-Host "Select AI backends to install (each can be installed independently):" -ForegroundColor White

    foreach ($def in $Definitions) {
        $default = if ([string]$def.key -eq "openai-whisper") { "Y" } else { "N" }
        $answer = Read-Host "  Install $($def.display)? [Y/N] (default: $default)"
        if ([string]::IsNullOrWhiteSpace("$answer")) {
            $answer = $default
        }
        if ("$answer" -match "^[Yy]") {
            $selected.Add([string]$def.key)
        }
    }

    return @($selected.ToArray() | Select-Object -Unique)
}

function Get-InstalledAiBackends {
    param(
        [string]$PythonExe,
        [array]$Definitions
    )

    $installed = @()
    $probePrevErrorAction = $ErrorActionPreference
    foreach ($def in $Definitions) {
        $importName = [string]$def.import
        $probeCode = "import importlib.util as u,sys; sys.exit(0 if u.find_spec('$importName') is not None else 1)"
        $ErrorActionPreference = 'Continue'
        & $PythonExe -c $probeCode 2>$null | Out-Null
        $ErrorActionPreference = $probePrevErrorAction
        if ($LASTEXITCODE -eq 0) {
            $installed += [string]$def.key
        }
    }
    $ErrorActionPreference = $probePrevErrorAction
    return @($installed | Select-Object -Unique)
}

function Get-EspeakNgInstallDir {
    # Check common Windows install locations first.
    $candidates = @(
        (Join-Path $env:ProgramFiles "eSpeak NG"),
        (Join-Path ${env:ProgramFiles(x86)} "eSpeak NG"),
        (Join-Path $env:ProgramFiles "eSpeak-NG"),
        (Join-Path ${env:ProgramFiles(x86)} "eSpeak-NG"),
        (Join-Path $env:ProgramFiles "eSpeak"),
        (Join-Path ${env:ProgramFiles(x86)} "eSpeak"),
        (Join-Path $env:LOCALAPPDATA "Programs\eSpeak NG")
    )

    $dllNames = @("libespeak-ng.dll", "espeak-ng.dll", "espeak.dll", "libespeak.dll")
    foreach ($dir in $candidates) {
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($dllName in $dllNames) {
            if (Test-Path -LiteralPath (Join-Path $dir $dllName)) {
                return $dir
            }
        }
    }

    # Fall back to registry uninstall entries (including DisplayIcon/UninstallString parse).
    $regDir = Get-ToolInstallDirFromRegistry -DisplayNamePattern "(?i)^eSpeak(\s|-)?(NG)?(\s|$)"
    if ($regDir -and (Test-Path -LiteralPath $regDir)) {
        return $regDir
    }

    # Winget user installs may be under LocalAppData\Microsoft\WinGet\Packages.
    $wingetPkgBase = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages"
    if (Test-Path -LiteralPath $wingetPkgBase) {
        $pkgDirs = Get-ChildItem -Path $wingetPkgBase -Directory -Filter "*eSpeak*" -ErrorAction SilentlyContinue
        foreach ($dir in $pkgDirs) {
            foreach ($dllName in $dllNames) {
                $match = Get-ChildItem -Path $dir.FullName -Filter $dllName -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($match) {
                    return (Split-Path -Parent $match.FullName)
                }
            }
        }
    }

    # Last resort: infer location from command discovery.
    foreach ($cmdName in @("espeak-ng", "espeak")) {
        try {
            $cmd = Get-Command $cmdName -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($cmd -and $cmd.Path) {
                $dir = Split-Path -Parent ([string]$cmd.Path)
                if ($dir -and (Test-Path -LiteralPath $dir)) {
                    return $dir
                }
            }
        } catch { }
    }

    return ""
}

function Ensure-EspeakNgForAeneas {
    # Installs eSpeak-NG and generates an 'espeak.lib' import library so aeneas
    # can compile its CEW C extension against the eSpeak-NG DLL.
    # Returns $true when espeak.lib is set up and CEW build can proceed;
    # caller should set AENEAS_WITH_CEW=False when this returns $false.

    Write-Host "  Setting up eSpeak-NG for aeneas..." -NoNewline -ForegroundColor Gray

    $espeakDir = Get-EspeakNgInstallDir

    if (-not $espeakDir) {
        Write-Host "" -ForegroundColor Gray
        if (Test-CommandAvailable "winget") {
            Write-Host "  Installing eSpeak-NG via winget..." -ForegroundColor Yellow
            foreach ($wingetId in @("eSpeak-NG.eSpeak-NG", "espeak.espeak")) {
                $null = & winget install --id $wingetId --exact --silent --accept-source-agreements --accept-package-agreements 2>&1
                $espeakDir = Get-EspeakNgInstallDir
                if ($espeakDir) { break }
            }
        }
        if (-not $espeakDir -and (Test-CommandAvailable "choco")) {
            Write-Host "  Installing eSpeak-NG via Chocolatey..." -ForegroundColor Yellow
            foreach ($chocoId in @("espeak-ng", "espeak")) {
                $null = & choco install $chocoId -y 2>&1
                $espeakDir = Get-EspeakNgInstallDir
                if ($espeakDir) { break }
            }
        }
        Write-Host "  Setting up eSpeak-NG for aeneas..." -NoNewline -ForegroundColor Gray
    }

    if (-not $espeakDir) {
        Write-Host " not found" -ForegroundColor Yellow
        return $false
    }

    # Make espeak-ng.exe callable in this session.
    if ($env:Path -notlike "*$espeakDir*") {
        $env:Path = "$espeakDir;$env:Path"
    }

    # Locate the eSpeak-NG DLL (try common name variants).
    $dll = ""
    foreach ($dllName in @("libespeak-ng.dll", "espeak-ng.dll", "espeak.dll", "libespeak.dll")) {
        $candidate = Join-Path $espeakDir $dllName
        if (Test-Path -LiteralPath $candidate) { $dll = $candidate; break }
    }

    if (-not $dll) {
        Write-Host " DLL not found, building without CEW" -ForegroundColor Yellow
        return $false
    }

    # Need MSVC dumpbin + lib.exe to generate the import library.
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        Write-Host " MSVC not found, building without CEW" -ForegroundColor Yellow
        return $false
    }

    try {
        $vsPath = [string](& $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null)
        if ([string]::IsNullOrWhiteSpace($vsPath) -or -not (Test-Path -LiteralPath $vsPath)) {
            Write-Host " MSVC install not on disk, building without CEW" -ForegroundColor Yellow
            return $false
        }

        $msvcRoot = Join-Path $vsPath "VC\Tools\MSVC"
        $versions  = Get-ChildItem -Path $msvcRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        if (-not $versions) {
            Write-Host " MSVC version dir not found, building without CEW" -ForegroundColor Yellow
            return $false
        }

        $toolBin = Join-Path $versions[0].FullName "bin\HostX64\x64"
        $dumpbin = Join-Path $toolBin "dumpbin.exe"
        $libExe  = Join-Path $toolBin "lib.exe"
        if (-not (Test-Path $dumpbin) -or -not (Test-Path $libExe)) {
            Write-Host " dumpbin/lib.exe missing, building without CEW" -ForegroundColor Yellow
            return $false
        }

        # Dump exports from the DLL to build a .def file.
        $exports = @(& $dumpbin /exports $dll 2>$null)
        if ($LASTEXITCODE -ne 0) {
            Write-Host " dumpbin failed, building without CEW" -ForegroundColor Yellow
            return $false
        }

        $dllBaseName = [System.IO.Path]::GetFileNameWithoutExtension($dll)
        $defLines = New-Object System.Collections.Generic.List[string]
        $defLines.Add("LIBRARY $dllBaseName")
        $defLines.Add("EXPORTS")
        $inExportTable = $false
        foreach ($line in $exports) {
            $s = [string]$line
            if ($s -match "ordinal\s+hint\s+RVA\s+name") { $inExportTable = $true; continue }
            if ($inExportTable) {
                if ($s -match "^\s*Summary") { break }
                if ($s -match "^\s+\d+\s+\w+\s+[0-9A-Fa-f]+\s+(\S+)") {
                    $fn = $matches[1]
                    if ($fn -and $fn -notmatch "^\[") { $defLines.Add("    $fn") }
                }
            }
        }

        if ($defLines.Count -le 2) {
            Write-Host " no exports parsed, building without CEW" -ForegroundColor Yellow
            return $false
        }

        # Generate espeak.lib in a temp directory and add it to LIB search path.
        $libOutDir = Join-Path $env:TEMP "espeak-importlib"
        if (-not (Test-Path -LiteralPath $libOutDir)) {
            New-Item -ItemType Directory -Path $libOutDir -Force | Out-Null
        }
        $defPath = Join-Path $libOutDir "espeak.def"
        $libPath = Join-Path $libOutDir "espeak.lib"
        $defLines | Set-Content -Path $defPath -Encoding ASCII

        $null = & $libExe "/def:$defPath" /machine:x64 "/out:$libPath" 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $libPath)) {
            Write-Host " lib.exe failed, building without CEW" -ForegroundColor Yellow
            return $false
        }

        # Add generated lib dir to MSVC linker search path.
        $env:LIB = if ($env:LIB) { "$libOutDir;$env:LIB" } else { $libOutDir }

        # Add eSpeak-NG include paths if present.
        foreach ($incDir in @((Join-Path $espeakDir "include"), (Join-Path $espeakDir "include\espeak"))) {
            if (Test-Path -LiteralPath $incDir) {
                $env:INCLUDE = if ($env:INCLUDE) { "$incDir;$env:INCLUDE" } else { $incDir }
            }
        }

        Write-Host " ready (espeak.lib generated)" -ForegroundColor Green
        return $true

    } catch {
        Write-Host " error: $($_.Exception.Message), building without CEW" -ForegroundColor Yellow
        return $false
    }
}

function Install-AiBackends {
    param(
        [string]$PythonExe,
        [array]$Definitions,
        [string[]]$SelectedBackends
    )

    $selectedSet = @{}
    foreach ($k in $SelectedBackends) {
        $selectedSet[[string]$k] = $true
    }

    $selectedDefs = @($Definitions | Where-Object { $selectedSet.ContainsKey([string]$_.key) })
    if ($selectedDefs.Count -eq 0) {
        return @{ installed = @(); failed = @(); attempted_packages = @() }
    }

    $packageInstallOrder = New-Object System.Collections.Generic.List[string]
    foreach ($def in $selectedDefs) {
        foreach ($pkg in [array]$def.packages) {
            if (-not $packageInstallOrder.Contains([string]$pkg)) {
                $packageInstallOrder.Add([string]$pkg)
            }
        }
    }

    # aeneas build requires setuptools/wheel/numpy present in the current environment first.
    if ($packageInstallOrder.Contains("aeneas")) {
        if (-not $packageInstallOrder.Contains("numpy")) {
            $packageInstallOrder.Insert(0, "numpy")
        }
        if (-not $packageInstallOrder.Contains("wheel")) {
            $packageInstallOrder.Insert(0, "wheel")
        }
        if (-not $packageInstallOrder.Contains("setuptools")) {
            $packageInstallOrder.Insert(0, "setuptools")
        }
    }

    Write-Host "Installing selected AI backend packages..." -ForegroundColor White
    $aeneasNoCew = $false
    if ($packageInstallOrder.Contains("aeneas")) {
        $espeakCewReady = Ensure-EspeakNgForAeneas
        if (-not $espeakCewReady) {
            $aeneasNoCew = $true
            Write-Host "  aeneas will be installed without the CEW C extension (eSpeak-NG unavailable)." -ForegroundColor DarkYellow
            Write-Host "  aeneas will still be functional via the subprocess backend." -ForegroundColor DarkYellow
        }
    }

    foreach ($pkg in $packageInstallOrder) {
        Write-Host "  Installing $pkg..." -NoNewline -ForegroundColor Gray
        $pkgExtraArgs = @()
        if ($pkg -eq "aeneas") {
            $pkgExtraArgs += "--no-build-isolation"
            if ($aeneasNoCew) {
                # AENEAS_WITH_CEW is the variable aeneas setup.py checks (not FORCE);
                # setting it to False prevents the C extension from being compiled.
                $env:AENEAS_WITH_CEW = "False"
            }
        }
        $pkgExit = Invoke-PipInstall -PythonExe $PythonExe -Packages @($pkg) -ExtraArgs $pkgExtraArgs
        if ($pkg -eq "aeneas") {
            Remove-Item Env:\AENEAS_WITH_CEW -ErrorAction SilentlyContinue
        }
        if ($pkgExit -eq 0) {
            Write-Host " [ " -NoNewline -ForegroundColor White
            Write-Host "OK" -NoNewline -ForegroundColor Green
            Write-Host " ]" -ForegroundColor White
        } else {
            Write-Host " [ " -NoNewline -ForegroundColor White
            Write-Host " failed" -ForegroundColor Yellow
            Write-Host " ]" -ForegroundColor White
        }
    }

    $installedBackends = @()
    $failedBackends = @()
    $probePrevErrorAction = $ErrorActionPreference
    foreach ($def in $selectedDefs) {
        $importName = [string]$def.import
        $checkCode = "import warnings,importlib,sys; warnings.filterwarnings('ignore'); importlib.import_module('$importName'); print('ok')"
        $ErrorActionPreference = 'Continue'
        $probe = & $PythonExe -c $checkCode 2>$null
        $ErrorActionPreference = $probePrevErrorAction
        if ($LASTEXITCODE -eq 0 -and $probe -match "ok") {
            $installedBackends += [string]$def.key
        } else {
            $failedBackends += [string]$def.key
        }
    }
    $ErrorActionPreference = $probePrevErrorAction

    return @{
        installed = @($installedBackends | Select-Object -Unique)
        failed = @($failedBackends | Select-Object -Unique)
        attempted_packages = @($packageInstallOrder.ToArray())
    }
}

function Ensure-SpeechBrainAudioSupport {
    param([string]$PythonExe)

    Write-Host "Ensuring SpeechBrain audio backend dependencies (soundfile)..." -ForegroundColor White
    Invoke-PipInstall -PythonExe $PythonExe -Packages @("soundfile") -Upgrade | Out-Null

    $verifyCode = @"
import sys
import warnings

warnings.filterwarnings('ignore', message='.*SpeechBrain could not find any working torchaudio backend.*')
warnings.filterwarnings('ignore', module='speechbrain.utils.torch_audio_backend')

ok = True

try:
    import soundfile as sf
    _ = sf.available_formats()
except Exception as exc:
    print(f'SOUNDFILE_ERROR: {exc}')
    ok = False

try:
    from speechbrain.dataio import audio_io
except Exception as exc:
    # audio_io was removed in newer SpeechBrain versions; soundfile working above is sufficient.
    pass

sys.exit(0 if ok else 1)
"@

    $probePrevErrorAction = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $probe = & $PythonExe -c $verifyCode 2>&1
    $ErrorActionPreference = $probePrevErrorAction
    if ($LASTEXITCODE -eq 0) {
        Write-StatusOK "SpeechBrain audio backend (soundfile)"
        return $true
    }

    Write-Host "SpeechBrain installed, but soundfile-based audio loading is not fully ready." -ForegroundColor Yellow
    foreach ($line in $probe) {
        if ($line) {
            Write-Host "  $line" -ForegroundColor DarkYellow
        }
    }
    $ErrorActionPreference = $probePrevErrorAction
    return $false
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

    function Test-DotNetDesktopRuntime {
        # Check for .NET Desktop Runtime 6+ (required by HandBrake 1.6+; version 8 for 1.7+)
        if (Get-Command "dotnet" -ErrorAction SilentlyContinue) {
            try {
                $runtimes = & dotnet --list-runtimes 2>$null
                if ($runtimes -match "Microsoft\.WindowsDesktop\.App\s+[6-9]\.") {
                    return $true
                }
            } catch { }
        }
        # Registry fallback — .NET 5+ writes version keys here
        $regBase = "HKLM:\SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedfx\Microsoft.WindowsDesktop.App"
        if (Test-Path $regBase) {
            $versions = Get-ChildItem $regBase -ErrorAction SilentlyContinue
            foreach ($ver in $versions) {
                if ($ver.PSChildName -match "^[6-9]\.") {
                    return $true
                }
            }
        }
        return $false
    }

function Add-MsvcClToProcessPath {
    # Locates cl.exe inside the detected VS/Build Tools install and adds its directory
    # to the current process PATH so it is callable in this session.
    # Returns $true if cl.exe is now reachable.
    if (Get-Command "cl.exe" -ErrorAction SilentlyContinue) {
        return $true
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) { return $false }

    try {
        $installPath = [string](& $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null)
        if ([string]::IsNullOrWhiteSpace($installPath) -or -not (Test-Path -LiteralPath $installPath)) {
            return $false
        }

        $msvcRoot = Join-Path $installPath "VC\Tools\MSVC"
        if (-not (Test-Path -LiteralPath $msvcRoot)) { return $false }

        $versions = Get-ChildItem -Path $msvcRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
        foreach ($ver in $versions) {
            $clPath = Join-Path $ver.FullName "bin\HostX64\x64\cl.exe"
            if (Test-Path -LiteralPath $clPath) {
                $clDir = Split-Path -Parent $clPath
                if ($env:Path -notlike "*$clDir*") {
                    $env:Path = "$clDir;$env:Path"
                    Write-Host "Added MSVC compiler directory to session PATH: $clDir" -ForegroundColor DarkGray
                }
                return [bool](Get-Command "cl.exe" -ErrorAction SilentlyContinue)
            }
        }
    } catch {
        # Best effort only.
    }

    return $false
}

function Test-MsvcBuildTools {
    # aeneas requires a native C/C++ compiler on Windows.
    if (Get-Command "cl.exe" -ErrorAction SilentlyContinue) {
        return $true
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        try {
            $installPath = [string](& $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath 2>$null)
            # Only trust the detection if the installation directory actually exists on disk.
            # A stale registry entry (e.g. pointing to an offline drive) will make vswhere
            # report an installationPath that cannot be used.
            if (-not [string]::IsNullOrWhiteSpace($installPath) -and (Test-Path -LiteralPath $installPath)) {
                # Installation is on disk - try to make cl.exe reachable in this session.
                Add-MsvcClToProcessPath | Out-Null
                return $true
            }
        } catch {
            # Ignore detection errors and report as unavailable.
        }
    }

    return $false
}

function Get-VswherePath {
    $candidate = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $candidate) {
        return $candidate
    }
    return ""
}

function Get-VisualStudioInstalledInstanceCount {
    $vswhere = Get-VswherePath
    if (-not $vswhere) {
        return -1
    }

    try {
        $raw = & $vswhere -all -products * -format json 2>$null
        if (-not $raw) {
            return 0
        }

        $parsed = $raw | ConvertFrom-Json
        if ($parsed -is [array]) {
            return $parsed.Count
        }

        if ($null -ne $parsed) {
            return 1
        }
    } catch {
        return -1
    }

    return 0
}

function Get-VisualStudioInstallerFailureContext {
    $context = @{
        has_corrupt_state = $false
        saw_stale_cache_reference = $false
        log_files = @()
    }

    $candidateRoots = @()
    if ($script:ChocoCacheLocation) {
        $candidateRoots += [string]$script:ChocoCacheLocation
    }
    $candidateRoots += @(
        (Join-Path $PSScriptRoot ".pip-tmp\chocolatey"),
        (Join-Path $env:TEMP "chocolatey"),
        $env:TEMP
    )

    $logFiles = @()
    foreach ($root in $candidateRoots) {
        if (-not $root) { continue }
        if (-not (Test-Path -LiteralPath $root)) { continue }

        try {
            $logFiles += Get-ChildItem -Path $root -Filter "dd_*" -File -Recurse -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 8
        } catch {
            # Best effort only.
        }
    }

    $logFiles = @($logFiles | Sort-Object FullName -Unique)
    $context.log_files = @($logFiles | ForEach-Object { $_.FullName })

    foreach ($log in $logFiles) {
        try {
            $content = Get-Content -Path $log.FullName -Raw -ErrorAction SilentlyContinue
            if (-not $content) { continue }

            if ($content -match "Failed to deserialize instance state|Failed to read instance|No valid product defined") {
                $context.has_corrupt_state = $true
            }

            if ($content -match "Could not find file '.*installCache|_Instances\\[A-Za-z0-9]{8}") {
                $context.saw_stale_cache_reference = $true
            }
        } catch {
            # Best effort only.
        }
    }

    return $context
}

function Invoke-VisualStudioInstallerRecovery {
    $instanceCount = Get-VisualStudioInstalledInstanceCount
    if ($instanceCount -lt 0) {
        Write-Host "Could not determine existing Visual Studio instances; skipping automatic installer cleanup for safety." -ForegroundColor DarkYellow
        Write-Host "If Build Tools continues to fail, run Visual Studio cleanup manually: https://aka.ms/vs/cleanup" -ForegroundColor DarkYellow
        return $false
    }
    if ($instanceCount -gt 0) {
        Write-Host "Detected existing Visual Studio instance(s); skipping automatic installer cleanup to avoid impacting installed products." -ForegroundColor DarkYellow
        Write-Host "If Build Tools continues to fail, run Visual Studio cleanup manually: https://aka.ms/vs/cleanup" -ForegroundColor DarkYellow
        return $false
    }

    if (-not (Test-IsAdministrator)) {
        Write-Host "Skipping Visual Studio installer recovery because Administrator privileges are required." -ForegroundColor DarkYellow
        return $false
    }

    $cleanupExe = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\InstallCleanup.exe"
    if (-not (Test-Path -LiteralPath $cleanupExe)) {
        Write-Host "Visual Studio InstallCleanup.exe not found; cannot run automatic installer recovery." -ForegroundColor DarkYellow
        return $false
    }

    # InstallCleanup writes dd_cleanup logs into TEMP; ensure TEMP/TMP roots exist.
    foreach ($tmpVar in @("TEMP", "TMP")) {
        $tmpPath = [Environment]::GetEnvironmentVariable($tmpVar)
        if ($tmpPath -and -not (Test-Path -LiteralPath $tmpPath)) {
            try {
                New-Item -ItemType Directory -Path $tmpPath -Force | Out-Null
            } catch {
                # Best effort only.
            }
        }
    }

    $cleanupWorkingDirectory = [Environment]::GetEnvironmentVariable("TEMP")
    if (-not $cleanupWorkingDirectory -or -not (Test-Path -LiteralPath $cleanupWorkingDirectory)) {
        $cleanupWorkingDirectory = $PSScriptRoot
    }

    Write-Host "Attempting Visual Studio installer recovery for stale/corrupt state..." -ForegroundColor Yellow
    try {
        $process = Start-Process -FilePath $cleanupExe -ArgumentList @("-f") -WorkingDirectory $cleanupWorkingDirectory -Wait -NoNewWindow -PassThru
        if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
            Write-Host "Visual Studio installer recovery exited with code $($process.ExitCode)." -ForegroundColor DarkYellow
            return $false
        }
    } catch {
        Write-Host "Visual Studio installer recovery failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $false
    }

    # Clear stale package metadata that may keep pointing to removed cache roots.
    $stalePackageRoot = Join-Path $env:ProgramData "Microsoft\VisualStudio\Packages"
    foreach ($subPath in @("_Instances", "_Channels")) {
        $target = Join-Path $stalePackageRoot $subPath
        try {
            if (Test-Path -LiteralPath $target) {
                Remove-PathRobust -Path $target | Out-Null
            }
        } catch {
            # Best effort only.
        }
    }

    Write-Host "Visual Studio installer recovery completed." -ForegroundColor Green
    return $true
}

function Clear-VisualStudioInstallerCacheOverrides {
    # Remove stale policy/setup cache path overrides that can force VS installer to
    # use unavailable locations (for example, disconnected drives).
    $registryTargets = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\VisualStudio\Setup",
        "HKCU:\SOFTWARE\Policies\Microsoft\VisualStudio\Setup",
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\Setup",
        "HKCU:\SOFTWARE\Microsoft\VisualStudio\Setup"
    )

    foreach ($keyPath in $registryTargets) {
        if (-not (Test-Path -LiteralPath $keyPath)) { continue }
        try {
            $props = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue
            if ($null -ne $props -and ($props.PSObject.Properties.Name -contains "CachePath")) {
                Remove-ItemProperty -Path $keyPath -Name "CachePath" -ErrorAction SilentlyContinue
            }
        } catch {
            # Best effort only.
        }
    }
}

function Set-VisualStudioInstallerCachePath {
    param(
        [string]$CachePath
    )

    if (-not $CachePath) {
        return
    }

    if (-not (Test-Path -LiteralPath $CachePath)) {
        try {
            New-Item -ItemType Directory -Path $CachePath -Force | Out-Null
        } catch {
            return
        }
    }

    $targetKeys = @(
        "HKLM:\SOFTWARE\Microsoft\VisualStudio\Setup",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\Setup"
    )

    foreach ($keyPath in $targetKeys) {
        try {
            if (-not (Test-Path -LiteralPath $keyPath)) {
                New-Item -Path $keyPath -Force | Out-Null
            }
            New-ItemProperty -Path $keyPath -Name "CachePath" -Value $CachePath -PropertyType String -Force | Out-Null
        } catch {
            # Best effort only.
        }
    }
}

function Get-StableVsInstallerCachePath {
    $basePath = [string]$env:ProgramData
    if ([string]::IsNullOrWhiteSpace($basePath)) {
        $basePath = [string]$env:LOCALAPPDATA
    }
    if ([string]::IsNullOrWhiteSpace($basePath)) {
        $basePath = [string]$env:TEMP
    }
    if ([string]::IsNullOrWhiteSpace($basePath)) {
        $basePath = $PSScriptRoot
    }

    return (Join-Path $basePath "SubtitleTool\vs-installer-cache")
}

function Install-MsvcBuildToolsWithBootstrapper {
    param(
        [string]$CachePath
    )

    if (-not (Test-IsAdministrator)) {
        Write-Host "Skipping bootstrapper Build Tools install because installer is not running as Administrator." -ForegroundColor DarkYellow
        return $false
    }

    $bootstrapperUrl = "https://aka.ms/vs/17/release/vs_BuildTools.exe"
    $bootstrapperExe = Join-Path $env:TEMP "vs_BuildTools.exe"

    try {
        if (Test-Path -LiteralPath $bootstrapperExe) {
            Remove-Item -LiteralPath $bootstrapperExe -Force -ErrorAction SilentlyContinue
        }
        Invoke-WebRequest -Uri $bootstrapperUrl -OutFile $bootstrapperExe -UseBasicParsing
    } catch {
        Write-Host "Failed to download Visual Studio Build Tools bootstrapper: $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $false
    }

    $args = @(
        "--quiet",
        "--wait",
        "--norestart",
        "--nocache",
        "--add", "Microsoft.VisualStudio.Workload.VCTools",
        "--includeRecommended"
    )

    if ($CachePath) {
        $args += "--path", "cache=$CachePath"
    }

    try {
        Write-Host "Installing Microsoft C++ Build Tools via Visual Studio bootstrapper fallback..." -ForegroundColor Yellow
        $proc = Start-Process -FilePath $bootstrapperExe -ArgumentList $args -Wait -NoNewWindow -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            return $true
        }

        Write-Host "Visual Studio bootstrapper exited with code $($proc.ExitCode)." -ForegroundColor DarkYellow
        return $false
    } catch {
        Write-Host "Visual Studio bootstrapper install failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $false
    }
}

function Install-MsvcBuildTools {
    param(
        [ValidateSet("auto", "winget", "choco")]
        [string]$Method = "auto"
    )

    $attempted = @()
    $usedMethod = ""
    $success = $false
    $methodsToTry = if ($Method -eq "auto") { @("winget", "choco", "bootstrapper") } else { @($Method) }
    if ($methodsToTry -contains "auto") {
        $methodsToTry = @("winget", "choco", "bootstrapper")
    }
    $stableVsCache = Get-StableVsInstallerCachePath

    if (-not (Test-Path -LiteralPath $stableVsCache)) {
        New-Item -ItemType Directory -Path $stableVsCache -Force | Out-Null
    }

    for ($pass = 1; $pass -le 2 -and -not $success; $pass++) {
        Clear-VisualStudioInstallerCacheOverrides
        Set-VisualStudioInstallerCachePath -CachePath $stableVsCache

        if ($pass -eq 2) {
            if ($Method -ne "auto") { break }
            $failureContext = Get-VisualStudioInstallerFailureContext
            if (-not ($failureContext.has_corrupt_state -or $failureContext.saw_stale_cache_reference)) {
                break
            }

            Write-Host "Detected Visual Studio installer state/cache corruption from recent logs; running one-time recovery and retry." -ForegroundColor DarkYellow
            if (-not (Invoke-VisualStudioInstallerRecovery)) {
                break
            }
            $attempted += "vs_installer_recovery"
        }

        foreach ($candidateMethod in $methodsToTry) {
            if ($success) { break }

            if ($candidateMethod -eq "winget") {
                if (-not (Test-CommandAvailable "winget")) { continue }
                $attempted += "winget"
                if ($script:WingetInstallLocation) {
                    Write-Host "Ignoring custom winget install location for Build Tools (best reliability path)." -ForegroundColor DarkYellow
                }
                Write-Host "Installing Microsoft C++ Build Tools via winget (this can take several minutes)..." -ForegroundColor Yellow
                try {
                    $overrideArgs = '"--quiet --wait --norestart --nocache --add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"'
                    $wingetArgs = @(
                        "install", "--id", "Microsoft.VisualStudio.2022.BuildTools", "--exact"
                    ) + (Get-WingetInstallBaseArgs) + @("--override", $overrideArgs)
                    if (Invoke-WingetInstallCommand -CommandArgs $wingetArgs -AllowRetryWithoutLocation -DisableCustomLocation) {
                        $usedMethod = "winget"
                        $success = $true
                        break
                    }
                } catch {
                    # Try next method.
                }
            } elseif ($candidateMethod -eq "choco") {
                if (-not (Test-CommandAvailable "choco")) { continue }
                if (-not (Test-IsAdministrator)) {
                    Write-Host "Skipping Chocolatey Build Tools install because installer is not running as Administrator." -ForegroundColor DarkYellow
                    continue
                }
                $attempted += "choco"
                if ($script:ChocoCacheLocation) {
                    Write-Host "Ignoring custom Chocolatey cache location for Build Tools (best reliability path)." -ForegroundColor DarkYellow
                }
                Write-Host "Installing Microsoft C++ Build Tools via Chocolatey (this can take several minutes)..." -ForegroundColor Yellow
                try {
                    $chocoPackageArgs = "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended --passive --norestart"
                    $chocoPackageArgsQuoted = '"' + $chocoPackageArgs + '"'
                    $chocoArgs = @("install", "visualstudio2022buildtools") + (Get-ChocoCommonArgs -DisableCustomCache) + @("--package-parameters", $chocoPackageArgsQuoted)
                    $process = Start-Process choco -ArgumentList $chocoArgs -Wait -NoNewWindow -PassThru
                    if ($process.ExitCode -eq 0) {
                        $usedMethod = "choco"
                        $success = $true
                        break
                    }
                } catch {
                    # Try next method.
                }
            } elseif ($candidateMethod -eq "bootstrapper") {
                $attempted += "bootstrapper"
                if (Install-MsvcBuildToolsWithBootstrapper -CachePath $stableVsCache) {
                    $usedMethod = "bootstrapper"
                    $success = $true
                    break
                }
            }
        }
    }

    if ($success) {
        Refresh-ProcessPathFromRegistry
        Start-Sleep -Seconds 2
        $success = Test-MsvcBuildTools
    }

    return @{
        success = $success
        install_method = $usedMethod
        attempted = $attempted
    }
}

function Install-VCRedistWithWinget {
    Write-Host "Installing Visual C++ Redistributable via winget..." -ForegroundColor Yellow
    $args = @("install", "--id", "Microsoft.VCRedist.2015+.x64", "--exact") + (Get-WingetInstallBaseArgs)
    if (Invoke-WingetInstallCommand -CommandArgs $args -AllowRetryWithoutLocation) {
        Write-Host "Visual C++ Redistributable ready via winget." -ForegroundColor Green
        return $true
    }
    return $false
}

function Install-VCRedistWithChoco {
    Write-Host "Installing Visual C++ Redistributable via chocolatey..." -ForegroundColor Yellow
    try {
        $chocoArgs = @("install", "vcredist-all") + (Get-ChocoCommonArgs)
        $process = Start-Process choco -ArgumentList $chocoArgs -Wait -NoNewWindow -PassThru
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
    $usedMethod = ""
    $attempted = @()
    
    # Try winget
    if (Test-CommandAvailable "winget") {
        $attempted += "winget"
        if (Install-VCRedistWithWinget) {
            $installed = $true
            $usedMethod = "winget"
        }
    }
    
    # Try chocolatey if winget failed
    if (-not $installed -and (Test-CommandAvailable "choco")) {
        $attempted += "choco"
        if (Install-VCRedistWithChoco) {
            $installed = $true
            $usedMethod = "choco"
        }
    }
    
    # Fall back to manual installation
    if (-not $installed) {
        $attempted += "manual"
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
            $usedMethod = "manual"
        } catch {
            Write-Host "Failed to install VC++ Redistributable: $($_.Exception.Message)" -ForegroundColor Red
            return @{
                success = $false
                install_method = ""
                attempted = @($attempted)
            }
        }
    }
    
    # Verify installation by checking registry again
    Start-Sleep -Seconds 1
    if (-not (Test-VCRedist)) {
        Write-Host "Warning: VC++ installation verification failed. May need manual restart." -ForegroundColor Yellow
        return @{
            success = $false
            install_method = $usedMethod
            attempted = @($attempted)
        }
    }
    
    return @{
        success = [bool]$installed
        install_method = $usedMethod
        attempted = @($attempted)
    }
}

    function Install-DotNetRuntime {
        # Installs .NET Desktop Runtime 8 (required by HandBrake 1.7+)
        $installed = $false
        $usedMethod = ""
        $attempted = @()

        if (Test-CommandAvailable "winget") {
            $attempted += "winget"
            Write-Host "Installing .NET Desktop Runtime 8 via winget..." -ForegroundColor Yellow
            $wArgs = @("install", "--id", "Microsoft.DotNet.DesktopRuntime.8", "--exact") + (Get-WingetInstallBaseArgs)
            if (Invoke-WingetInstallCommand -CommandArgs $wArgs -AllowRetryWithoutLocation) {
                $installed = $true
                $usedMethod = "winget"
            }
        }

        if (-not $installed -and (Test-CommandAvailable "choco")) {
            $attempted += "choco"
            Write-Host "Installing .NET Desktop Runtime via Chocolatey..." -ForegroundColor Yellow
            try {
                $cArgs = @("install", "dotnet-desktopruntime", "-y") + (Get-ChocoCommonArgs)
                $proc = Start-Process choco -ArgumentList $cArgs -Wait -NoNewWindow -PassThru
                if ($proc.ExitCode -eq 0) {
                    $installed = $true
                    $usedMethod = "choco"
                }
            } catch { }
        }

        if (-not $installed) {
            $attempted += "manual"
            Write-Host "Downloading .NET Desktop Runtime 8 manually..." -ForegroundColor Yellow
            $dotnetUrl = "https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe"
            $dotnetInst = Join-Path $env:TEMP "dotnet-desktop-runtime-8-x64.exe"
            try {
                if (Test-Path $dotnetInst) { Remove-Item $dotnetInst -Force -ErrorAction SilentlyContinue }
                Invoke-WebRequest -Uri $dotnetUrl -OutFile $dotnetInst -UseBasicParsing
                $proc = Start-Process -FilePath $dotnetInst -ArgumentList "/install", "/quiet", "/norestart" -Wait -PassThru
                Start-Sleep -Seconds 2
                Remove-Item $dotnetInst -ErrorAction SilentlyContinue
                if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
                    $installed = $true
                    $usedMethod = "manual"
                }
            } catch {
                Write-Host "Failed to install .NET Desktop Runtime: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        return @{
            success        = [bool]$installed
            install_method = $usedMethod
            attempted      = @($attempted)
        }
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

function Should-RequestElevation {
    if (-not $script:AutoElevationEnabled) {
        return $false
    }

    if (Test-IsAdministrator) {
        return $false
    }

    if ($script:WingetInstallScope -eq "machine") {
        return $true
    }

    if ($PythonInstallMethod -eq "choco" -or $FfmpegInstallMethod -eq "choco" -or $ToolInstallMethod -eq "choco") {
        return $true
    }

    if ($InstallAiAeneas) {
        return $true
    }

    return $false
}

function Request-ElevationAndRelaunch {
    if (-not (Should-RequestElevation)) {
        return
    }

    Write-Host "Administrator privileges are recommended for selected install options." -ForegroundColor Yellow
    $answer = Read-Host "Relaunch installer as Administrator now? [Y/N] (default: Y)"
    if ([string]::IsNullOrWhiteSpace("$answer")) { $answer = "Y" }
    if ($answer -notmatch '^[Yy]') {
        Write-Host "Continuing without elevation. Some installs may be skipped or fail." -ForegroundColor Yellow
        return
    }

    $relaunchArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath) + (Build-ScriptRelaunchArgs)
    try {
        Start-Process -FilePath "powershell.exe" -ArgumentList $relaunchArgs -Verb RunAs | Out-Null
        exit 0
    } catch {
        Write-Host "Elevation request was cancelled or failed. Continuing without elevation." -ForegroundColor Yellow
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
    $wingetArgs = @("install", "--id", "Python.Python.3.11", "--exact") + (Get-WingetInstallBaseArgs)
    if ($script:WingetInstallScope -eq "default") {
        $wingetArgs += "--scope", "user"
    }
    $ok = Invoke-WingetInstallCommand -CommandArgs $wingetArgs -AllowRetryWithoutLocation
    if (-not $ok) {
        throw "winget Python install failed."
    }
}

function Install-PythonWithChoco {
    Write-Host "Installing Python 3.11 using Chocolatey..."
    $chocoArgs = @("install", "python311") + (Get-ChocoCommonArgs)
    $process = Start-Process choco -ArgumentList $chocoArgs -Wait -NoNewWindow -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Chocolatey Python install failed."
    }
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
    $usedMethod = ""

    if ($Method -eq "winget") {
        if (-not (Test-CommandAvailable "winget")) { throw "winget not found." }
        Install-PythonWithWinget
        $attempted += "winget"
        $usedMethod = "winget"
    } elseif ($Method -eq "choco") {
        if (-not (Test-CommandAvailable "choco")) { throw "choco not found." }
        Install-PythonWithChoco
        $attempted += "choco"
        $usedMethod = "choco"
    } elseif ($Method -eq "scoop") {
        if (-not (Test-CommandAvailable "scoop")) { throw "scoop not found." }
        Install-PythonWithScoop
        $attempted += "scoop"
        $usedMethod = "scoop"
    } else {
        if (Test-CommandAvailable "winget") {
            try {
                Install-PythonWithWinget
                $attempted += "winget"

                $resolvedAfterWinget = Wait-ForPythonCommand -Seconds 30
                if ($resolvedAfterWinget) {
                    $usedMethod = "winget"
                    return @{
                        success = $true
                        install_method = $usedMethod
                        attempted = @($attempted)
                    }
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
                    $usedMethod = "manual"
                    return @{
                        success = $true
                        install_method = $usedMethod
                        attempted = @($attempted)
                    }
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
                if (Find-PythonCommand) {
                    $usedMethod = "choco"
                }
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

    return @{
        success = $true
        install_method = $usedMethod
        attempted = @($attempted)
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

# ---------------------------------------------------------------------------
# UV - fast Python package installer (https://github.com/astral-sh/uv)
# Install-UV:   downloads and installs UV if not already present; returns
#               the path to the uv executable, or "" on failure.
# Invoke-PipInstall: wraps uv/pip installs so the rest of the script
#               doesn't care which backend is in use.
# ---------------------------------------------------------------------------
function Install-UV {
    # Step 1: already on PATH
    $uvCmd = Get-Command "uv" -ErrorAction SilentlyContinue
    if ($uvCmd) { return $uvCmd.Source }

    # Step 2: not on PATH yet but binary exists at standard install locations
    $candidates = @(
        (Join-Path $env:USERPROFILE ".local\bin\uv.exe"),
        (Join-Path $env:USERPROFILE ".cargo\bin\uv.exe")
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) {
            Add-UserPathEntry -PathEntry (Split-Path $c -Parent) | Out-Null
            Refresh-ProcessPathFromRegistry
            return $c
        }
    }

    # Step 3: download and run the official PowerShell installer
    try {
        $installScript = Invoke-RestMethod "https://astral.sh/uv/install.ps1" `
            -UseBasicParsing -TimeoutSec 60
        # Redirect all streams so UV's own installer messages don't clutter output
        & ([scriptblock]::Create($installScript)) *>&1 | Out-Null
        Refresh-ProcessPathFromRegistry
    } catch {
        return ""
    }

    # Step 4: re-check after install
    $uvCmd = Get-Command "uv" -ErrorAction SilentlyContinue
    if ($uvCmd) { return $uvCmd.Source }

    foreach ($c in $candidates) {
        if (Test-Path $c) {
            Add-UserPathEntry -PathEntry (Split-Path $c -Parent) | Out-Null
            Refresh-ProcessPathFromRegistry
            return $c
        }
    }

    return ""
}

# Invoke-PipInstall: install packages via uv (preferred) or pip (fallback).
# When $script:QuietOutput is true, all install noise is suppressed.
function Invoke-PipInstall {
    param(
        [string]$PythonExe,
        [string[]]$Packages = @(),
        [string]$RequirementsFile = "",
        [switch]$Upgrade,
        [string[]]$ExtraArgs = @()
    )

    $pkgArgs = @()
    if ($Upgrade) { $pkgArgs += "--upgrade" }
    $pkgArgs += $ExtraArgs
    if ($RequirementsFile) {
        $pkgArgs += "-r", $RequirementsFile
    } else {
        $pkgArgs += $Packages
    }

    $showOutput = -not $script:QuietOutput
    $backend = [string]$script:PythonPackageInstallBackend
    if (-not $backend) { $backend = "auto" }
    $nativePrevErrorAction = $ErrorActionPreference
    $uvOnPath = [bool](Get-Command "uv" -ErrorAction SilentlyContinue)
    $uvReady = [bool](($script:UvExe -and (Test-Path $script:UvExe)) -or $uvOnPath)
    $script:LastPackageInstallBackendUsed = ""
    $script:LastPackageInstallError = ""

    if ($backend -eq "uv" -and -not $uvReady) {
        $script:LastPackageInstallBackendUsed = "uv"
        $script:LastPackageInstallError = "UV backend selected, but uv executable is unavailable."
        if ($showOutput) {
            Write-Host "    UV backend selected, but uv executable is unavailable." -ForegroundColor Yellow
        }
        return 1
    }

    $useUv = $false
    if ($backend -eq "uv") {
        $useUv = $true
    } elseif ($backend -eq "auto" -and $uvReady) {
        $useUv = $true
    }

    if ($useUv) {
        $cmdArgs = @("pip", "install", "--python", $PythonExe) + $pkgArgs
        $script:LastPackageInstallBackendUsed = if ($backend -eq "uv") { "uv" } else { "uv(auto)" }

        $uvCandidates = New-Object System.Collections.Generic.List[string]
        if ($script:UvExe -and (Test-Path $script:UvExe)) {
            $uvCandidates.Add([string]$script:UvExe)
        }
        $uvCmdObj = Get-Command "uv" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($uvCmdObj) {
            $uvPathCandidate = [string]$(if ($uvCmdObj.Source) { $uvCmdObj.Source } elseif ($uvCmdObj.Path) { $uvCmdObj.Path } else { "uv" })
            $already = $false
            foreach ($candidate in $uvCandidates) {
                if ([string]::Equals($candidate, $uvPathCandidate, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $already = $true
                    break
                }
            }
            if (-not $already) {
                $uvCandidates.Add($uvPathCandidate)
            }
        }
        if ($uvCandidates.Count -eq 0) {
            $uvCandidates.Add("uv")
        }

        $uvOutput = @()
        $uvExitCode = 1
        $uvSucceeded = $false
        $uvExceptions = New-Object System.Collections.Generic.List[string]
        Write-VerboseLog -Lines @("CMD: $uvBin $($cmdArgs -join ' ')") -Prefix " [uv-cmd] "

        foreach ($uvBin in $uvCandidates) {
            try {
                if ($showOutput) {
                    $ErrorActionPreference = 'Continue'
                    $uvOutput = @(& $uvBin @cmdArgs 2>&1)
                    $ErrorActionPreference = $nativePrevErrorAction
                    $firstLine = $true
                    foreach ($line in $uvOutput) {
                        $lineStr = [string]$line
                        if ($lineStr -match "Resolved|Prepared|Installed|Audited|Downloading|warning:|error") {
                            if ($firstLine) { Write-Host ""; $firstLine = $false }
                            Write-Host "    $lineStr" -ForegroundColor DarkGray
                        }
                    }
                } else {
                    $ErrorActionPreference = 'Continue'
                    $uvOutput = @(& $uvBin @cmdArgs 2>&1)
                    $ErrorActionPreference = $nativePrevErrorAction
                }

                $uvExitCode = $LASTEXITCODE
                Write-VerboseLog -Lines $uvOutput -Prefix " [uv] "
                if ($uvExitCode -eq 0) {
                    $uvSucceeded = $true
                    break
                }
            } catch {
                $ErrorActionPreference = $nativePrevErrorAction
                $msg = "[$uvBin] $($_.Exception.Message)"
                $uvExceptions.Add($msg)
                if ($showOutput) {
                    Write-Host "    UV invocation exception: $msg" -ForegroundColor DarkYellow
                }
                $uvExitCode = 1
            }
        }

        if ($uvSucceeded) {
            $script:LastPackageInstallError = ""
            return 0
        }

        $uvTail = ($uvOutput | Select-Object -Last 12) -join [Environment]::NewLine
        $exceptionsText = if ($uvExceptions.Count -gt 0) { ($uvExceptions -join " | ") } else { "none" }
        $script:LastPackageInstallError = "UV failed (exit $uvExitCode). exceptions=$exceptionsText. $uvTail"

        if ($showOutput) {
            Write-Host "    UV install failed (exit $uvExitCode)." -ForegroundColor Yellow
            if ($uvExceptions.Count -gt 0) {
                Write-Host "    UV invocation exceptions: $exceptionsText" -ForegroundColor DarkYellow
            }
            $tail = $uvOutput | Select-Object -Last 12
            foreach ($line in $tail) {
                if ($line) {
                    Write-Host "    $line" -ForegroundColor DarkYellow
                }
            }
        }

        # In auto mode, retry with pip before giving up.
        if ($backend -eq "auto") {
            if ($showOutput) {
                Write-Host "    Retrying with pip fallback..." -ForegroundColor Yellow
            }
            $script:LastPackageInstallBackendUsed = "pip(fallback)"
        } else {
            return $uvExitCode
        }
    }

    $pipArgs = @("-m", "pip", "install") + $pkgArgs
    $script:LastPackageInstallBackendUsed = if ($script:LastPackageInstallBackendUsed) { $script:LastPackageInstallBackendUsed } else { "pip" }
    Write-VerboseLog -Lines @("CMD: $PythonExe $($pipArgs -join ' ')") -Prefix " [pip-cmd] "
    if ($showOutput) {
        $firstLine = $true
        $ErrorActionPreference = 'Continue'
        $pipOutput = @(& $PythonExe @pipArgs 2>&1)
        $ErrorActionPreference = $nativePrevErrorAction
        Write-VerboseLog -Lines $pipOutput -Prefix " [pip] "
        foreach ($line in $pipOutput) {
            $lineStr = [string]$line
            if ($lineStr -match "Successfully installed|Requirement already satisfied|Collecting|Downloading|ERROR:") {
                if ($firstLine) { Write-Host ""; $firstLine = $false }
                Write-Host "    $lineStr" -ForegroundColor DarkGray
            }
        }
        if ($LASTEXITCODE -ne 0) {
            $pipTail = ($pipOutput | Select-Object -Last 12) -join [Environment]::NewLine
            $script:LastPackageInstallError = "PIP failed (exit $LASTEXITCODE). $pipTail"
        } else {
            $script:LastPackageInstallError = ""
        }
    } else {
        $ErrorActionPreference = 'Continue'
        $pipOutput = @(& $PythonExe @pipArgs 2>&1)
        $ErrorActionPreference = $nativePrevErrorAction
        Write-VerboseLog -Lines $pipOutput -Prefix " [pip] "
        if ($LASTEXITCODE -ne 0) {
            $pipTail = ($pipOutput | Select-Object -Last 12) -join [Environment]::NewLine
            $script:LastPackageInstallError = "PIP failed (exit $LASTEXITCODE). $pipTail"
        } else {
            $script:LastPackageInstallError = ""
        }
    }

    $ErrorActionPreference = $nativePrevErrorAction

    return $LASTEXITCODE
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$requirementsPath = Join-Path $scriptDir "requirements.txt"
$requirementsAIPath = Join-Path $scriptDir "requirements_ai.txt"
$ffmpegInstallerPath = Join-Path $scriptDir "install_ffmpeg_windows.ps1"
$subtitleToolPath = Join-Path $scriptDir "subtitle_tool.py"
$venvPath = Join-Path $scriptDir "venv"
$manifestPath = Join-Path $scriptDir ".install_manifest.json"
$installLogPath = Join-Path $scriptDir "install_all_windows.log"

Set-Location $scriptDir

# Start transcript so every line of output is captured to a log file.
try {
    Start-Transcript -Path $installLogPath -Append -Force | Out-Null
} catch {
    Write-Host "Note: could not start install transcript: $($_.Exception.Message)" -ForegroundColor DarkYellow
}

# ---------------------------------------------------------------------------
# Verbose detail log — collects all raw subprocess output and timestamped
# phase banners that would otherwise be absent from the transcript.
# ---------------------------------------------------------------------------
$verboseLogPath = Join-Path $scriptDir "install_verbose.log"
$script:verboseLogPath = $verboseLogPath
try {
    $vLogHeader = @(
        "=================================================================",
        "  Subtitle Tool Installer — Verbose Log",
        "  Date    : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "  Script  : $($MyInvocation.MyCommand.Path)",
        "  User    : $([System.Environment]::UserName)  Machine: $([System.Environment]::MachineName)",
        "  OS      : $([System.Environment]::OSVersion.VersionString)",
        "  PS      : $($PSVersionTable.PSVersion)",
        "=================================================================",
        ""
    )
    Set-Content -Path $verboseLogPath -Value ($vLogHeader -join "`n") -Encoding UTF8 -ErrorAction Stop
} catch {
    Write-Host "Note: could not initialise verbose log: $($_.Exception.Message)" -ForegroundColor DarkYellow
    $script:verboseLogPath = ""
}

# Environment snapshot written immediately after log is created
Write-VerboseLogBanner "Environment"
Write-VerboseLog -Lines @(
    "Script dir  : $scriptDir",
    "Script path : $($MyInvocation.MyCommand.Path)",
    "Working dir : $(Get-Location)",
    "Transcript  : $installLogPath",
    "Verbose log : $verboseLogPath",
    "Script drive: $(Split-Path -Qualifier $scriptDir)  System drive: $([System.Environment]::GetEnvironmentVariable('SystemDrive'))",
    "Manifest    : $manifestPath",
    "Venv path   : $venvPath",
    "App script  : $subtitleToolPath",
    "PS edition  : $($PSVersionTable.PSEdition)  PS version: $($PSVersionTable.PSVersion)",
    "OS          : $([System.Environment]::OSVersion.VersionString)",
    "CPU cores   : $([System.Environment]::ProcessorCount)",
    "User        : $([System.Environment]::UserName)  Admin: $(([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))",
    "Python cmd  : $((Get-Command 'python' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue))",
    "Py launcher : $((Get-Command 'py' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue))",
    "TMPDIR      : $env:TMPDIR",
    "TEMP        : $env:TEMP",
    "TMP         : $env:TMP",
    "PIP_CACHE   : $env:PIP_CACHE_DIR",
    "PATH(process): $env:Path",
    "PATH(machine): $([System.Environment]::GetEnvironmentVariable('Path', 'Machine'))",
    "PATH(user)   : $([System.Environment]::GetEnvironmentVariable('Path', 'User'))",
    ""
) -Prefix "  "

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

Show-InteractiveInstallerMenu

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
        $aiPackages = @(
            "torch", "torchvision", "torchaudio",
            "openai-whisper", "whisper", "pysubs2",
            "faster-whisper", "whisperx", "stable-ts", "whisper-timestamped",
            "speechbrain", "vosk", "aeneas",
            "cinemagoer", "imdbpy"
        )

        Write-Host "Uninstalling AI packages from venv..." -ForegroundColor White
        foreach ($pkg in $aiPackages) {
            Write-Host "  Removing $pkg..." -NoNewline -ForegroundColor Gray
            $result = & $venvPy -m pip uninstall -y $pkg 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host " [ " -NoNewline -ForegroundColor White
                Write-Host "OK" -NoNewline -ForegroundColor Green
                Write-Host " ]" -ForegroundColor White
            } else {
                # pip exits non-zero when the package isn't installed - that's fine.
                Write-Host " (not installed, skipping)" -ForegroundColor DarkGray
            }
        }

        # Also remove stale torch-related packages that may linger after uninstall
        Write-Host "  Cleaning up torch dependencies..." -NoNewline -ForegroundColor Gray
        & $venvPy -m pip uninstall -y filelock sympy networkx jinja2-markupsafe 2>&1 | Out-Null
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White

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
    $mMkvtoolnix = if ($hasMfst) { Get-MfstProp $m "mkvtoolnix" } else { $null }
    $mHandBrake = if ($hasMfst) { Get-MfstProp $m "handbrake" } else { $null }
    $mMakeMKV = if ($hasMfst) { Get-MfstProp $m "makemkv" } else { $null }
        $mDotNetRuntime = if ($hasMfst) { Get-MfstProp $m "dotnet_runtime" } else { $null }

    $pyPreExisted  = if ($mPython)   { [bool](Get-MfstProp $mPython   "pre_existed" $true) } else { $true }
    $pyMethod      = if ($mPython)   { [string](Get-MfstProp $mPython   "install_method" "") } else { "" }
    $vcPreExisted  = if ($mVCRedist) { [bool](Get-MfstProp $mVCRedist "pre_existed" $true) } else { $true }
    $vcMethod      = if ($mVCRedist) { [string](Get-MfstProp $mVCRedist "install_method" "") } else { "" }
    $ffPreExisted  = if ($mFfmpeg)   { [bool](Get-MfstProp $mFfmpeg   "pre_existed" $true) } else { $true }
    $ffMethod      = if ($mFfmpeg)   { [string](Get-MfstProp $mFfmpeg   "install_method" "") } else { "" }
    $mkvPreExisted = if ($mMkvtoolnix) { [bool](Get-MfstProp $mMkvtoolnix "pre_existed" $true) } else { $true }
    $mkvMethod = if ($mMkvtoolnix) { [string](Get-MfstProp $mMkvtoolnix "install_method" "") } else { "" }
    $hbPreExisted = if ($mHandBrake) { [bool](Get-MfstProp $mHandBrake "pre_existed" $true) } else { $true }
    $hbMethod = if ($mHandBrake) { [string](Get-MfstProp $mHandBrake "install_method" "") } else { "" }
    $mmPreExisted = if ($mMakeMKV) { [bool](Get-MfstProp $mMakeMKV "pre_existed" $true) } else { $true }
    $mmMethod = if ($mMakeMKV) { [string](Get-MfstProp $mMakeMKV "install_method" "") } else { "" }
        $dotnetPreExisted = if ($mDotNetRuntime) { [bool](Get-MfstProp $mDotNetRuntime "pre_existed" $true) } else { $true }
        $dotnetMethod = if ($mDotNetRuntime) { [string](Get-MfstProp $mDotNetRuntime "install_method" "") } else { "" }

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
        if (-not $mkvPreExisted -and $mkvMethod) {
            Write-Host "  - MKVToolNix     will be uninstalled via $mkvMethod" -ForegroundColor Cyan
        } else {
            Write-Host "  - MKVToolNix     was already installed before this script - will NOT be touched" -ForegroundColor Gray
        }
        if (-not $hbPreExisted -and $hbMethod) {
            Write-Host "  - HandBrakeCLI   will be uninstalled via $hbMethod" -ForegroundColor Cyan
        } else {
            Write-Host "  - HandBrakeCLI   was already installed before this script - will NOT be touched" -ForegroundColor Gray
        }
        if (-not $mmPreExisted -and $mmMethod) {
            Write-Host "  - MakeMKV        will be uninstalled via $mmMethod" -ForegroundColor Cyan
        } else {
            Write-Host "  - MakeMKV        was already installed before this script - will NOT be touched" -ForegroundColor Gray
        }
        if (-not $dotnetPreExisted -and $dotnetMethod) {
            Write-Host "  - .NET Runtime   will be uninstalled via $dotnetMethod" -ForegroundColor Cyan
        } else {
            Write-Host "  - .NET Runtime   was already installed before this script - will NOT be touched" -ForegroundColor Gray
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

        # --- MKVToolNix ---
        if (-not $mkvPreExisted -and $mkvMethod) {
            Write-Host ""
            Write-Host "Uninstalling MKVToolNix via $mkvMethod..." -ForegroundColor White
            switch ($mkvMethod) {
                "winget" {
                    winget uninstall --id MoritzBunkus.MKVToolNix --silent 2>&1 | Out-Null
                    Write-Host "  MKVToolNix uninstall attempted." -ForegroundColor Green
                }
                "choco" {
                    choco uninstall mkvtoolnix -y 2>&1 | Out-Null
                    Write-Host "  MKVToolNix uninstalled via Chocolatey." -ForegroundColor Green
                }
                "scoop" {
                    scoop uninstall mkvtoolnix 2>&1 | Out-Null
                    Write-Host "  MKVToolNix uninstalled via Scoop." -ForegroundColor Green
                }
                default {
                    Write-Host "  Install method '$mkvMethod' - please remove MKVToolNix manually." -ForegroundColor Yellow
                }
            }
        }

        # --- HandBrake ---
        if (-not $hbPreExisted -and $hbMethod) {
            Write-Host ""
            Write-Host "Uninstalling HandBrake via $hbMethod..." -ForegroundColor White
            switch ($hbMethod) {
                "winget" {
                    winget uninstall --id HandBrake.HandBrakeCLI --silent 2>&1 | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        # Legacy fallback in case old manifest tracked GUI package.
                        winget uninstall --id HandBrake.HandBrake --silent 2>&1 | Out-Null
                    }
                    Write-Host "  HandBrake uninstall attempted." -ForegroundColor Green
                }
                "choco" {
                    choco uninstall handbrake-cli -y 2>&1 | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        # Legacy fallback in case old manifest tracked GUI package.
                        choco uninstall handbrake -y 2>&1 | Out-Null
                    }
                    Write-Host "  HandBrake uninstalled via Chocolatey." -ForegroundColor Green
                }
                "scoop" {
                    scoop uninstall handbrake-cli 2>&1 | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        # Legacy fallback in case old manifest tracked GUI package.
                        scoop uninstall handbrake 2>&1 | Out-Null
                    }
                    Write-Host "  HandBrake uninstalled via Scoop." -ForegroundColor Green
                }
                default {
                    Write-Host "  Install method '$hbMethod' - please remove HandBrake manually." -ForegroundColor Yellow
                }
            }
        }

        # --- MakeMKV ---
        if (-not $mmPreExisted -and $mmMethod) {
            Write-Host ""
            Write-Host "Uninstalling MakeMKV via $mmMethod..." -ForegroundColor White
            switch ($mmMethod) {
                "winget" {
                    winget uninstall --id GuinpinSoft.MakeMKV --silent 2>&1 | Out-Null
                    Write-Host "  MakeMKV uninstall attempted." -ForegroundColor Green
                }
                "choco" {
                    choco uninstall makemkv -y 2>&1 | Out-Null
                    Write-Host "  MakeMKV uninstalled via Chocolatey." -ForegroundColor Green
                }
                "scoop" {
                    scoop uninstall makemkv 2>&1 | Out-Null
                    Write-Host "  MakeMKV uninstalled via Scoop." -ForegroundColor Green
                }
                default {
                    Write-Host "  Install method '$mmMethod' - please remove MakeMKV manually." -ForegroundColor Yellow
                }
            }
        }

        # --- .NET Desktop Runtime ---
        if (-not $dotnetPreExisted -and $dotnetMethod) {
            Write-Host ""
            Write-Host "Uninstalling .NET Desktop Runtime via $dotnetMethod..." -ForegroundColor White
            switch ($dotnetMethod) {
                "winget" {
                    winget uninstall --id Microsoft.DotNet.DesktopRuntime.8 --silent 2>&1 | Out-Null
                    Write-Host "  .NET Desktop Runtime uninstalled." -ForegroundColor Green
                }
                "choco" {
                    choco uninstall dotnet-desktopruntime -y 2>&1 | Out-Null
                    Write-Host "  .NET Desktop Runtime uninstalled via Chocolatey." -ForegroundColor Green
                }
                default {
                    Write-Host "  Install method '$dotnetMethod' - please remove .NET Desktop Runtime manually via Apps & Features." -ForegroundColor Yellow
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

Request-ElevationAndRelaunch

# Ensure we have a package manager
Ensure-PackageManager

# Log available package managers
Write-VerboseLogBanner "Package manager availability"
Write-VerboseLog -Lines @(
    "winget : $(if (Test-CommandAvailable 'winget') { $((& winget --version 2>$null) -replace '\s','') } else { 'not found' })",
    "choco  : $(if (Test-CommandAvailable 'choco') { $((& choco --version 2>$null | Select-Object -First 1) -replace '\s','') } else { 'not found' })",
    "scoop  : $(if (Test-CommandAvailable 'scoop') { 'found' } else { 'not found' })"
) -Prefix "  "

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
$script:MkvtoolnixPreExisted = $true
$script:MkvtoolnixInstallUsed = ""
$script:HandBrakePreExisted = $true
$script:HandBrakeInstallUsed = ""
$script:MakeMKVPreExisted = $true
$script:MakeMKVInstallUsed = ""
$script:DotNetRuntimePreExisted  = $true
$script:DotNetRuntimeInstallUsed = ""

Write-VerboseLogBanner "Python detection"
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
    $pythonInstallResult = Install-Python -Method $PythonInstallMethod
    $script:PythonInstallUsed = [string]$pythonInstallResult.install_method
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
# Log Python detection result
$_pyDetectedVersion = Get-PythonCommandVersion $pythonCmd 2>$null
Write-VerboseLog -Lines @(
    "Status  : $(if ($script:PythonPreExisted) { 'pre-existed' } else { 'installed by script (method: ' + $script:PythonInstallUsed + ')' })",
    "Command : $pythonCmd",
    "Version : $_pyDetectedVersion"
) -Prefix "  "

# Refresh PATH in current session in case installer changed machine/user PATH.
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
            [System.Environment]::GetEnvironmentVariable("Path", "User")

# Check if virtual environment exists and uses a supported Python version.
Write-VerboseLogBanner "Virtual environment"
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
Write-VerboseLog -Lines @(
    "Venv path   : $venvPath",
    "Python exe  : $venvPythonCmd",
    "Python ver  : $venvVersion",
    "Created now : $createVenv"
) -Prefix "  "

# Configure package installer backend: auto | uv | pip
$script:PythonPackageInstallBackend = [string]$script:PythonPackageInstallBackend
if (-not $script:PythonPackageInstallBackend) { $script:PythonPackageInstallBackend = "auto" }

if ($script:PythonPackageInstallBackend -eq "pip") {
    Write-Host "Package backend: pip (UV disabled by selection)" -ForegroundColor Cyan
    $script:UvExe = ""
} else {
    Write-Host "Setting up UV package manager" -NoNewline -ForegroundColor White
    Write-Host "..." -NoNewline -ForegroundColor White
    $script:UvExe = Install-UV
    if ($script:UvExe) {
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White
        Write-Host "  UV: $($script:UvExe)" -ForegroundColor DarkGray
        # Avoid noisy hardlink warnings on Windows when cache and venv are on different volumes.
        if (-not $env:UV_LINK_MODE) {
            $env:UV_LINK_MODE = "copy"
        }
        if ($script:PythonPackageInstallBackend -eq "uv") {
            Write-Host "Package backend: uv (forced)" -ForegroundColor Cyan
        } else {
            Write-Host "Package backend: auto (uv active)" -ForegroundColor Cyan
        }
    } else {
        if ($script:PythonPackageInstallBackend -eq "uv") {
            Write-Host " [ " -NoNewline -ForegroundColor White
            Write-Host "FAIL" -NoNewline -ForegroundColor Red
            Write-Host " ]" -ForegroundColor White
            throw "UV backend was selected, but uv could not be installed or located on PATH."
        }
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "SKIP" -NoNewline -ForegroundColor Yellow
        Write-Host " ] (UV unavailable, falling back to pip)" -ForegroundColor White
        Write-Host "Package backend: auto (pip fallback)" -ForegroundColor Cyan
    }
}
Write-VerboseLogBanner "Package backend"
Write-VerboseLog -Lines @(
    "UV exe      : $(if ($script:UvExe) { $script:UvExe } else { 'not available' })",
    "Backend     : $(if ($script:UvExe) { 'uv (active)' } else { 'pip (fallback)' })",
    "UV_LINK_MODE: $($env:UV_LINK_MODE)"
) -Prefix "  "

if ($script:UvExe -and (Test-Path $script:UvExe)) {
    Write-Host "Skipping pip/setuptools/wheel upgrade (UV backend active)." -ForegroundColor Cyan
} else {
    Write-Host "Upgrading pip/setuptools/wheel" -NoNewline -ForegroundColor White
    $_pipUpgradeExit = Invoke-PipInstall -PythonExe $venvPythonCmd -Packages @("pip", "setuptools", "wheel") -Upgrade
    if ($_pipUpgradeExit -ne 0) {
        Write-Host "" # end NoNewline if output was quiet
        Write-StatusFail "pip upgrade"
        $backendUsed = if ($script:LastPackageInstallBackendUsed) { $script:LastPackageInstallBackendUsed } else { "unknown" }
        $detail = if ($script:LastPackageInstallError) { $script:LastPackageInstallError } else { "No backend error details captured." }
        throw "Failed to upgrade pip (backend=$backendUsed). $detail"
    }
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "OK" -NoNewline -ForegroundColor Green
    Write-Host " ]" -ForegroundColor White
}

if (-not (Test-Path $requirementsPath)) {
    throw "requirements.txt not found at $requirementsPath"
}

Write-VerboseLogBanner "Installing Python dependencies (requirements.txt)"
Write-Host "Installing Python dependencies" -NoNewline -ForegroundColor White
$_depsExit = Invoke-PipInstall -PythonExe $venvPythonCmd -RequirementsFile $requirementsPath
$script:PipBackendForRequirements = if ($script:LastPackageInstallBackendUsed) { $script:LastPackageInstallBackendUsed } else { "unknown" }
if ($_depsExit -ne 0) {
    Write-Host "" # end NoNewline if output was quiet
    Write-StatusFail "Python dependencies"
    $backendUsed = $script:PipBackendForRequirements
    $detail = if ($script:LastPackageInstallError) { $script:LastPackageInstallError } else { "No backend error details captured." }
    throw "Failed to install Python dependencies (backend=$backendUsed). $detail"
}
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White
# Record installed package list to verbose log immediately after core install
try {
    $pkgListJson = & $venvPythonCmd -m pip list --format=json 2>$null
    if ($LASTEXITCODE -eq 0 -and $pkgListJson) {
        $pkgList = ($pkgListJson | ConvertFrom-Json) | ForEach-Object { "$($_.name)==$($_.version)" }
        Write-VerboseLogBanner "Installed packages after requirements.txt"
        Write-VerboseLog -Lines $pkgList -Prefix "  "
        $script:InstalledPackageList = @($pkgList)
    }
} catch {
    Write-VerboseLog -Lines @("pip list failed: $($_.Exception.Message)") -Prefix "  "
}

$aiDefinitions = Get-AiBackendDefinitions
$script:InstalledAiBackends = @()
$script:FailedAiBackends = @()
$script:InstalledOptionalAiBackends = @()
$script:FailedOptionalAiBackends = @()

$aiSettingOverride = $null

Write-Host ""
Write-Host "Checking for installed AI backends..." -NoNewline -ForegroundColor White
$existingAiBackends = Get-InstalledAiBackends -PythonExe $venvPythonCmd -Definitions $aiDefinitions
if ($existingAiBackends.Count -gt 0) {
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "FOUND" -NoNewline -ForegroundColor Green
    Write-Host " ] $($existingAiBackends -join ', ')" -ForegroundColor White
    $script:InstalledAiBackends = @($existingAiBackends)
    $script:InstalledOptionalAiBackends = @($existingAiBackends | Where-Object { $_ -ne "openai-whisper" })
} else {
    Write-Host " [ " -NoNewline -ForegroundColor White
    Write-Host "NONE" -NoNewline -ForegroundColor Yellow
    Write-Host " ]" -ForegroundColor White
}

$selectionFromFlags = Get-AiBackendSelectionFromFlags -Definitions $aiDefinitions
$selectedAiBackends = @()
if ([bool]$selectionFromFlags.has_flags) {
    $selectedAiBackends = @($selectionFromFlags.selected)
    if ($selectedAiBackends.Count -gt 0) {
        Write-Host "AI backend selection from flags: $($selectedAiBackends -join ', ')" -ForegroundColor Cyan
    }
}

if ($selectedAiBackends.Count -eq 0 -and $existingAiBackends.Count -eq 0 -and -not $SkipAiSelectionPrompt) {
    Write-Host ""
    Write-Host "=== AI Backend Selection ===" -ForegroundColor Cyan
    Write-Host "Install only the AI backends you want. The original Whisper backend is available as 'OpenAI Whisper (original)'." -ForegroundColor White
    $selectedAiBackends = Prompt-AiBackendSelections -Definitions $aiDefinitions
}

if ($selectedAiBackends.Count -gt 0) {
    $selectedSet = @{}
    foreach ($k in $selectedAiBackends) { $selectedSet[[string]$k] = $true }

    if ($selectedSet.ContainsKey("aeneas") -and -not (Test-MsvcBuildTools)) {
        Write-Host "Aeneas requires Microsoft C++ Build Tools (MSVC v14+)." -ForegroundColor Yellow
        $buildToolsMethod = if ($ToolInstallMethod -eq "scoop") { "auto" } else { $ToolInstallMethod }
        $buildToolsResult = Install-MsvcBuildTools -Method $buildToolsMethod
        if ($buildToolsResult.success) {
            Add-MsvcClToProcessPath | Out-Null
            Write-Host "Microsoft C++ Build Tools ready. Continuing with aeneas installation." -ForegroundColor Green
        } else {
            $attemptText = if ($buildToolsResult.attempted.Count -gt 0) { $buildToolsResult.attempted -join ", " } else { "none" }
            $selectedAiBackends = @($selectedAiBackends | Where-Object { $_ -ne "aeneas" })
            $selectedSet = @{}
            foreach ($k in $selectedAiBackends) { $selectedSet[[string]$k] = $true }
            Write-Host "Skipping aeneas: Microsoft C++ Build Tools were not installed (attempted: $attemptText)." -ForegroundColor Yellow
            Write-Host "Install manually from https://visualstudio.microsoft.com/visual-cpp-build-tools/ and re-run installer to enable aeneas." -ForegroundColor DarkYellow
        }
    }

    $selectedDefs = @($aiDefinitions | Where-Object { $selectedSet.ContainsKey([string]$_.key) })

    $requiresVc = $false
    foreach ($def in $selectedDefs) {
        if ([bool]$def.needs_vcredist) {
            $requiresVc = $true
            break
        }
    }

    if ($requiresVc) {
        Write-Host "Checking Visual C++ Redistributable..." -NoNewline -ForegroundColor White
        if (-not (Test-VCRedist)) {
            $script:VCRedistPreExisted = $false
            Write-Host " NOT FOUND" -ForegroundColor Yellow
            Write-Host "Installing VC++ Redistributable for selected AI backends..." -ForegroundColor White
            $vcInstallResult = Install-VCRedist
            $script:VCRedistInstallUsed = [string]$vcInstallResult.install_method
            if (-not [bool]$vcInstallResult.success) {
                Write-Host "WARNING: VC++ install failed. Torch-based backends may fail to install." -ForegroundColor Yellow
                Write-VerboseLog -Lines @("VC++ Redist  : INSTALL FAILED (tried: $($vcInstallResult.attempted -join ', '))") -Prefix "  "
            } else {
                Write-StatusOK "Visual C++ Redistributable installed"
                Write-VerboseLog -Lines @("VC++ Redist  : installed via $script:VCRedistInstallUsed") -Prefix "  "
            }
        } else {
            Write-Host " [ " -NoNewline -ForegroundColor White
            Write-Host "OK" -NoNewline -ForegroundColor Green
            Write-Host " ]" -ForegroundColor White
            Write-VerboseLog -Lines @("VC++ Redist  : pre-existed (not installed by this script)") -Prefix "  "
        }
    }

    Write-VerboseLogBanner "Installing AI backend packages"
    $aiInstallResult = Install-AiBackends -PythonExe $venvPythonCmd -Definitions $aiDefinitions -SelectedBackends $selectedAiBackends
    $script:PipBackendForAI = if ($script:LastPackageInstallBackendUsed) { $script:LastPackageInstallBackendUsed } else { "unknown" }
    $script:InstalledAiBackends = @($aiInstallResult.installed)
    $script:FailedAiBackends = @($aiInstallResult.failed)
    $script:InstalledOptionalAiBackends = @($script:InstalledAiBackends | Where-Object { $_ -ne "openai-whisper" })
    $script:FailedOptionalAiBackends = @($script:FailedAiBackends | Where-Object { $_ -ne "openai-whisper" })

    if ($script:InstalledAiBackends.Count -gt 0) {
        Write-Host "Installed AI backends: $($script:InstalledAiBackends -join ', ')" -ForegroundColor Green
    }
    if ($script:FailedAiBackends.Count -gt 0) {
        Write-Host "Backends not ready after install: $($script:FailedAiBackends -join ', ')" -ForegroundColor Yellow
    }
}

$speechBrainRequested = ($selectedAiBackends -contains "speechbrain") -or ($existingAiBackends -contains "speechbrain")
if ($speechBrainRequested) {
    $speechBrainAudioReady = Ensure-SpeechBrainAudioSupport -PythonExe $venvPythonCmd
    if (-not $speechBrainAudioReady) {
        Write-Host "SpeechBrain may still fail to read some audio files. If needed, convert source audio to WAV/FLAC with ffmpeg." -ForegroundColor Yellow
    }
}

$finalAiBackends = Get-InstalledAiBackends -PythonExe $venvPythonCmd -Definitions $aiDefinitions
if ($finalAiBackends.Count -gt 0) {
    $aiSettingOverride = $true
    Write-Host "AI mode enabled (detected backend(s): $($finalAiBackends -join ', '))." -ForegroundColor Green
} else {
    $aiSettingOverride = $false
    if (-not [bool]$selectionFromFlags.has_flags -and -not $SkipAiSelectionPrompt) {
        Write-Host "No AI backend selected or installed. Continuing without AI mode." -ForegroundColor Yellow
    } else {
        Write-Host "No AI backend detected. Continuing without AI mode." -ForegroundColor Yellow
    }
}

Write-Host ""

if (-not (Test-Path $ffmpegInstallerPath)) {
    throw "ffmpeg installer script not found at $ffmpegInstallerPath"
}

# Track whether ffmpeg pre-existed before this run
if (-not (Test-CommandAvailable "ffmpeg") -or -not (Test-CommandAvailable "ffprobe")) {
    $script:FfmpegPreExisted = $false
}

Write-Host "Installing ffmpeg/ffprobe" -NoNewline -ForegroundColor White
$ffmpegMethodOut = Join-Path $env:TEMP "subtitle_ffmpeg_install_method.txt"
if (Test-Path -LiteralPath $ffmpegMethodOut) {
    Remove-Item -LiteralPath $ffmpegMethodOut -Force -ErrorAction SilentlyContinue
}
Write-VerboseLogBanner "Installing ffmpeg/ffprobe"
$ffmpegRawOut = @(powershell -NoProfile -ExecutionPolicy Bypass -File "$ffmpegInstallerPath" -Method $FfmpegInstallMethod -MethodOutFile "$ffmpegMethodOut" 2>&1)
Write-VerboseLog -Lines $ffmpegRawOut -Prefix "  [ffmpeg-installer] "
if ($LASTEXITCODE -ne 0) {
    Write-StatusFail "ffmpeg installation"
    throw "ffmpeg installation failed."
}
if (-not $script:FfmpegPreExisted -and (Test-Path -LiteralPath $ffmpegMethodOut)) {
    try {
        $methodFromInstaller = [string](Get-Content -LiteralPath $ffmpegMethodOut -Raw -ErrorAction SilentlyContinue)
        if ($methodFromInstaller) {
            $script:FfmpegInstallUsed = $methodFromInstaller.Trim()
        }
    } catch {
        # Keep existing value on read failure.
    }
}
if (-not $script:FfmpegPreExisted -and -not $script:FfmpegInstallUsed) {
    # Fallback only if child installer could not report method for some reason.
    if ($FfmpegInstallMethod -in @("winget", "choco", "scoop")) {
        $script:FfmpegInstallUsed = [string]$FfmpegInstallMethod
    } else {
        $script:FfmpegInstallUsed = "auto"
    }
}
Write-Host "..." -NoNewline -ForegroundColor White
Write-Host " [ " -NoNewline -ForegroundColor White
Write-Host "OK" -NoNewline -ForegroundColor Green
Write-Host " ]" -ForegroundColor White
Write-VerboseLog -Lines @(
    "ffmpeg pre-existed : $script:FfmpegPreExisted",
    "install method used: $(if ($script:FfmpegInstallUsed) { $script:FfmpegInstallUsed } else { 'N/A (pre-existed)' })",
    "ffmpeg path        : $((Get-Command 'ffmpeg' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue))",
    "ffprobe path       : $((Get-Command 'ffprobe' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue))"
) -Prefix "  "

Write-Host ""
Write-Host "=== MKVToolNix/HandBrakeCLI/MakeMKV Check (Optional) ==="
$videoToolsBefore = Get-OptionalVideoToolStatus
Write-VerboseLogBanner "Optional video tools (pre-install status)"
foreach ($toolKey in @("mkvtoolnix", "handbrake", "makemkv")) {
    $t = $videoToolsBefore[$toolKey]
    Write-VerboseLog -Lines @(
        "$([string]$t.display):",
        "  found    : $([bool]$t.found)",
        "  on PATH  : $([bool]$t.command_on_path)",
        "  path     : $(if ($t.path) { [string]$t.path } else { 'N/A' })"
    ) -Prefix "  "
}

$script:MkvtoolnixPreExisted = [bool]$videoToolsBefore["mkvtoolnix"].found
    $script:HandBrakePreExisted  = [bool]$videoToolsBefore["handbrake"].found
    $script:MakeMKVPreExisted    = [bool]$videoToolsBefore["makemkv"].found
    $script:MkvtoolnixFinalPath  = [string]$videoToolsBefore["mkvtoolnix"].path
    $script:HandBrakeFinalPath   = [string]$videoToolsBefore["handbrake"].path
    $script:MakeMKVFinalPath     = [string]$videoToolsBefore["makemkv"].path
if ($script:AutoPathBridgeEnabled) {
    foreach ($toolKey in @("mkvtoolnix", "handbrake", "makemkv")) {
        $tool = $videoToolsBefore[$toolKey]
        if ([bool]$tool.found -and -not [bool]$tool.command_on_path) {
            Ensure-ToolCliReachable `
                -ToolCommand ([string]$tool.command) `
                -DetectedBinaryPath ([string]$tool.path) `
                -ToolDisplayName ([string]$tool.display) | Out-Null
        }
    }
    $videoToolsBefore = Get-OptionalVideoToolStatus
}

foreach ($toolKey in @("mkvtoolnix", "handbrake", "makemkv")) {
    $tool = $videoToolsBefore[$toolKey]
    Write-Host "$($tool.display)" -NoNewline -ForegroundColor White
    Write-Host "..." -NoNewline -ForegroundColor White
    if ($tool.found -and $tool.command_on_path) {
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White
    } elseif ($tool.found) {
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "found, adding to PATH" -NoNewline -ForegroundColor Yellow
        Write-Host " ] ($($tool.path))" -ForegroundColor White
    } else {
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "NOT FOUND" -NoNewline -ForegroundColor Yellow
        Write-Host " ]" -ForegroundColor White
    }
}

$missingToolKeys = @("mkvtoolnix", "handbrake", "makemkv") | Where-Object {
    -not [bool]$videoToolsBefore[$_].found
}

if ($missingToolKeys.Count -gt 0) {
    $selectedMissingToolKeys = @()
    $selectionResolved = $false
    $remainingToolKeys = @($missingToolKeys)

    $autoInstallSet = @{}
    foreach ($toolKey in @($script:OptionalToolAutoInstallKeys)) {
        if ($toolKey) {
            $autoInstallSet[[string]$toolKey] = $true
        }
    }
    foreach ($toolKey in $missingToolKeys) {
        if ($autoInstallSet.ContainsKey($toolKey)) {
            $selectedMissingToolKeys += $toolKey
        }
    }
    $selectedMissingToolKeys = @($selectedMissingToolKeys | Select-Object -Unique)
    if ($selectedMissingToolKeys.Count -gt 0) {
        Write-Host "Installer profile requested auto-install for: $($selectedMissingToolKeys -join ', ')." -ForegroundColor Cyan
    }

    $remainingToolKeys = @($missingToolKeys | Where-Object { $selectedMissingToolKeys -notcontains $_ })
    if ($remainingToolKeys.Count -eq 0) {
        $selectionResolved = $true
    } elseif ($script:OptionalToolInstallMode -eq "none") {
        Write-Host "Installer profile requested skipping remaining missing tools: $($remainingToolKeys -join ', ')." -ForegroundColor Cyan
        $selectionResolved = $true
    }

    if (-not $selectionResolved -and (Test-ClickSelectionAvailable)) {
        $toolSelectionDefs = @()
        foreach ($toolKey in $remainingToolKeys) {
            $tool = $videoToolsBefore[$toolKey]
            $toolSelectionDefs += @{ key = [string]$toolKey; display = [string]$tool.display }
        }

        $mouseToolSelection = Select-DefinitionKeysWithMouse `
            -Definitions $toolSelectionDefs `
            -PreselectedKeys $remainingToolKeys `
            -Title "Select MKVToolNix, HandBrakeCLI, MakeMKV to install (optional)"

        if ([bool]$mouseToolSelection.used_gui -and -not [bool]$mouseToolSelection.cancelled) {
            $selectedMissingToolKeys += @($mouseToolSelection.selected)
            $selectedMissingToolKeys = @($selectedMissingToolKeys | Select-Object -Unique)
            $selectionResolved = $true
        } elseif ([bool]$mouseToolSelection.cancelled) {
            Write-Host "Selection window cancelled. Proceeding with any preselected auto-install tools only." -ForegroundColor Yellow
            $selectionResolved = $true
        }
    }

    if (-not $selectionResolved) {
        Write-Host ""
        Write-Host "Install missing MKVToolNix, HandBrakeCLI, and MakeMKV tools now? (optional)" -ForegroundColor Yellow
        Write-Host "These power MKVToolNix/HandBrake/MakeMKV backend options in the app." -ForegroundColor White
        Write-Host "  [Y] Yes (install all missing)" -ForegroundColor White
        Write-Host "  [N] No" -ForegroundColor White
        $installOptionalTools = Read-Host "Your choice [Y/N]"
        if ($installOptionalTools -match "^[Yy]") {
            $selectedMissingToolKeys += @($remainingToolKeys)
            $selectedMissingToolKeys = @($selectedMissingToolKeys | Select-Object -Unique)
        }
    }

    if ($selectedMissingToolKeys.Count -gt 0) {
        foreach ($toolKey in $selectedMissingToolKeys) {
            $tool = $videoToolsBefore[$toolKey]
            $toolInstallMethod = $ToolInstallMethod
            if ($toolKey -eq "mkvtoolnix") {
                $toolInstallMethod = $script:MkvtoolnixInstallMethod
            } elseif ($toolKey -eq "handbrake") {
                $toolInstallMethod = $script:HandBrakeInstallMethod
            } elseif ($toolKey -eq "makemkv") {
                $toolInstallMethod = $script:MakeMKVInstallMethod
            }

            $installResult = Install-OptionalSystemTool `
                -DisplayName ([string]$tool.display) `
                -WingetId ([string]$tool.winget_id) `
                -ChocoId ([string]$tool.choco_id) `
                -ScoopId ([string]$tool.scoop_id) `
                -Method $toolInstallMethod

            $videoToolsAfter = Get-OptionalVideoToolStatus
            if ([bool]$videoToolsAfter[$toolKey].found -and -not [bool]$videoToolsAfter[$toolKey].command_on_path) {
                Ensure-ToolCliReachable `
                    -ToolCommand ([string]$videoToolsAfter[$toolKey].command) `
                    -DetectedBinaryPath ([string]$videoToolsAfter[$toolKey].path) `
                    -ToolDisplayName ([string]$videoToolsAfter[$toolKey].display) | Out-Null
                $videoToolsAfter = Get-OptionalVideoToolStatus
            }

            if ([bool]$videoToolsAfter[$toolKey].found -and [bool]$videoToolsAfter[$toolKey].command_on_path) {
                Write-StatusOK "$($tool.display) ready"
                $usedMethod = [string]$installResult.install_method
                if ($toolKey -eq "mkvtoolnix") {
                    $script:MkvtoolnixInstallUsed = $usedMethod
                    $script:MkvtoolnixFinalPath   = [string]$videoToolsAfter[$toolKey].path
                } elseif ($toolKey -eq "handbrake") {
                    $script:HandBrakeInstallUsed = $usedMethod
                    $script:HandBrakeFinalPath   = [string]$videoToolsAfter[$toolKey].path
                } elseif ($toolKey -eq "makemkv") {
                    $script:MakeMKVInstallUsed = $usedMethod
                    $script:MakeMKVFinalPath   = [string]$videoToolsAfter[$toolKey].path
                }
                Write-VerboseLog -Lines @("$($tool.display): installed via $usedMethod  path=$([string]$videoToolsAfter[$toolKey].path)") -Prefix "  "
            } else {
                $attemptedText = ""
                if ($installResult.attempted.Count -gt 0) {
                    $attemptedText = " Tried: $($installResult.attempted -join ', ')."
                }
                if ([bool]$videoToolsAfter[$toolKey].found) {
                    Write-Host "$($tool.display) found but CLI command still not reachable.$attemptedText" -ForegroundColor Yellow
                } else {
                    Write-Host "$($tool.display) not detected after install.$attemptedText" -ForegroundColor Yellow
                }
                Write-Host "Configure its path later in the GUI Tooling & Diagnostics tab." -ForegroundColor DarkYellow
            }
        }
    } else {
        Write-Host "Skipping MKVToolNix/HandBrakeCLI/MakeMKV install (optional)." -ForegroundColor Yellow
    }
}

    # --- .NET Desktop Runtime (HandBrake dependency) ---
    $currentVideoStatus = Get-OptionalVideoToolStatus
    if ([bool]$currentVideoStatus["handbrake"].found) {
        Write-Host ".NET Desktop Runtime" -NoNewline -ForegroundColor White
        Write-Host "..." -NoNewline -ForegroundColor White
        if (Test-DotNetDesktopRuntime) {
            Write-Host " [ " -NoNewline -ForegroundColor White
            Write-Host "OK" -NoNewline -ForegroundColor Green
            Write-Host " ]" -ForegroundColor White
        } else {
            $script:DotNetRuntimePreExisted = $false
            Write-Host " NOT FOUND" -ForegroundColor Yellow
            Write-Host "HandBrake requires .NET Desktop Runtime 8. Installing..." -ForegroundColor White
            $dotnetResult = Install-DotNetRuntime
            $script:DotNetRuntimeInstallUsed = [string]$dotnetResult.install_method
            if ([bool]$dotnetResult.success) {
                Write-StatusOK ".NET Desktop Runtime installed"
            } else {
                $attemptText = if ($dotnetResult.attempted.Count -gt 0) { $dotnetResult.attempted -join ", " } else { "none" }
                Write-Host "WARNING: .NET Desktop Runtime install failed (tried: $attemptText)." -ForegroundColor Yellow
                Write-Host "Install manually: https://dotnet.microsoft.com/download/dotnet/8.0" -ForegroundColor DarkYellow
            }
        }
    }

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

# Verify optional backend binaries and show paths/fallback guidance.
$videoToolVerify = Get-OptionalVideoToolStatus
foreach ($toolKey in @("mkvtoolnix", "handbrake", "makemkv")) {
    $tool = $videoToolVerify[$toolKey]
    Write-Host "$($tool.display)" -NoNewline -ForegroundColor White
    if ([bool]$tool.found -and [bool]$tool.command_on_path) {
        Write-Host "..." -NoNewline -ForegroundColor White
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ] $($tool.path)" -ForegroundColor White
    } elseif ([bool]$tool.found) {
        Write-Host "..." -NoNewline -ForegroundColor White
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "PARTIAL" -NoNewline -ForegroundColor Yellow
        Write-Host " ] $($tool.path)" -ForegroundColor White
        Write-Host "  Binary exists but CLI command is not callable from PATH yet." -ForegroundColor DarkYellow
    } else {
        Write-Host "..." -NoNewline -ForegroundColor White
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OPTIONAL" -NoNewline -ForegroundColor Yellow
        Write-Host " ]" -ForegroundColor White
        Write-Host "  Not detected. Install with package manager or set path in Tooling & Diagnostics." -ForegroundColor DarkYellow
    }
}

# Verify AI backend module availability and summarize readiness.
$aiBackendChecks = @(
    @{label = "OpenAI Whisper"; import = "whisper"},
    @{label = "faster-whisper"; import = "faster_whisper"},
    @{label = "WhisperX"; import = "whisperx"},
    @{label = "stable-ts"; import = "stable_whisper"},
    @{label = "whisper-timestamped"; import = "whisper_timestamped"},
    @{label = "SpeechBrain"; import = "speechbrain"},
    @{label = "Vosk"; import = "vosk"},
    @{label = "Aeneas"; import = "aeneas"}
)

$availableAiBackends = @()
foreach ($check in $aiBackendChecks) {
    $label = [string]$check.label
    $moduleName = [string]$check.import
    Write-Host "$label backend" -NoNewline -ForegroundColor White
    & $venvPythonCmd -c "import importlib.util as u,sys; sys.exit(0 if u.find_spec('$moduleName') is not None else 1)" 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $availableAiBackends += $label
        Write-Host "..." -NoNewline -ForegroundColor White
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OK" -NoNewline -ForegroundColor Green
        Write-Host " ]" -ForegroundColor White
    } else {
        Write-Host "..." -NoNewline -ForegroundColor White
        Write-Host " [ " -NoNewline -ForegroundColor White
        Write-Host "OPTIONAL" -NoNewline -ForegroundColor Yellow
        Write-Host " ]" -ForegroundColor White
    }
}

if ($availableAiBackends.Count -gt 0) {
    Write-Host "Detected AI backends: $($availableAiBackends -join ', ')" -ForegroundColor Green
} else {
    Write-Host "No AI backends detected." -ForegroundColor Yellow
}
Write-StatusOK "All components installed and verified"

# ---------------------------------------------------------------------------
# Write install manifest so the uninstaller knows what we set up.
# ---------------------------------------------------------------------------
$newManifest = @{
    install_date        = (Get-Date -Format "yyyy-MM-dd HH:mm")
    script_version      = "1.1"
    python = @{
        pre_existed     = $script:PythonPreExisted
        install_method  = if ($script:PythonPreExisted) { "" } else { $script:PythonInstallUsed }
        version         = if ($venvVersion) { $venvVersion } else { "" }
    }
    vcredist = @{
        pre_existed     = $script:VCRedistPreExisted
        install_method  = if ($script:VCRedistPreExisted) { "" } else { $script:VCRedistInstallUsed }
    }
        dotnet_runtime = @{
            pre_existed     = $script:DotNetRuntimePreExisted
            install_method  = if ($script:DotNetRuntimePreExisted) { "" } else { $script:DotNetRuntimeInstallUsed }
        }
    ffmpeg = @{
        pre_existed     = $script:FfmpegPreExisted
        install_method  = if ($script:FfmpegPreExisted) { "" } else { $script:FfmpegInstallUsed }
    }
    mkvtoolnix = @{
        pre_existed     = $script:MkvtoolnixPreExisted
        install_method  = if ($script:MkvtoolnixPreExisted) { "" } else { $script:MkvtoolnixInstallUsed }
        is_installed    = [bool](Get-OptionalVideoToolStatus)["mkvtoolnix"].found
        path            = if ($script:MkvtoolnixFinalPath) { $script:MkvtoolnixFinalPath } else { [string](Get-OptionalVideoToolStatus)["mkvtoolnix"].path }
    }
    handbrake = @{
        pre_existed     = $script:HandBrakePreExisted
        install_method  = if ($script:HandBrakePreExisted) { "" } else { $script:HandBrakeInstallUsed }
        is_installed    = [bool](Get-OptionalVideoToolStatus)["handbrake"].found
        path            = if ($script:HandBrakeFinalPath) { $script:HandBrakeFinalPath } else { [string](Get-OptionalVideoToolStatus)["handbrake"].path }
    }
    makemkv = @{
        pre_existed     = $script:MakeMKVPreExisted
        install_method  = if ($script:MakeMKVPreExisted) { "" } else { $script:MakeMKVInstallUsed }
        is_installed    = [bool](Get-OptionalVideoToolStatus)["makemkv"].found
        path            = if ($script:MakeMKVFinalPath) { $script:MakeMKVFinalPath } else { [string](Get-OptionalVideoToolStatus)["makemkv"].path }
    }
    venv = @{
        path            = $venvPath
        created_by_script = $true
    }
    packages = @{
        pip_backend         = $script:PipBackendForRequirements
        ai_pip_backend      = if ($script:PipBackendForAI) { $script:PipBackendForAI } else { "" }
        installed_packages  = if ($script:InstalledPackageList) { @($script:InstalledPackageList) } else { @() }
    }
    ai = @{
        installed       = ($aiSettingOverride -eq $true)
        installed_backends = @($script:InstalledAiBackends)
        failed_backends = @($script:FailedAiBackends)
        optional_backends_installed = @($script:InstalledOptionalAiBackends)
        optional_backends_failed = @($script:FailedOptionalAiBackends)
    }
}
Save-InstallManifest -Manifest $newManifest
Write-Host "Install manifest saved to: $manifestPath" -ForegroundColor Gray

Write-VerboseLogBanner "Install summary"
Write-VerboseLog -Lines @(
    "--- System components ---",
    "Python       : $(if ($script:PythonPreExisted) { 'pre-existed (' + $pythonCmd + ')' } else { 'installed via ' + $script:PythonInstallUsed })",
    "Venv         : $venvPath  (Python $venvVersion)",
    "VC++ Redist  : $(if ($script:VCRedistPreExisted) { 'pre-existed' } else { 'installed via ' + $script:VCRedistInstallUsed })",
    "ffmpeg       : $(if ($script:FfmpegPreExisted) { 'pre-existed' } else { 'installed via ' + $script:FfmpegInstallUsed })",
    "MKVToolNix   : $(if ($script:MkvtoolnixPreExisted) { 'pre-existed' } else { if ($script:MkvtoolnixInstallUsed) { 'installed via ' + $script:MkvtoolnixInstallUsed } else { 'not installed' } }) $(if ($script:MkvtoolnixFinalPath) { '(' + $script:MkvtoolnixFinalPath + ')' })",
    "HandBrake    : $(if ($script:HandBrakePreExisted) { 'pre-existed' } else { if ($script:HandBrakeInstallUsed) { 'installed via ' + $script:HandBrakeInstallUsed } else { 'not installed' } }) $(if ($script:HandBrakeFinalPath) { '(' + $script:HandBrakeFinalPath + ')' })",
    "MakeMKV      : $(if ($script:MakeMKVPreExisted) { 'pre-existed' } else { if ($script:MakeMKVInstallUsed) { 'installed via ' + $script:MakeMKVInstallUsed } else { 'not installed' } }) $(if ($script:MakeMKVFinalPath) { '(' + $script:MakeMKVFinalPath + ')' })",
    ".NET Runtime : $(if ($script:DotNetRuntimePreExisted) { 'pre-existed' } else { 'installed via ' + $script:DotNetRuntimeInstallUsed })",
    "",
    "--- Package backend ---",
    "Requirements : $script:PipBackendForRequirements",
    "AI packages  : $(if ($script:PipBackendForAI) { $script:PipBackendForAI } else { 'N/A' })",
    "",
    "--- AI backends ---",
    "Installed    : $(if ($script:InstalledAiBackends.Count -gt 0) { $script:InstalledAiBackends -join ', ' } else { 'none' })",
    "Failed       : $(if ($script:FailedAiBackends.Count -gt 0) { $script:FailedAiBackends -join ', ' } else { 'none' })"
) -Prefix "  "

Write-VerboseLogBanner "Install complete"
try { Stop-Transcript | Out-Null } catch { }
Write-Host "Install log saved to: $installLogPath" -ForegroundColor Gray
if ($script:verboseLogPath -and (Test-Path $script:verboseLogPath)) {
    Write-Host "Verbose log saved to: $($script:verboseLogPath)" -ForegroundColor Gray
}

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

$previousPythonWarnings = $env:PYTHONWARNINGS
try {
    $env:PYTHONWARNINGS = "ignore::UserWarning:speechbrain.utils.torch_audio_backend"
    & $venvPythonCmd @guiArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Subtitle Tool failed to launch correctly (exit code: $LASTEXITCODE)."
    }
} finally {
    if ($null -eq $previousPythonWarnings) {
        Remove-Item Env:PYTHONWARNINGS -ErrorAction SilentlyContinue
    } else {
        $env:PYTHONWARNINGS = $previousPythonWarnings
    }
}

Cleanup-InstallerArtifacts

if (-not $NoPause) {
    Write-Host ""
    Read-Host "Installation finished. Press Enter to close"
}

exit 0
