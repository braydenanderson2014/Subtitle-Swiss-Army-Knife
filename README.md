# Subtitle Tool (PyQt + API)

This tool scans video folders, removes embedded subtitle streams, and can add subtitles back from sidecar subtitle files.

## Features

- PyQt GUI to target one or more folders.
- Scan mode to inspect embedded subtitle streams and sidecar subtitles.
- Remove mode to strip embedded subtitle streams.
- Include mode to embed subtitle sidecar files back into the video container.
- Extract mode to export embedded subtitles to separate files.
- **Format Conversion**: Convert videos between MKV and MP4 formats.
- **Media Organization**: Automatically organize movies and TV shows.
- **Metadata Repair**: Fix corrupted video containers.
- **External Backend Support**: Use FFmpeg, MKVToolNix (`mkvmerge`), HandBrakeCLI, and MakeMKV (`makemkvcon`) where applicable.
- **Command Diagnostics**: Adjustable command feedback and ffmpeg/ffprobe log levels for easier troubleshooting of corrupt/truncated files.
- **AI Subtitle Generation** (Optional): Generate subtitles from video audio using multiple backends (Whisper, faster-whisper, WhisperX, stable-ts, whisper-timestamped, SpeechBrain, Vosk, Text-to-Timestamps heuristic).
- **AI Audio Language Tagging** (Optional): Detect audio-stream language with Whisper and write stream language metadata tags.
- FastAPI service mode with background job execution.

## Requirements

- Python 3.10+
- `ffmpeg` and `ffprobe` on your `PATH`

## Windows Install (Recommended)

Run one of these from the repository root.

Option 1: Full installer via PowerShell (recommended)

```powershell
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1"
```

Option 2: Full installer via BAT launcher (same functionality, easier double-click)

```bat
"Python Projects\Subtitle\install_all_windows.bat"
```

The full installer will:
- Detect/install a supported Python version (3.10-3.12, prefers 3.11)
- Create or repair the local virtual environment
- Install core dependencies from `requirements.txt`
- Optionally install AI dependencies from `requirements_ai.txt`
- Install/verify `ffmpeg` and `ffprobe`
- Check MKVToolNix (`mkvmerge`), HandBrakeCLI, and MakeMKV (`makemkvcon`) and optionally install missing tools
- Use click-select menus (Out-GridView) for AI/tool selection when available, with keyboard fallback
- Open a WinUtil-style interactive control panel on Windows with clickable controls and keyboard shortcuts
- In that control panel, MKVToolNix, HandBrakeCLI, and MakeMKV are configured separately (method + auto-install-if-missing)
- Automatically bridge detected GUI-installed tools into CLI usage by adding PATH entries (and creating a MakeMKV shim when needed)
- Optionally install extended AI backend packages (faster-whisper, WhisperX, stable-ts, whisper-timestamped, SpeechBrain, Vosk, Aeneas)
- Verify key packages and launch the GUI

Installer control panel shortcuts (Windows Forms mode):

- `Ctrl+R`: Recommended preset
- `Ctrl+F`: Full AI preset
- `Ctrl+M`: Minimal preset
- `Enter`: Continue install
- `Esc`: Cancel

Useful installer options (PowerShell form):

```powershell
# Force Python install method: auto | winget | choco | scoop
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -PythonInstallMethod winget

# Force ffmpeg install method: auto | winget | choco | scoop
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -FfmpegInstallMethod winget

# Force MKVToolNix/HandBrakeCLI/MakeMKV install method (optional): auto | winget | choco | scoop
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -ToolInstallMethod winget

# Preselect AI backend installs (per-backend flags)
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -InstallAiOpenAIWhisper -InstallAiVosk

# Install all supported AI backends without per-item prompts
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -InstallAiAll -SkipAiSelectionPrompt

# Skip the interactive setup menu (automation/CI)
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -NoMenu

# Show interactive menu explicitly
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -InteractiveMenu

# Disable automatic CLI PATH bridge behavior for detected GUI tool installs
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -DisableAutoPathBridge

# Disable click-select picker windows and always use keyboard prompts
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -DisableClickSelection

# Do not pause at end
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -NoPause

# Keep temporary installer artifacts/cache folders
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1" -KeepInstallArtifacts
```

The BAT launcher passes arguments through, so these also work:

```bat
"Python Projects\Subtitle\install_all_windows.bat" -NoPause
"Python Projects\Subtitle\install_all_windows.bat" -PythonInstallMethod winget -FfmpegInstallMethod winget -ToolInstallMethod winget
"Python Projects\Subtitle\install_all_windows.bat" -InstallAiOpenAIWhisper -InstallAiWhisperX -InstallAiAeneas
"Python Projects\Subtitle\install_all_windows.bat" -NoMenu -InstallAiAll -SkipAiSelectionPrompt
```

AI backend selection flags:

- `-InstallAiOpenAIWhisper`
- `-InstallAiFasterWhisper`
- `-InstallAiWhisperX`
- `-InstallAiStableTs`
- `-InstallAiWhisperTimestamped`
- `-InstallAiSpeechBrain`
- `-InstallAiVosk`
- `-InstallAiAeneas`
- `-InstallAiAll`
- `-SkipAiSelectionPrompt`
- `-InteractiveMenu`
- `-NoMenu`
- `-DisableAutoPathBridge`
- `-DisableClickSelection`

When auto PATH bridge is enabled, the installer also writes user-level tool path variables when detected:

- `SUBTITLE_TOOL_MKVMERGE_BIN`
- `SUBTITLE_TOOL_HANDBRAKE_BIN`
- `SUBTITLE_TOOL_MAKEMKVCON_BIN`

### ffmpeg-Only Install (Optional)

Use this only if you already have Python and dependencies installed and just need ffmpeg/ffprobe.

```powershell
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_ffmpeg_windows.ps1"
```

```bat
"Python Projects\Subtitle\install_ffmpeg_windows.bat"
```

You can force method selection here as well:

```powershell
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_ffmpeg_windows.ps1" -Method winget
```

## Manual Install

If you prefer manual setup, install Python dependencies with:

```bash
pip install -r "Python Projects/Subtitle/requirements.txt"

# Optional: Install AI subtitle generation (Whisper AI + ~1-2GB disk space)
pip install -r "Python Projects/Subtitle/requirements_ai.txt"

# Optional: Install extra AI backends supported by the app
pip install faster-whisper whisperx stable-ts whisper-timestamped speechbrain vosk aeneas
```

Install ffmpeg on Windows manually via installer script:

```powershell
# Auto-detect winget/choco/scoop
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_ffmpeg_windows.ps1"

# Or force one method
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_ffmpeg_windows.ps1" -Method winget
```

## GUI Usage

```bash
# Launch with default settings (uses saved 'use_ai' preference)
python "Python Projects/Subtitle/subtitle_ui.py"

# Disable AI features and save preference
python "Python Projects/Subtitle/subtitle_ui.py" --no-ai

# Enable AI features and save preference
python "Python Projects/Subtitle/subtitle_ui.py" --use-ai

# Clear saved UI state/memory
python "Python Projects/Subtitle/subtitle_ui.py" --clear
```

Or use the quick launcher:

```bat
"Python Projects\Subtitle\launch_gui.bat"
```

**Note**: The `--use-ai` and `--no-ai` flags save your preference to settings. You only need to use them once to change the setting.

In GUI mode:

- Add one or more folders.
- Optionally add specific target video files (supports drag/drop).
- Optionally map manual subtitle files to a selected target video (supports drag/drop).
- Optionally enable scan filtering to only show files that have embedded subtitles.
- `Scan Videos` to preview subtitle availability.
- `Remove Embedded Subtitles` to strip subtitle streams.
- `Include Subtitles Back In` to add sidecar subtitle files.
- Use `Video Tools (Swiss Army Knife)`:
  - `Operations` tab for conversion/repair/organize actions
  - `Tooling & Diagnostics` tab to configure executable paths, choose conversion/repair backends, and refresh tool status
- `Open Help` to view the integrated help documentation in-app.
- `Show Tutorial` to launch an interactive walkthrough of all features.

### Organize Media JSON Rules (Optional)

`Organize Media` now supports an optional JSON rules file so you can clean torrent-style names and control TV episode naming.

- In the GUI, set **Organize Rules JSON (optional)** to your rules file.
- Leave it blank to use the default built-in behavior.
- A starter template is included at:
  - `Python Projects/Subtitle/organize_media_rules.example.json`
- Additional presets are included at:
  - `Python Projects/Subtitle/json_examples/media_rules_balanced.json`
  - `Python Projects/Subtitle/json_examples/media_rules_aggressive_scene.json`
  - `Python Projects/Subtitle/json_examples/media_rules_minimal_cleanup.json`
  - `Python Projects/Subtitle/json_examples/media_rules_movie_year_focus.json`
  - `Python Projects/Subtitle/json_examples/media_rules_tv_episode_focus.json`

Example behavior with rules:
- `Show.Name.S01E02.1080p.WEBRip.x265-GRP.mkv` → `Show Name - S01E02.mkv`
- `Movie.Title.2024.1080p.BluRay.x264-YTS.mkv` (folder movie) → `Movie Title 2024.mkv`

**First Run**: The tutorial automatically prompts on first launch to help new users get started.

## AI Subtitle Generation (Optional)

The tool includes optional AI-powered subtitle generation with selectable backend engines:

- **100% Local**: Runs entirely on your machine - no internet or API keys needed
- **Backends**: Auto, OpenAI Whisper, faster-whisper, WhisperX, stable-ts, whisper-timestamped, SpeechBrain, Vosk, Text-to-Timestamps (heuristic)
- **Multiple Model Sizes**: For Whisper-family backends, choose from 7 models: tiny, base, small, medium, large, large-v2, large-v3
- **90+ Languages**: Automatic language detection or manual specification
- **Disk Space**: ~3-4GB for PyTorch base, models 72MB-2.9GB each (total up to ~10GB with all models)

Supported AI tooling in the app:

- `faster-whisper`
- `WhisperX`
- `stable-ts`
- `whisper-timestamped`
- `SpeechBrain`
- `Vosk`
- `Aeneas` (subtitle sync backend)
- `Text to Timestamps` (built-in heuristic backend)
- `pysubs2` (sometimes searched as `pysub2`)

### Requirements for AI Features

**Windows Requirements:**
- Visual C++ Redistributable 2015-2022 (x64) - automatically installed by installer
  - Manual download: https://aka.ms/vs/17/release/vc_redist.x64.exe
- ~10GB free disk space (base libraries + models)
- Stable internet for initial download

**Common Issue - DLL Error:**
If you see `Error loading "c10.dll"` or similar:
1. Install Visual C++ Redistributable (link above)
2. Restart your computer
3. Re-run the installer or: `pip install -r requirements_ai.txt`

### Installing AI Features

During installation, you'll be prompted whether to install AI libraries. If you skip it:

The installer now allows per-backend selection (including original OpenAI Whisper). It only enables `--use-ai` when at least one AI backend is detected as installed.

```bash
# Install AI libraries later
pip install -r "Python Projects/Subtitle/requirements_ai.txt"

# Then enable AI in settings
python "Python Projects/Subtitle/subtitle_ui.py" --use-ai
```

### Using AI Subtitle Generation

1. Add video files or folders
2. Select model size (tiny/base/small/medium/large/large-v2/large-v3)
3. Optionally specify language code (e.g., "en", "es", "fr")
4. Choose backend + model and click "Generate Subtitles"
5. SRT files are created next to your videos

**Note**: First run may download selected backend models. Processing time depends on backend/model and video length.

### AI Audio Language Detection + Metadata Tagging

Use Whisper to detect language for each audio stream, then write language tags back into container metadata.

- Supports `random snippets` mode (faster, default) and `whole stream` mode (slower, deeper analysis).
- Can preserve existing language tags or overwrite them.
- Works in GUI (`Detect + Tag Audio Language`) and CLI.

CLI example:

```bash
python "Python Projects/Subtitle/subtitle_cli.py" tag-audio-language --folders "D:\\Videos" --strategy snippets --snippets 3 --sample-seconds 25
```

### Model Sizes

- **tiny**: ~39M params, ~72MB download, fastest, least accurate
- **base**: ~74M params, ~140MB download, good balance (recommended)
- **small**: ~244M params, ~460MB download, better accuracy
- **medium**: ~769M params, ~1.5GB download, high accuracy
- **large**: ~1550M params, ~2.9GB download, best accuracy
- **large-v2**: ~1550M params, ~2.9GB download, improved large model
- **large-v3**: ~1550M params, ~2.9GB download, latest & best accuracy

## Integrated Help & Tutorial

The GUI includes:

- **Help Dialog**: Click `Open Help` to view comprehensive documentation in a built-in window. Content is from `SUBTITLE_TOOL_HELP.md`.
- **Interactive Tutorial**: Click `Show Tutorial` to launch a step-by-step walkthrough that highlights and explains each UI element. Features animated flashing borders that pulse to draw attention to the current element being explained.
- **First-Run Tutorial**: On first launch, you'll be prompted to take the tutorial. You can skip it and access it later via `Show Tutorial`.

Settings are stored in `.subtitle_tool_settings.json` in the same directory as the script.

## CLI Usage

Scan:

```bash
python "Python Projects/Subtitle/subtitle_cli.py" scan --folders "D:\\Videos" --only-with-embedded
```

Scan with custom tool binaries and verbose command feedback:

```bash
python "Python Projects/Subtitle/subtitle_cli.py" scan --folders "D:\\Videos" --ffmpeg-bin "C:\\Tools\\ffmpeg\\bin\\ffmpeg.exe" --ffprobe-bin "C:\\Tools\\ffmpeg\\bin\\ffprobe.exe" --command-feedback verbose --ffprobe-loglevel warning
```

Remove subtitles:

```bash
python "Python Projects/Subtitle/subtitle_cli.py" remove --folders "D:\\Videos" --suffix _nosubs
```

Include subtitles:

```bash
python "Python Projects/Subtitle/subtitle_cli.py" include --folders "D:\\Videos" --suffix _withsubs
```

## Full Windows Install Automation

Use the full installer to validate/install Python, install Python dependencies, install ffmpeg/ffprobe, and optionally install AI libraries:

```powershell
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_all_windows.ps1"
```

Or via batch wrapper:

```bat
"Python Projects\Subtitle\install_all_windows.bat"
```

**During installation** you'll be asked:
- Whether to install AI libraries (Whisper + pysubs2)
- If you choose "No", the tool launches with `--no-ai` flag automatically
- You can install AI libraries later and use `--use-ai` flag

## API Usage (Background Jobs)

Start API server:

```bash
python "Python Projects/Subtitle/subtitle_tool.py" api --host 127.0.0.1 --port 8891
```

Queue a remove job:

```bash
curl -X POST "http://127.0.0.1:8891/jobs/remove" \
  -H "Content-Type: application/json" \
  -d '{"folders":["D:/Videos"],"recursive":true,"overwrite":false,"output_suffix":"_nosubs","extract_for_restore":true}'
```

Check job status:

```bash
curl "http://127.0.0.1:8891/jobs/<job_id>"
```

## Notes

- Remove mode can optionally extract embedded subtitle streams to `*.embedded_subN.srt` before stripping.
- Include mode searches for sidecar subtitle files matching the video base name.
- For MP4-family output files, subtitle streams are encoded as `mov_text`.
