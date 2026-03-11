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
- **AI Subtitle Generation** (Optional): Generate subtitles from video audio using Whisper AI.
- FastAPI service mode with background job execution.

## Requirements

- Python 3.10+
- `ffmpeg` and `ffprobe` on your `PATH`

Install Python dependencies:

```bash
pip install -r "Python Projects/Subtitle/requirements.txt"

# Optional: Install AI subtitle generation (Whisper AI + ~1-2GB disk space)
pip install -r "Python Projects/Subtitle/requirements_ai.txt"
```

Install ffmpeg on Windows:

```powershell
# Auto-detect winget/choco/scoop
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_ffmpeg_windows.ps1"

# Or force one method
powershell -ExecutionPolicy Bypass -File "Python Projects/Subtitle/install_ffmpeg_windows.ps1" -Method winget
```

## GUI Usage

```bash
# Launch with default settings (uses saved 'use_ai' preference)
python "Python Projects/Subtitle/subtitle_tool.py" gui

# Disable AI features and save preference
python "Python Projects/Subtitle/subtitle_tool.py" gui --no-ai

# Enable AI features and save preference
python "Python Projects/Subtitle/subtitle_tool.py" gui --use-ai

# Clear saved UI state/memory
python "Python Projects/Subtitle/subtitle_tool.py" gui --clear
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
- `Open Help` to view the integrated help documentation in-app.
- `Show Tutorial` to launch an interactive walkthrough of all features.

**First Run**: The tutorial automatically prompts on first launch to help new users get started.

## AI Subtitle Generation (Optional)

The tool includes optional AI-powered subtitle generation using OpenAI's Whisper model:

- **100% Local**: Runs entirely on your machine - no internet or API keys needed
- **Multiple Model Sizes**: Choose from 7 models: tiny, base, small, medium, large, large-v2, large-v3
- **90+ Languages**: Automatic language detection or manual specification
- **Disk Space**: ~3-4GB for PyTorch base, models 72MB-2.9GB each (total up to ~10GB with all models)

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

```bash
# Install AI libraries later
pip install -r "Python Projects/Subtitle/requirements_ai.txt"

# Then enable AI in settings
python "Python Projects/Subtitle/subtitle_tool.py" gui --use-ai
```

### Using AI Subtitle Generation

1. Add video files or folders
2. Select model size (tiny/base/small/medium/large/large-v2/large-v3)
3. Optionally specify language code (e.g., "en", "es", "fr")
4. Click "Generate Subtitles"
5. SRT files are created next to your videos

**Note**: First run downloads the selected model. Processing time depends on model size and video length.

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
python "Python Projects/Subtitle/subtitle_tool.py" scan --folders "D:\\Videos" --only-with-embedded
```

Remove subtitles:

```bash
python "Python Projects/Subtitle/subtitle_tool.py" remove --folders "D:\\Videos" --suffix _nosubs
```

Include subtitles:

```bash
python "Python Projects/Subtitle/subtitle_tool.py" include --folders "D:\\Videos" --suffix _withsubs
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
