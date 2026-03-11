# Subtitle Tool Help

## What This Tool Does

This utility helps you manage subtitle streams in video files.

- Scan videos to see embedded subtitle streams and sidecar subtitle files.
- Remove embedded subtitle streams from videos.
- Include subtitle sidecar files back into videos.
- Extract embedded subtitle streams to standalone subtitle files.

The GUI supports folder scanning and explicit target-file workflows.

## UI Sections

### Target Folders

Use this list when you want to process whole directories.

- `Add Folder`: add a folder to process.
- `Remove Selected`: remove highlighted folders.
- `Clear`: remove all folders.

### Target Video Files (optional)

Use this list when you want to process specific files.

- Drag and drop video files directly into the list.
- `Add Files`: browse and add one or more files.
- `Remove Selected`: remove highlighted files.
- `Clear`: remove all target files.

If both folders and target files are set, the tool processes the combined unique set.

### Manual Subtitle Files for Selected Video (include mode)

Use this to force subtitle files for one selected target video.

- Select one video in `Target Video Files` first.
- Drag and drop subtitle files into the subtitle list, or click `Add Subtitle Files`.
- `Remove Selected`: remove highlighted subtitle files from that video mapping.
- `Clear Current Video List`: clear all mapped subtitle files for the selected video.

Supported subtitle sidecar formats:

- `.srt`, `.ass`, `.ssa`, `.vtt`, `.sub`, `.ttml`

### Options

- `Scan folders recursively`: include subfolders during folder-based discovery.
- `Overwrite original files`: replace originals instead of creating suffixed output files.
- `Extract embedded subtitles before removal (for restore)`: save embedded streams before stripping.
- `Export .txt copies for subtitles`: when extracting text-based subtitles, also emit plain text files.
- `Scan only files with embedded subtitles`: scan output excludes files that have no embedded subtitle streams.
- `Use only selected target video file(s)`: process only selected entries from the target file list.

Suffix settings:

- `Remove output suffix`: default `_nosubs`
- `Include output suffix`: default `_withsubs`
- `Extract output suffix`: default `.embedded_sub`

### Action Buttons

- `Scan Videos`: inspect files and report subtitle availability.
- `Remove Embedded Subtitles`: remove subtitle streams from videos.
- `Include Subtitles Back In`: embed sidecar subtitles into videos.
- `Extract Embedded Subtitles`: export each subtitle stream to sidecar files.
- `Open Help`: opens this help document.

## Processing Behavior

### Scan

Reports, per file:

- number of embedded subtitle streams
- number of matching sidecar subtitle files

### Remove Embedded Subtitles

- Keeps video/audio streams.
- Removes subtitle streams.
- Optionally extracts subtitle streams first.

### Include Subtitles Back In

- Drops existing embedded subtitle streams.
- Adds discovered or manually specified sidecar subtitles.
- For MP4/M4V/MOV outputs, subtitle codec is `mov_text` for compatibility.

### Extract Embedded Subtitles

- Exports each subtitle stream.
- Uses an extension based on subtitle codec when possible.
- Can export `.txt` versions for text subtitle formats.

## Typical Workflows

### Process only one file

1. Add a file to `Target Video Files`.
2. Select the file.
3. Enable `Use only selected target video file(s)`.
4. Click the action you want.

### Include custom subtitle files for one video

1. Add/select a target video file.
2. Add subtitle files in `Manual Subtitle Files for Selected Video`.
3. Run `Include Subtitles Back In`.

### Scan only files that have embedded subtitles

1. Enable `Scan only files with embedded subtitles`.
2. Click `Scan Videos`.

## CLI Quick Reference

From the repository root:

```bash
python "Python Projects/Subtitle/subtitle_tool.py" gui
python "Python Projects/Subtitle/subtitle_tool.py" scan --folders "D:\Videos" --only-with-embedded
python "Python Projects/Subtitle/subtitle_tool.py" remove --folders "D:\Videos" --suffix _nosubs
python "Python Projects/Subtitle/subtitle_tool.py" include --folders "D:\Videos" --suffix _withsubs
python "Python Projects/Subtitle/subtitle_tool.py" extract --folders "D:\Videos" --suffix .embedded_sub
```

## Troubleshooting

- If ffmpeg/ffprobe is missing, install ffmpeg first.
- If GUI does not start, install Python dependencies from `requirements.txt`.
- If include mode skips files, check sidecar file naming or manual mapping for selected videos.
