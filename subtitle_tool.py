#!/usr/bin/env python3
"""
Subtitle utility with a PyQt UI and FastAPI background API.

Features:
- Scan one or more folders for video files and subtitle availability.
- Remove embedded subtitle streams from video containers.
- Re-embed subtitles from sidecar subtitle files.
- Run as a GUI application or as a background HTTP API service.

Notes:
- This tool relies on ffmpeg and ffprobe binaries being installed on the host.
- For MP4 outputs, subtitle streams are encoded as mov_text for compatibility.
"""

from __future__ import annotations

import argparse
import importlib
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree

try:
    from fastapi import FastAPI, HTTPException
    from pydantic import BaseModel, Field
except ImportError:
    FastAPI = None  # type: ignore[assignment]

    class HTTPException(Exception):
        def __init__(self, status_code: int = 500, detail: str = "") -> None:
            super().__init__(detail)
            self.status_code = status_code
            self.detail = detail

    class BaseModel:
        pass

    def Field(*, default=None, default_factory=None, **_kwargs):
        if default_factory is not None:
            return default_factory()
        return default

try:
    import uvicorn
except ImportError:
    uvicorn = None  # type: ignore[assignment]

try:
    from PyQt6.QtCore import QThread, Qt, QTimer, pyqtSignal
    from PyQt6.QtGui import QFont, QPalette, QColor
    from PyQt6.QtWidgets import (
        QAbstractItemView,
        QApplication,
        QCheckBox,
        QComboBox,
        QDialog,
        QDialogButtonBox,
        QFileDialog,
        QFrame,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QProgressBar,
        QScrollArea,
        QTextBrowser,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )
except ImportError:
    QApplication = None  # type: ignore[assignment]

try:
    import whisper
except (ImportError, OSError) as e:
    # OSError can be raised if PyTorch DLLs fail to load (e.g., missing VC++ Redistributable)
    whisper = None  # type: ignore[assignment]

try:
    import pysubs2
except ImportError:
    pysubs2 = None  # type: ignore[assignment]

try:
    from imdb import Cinemagoer as _Cinemagoer
    _CINEMAGOER_AVAILABLE = True
except ImportError:
    _Cinemagoer = None  # type: ignore[assignment]
    _CINEMAGOER_AVAILABLE = False


def probe_ai_runtime() -> Tuple[bool, List[str], Dict[str, str]]:
    """Probe AI dependencies in the *current* interpreter at runtime."""
    global whisper, pysubs2

    missing: List[str] = []
    details: Dict[str, str] = {}

    try:
        torch_mod = importlib.import_module("torch")
        details["torch"] = str(getattr(torch_mod, "__version__", "installed"))
    except Exception as exc:
        missing.append("openai-whisper / torch")
        details["torch_error"] = f"{type(exc).__name__}: {exc}"

    try:
        whisper_mod = importlib.import_module("whisper")
        whisper = whisper_mod  # type: ignore[assignment]
        details["whisper"] = str(getattr(whisper_mod, "__version__", "installed"))
    except Exception as exc:
        if "openai-whisper / torch" not in missing:
            missing.append("openai-whisper / torch")
        details["whisper_error"] = f"{type(exc).__name__}: {exc}"

    try:
        pysubs2_mod = importlib.import_module("pysubs2")
        pysubs2 = pysubs2_mod  # type: ignore[assignment]
        details["pysubs2"] = str(getattr(pysubs2_mod, "VERSION", "installed"))
    except Exception as exc:
        missing.append("pysubs2")
        details["pysubs2_error"] = f"{type(exc).__name__}: {exc}"

    return len(missing) == 0, missing, details

VIDEO_EXTENSIONS = {
    ".mp4",
    ".m4v",
    ".mov",
    ".mkv",
    ".avi",
    ".wmv",
    ".flv",
    ".webm",
    ".mpg",
    ".mpeg",
    ".ts",
    ".m2ts",
}

SUBTITLE_EXTENSIONS = {".srt", ".ass", ".ssa", ".vtt", ".sub", ".ttml"}
MP4_FAMILY = {".mp4", ".m4v", ".mov"}
TEXT_SUBTITLE_CODECS = {
    "subrip",
    "ass",
    "ssa",
    "webvtt",
    "mov_text",
    "text",
    "ttml",
}
SUBTITLE_CODEC_EXT = {
    "subrip": ".srt",
    "ass": ".ass",
    "ssa": ".ssa",
    "webvtt": ".vtt",
    "mov_text": ".srt",
    "text": ".srt",
    "ttml": ".ttml",
    "dvd_subtitle": ".sub",
    "hdmv_pgs_subtitle": ".sup",
    "pgs": ".sup",
}
HELP_DOC_NAME = "SUBTITLE_TOOL_HELP.md"
SETTINGS_FILE = ".subtitle_tool_settings.json"


@dataclass
class ScanRecord:
    path: str
    embedded_subtitle_streams: int
    sidecar_subtitles: List[str]


@dataclass
class OperationSummary:
    action: str
    scanned: int = 0
    processed: int = 0
    skipped: int = 0
    failed: int = 0
    details: List[Dict[str, str]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, object]:
        return {
            "action": self.action,
            "scanned": self.scanned,
            "processed": self.processed,
            "skipped": self.skipped,
            "failed": self.failed,
            "details": self.details,
        }


class SubtitleProcessor:
    def __init__(
        self,
        ffmpeg_bin: Optional[str] = None,
        ffprobe_bin: Optional[str] = None,
        log_callback: Optional[Callable[[str], None]] = None,
    ) -> None:
        self.ffmpeg_bin = ffmpeg_bin or "ffmpeg"
        self.ffprobe_bin = ffprobe_bin or "ffprobe"
        self.log_callback = log_callback
        # Cache for IMDB episode name lookups - avoids repeated network requests
        # Key: "show_name_lower|season|episode", Value: episode title or None
        self._episode_name_cache: Dict[str, Optional[str]] = {}

    def _log(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)

    def check_dependencies(self) -> Dict[str, object]:
        ffmpeg = shutil.which(self.ffmpeg_bin)
        ffprobe = shutil.which(self.ffprobe_bin)
        return {
            "ffmpeg_found": bool(ffmpeg),
            "ffprobe_found": bool(ffprobe),
            "ffmpeg_path": ffmpeg or "",
            "ffprobe_path": ffprobe or "",
        }

    def _iter_video_files(self, folders: List[str], recursive: bool) -> Iterable[Path]:
        seen: set[str] = set()
        for folder in folders:
            root = Path(folder).expanduser().resolve()
            if not root.exists() or not root.is_dir():
                self._log(f"Skipping invalid folder: {root}")
                continue

            iterator = root.rglob("*") if recursive else root.glob("*")
            for entry in iterator:
                if not entry.is_file():
                    continue
                if entry.suffix.lower() not in VIDEO_EXTENSIONS:
                    continue
                key = str(entry)
                if key in seen:
                    continue
                seen.add(key)
                yield entry

    def _normalize_video_file(self, value: str) -> Optional[Path]:
        path = Path(value).expanduser().resolve()
        if not path.exists() or not path.is_file():
            self._log(f"Skipping invalid file: {path}")
            return None
        if path.suffix.lower() not in VIDEO_EXTENSIONS:
            self._log(f"Skipping non-video file: {path}")
            return None
        return path

    def _iter_target_videos(
        self,
        folders: List[str],
        recursive: bool,
        target_files: Optional[List[str]] = None,
    ) -> Iterable[Path]:
        seen: set[str] = set()
        for video in self._iter_video_files(folders, recursive):
            key = str(video)
            if key in seen:
                continue
            seen.add(key)
            yield video

        for raw in target_files or []:
            normalized = self._normalize_video_file(raw)
            if normalized is None:
                continue
            key = str(normalized)
            if key in seen:
                continue
            seen.add(key)
            yield normalized

    def _find_sidecar_subtitles(self, video_path: Path) -> List[Path]:
        candidates: List[Path] = []
        stem = video_path.stem
        for ext in SUBTITLE_EXTENSIONS:
            exact = video_path.with_suffix(ext)
            if exact.exists() and exact.is_file():
                candidates.append(exact)

        for item in sorted(video_path.parent.glob(f"{stem}.*")):
            if item == video_path:
                continue
            suffix = item.suffix.lower()
            if suffix in SUBTITLE_EXTENSIONS and item not in candidates:
                candidates.append(item)

        for item in sorted(video_path.parent.glob(f"{stem}.embedded_sub*.srt")):
            if item not in candidates:
                candidates.append(item)

        return sorted(candidates)

    def _run_command(self, args: List[str]) -> subprocess.CompletedProcess[str]:
        self._log("Running: " + " ".join(args))
        return subprocess.run(
            args,
            capture_output=True,
            text=True,
            check=False,
        )

    def _probe_subtitle_streams(self, video_path: Path) -> List[Dict[str, object]]:
        cmd = [
            self.ffprobe_bin,
            "-v",
            "error",
            "-select_streams",
            "s",
            "-show_entries",
            "stream=index,codec_name:stream_tags=language,title",
            "-of",
            "json",
            str(video_path),
        ]
        result = self._run_command(cmd)
        if result.returncode != 0:
            self._log(f"ffprobe failed for {video_path}: {result.stderr.strip()}")
            return []

        try:
            payload = json.loads(result.stdout or "{}")
            streams = payload.get("streams", [])
            if isinstance(streams, list):
                return streams
        except json.JSONDecodeError:
            self._log(f"Failed parsing ffprobe output for {video_path}")
        return []

    def _subtitle_extension_for_codec(self, codec_name: Optional[str]) -> str:
        if not codec_name:
            return ".srt"
        return SUBTITLE_CODEC_EXT.get(codec_name.lower(), ".srt")

    def _is_text_subtitle_codec(self, codec_name: Optional[str]) -> bool:
        return (codec_name or "").lower() in TEXT_SUBTITLE_CODECS

    @staticmethod
    def _strip_subtitle_tags(text: str) -> str:
        text = re.sub(r"<[^>]+>", "", text)
        text = re.sub(r"\{.*?\}", "", text)
        return text

    def _plain_text_from_subtitle(self, subtitle_path: Path, codec_name: Optional[str]) -> str:
        try:
            content = subtitle_path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            return ""

        ext = subtitle_path.suffix.lower()
        if ext in {".ass", ".ssa"}:
            lines: List[str] = []
            for line in content.splitlines():
                if not line.strip().lower().startswith("dialogue:"):
                    continue
                parts = line.split(",", 9)
                text = parts[9] if len(parts) >= 10 else line
                text = text.replace("\\N", " ").replace("\\n", " ")
                text = self._strip_subtitle_tags(text).strip()
                if text:
                    lines.append(text)
            return "\n".join(lines)

        if ext == ".ttml":
            try:
                root = ElementTree.fromstring(content)
                texts = [node.text.strip() for node in root.iter() if node.text and node.text.strip()]
                return "\n".join(texts)
            except ElementTree.ParseError:
                pass

        lines = []
        for line in content.splitlines():
            raw = line.strip()
            if not raw:
                continue
            if raw.isdigit():
                continue
            if "-->" in raw:
                continue
            if raw.upper().startswith("WEBVTT"):
                continue
            if re.match(r"^\d{2}:\d{2}:\d{2}[,\.]\d{3}", raw):
                continue
            cleaned = self._strip_subtitle_tags(raw).strip()
            if cleaned:
                lines.append(cleaned)
        return "\n".join(lines)

    def _write_plaintext_version(self, subtitle_path: Path, codec_name: Optional[str]) -> Optional[Path]:
        if not self._is_text_subtitle_codec(codec_name):
            return None
        text = self._plain_text_from_subtitle(subtitle_path, codec_name)
        if not text.strip():
            return None
        txt_path = subtitle_path.with_suffix(".txt")
        try:
            txt_path.write_text(text, encoding="utf-8")
            return txt_path
        except OSError:
            return None

    def scan_videos(
        self,
        folders: List[str],
        recursive: bool = True,
        target_files: Optional[List[str]] = None,
        only_with_embedded: bool = False,
    ) -> List[ScanRecord]:
        output: List[ScanRecord] = []
        for video in self._iter_target_videos(folders, recursive, target_files=target_files):
            streams = self._probe_subtitle_streams(video)
            if only_with_embedded and not streams:
                continue
            sidecars = self._find_sidecar_subtitles(video)
            output.append(
                ScanRecord(
                    path=str(video),
                    embedded_subtitle_streams=len(streams),
                    sidecar_subtitles=[str(p) for p in sidecars],
                )
            )
        return output

    def _build_output_paths(self, source: Path, suffix: str, overwrite: bool) -> tuple[Path, Optional[Path]]:
        if overwrite:
            temp_output = source.with_name(f"{source.stem}.tmp_subtitle_tool{source.suffix}")
            return temp_output, source

        desired = source.with_name(f"{source.stem}{suffix}{source.suffix}")
        if not desired.exists():
            return desired, None

        index = 1
        while True:
            candidate = source.with_name(f"{source.stem}{suffix}_{index}{source.suffix}")
            if not candidate.exists():
                return candidate, None
            index += 1

    def _extract_subtitles_for_restore(self, video: Path, stream_count: int) -> List[Path]:
        extracted: List[Path] = []
        for stream_idx in range(stream_count):
            out_file = video.with_name(f"{video.stem}.embedded_sub{stream_idx + 1}.srt")
            cmd = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                "error",
                "-nostats",
                "-i",
                str(video),
                "-map",
                f"0:s:{stream_idx}",
                str(out_file),
            ]
            result = self._run_command(cmd)
            if result.returncode == 0:
                extracted.append(out_file)
            else:
                self._log(
                    f"Could not extract subtitle stream {stream_idx} from {video.name}. "
                    f"This can happen with image-based subtitles."
                )
        return extracted

    def extract_embedded_subtitles(
        self,
        folders: List[str],
        recursive: bool = True,
        overwrite: bool = False,
        output_suffix: str = ".embedded_sub",
        export_txt: bool = True,
        target_files: Optional[List[str]] = None,
    ) -> OperationSummary:
        summary = OperationSummary(action="extract")

        for video in self._iter_target_videos(folders, recursive, target_files=target_files):
            summary.scanned += 1
            streams = self._probe_subtitle_streams(video)
            if not streams:
                summary.skipped += 1
                summary.details.append({"file": str(video), "status": "skipped", "reason": "no subtitle streams"})
                continue

            extracted_count = 0
            skipped_count = 0
            failed_count = 0
            txt_count = 0

            for stream_idx, stream in enumerate(streams):
                codec_name = str(stream.get("codec_name") or "")
                tags = stream.get("tags") or {}
                language = ""
                if isinstance(tags, dict):
                    language = str(tags.get("language") or "").strip()
                language_slug = f".{language}" if language else ""
                ext = self._subtitle_extension_for_codec(codec_name)
                out_file = video.with_name(
                    f"{video.stem}{output_suffix}{stream_idx + 1}{language_slug}{ext}"
                )

                if out_file.exists() and not overwrite:
                    skipped_count += 1
                    continue

                cmd = [
                    self.ffmpeg_bin,
                    "-y" if overwrite else "-n",
                    "-loglevel",
                    "error",
                    "-nostats",
                    "-i",
                    str(video),
                    "-map",
                    f"0:s:{stream_idx}",
                ]

                if codec_name and not self._is_text_subtitle_codec(codec_name):
                    cmd.extend(["-c:s", "copy"])

                cmd.append(str(out_file))
                result = self._run_command(cmd)

                if result.returncode == 0 and out_file.exists():
                    extracted_count += 1
                    if export_txt and self._is_text_subtitle_codec(codec_name):
                        if self._write_plaintext_version(out_file, codec_name):
                            txt_count += 1
                else:
                    failed_count += 1
                    self._log(
                        f"Failed extracting subtitle stream {stream_idx + 1} from {video.name}: "
                        f"{result.stderr.strip()}"
                    )

            if extracted_count > 0:
                summary.processed += 1
                reason = f"extracted {extracted_count} subtitle stream(s)"
                if txt_count:
                    reason += f", wrote {txt_count} .txt file(s)"
                if skipped_count:
                    reason += f", skipped {skipped_count} existing"
                if failed_count:
                    reason += f", failed {failed_count}"
                summary.details.append({"file": str(video), "status": "processed", "reason": reason})
            elif skipped_count > 0:
                summary.skipped += 1
                summary.details.append(
                    {"file": str(video), "status": "skipped", "reason": "all subtitle outputs exist"}
                )
            else:
                summary.failed += 1
                summary.details.append({"file": str(video), "status": "failed", "reason": "extraction failed"})

        return summary

    def remove_embedded_subtitles(
        self,
        folders: List[str],
        recursive: bool = True,
        overwrite: bool = False,
        output_suffix: str = "_nosubs",
        extract_for_restore: bool = True,
        target_files: Optional[List[str]] = None,
    ) -> OperationSummary:
        summary = OperationSummary(action="remove")

        for video in self._iter_target_videos(folders, recursive, target_files=target_files):
            summary.scanned += 1
            streams = self._probe_subtitle_streams(video)
            stream_count = len(streams)

            if stream_count == 0:
                summary.skipped += 1
                summary.details.append({"file": str(video), "status": "skipped", "reason": "no subtitle streams"})
                continue

            if extract_for_restore:
                extracted = self._extract_subtitles_for_restore(video, stream_count)
                if extracted:
                    self._log(f"Extracted {len(extracted)} subtitle stream(s) for restore: {video.name}")

            output_path, replace_target = self._build_output_paths(video, output_suffix, overwrite)
            cmd = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                "error",
                "-nostats",
                "-i",
                str(video),
                "-map",
                "0",
                "-map",
                "-0:s",
                "-c",
                "copy",
                str(output_path),
            ]
            result = self._run_command(cmd)

            if result.returncode != 0:
                summary.failed += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "failed",
                        "reason": result.stderr.strip() or "ffmpeg failed",
                    }
                )
                continue

            if replace_target:
                output_path.replace(replace_target)

            summary.processed += 1
            summary.details.append(
                {
                    "file": str(video),
                    "status": "processed",
                    "reason": f"removed {stream_count} subtitle stream(s)",
                }
            )

        return summary

    def include_subtitles(
        self,
        folders: List[str],
        recursive: bool = True,
        overwrite: bool = False,
        output_suffix: str = "_withsubs",
        target_files: Optional[List[str]] = None,
        manual_sidecars: Optional[Dict[str, List[str]]] = None,
    ) -> OperationSummary:
        summary = OperationSummary(action="include")
        normalized_sidecars: Dict[str, List[Path]] = {}
        for video_key, sidecar_paths in (manual_sidecars or {}).items():
            sidecars: List[Path] = []
            for candidate in sidecar_paths:
                path = Path(candidate).expanduser().resolve()
                if path.exists() and path.is_file() and path.suffix.lower() in SUBTITLE_EXTENSIONS:
                    sidecars.append(path)
            if sidecars:
                normalized_sidecars[str(Path(video_key).expanduser().resolve())] = sidecars

        for video in self._iter_target_videos(folders, recursive, target_files=target_files):
            summary.scanned += 1
            manual = normalized_sidecars.get(str(video))
            sidecars = manual if manual else self._find_sidecar_subtitles(video)
            if not sidecars:
                summary.skipped += 1
                summary.details.append({"file": str(video), "status": "skipped", "reason": "no sidecar subtitles"})
                continue

            output_path, replace_target = self._build_output_paths(video, output_suffix, overwrite)

            cmd: List[str] = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                "error",
                "-nostats",
                "-i",
                str(video),
            ]

            for sub in sidecars:
                cmd.extend(["-i", str(sub)])

            # Drop existing subtitle streams and then map sidecars.
            cmd.extend(["-map", "0", "-map", "-0:s"])
            for idx in range(len(sidecars)):
                cmd.extend(["-map", str(idx + 1)])

            cmd.extend(["-c:v", "copy", "-c:a", "copy", "-c:d", "copy"])
            subtitle_codec = "mov_text" if video.suffix.lower() in MP4_FAMILY else "srt"
            cmd.extend(["-c:s", subtitle_codec])

            for idx in range(len(sidecars)):
                cmd.extend([f"-metadata:s:s:{idx}", "language=eng"])

            cmd.append(str(output_path))
            result = self._run_command(cmd)

            if result.returncode != 0:
                summary.failed += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "failed",
                        "reason": result.stderr.strip() or "ffmpeg failed",
                    }
                )
                continue

            if replace_target:
                output_path.replace(replace_target)

            summary.processed += 1
            summary.details.append(
                {
                    "file": str(video),
                    "status": "processed",
                    "reason": f"embedded {len(sidecars)} sidecar subtitle file(s)",
                }
            )

        return summary
    
    def convert_format(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        target_format: str,  # "mkv" or "mp4"
        overwrite: bool = False,
        output_suffix: str = "_converted",
    ) -> OperationSummary:
        """Convert video files between mkv and mp4 formats while preserving all streams."""
        summary = OperationSummary(action=f"convert_to_{target_format}")
        
        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)
        
        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)
        
        if not videos:
            self._log("No video files found to convert")
            return summary
        
        for video in videos:
            current_ext = video.suffix.lower()
            target_ext = f".{target_format}"
            
            # Skip if already target format
            if current_ext == target_ext:
                summary.skipped += 1
                summary.details.append({
                    "file": str(video),
                    "status": "skipped",
                    "reason": f"already {target_format} format"
                })
                continue
            
            output_path = video.with_name(f"{video.stem}{output_suffix}{target_ext}")
            replace_target = None
            
            if not overwrite and output_path.exists():
                summary.skipped += 1
                summary.details.append({
                    "file": str(video),
                    "status": "skipped",
                    "reason": "output exists and overwrite=False"
                })
                continue
            
            if overwrite and not output_suffix:
                replace_target = video
            
            self._log(f"Converting {video.name} to {target_format.upper()}...")
            
            # Build ffmpeg command for format conversion
            cmd = [self.ffmpeg_bin, "-i", str(video)]
            
            # For MP4, use specific codecs
            if target_format == "mp4":
                cmd.extend(["-c:v", "copy", "-c:a", "copy", "-c:s", "mov_text"])
            else:
                # For MKV, just copy everything
                cmd.extend(["-c", "copy"])
            
            cmd.extend(["-y", str(output_path)])
            result = self._run_command(cmd)
            
            if result.returncode != 0:
                summary.failed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "failed",
                    "reason": result.stderr.strip() or "conversion failed"
                })
                continue
            
            if replace_target:
                output_path.replace(replace_target)
            
            summary.processed += 1
            summary.details.append({
                "file": str(video),
                "status": "converted",
                "reason": f"converted from {current_ext} to {target_ext}"
            })
        
        return summary

    def _load_organize_rules(self, config_path: Optional[str]) -> Dict[str, object]:
        if not config_path:
            return {}

        path = Path(config_path).expanduser().resolve()
        if not path.exists() or not path.is_file():
            self._log(f"Organize config not found: {path}. Using built-in behavior.")
            return {}

        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            self._log(f"Failed to load organize config {path}: {exc}. Using built-in behavior.")
            return {}

        if not isinstance(payload, dict):
            self._log(f"Organize config must be a JSON object: {path}. Using built-in behavior.")
            return {}

        self._log(f"Loaded organize config: {path}")
        return payload

    def _clean_media_name(self, value: str, rules: Dict[str, object]) -> str:
        cleaned = value

        if bool(rules.get("normalize_separators", False)):
            cleaned = re.sub(r"[._]+", " ", cleaned)

        if bool(rules.get("strip_bracketed", False)):
            cleaned = re.sub(r"\[[^\]]*\]", " ", cleaned)
            cleaned = re.sub(r"\([^\)]*\)", " ", cleaned)
            cleaned = re.sub(r"\{[^\}]*\}", " ", cleaned)

        cutoff_tokens = rules.get("cutoff_tokens", [])
        if isinstance(cutoff_tokens, list) and cutoff_tokens:
            cutoff_index = len(cleaned)
            for token in cutoff_tokens:
                if not isinstance(token, str) or not token.strip():
                    continue
                token_text = token.strip()
                try:
                    match = re.search(token_text, cleaned, re.IGNORECASE)
                except re.error:
                    match = re.search(re.escape(token_text), cleaned, re.IGNORECASE)
                if match and match.start() < cutoff_index:
                    cutoff_index = match.start()
            cleaned = cleaned[:cutoff_index]

        cleanup_regex = rules.get("cleanup_regex", [])
        if isinstance(cleanup_regex, list):
            for entry in cleanup_regex:
                if not isinstance(entry, dict):
                    continue
                pattern = entry.get("pattern")
                replace = entry.get("replace", "")
                if not isinstance(pattern, str) or not pattern:
                    continue
                if not isinstance(replace, str):
                    replace = str(replace)
                try:
                    cleaned = re.sub(pattern, replace, cleaned)
                except re.error:
                    continue

        cleaned = re.sub(r"\s+", " ", cleaned).strip(" ._-")
        return cleaned or value.strip()

    def _extract_tv_episode_info(self, file_stem: str, tv_rules: Dict[str, object]) -> Optional[Dict[str, int]]:
        configured_patterns = tv_rules.get("patterns", [])
        pattern_entries: List[Dict[str, object]] = []

        if isinstance(configured_patterns, list):
            for entry in configured_patterns:
                if isinstance(entry, str):
                    pattern_entries.append({
                        "pattern": entry,
                        "season_group": "season",
                        "episode_group": "episode",
                    })
                elif isinstance(entry, dict):
                    pattern_entries.append({
                        "pattern": entry.get("pattern"),
                        "season_group": entry.get("season_group", "season"),
                        "episode_group": entry.get("episode_group", "episode"),
                    })

        if not pattern_entries:
            pattern_entries = [{"pattern": r"([Ss]\d{2}[Ee]\d{2})", "season_group": None, "episode_group": None}]

        for entry in pattern_entries:
            pattern = entry.get("pattern")
            if not isinstance(pattern, str) or not pattern:
                continue

            try:
                regex = re.compile(pattern, re.IGNORECASE)
            except re.error:
                continue

            match = regex.search(file_stem)
            if not match:
                continue

            season_group = entry.get("season_group")
            episode_group = entry.get("episode_group")

            season: Optional[int] = None
            episode: Optional[int] = None

            if isinstance(season_group, str) and isinstance(episode_group, str):
                try:
                    season = int(match.group(season_group))
                    episode = int(match.group(episode_group))
                except (IndexError, KeyError, TypeError, ValueError):
                    season = None
                    episode = None

            if season is None or episode is None:
                full_match = match.group(0)
                sxe = re.search(r"[Ss](\d{1,2})[Ee](\d{1,3})", full_match)
                if sxe:
                    season = int(sxe.group(1))
                    episode = int(sxe.group(2))
                elif match.lastindex and match.lastindex >= 2:
                    try:
                        season = int(match.group(1))
                        episode = int(match.group(2))
                    except (TypeError, ValueError):
                        season = None
                        episode = None

            if season is None or episode is None:
                continue

            return {
                "season": season,
                "episode": episode,
                "match_start": int(match.start()),
            }

        return None

    def _render_tv_stem(self, template: str, season: int, episode: int, clean_name: str, episode_name: Optional[str] = None) -> str:
        season_episode = f"S{season:02d}E{episode:02d}"
        ep_name_val = episode_name or ""
        try:
            rendered = template.format(
                season=season,
                episode=episode,
                season_episode=season_episode,
                clean_name=clean_name,
                episode_name=ep_name_val,
            )
        except Exception:
            rendered = season_episode

        # If episode_name was empty, clean up orphaned trailing separators like " - "
        if not ep_name_val:
            rendered = re.sub(r"(\s*-\s*){2,}", " - ", rendered)  # collapse double-dash
            rendered = re.sub(r"[\s\-_]+$", "", rendered)          # trim trailing separators

        rendered = re.sub(r"\s+", " ", rendered).strip(" ._-")
        return rendered or season_episode

    def _lookup_episode_name(self, show_name: str, season: int, episode: int) -> Optional[str]:
        """Look up an episode's title from IMDB using cinemagoer.

        Requires ``pip install cinemagoer``.  Results are cached in
        ``self._episode_name_cache`` for the lifetime of this processor
        instance so that multiple files from the same series only hit the
        network once per series (the episode list is fetched in bulk).
        """
        if not _CINEMAGOER_AVAILABLE:
            self._log(
                "cinemagoer is not installed; IMDB lookup skipped. "
                "Install it with:  pip install cinemagoer"
            )
            return None

        cache_key = f"{show_name.lower()}|{season}|{episode}"
        if cache_key in self._episode_name_cache:
            return self._episode_name_cache[cache_key]

        result: Optional[str] = None
        try:
            ia = _Cinemagoer()
            search_results = ia.search_movie(show_name)

            # Prefer an explicit TV series match in the first few results
            series = None
            for r in search_results[:5]:
                if r.get("kind") in ("tv series", "tv mini series"):
                    series = r
                    break
            if series is None and search_results:
                series = search_results[0]

            if series is None:
                self._log(f"IMDB: No results for '{show_name}'")
                self._episode_name_cache[cache_key] = None
                return None

            self._log(f"IMDB: Fetching episodes for '{series.get('title', show_name)}' …")
            ia.update(series, "episodes")
            episodes_by_season = series.get("episodes", {})

            if season in episodes_by_season and episode in episodes_by_season[season]:
                ep = episodes_by_season[season][episode]
                result = ep.get("title") or None
                if result:
                    self._log(
                        f"IMDB: S{season:02d}E{episode:02d} of '{show_name}' → '{result}'"
                    )
            else:
                self._log(
                    f"IMDB: S{season:02d}E{episode:02d} not found for '{show_name}'"
                )

        except Exception as exc:
            self._log(
                f"IMDB lookup error for '{show_name}' S{season:02d}E{episode:02d}: {exc}"
            )

        self._episode_name_cache[cache_key] = result
        return result

    def organize_media(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        organize_movies: bool = True,
        organize_tv: bool = True,
        organize_config_path: Optional[str] = None,
    ) -> OperationSummary:
        """Organize media files - move movies up one level and normalize TV episode names."""
        summary = OperationSummary(action="organize_media")

        rules = self._load_organize_rules(organize_config_path)
        movie_rules = rules.get("movie_name", {}) if isinstance(rules.get("movie_name"), dict) else {}
        tv_rules = rules.get("tv_name", {}) if isinstance(rules.get("tv_name"), dict) else {}
        tv_template = str(tv_rules.get("template", "{season_episode}"))

        # IMDB lookup config (optional) - requires: pip install cinemagoer
        _imdb_cfg = tv_rules.get("imdb_lookup", {})
        _imdb_enabled = isinstance(_imdb_cfg, dict) and bool(_imdb_cfg.get("enabled", False))
        
        # Process folders
        for folder_str in folders:
            folder = Path(folder_str).expanduser().resolve()
            if not folder.exists() or not folder.is_dir():
                continue
            
            # Organize movies - look for folders with single video file
            if organize_movies:
                for subfolder in folder.iterdir():
                    if not subfolder.is_dir():
                        continue
                    
                    video_files = [f for f in subfolder.iterdir() 
                                 if f.is_file() and f.suffix.lower() in VIDEO_EXTENSIONS]
                    
                    if not video_files:
                        continue
                    
                    # Check if it looks like a TV show
                    is_tv = any(
                        self._extract_tv_episode_info(f.stem, tv_rules) is not None
                        for f in video_files
                    )
                    if is_tv:
                        continue
                    
                    # Move movie file(s) up one level
                    for video_file in video_files:
                        try:
                            file_ext = video_file.suffix
                            movie_name = subfolder.name
                            if movie_rules:
                                movie_name = self._clean_media_name(movie_name, movie_rules)

                            new_filename = f"{movie_name}{file_ext}"
                            new_path = folder / new_filename
                            
                            # Handle conflicts
                            counter = 1
                            while new_path.exists():
                                new_filename = f"{movie_name}_{counter}{file_ext}"
                                new_path = folder / new_filename
                                counter += 1
                            
                            self._log(f"Moving {video_file.name} -> {new_filename}")
                            shutil.move(str(video_file), str(new_path))
                            summary.processed += 1
                            summary.details.append({
                                "file": str(video_file),
                                "status": "organized",
                                "reason": f"moved to {new_filename}"
                            })
                        except Exception as e:
                            summary.failed += 1
                            summary.details.append({
                                "file": str(video_file),
                                "status": "failed",
                                "reason": str(e)
                            })
                    
                    # Try to remove empty folder
                    try:
                        if not any(subfolder.iterdir()):
                            os.rmdir(subfolder)
                    except:
                        pass
            
            # Organize TV shows - rename to S##E## format
            if organize_tv:
                for root, dirs, files in os.walk(folder):
                    root_path = Path(root)
                    video_files = [f for f in files if Path(f).suffix.lower() in VIDEO_EXTENSIONS]
                    
                    for video_file in video_files:
                        stem = Path(video_file).stem
                        episode_info = self._extract_tv_episode_info(stem, tv_rules)
                        if not episode_info:
                            continue
                        
                        try:
                            old_path = root_path / video_file
                            file_ext = Path(video_file).suffix
                            season = int(episode_info["season"])
                            episode = int(episode_info["episode"])
                            match_start = int(episode_info["match_start"])

                            if tv_rules:
                                prefix = stem[:match_start].strip()
                                clean_source = prefix or root_path.name
                                clean_name = self._clean_media_name(clean_source, tv_rules)

                                # Optional IMDB lookup for real episode title
                                episode_name: Optional[str] = None
                                if _imdb_enabled:
                                    episode_name = self._lookup_episode_name(clean_name, season, episode)

                                new_stem = self._render_tv_stem(tv_template, season, episode, clean_name, episode_name)
                            else:
                                new_stem = f"S{season:02d}E{episode:02d}"

                            new_filename = f"{new_stem}{file_ext}"
                            new_path = root_path / new_filename
                            
                            # Skip if already named correctly
                            if old_path == new_path:
                                continue
                            
                            # Handle conflicts
                            counter = 1
                            while new_path.exists():
                                new_filename = f"{new_stem}_{counter}{file_ext}"
                                new_path = root_path / new_filename
                                counter += 1
                            
                            self._log(f"Renaming {video_file} -> {new_filename}")
                            os.rename(str(old_path), str(new_path))
                            summary.processed += 1
                            summary.details.append({
                                "file": str(old_path),
                                "status": "renamed",
                                "reason": f"renamed to {new_filename}"
                            })
                        except Exception as e:
                            summary.failed += 1
                            summary.details.append({
                                "file": str(old_path),
                                "status": "failed",
                                "reason": str(e)
                            })
        
        summary.scanned = summary.processed + summary.failed + summary.skipped
        return summary
    
    def repair_metadata(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        create_backup: bool = True,
    ) -> OperationSummary:
        """Repair corrupted video metadata by rebuilding containers with ffmpeg."""
        summary = OperationSummary(action="repair_metadata")
        
        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)
        
        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)
        
        if not videos:
            self._log("No video files found to repair")
            return summary
        
        # Create backup directory if needed
        backup_dir = None
        if create_backup:
            backup_dir = Path.cwd() / "media_repair_backups"
            backup_dir.mkdir(exist_ok=True)
        
        for video in videos:
            # Skip very small files
            if video.stat().st_size < 10_000_000:  # 10 MB
                summary.skipped += 1
                summary.details.append({
                    "file": str(video),
                    "status": "skipped",
                    "reason": "file too small (possibly incomplete)"
                })
                continue
            
            self._log(f"Repairing {video.name}...")
            
            # Create backup if requested
            if create_backup and backup_dir:
                backup_path = backup_dir / video.name
                if not backup_path.exists():
                    try:
                        self._log(f"  Creating backup...")
                        shutil.copy2(str(video), str(backup_path))
                    except Exception as e:
                        self._log(f"  Warning: Could not create backup: {e}")
            
            # Create temp file for repair
            temp_file = video.with_name(f"{video.stem}_repair_temp{video.suffix}")
            
            try:
                # Use aggressive error handling to rebuild container
                cmd = [
                    self.ffmpeg_bin,
                    "-fflags", "+genpts",  # Generate presentation timestamps
                    "-err_detect", "ignore_err",  # Ignore errors
                    "-i", str(video),
                    "-c", "copy",  # Copy all streams
                    "-y",
                    str(temp_file)
                ]
                
                result = self._run_command(cmd)
                
                # Check if output file was created and has reasonable size
                if temp_file.exists() and temp_file.stat().st_size > 1000:
                    # Replace original with repaired version
                    os.remove(str(video))
                    os.rename(str(temp_file), str(video))
                    summary.processed += 1
                    summary.details.append({
                        "file": str(video),
                        "status": "repaired",
                        "reason": "container rebuilt successfully"
                    })
                    self._log(f"  Successfully repaired {video.name}")
                else:
                    if temp_file.exists():
                        os.remove(str(temp_file))
                    summary.failed += 1
                    summary.details.append({
                        "file": str(video),
                        "status": "failed",
                        "reason": "repair output invalid or too small"
                    })
            
            except Exception as e:
                if temp_file.exists():
                    try:
                        os.remove(str(temp_file))
                    except:
                        pass
                summary.failed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "failed",
                    "reason": str(e)
                })
        
        return summary
    
    def generate_subtitles(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        model_size: str = "base",
        output_format: str = "srt",
        language: Optional[str] = None,
    ) -> OperationSummary:
        """Generate subtitles from video audio using Whisper AI."""
        summary = OperationSummary(action="generate_subtitles")
        
        if whisper is None:
            self._log("ERROR: openai-whisper not installed. Run: pip install openai-whisper")
            summary.failed = 1
            summary.details.append({
                "file": "N/A",
                "status": "failed",
                "reason": "Whisper library not installed"
            })
            return summary
        
        if pysubs2 is None:
            self._log("ERROR: pysubs2 not installed. Run: pip install pysubs2")
            summary.failed = 1
            summary.details.append({
                "file": "N/A",
                "status": "failed",
                "reason": "pysubs2 library not installed"
            })
            return summary
        
        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)
        
        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)
        
        if not videos:
            self._log("No video files found to generate subtitles")
            return summary
        
        # Load Whisper model
        try:
            self._log(f"Loading Whisper model: {model_size}...")
            model = whisper.load_model(model_size)
            self._log(f"Model loaded successfully")
        except Exception as e:
            self._log(f"ERROR loading Whisper model: {e}")
            summary.failed = len(videos)
            for video in videos:
                summary.details.append({
                    "file": str(video),
                    "status": "failed",
                    "reason": f"Failed to load model: {e}"
                })
            return summary
        
        for video in videos:
            self._log(f"Generating subtitles for {video.name}...")
            
            # Check if subtitle already exists
            output_path = video.with_suffix(f".{output_format}")
            if output_path.exists():
                summary.skipped += 1
                summary.details.append({
                    "file": str(video),
                    "status": "skipped",
                    "reason": f"subtitle file already exists: {output_path.name}",
                    "output_path": str(output_path)
                })
                continue
            
            try:
                # Transcribe video
                self._log(f"  Transcribing audio (this may take a while)...")
                transcribe_options = {"task": "transcribe"}
                if language:
                    transcribe_options["language"] = language
                
                result = model.transcribe(str(video), **transcribe_options)
                
                # Create subtitle file using pysubs2
                subs = pysubs2.SSAFile()
                for segment in result["segments"]:
                    event = pysubs2.SSAEvent(
                        start=int(segment["start"] * 1000),  # Convert to milliseconds
                        end=int(segment["end"] * 1000),
                        text=segment["text"].strip()
                    )
                    subs.append(event)
                
                # Save subtitle file
                subs.save(str(output_path))
                
                detected_lang = result.get("language", "unknown")
                self._log(f"  Generated {output_path.name} (language: {detected_lang})")
                self._log(f"  Saved subtitle to: {output_path}")
                
                summary.processed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "generated",
                    "reason": f"created {output_path.name} with {len(result['segments'])} segments",
                    "output_path": str(output_path)
                })
            
            except Exception as e:
                self._log(f"  ERROR: {e}")
                summary.failed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "failed",
                    "reason": str(e)
                })
        
        return summary


class JobPayload(BaseModel):
    folders: List[str] = Field(default_factory=list)
    target_files: List[str] = Field(default_factory=list)
    manual_sidecars: Dict[str, List[str]] = Field(default_factory=dict)
    recursive: bool = True
    overwrite: bool = False
    output_suffix: str = ""
    extract_for_restore: bool = True
    export_txt: bool = True
    scan_only_embedded: bool = False


@dataclass
class JobRecord:
    id: str
    action: str
    status: str
    created_at: str
    updated_at: str
    result: Optional[Dict[str, object]] = None
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, object]:
        return {
            "id": self.id,
            "action": self.action,
            "status": self.status,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "result": self.result,
            "error": self.error,
        }


class JobManager:
    def __init__(self, processor: SubtitleProcessor) -> None:
        self.processor = processor
        self.executor = ThreadPoolExecutor(max_workers=2)
        self.lock = threading.Lock()
        self.jobs: Dict[str, JobRecord] = {}

    @staticmethod
    def _now() -> str:
        return datetime.utcnow().isoformat() + "Z"

    def submit(self, action: str, payload: JobPayload) -> JobRecord:
        if not payload.folders and not payload.target_files:
            raise ValueError("folders and target_files cannot both be empty")

        job_id = uuid.uuid4().hex
        job = JobRecord(
            id=job_id,
            action=action,
            status="queued",
            created_at=self._now(),
            updated_at=self._now(),
        )

        with self.lock:
            self.jobs[job_id] = job

        self.executor.submit(self._run_job, job_id, action, payload)
        return job

    def _run_job(self, job_id: str, action: str, payload: JobPayload) -> None:
        self._update(job_id, status="running")
        try:
            if action == "scan":
                rows = self.processor.scan_videos(
                    payload.folders,
                    recursive=payload.recursive,
                    target_files=payload.target_files,
                    only_with_embedded=payload.scan_only_embedded,
                )
                result: Dict[str, object] = {
                    "action": "scan",
                    "count": len(rows),
                    "files": [
                        {
                            "path": r.path,
                            "embedded_subtitle_streams": r.embedded_subtitle_streams,
                            "sidecar_subtitles": r.sidecar_subtitles,
                        }
                        for r in rows
                    ],
                }
            elif action == "remove":
                summary = self.processor.remove_embedded_subtitles(
                    folders=payload.folders,
                    recursive=payload.recursive,
                    overwrite=payload.overwrite,
                    output_suffix=payload.output_suffix or "_nosubs",
                    extract_for_restore=payload.extract_for_restore,
                    target_files=payload.target_files,
                )
                result = summary.to_dict()
            elif action == "include":
                summary = self.processor.include_subtitles(
                    folders=payload.folders,
                    recursive=payload.recursive,
                    overwrite=payload.overwrite,
                    output_suffix=payload.output_suffix or "_withsubs",
                    target_files=payload.target_files,
                    manual_sidecars=payload.manual_sidecars,
                )
                result = summary.to_dict()
            elif action == "extract":
                summary = self.processor.extract_embedded_subtitles(
                    folders=payload.folders,
                    recursive=payload.recursive,
                    overwrite=payload.overwrite,
                    output_suffix=payload.output_suffix or ".embedded_sub",
                    export_txt=payload.export_txt,
                    target_files=payload.target_files,
                )
                result = summary.to_dict()
            else:
                raise ValueError(f"Unsupported action: {action}")

            self._update(job_id, status="completed", result=result)
        except Exception as exc:
            self._update(job_id, status="failed", error=f"{exc}\n{traceback.format_exc()}")

    def _update(
        self,
        job_id: str,
        status: Optional[str] = None,
        result: Optional[Dict[str, object]] = None,
        error: Optional[str] = None,
    ) -> None:
        with self.lock:
            job = self.jobs[job_id]
            if status:
                job.status = status
            if result is not None:
                job.result = result
            if error is not None:
                job.error = error
            job.updated_at = self._now()

    def get(self, job_id: str) -> Optional[JobRecord]:
        with self.lock:
            return self.jobs.get(job_id)

    def list(self) -> List[JobRecord]:
        with self.lock:
            return list(self.jobs.values())


def create_api_app():
    if FastAPI is None:
        raise RuntimeError("FastAPI is not installed. Install requirements first.")

    processor = SubtitleProcessor()
    manager = JobManager(processor)
    app = FastAPI(title="Subtitle Tool API", version="1.0.0")

    @app.get("/health")
    def health() -> Dict[str, object]:
        return {
            "status": "ok",
            "dependencies": processor.check_dependencies(),
        }

    @app.get("/jobs")
    def list_jobs() -> Dict[str, object]:
        return {"jobs": [job.to_dict() for job in manager.list()]}

    @app.get("/jobs/{job_id}")
    def get_job(job_id: str) -> Dict[str, object]:
        job = manager.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="job not found")
        return job.to_dict()

    @app.post("/jobs/scan")
    def start_scan(payload: JobPayload) -> Dict[str, object]:
        try:
            job = manager.submit("scan", payload)
            return {"job_id": job.id, "status": job.status}
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/jobs/remove")
    def start_remove(payload: JobPayload) -> Dict[str, object]:
        try:
            job = manager.submit("remove", payload)
            return {"job_id": job.id, "status": job.status}
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/jobs/include")
    def start_include(payload: JobPayload) -> Dict[str, object]:
        try:
            job = manager.submit("include", payload)
            return {"job_id": job.id, "status": job.status}
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/jobs/extract")
    def start_extract(payload: JobPayload) -> Dict[str, object]:
        try:
            job = manager.submit("extract", payload)
            return {"job_id": job.id, "status": job.status}
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    return app


if QApplication is not None:

    class HelpDialog(QDialog):
        def __init__(self, parent: Optional[QWidget] = None) -> None:
            super().__init__(parent)
            self.setWindowTitle("Subtitle Tool Help")
            self.resize(900, 700)
            self._build_ui()

        def _build_ui(self) -> None:
            layout = QVBoxLayout(self)
            
            help_text = self._load_help_content()
            
            browser = QTextBrowser()
            browser.setOpenExternalLinks(True)
            browser.setMarkdown(help_text)
            layout.addWidget(browser)
            
            button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
            button_box.rejected.connect(self.close)
            layout.addWidget(button_box)

        def _load_help_content(self) -> str:
            help_path = Path(__file__).resolve().with_name(HELP_DOC_NAME)
            if help_path.exists():
                try:
                    return help_path.read_text(encoding="utf-8")
                except OSError:
                    pass
            return self._get_default_help()

        def _get_default_help(self) -> str:
            return """# Subtitle Tool Help

## Quick Start

1. Add folders or specific video files to process
2. Choose an action: Scan, Remove, Include, or Extract
3. Configure options as needed
4. Click the action button to start processing

## UI Sections

- **Target Folders**: Add directories to scan for video files
- **Target Video Files**: Add specific files (supports drag & drop)
- **Manual Subtitle Files**: Map specific subtitle files to videos
- **Options**: Configure processing behavior

## Actions

- **Scan**: Inspect videos for subtitle streams
- **Remove**: Strip embedded subtitle streams
- **Include**: Embed subtitle files into videos
- **Extract**: Export subtitle streams to files

For detailed information, see SUBTITLE_TOOL_HELP.md in the installation directory.
"""


    class TutorialOverlay(QWidget):
        """Semi-transparent overlay that highlights specific widgets during tutorial."""
        def __init__(self, parent: QWidget) -> None:
            super().__init__(parent)
            self.target_widget: Optional[QWidget] = None
            
            # Set up debug logging to file
            import os
            self.debug_log_path = os.path.join(os.path.dirname(__file__), 'tutorial_debug.log')
            
            # Make the widget transparent for mouse events but visible
            self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            # Don't use stylesheet - it interferes with custom painting
            self.setAutoFillBackground(False)
            
            # Animation state
            self.flash_phase = 0.0  # 0.0 to 1.0
            self.flash_direction = 1  # 1 for increasing, -1 for decreasing
            self.animation_timer = QTimer(self)
            self.animation_timer.timeout.connect(self._animate_flash)
            self.animation_timer.setInterval(30)  # ~33 FPS
            self.hide()
        
        def _log(self, message: str) -> None:
            """Log debug messages to file (minimal logging)"""
            # Only log errors and important events
            pass

        def highlight_widget(self, widget: Optional[QWidget]) -> None:
            self.target_widget = widget
            # Position overlay to cover entire parent
            if self.parent():
                self.setGeometry(self.parent().rect())
            self.show()
            self.raise_()
            self.flash_phase = 0.0
            self.flash_direction = 1
            
            # Always start animation and show overlay
            # If widget is None, we'll just show full overlay without cutout
            if widget is not None:
                self.animation_timer.start()
            else:
                # No widget to highlight, but keep overlay visible
                self.animation_timer.stop()
            
            self.update()

        def _animate_flash(self) -> None:
            """Update flash animation state."""
            # Pulse speed: larger value = faster
            step = 0.08
            self.flash_phase += step * self.flash_direction
            
            # Reverse direction at boundaries
            if self.flash_phase >= 1.0:
                self.flash_phase = 1.0
                self.flash_direction = -1
            elif self.flash_phase <= 0.0:
                self.flash_phase = 0.0
                self.flash_direction = 1
            
            self.update()

        def paintEvent(self, event) -> None:  # type: ignore[override]
            from PyQt6.QtGui import QPainter, QPen, QBrush, QPainterPath
            from PyQt6.QtCore import QRectF, Qt
            
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            # If we have a target widget, create cutout effect
            if self.target_widget is not None and self.target_widget.isVisible():
                # Get target widget position relative to overlay
                target_rect = self.target_widget.geometry()
                widget_parent = self.target_widget.parent()
                
                # Calculate position in overlay coordinates
                if widget_parent is not None and self.parent():
                    # Map from target widget's parent coordinates to overlay coordinates
                    top_left = widget_parent.mapTo(self.parent(), target_rect.topLeft())
                    target_rect.moveTo(top_left)
                
                # Add padding around highlighted area
                padding = 10
                highlight_rect = target_rect.adjusted(-padding, -padding, padding, padding)
                
                # Create path for the cutout effect using fill rule
                path = QPainterPath()
                path.setFillRule(Qt.FillRule.OddEvenFill)
                path.addRect(QRectF(self.rect()))
                path.addRect(QRectF(highlight_rect))
                
                # Fill with semi-transparent overlay (everything except highlighted area)
                painter.fillPath(path, QColor(0, 0, 0, 180))
                
                # Calculate animated values
                min_brightness = 120
                max_brightness = 255
                brightness = int(min_brightness + (max_brightness - min_brightness) * self.flash_phase)
                
                # Vary border width
                min_width = 5
                max_width = 8
                border_width = min_width + int((max_width - min_width) * self.flash_phase)
                
                # Draw main animated border
                pen = QPen(QColor(0, brightness, 255), border_width)
                painter.setPen(pen)
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.drawRect(highlight_rect)
                
                # Draw inner glow effect
                glow_alpha = int(200 * self.flash_phase)
                inner_pen = QPen(QColor(100, 220, 255, glow_alpha), 3)
                painter.setPen(inner_pen)
                inner_rect = highlight_rect.adjusted(5, 5, -5, -5)
                painter.drawRect(inner_rect)
                
                # Draw corner accent marks for extra emphasis
                corner_size = 20
                accent_alpha = int(255 * self.flash_phase)
                corner_pen = QPen(QColor(255, 255, 255, accent_alpha), 3)
                painter.setPen(corner_pen)
                
                # Top-left corner
                painter.drawLine(highlight_rect.left(), highlight_rect.top(), 
                               highlight_rect.left() + corner_size, highlight_rect.top())
                painter.drawLine(highlight_rect.left(), highlight_rect.top(), 
                               highlight_rect.left(), highlight_rect.top() + corner_size)
                
                # Top-right corner
                painter.drawLine(highlight_rect.right(), highlight_rect.top(), 
                               highlight_rect.right() - corner_size, highlight_rect.top())
                painter.drawLine(highlight_rect.right(), highlight_rect.top(), 
                               highlight_rect.right(), highlight_rect.top() + corner_size)
                
                # Bottom-left corner
                painter.drawLine(highlight_rect.left(), highlight_rect.bottom(), 
                               highlight_rect.left() + corner_size, highlight_rect.bottom())
                painter.drawLine(highlight_rect.left(), highlight_rect.bottom(), 
                               highlight_rect.left(), highlight_rect.bottom() - corner_size)
                
                # Bottom-right corner
                painter.drawLine(highlight_rect.right(), highlight_rect.bottom(), 
                               highlight_rect.right() - corner_size, highlight_rect.bottom())
                painter.drawLine(highlight_rect.right(), highlight_rect.bottom(), 
                               highlight_rect.right(), highlight_rect.bottom() - corner_size)
            else:
                # No target widget, just draw full overlay
                painter.fillRect(self.rect(), QColor(0, 0, 0, 180))


    class TutorialDialog(QDialog):
        """Interactive tutorial that walks through each UI element."""
        def __init__(self, main_window: "SubtitleToolWindow", parent: Optional[QWidget] = None) -> None:
            super().__init__(parent or main_window)
            self.main_window = main_window
            self.current_step = 0
            self.overlay: Optional[TutorialOverlay] = None
            self.setWindowTitle("Tutorial")
            self.setModal(False)
            self.resize(400, 250)
            self._build_ui()
            self._define_tutorial_steps()
            
        def _build_ui(self) -> None:
            layout = QVBoxLayout(self)
            
            # Create scroll area for content
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_area.setFrameShape(QFrame.Shape.NoFrame)
            
            # Create content widget for scroll area
            content_widget = QWidget()
            content_layout = QVBoxLayout(content_widget)
            
            self.step_label = QLabel()
            self.step_label.setWordWrap(True)
            font = QFont()
            font.setPointSize(10)
            self.step_label.setFont(font)
            content_layout.addWidget(self.step_label)
            
            self.description_label = QLabel()
            self.description_label.setWordWrap(True)
            content_layout.addWidget(self.description_label)
            
            content_layout.addStretch()
            
            # Set content widget in scroll area
            scroll_area.setWidget(content_widget)
            layout.addWidget(scroll_area)
            
            # Button layout stays outside scroll area at bottom
            button_layout = QHBoxLayout()
            self.prev_button = QPushButton("Previous")
            self.next_button = QPushButton("Next")
            self.finish_button = QPushButton("Finish")
            
            self.prev_button.clicked.connect(self._prev_step)
            self.next_button.clicked.connect(self._next_step)
            self.finish_button.clicked.connect(self._finish_tutorial)
            
            button_layout.addWidget(self.prev_button)
            button_layout.addStretch()
            button_layout.addWidget(self.next_button)
            button_layout.addWidget(self.finish_button)
            layout.addLayout(button_layout)

        def _define_tutorial_steps(self) -> None:
            self.steps = [
                {
                    "title": "Welcome to Subtitle Tool!",
                    "description": "This tutorial will walk you through the main features of the application. Click Next to continue.",
                    "widget": None,
                },
                {
                    "title": "Target Folders",
                    "description": "Add folders here to scan for video files. Use 'Add Folder' to browse, or process entire directories recursively.",
                    "widget": self.main_window.folder_list,
                },
                {
                    "title": "Target Video Files",
                    "description": "Add specific video files here. You can drag and drop files directly, or use 'Add Files' to browse. Great for processing individual files.",
                    "widget": self.main_window.target_file_list,
                },
                {
                    "title": "Manual Subtitle Assignment",
                    "description": "Select a video file above, then add subtitle files here to manually map subtitles to that specific video. Supports drag and drop.",
                    "widget": self.main_window.manual_subtitle_list,
                },
                {
                    "title": "Processing Options",
                    "description": "Configure how videos are processed: recursive scanning, overwrite behavior, subtitle extraction, and output file naming.",
                    "widget": self.main_window.recursive_checkbox,
                },
                {
                    "title": "Scan Videos",
                    "description": "Inspects videos to see which have embedded subtitle streams and which have matching sidecar subtitle files. No files are modified.",
                    "widget": self.main_window.scan_button,
                },
                {
                    "title": "Remove Embedded Subtitles",
                    "description": "Strips subtitle streams from video files. Optionally extracts them first for backup. Keeps video and audio streams intact.",
                    "widget": self.main_window.remove_button,
                },
                {
                    "title": "Include Subtitles Back In",
                    "description": "Embeds sidecar subtitle files into video containers. For MP4 files, uses mov_text codec for compatibility.",
                    "widget": self.main_window.include_button,
                },
                {
                    "title": "Extract Embedded Subtitles",
                    "description": "Exports embedded subtitle streams to separate files. Can also create plain text versions for easy preview.",
                    "widget": self.main_window.extract_button,
                },
                {
                    "title": "Swiss Army Knife Tools",
                    "description": "Additional video tools for format conversion, organization, and repair. All tools work with your selected folders and files.",
                    "widget": None,
                },
                {
                    "title": "Convert to MKV/MP4",
                    "description": "Convert videos between MKV and MP4 formats while preserving all streams (video, audio, subtitles). Great for device compatibility.",
                    "widget": self.main_window.convert_mkv_button,
                },
                {
                    "title": "Organize Media",
                    "description": "Automatically organize media files: move movies up one level and rename TV show episodes to S##E## format. Configure options with checkboxes.",
                    "widget": self.main_window.organize_button,
                },
                {
                    "title": "Repair Metadata",
                    "description": "Rebuild corrupted video containers using FFmpeg. Useful for fixing playback issues with torrented files. Creates backups by default.",
                    "widget": self.main_window.repair_button,
                },
                {
                    "title": "Generate Subtitles with Whisper AI",
                    "description": "Use Whisper AI to automatically generate subtitles from video audio. Choose from 7 model sizes (tiny to large-v3) for speed vs accuracy trade-off. Runs 100% locally - no internet or API keys needed! Note: Requires ~10GB disk space for full installation.",
                    "widget": self.main_window.generate_button if self.main_window.use_ai else None,
                },
                {
                    "title": "Organization Options",
                    "description": "Toggle whether to organize movies, TV shows, and whether to create backups during repair operations.",
                    "widget": self.main_window.organize_movies_checkbox,
                },
                {
                    "title": "Help & Tutorial",
                    "description": "Click 'Open Help' anytime to view detailed documentation. Use 'Show Tutorial' to see this walkthrough again.",
                    "widget": self.main_window.help_button,
                },
                {
                    "title": "Theme Toggle",
                    "description": "Switch between Light and Dark modes to suit your preference. Your choice is saved automatically.",
                    "widget": self.main_window.theme_toggle_button,
                },
                {
                    "title": "Show Tutorial Button",
                    "description": "Click this anytime to see the tutorial again if you need a refresher on any features.",
                    "widget": self.main_window.tutorial_button,
                },
                {
                    "title": "Error History",
                    "description": "View all logged errors here. Errors are tracked automatically and shown on startup. You can clear them from this dialog.",
                    "widget": self.main_window.error_history_button,
                },
                {
                    "title": "Activity Log",
                    "description": "All operations and results are logged here. Use this to track progress and troubleshoot issues.",
                    "widget": self.main_window.log_box,
                },
                {
                    "title": "Tutorial Complete!",
                    "description": "You're ready to start processing videos. Add folders or files, choose an action, and click the corresponding button. Happy subtitle managing!",
                    "widget": None,
                },
            ]

        def showEvent(self, event) -> None:  # type: ignore[override]
            super().showEvent(event)
            if not self.overlay:
                # Create overlay as child of main window's central widget
                central = self.main_window.centralWidget()
                if central:
                    self.overlay = TutorialOverlay(central)
                else:
                    self.overlay = TutorialOverlay(self.main_window)
            
            # Update overlay geometry to match parent
            if self.overlay.parent():
                self.overlay.setGeometry(self.overlay.parent().rect())
                self.overlay.raise_()
            
            self.current_step = 0
            self._show_step()

        def closeEvent(self, event) -> None:  # type: ignore[override]
            if self.overlay:
                self.overlay.animation_timer.stop()
                self.overlay.hide()
            super().closeEvent(event)

        def _show_step(self) -> None:
            if self.current_step < 0 or self.current_step >= len(self.steps):
                return
            
            step = self.steps[self.current_step]
            
            # Log step details
            import os
            import datetime
            log_path = os.path.join(os.path.dirname(__file__), 'tutorial_debug.log')
            with open(log_path, 'a') as f:
                widget = step.get("widget")
                f.write(f"\n[{datetime.datetime.now().strftime('%H:%M:%S.%f')[:-3]}] ===== Showing Step {self.current_step} =====\n")
                f.write(f"  Title: {step['title']}\n")
                f.write(f"  Widget from step.get('widget'): {widget}\n")
                f.write(f"  Widget type: {type(widget).__name__ if widget else 'None'}\n")
                if widget:
                    f.write(f"  Widget isVisible: {widget.isVisible()}\n")
                    f.write(f"  Widget geometry: {widget.geometry()}\n")
            
            self.step_label.setText(f"<b>Step {self.current_step + 1} of {len(self.steps)}: {step['title']}</b>")
            self.description_label.setText(step["description"])
            
            self.prev_button.setEnabled(self.current_step > 0)
            self.next_button.setEnabled(self.current_step < len(self.steps) - 1)
            self.finish_button.setEnabled(True)
            
            if self.overlay:
                # Ensure overlay size matches its parent
                if self.overlay.parent():
                    self.overlay.setGeometry(self.overlay.parent().rect())
                    self.overlay.raise_()
                
                widget = step.get("widget")
                self.overlay.highlight_widget(widget)
                
                # Scroll to widget if it's in the main window's scroll area
                if widget is not None and hasattr(self.main_window, 'scroll_area'):
                    # Ensure the widget is visible in the scroll area
                    self.main_window.scroll_area.ensureWidgetVisible(widget, 50, 50)
            
            # Keep tutorial dialog centered and visible
            self._position_centered()

        def _position_centered(self) -> None:
            """Position tutorial dialog in the center-top area of the screen."""
            if self.main_window:
                # Position relative to main window center-top
                main_window_rect = self.main_window.geometry()
                x = main_window_rect.center().x() - self.width() // 2
                y = main_window_rect.top() + 50  # 50px from top
                
                # Ensure it stays on screen
                screen = QApplication.primaryScreen().geometry()
                x = max(10, min(x, screen.right() - self.width() - 10))
                y = max(10, min(y, screen.bottom() - self.height() - 10))
                
                self.move(x, y)

        def _next_step(self) -> None:
            if self.current_step < len(self.steps) - 1:
                self.current_step += 1
                # Skip steps with disabled features (widget is None and step has a widget field)
                while (self.current_step < len(self.steps) - 1 and 
                       "widget" in self.steps[self.current_step] and
                       self.steps[self.current_step].get("widget") is None and
                       self.steps[self.current_step].get("title") != "Welcome to Subtitle Tool!" and
                       self.steps[self.current_step].get("title") != "Tutorial Complete!"):
                    self.current_step += 1
                self._show_step()

        def _prev_step(self) -> None:
            if self.current_step > 0:
                self.current_step -= 1
                # Skip steps with disabled features (widget is None and step has a widget field)
                while (self.current_step > 0 and 
                       "widget" in self.steps[self.current_step] and
                       self.steps[self.current_step].get("widget") is None and
                       self.steps[self.current_step].get("title") != "Welcome to Subtitle Tool!" and
                       self.steps[self.current_step].get("title") != "Tutorial Complete!"):
                    self.current_step -= 1
                self._show_step()

        def _finish_tutorial(self) -> None:
            if self.overlay:
                self.overlay.animation_timer.stop()
                self.overlay.hide()
            self.accept()


    class DragDropPathListWidget(QListWidget):
        files_dropped = pyqtSignal(list)

        def __init__(self, allowed_extensions: set[str], parent: Optional[QWidget] = None) -> None:
            super().__init__(parent)
            self.allowed_extensions = {ext.lower() for ext in allowed_extensions}
            self.setAcceptDrops(True)
            self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        def dragEnterEvent(self, event) -> None:  # type: ignore[override]
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return
            super().dragEnterEvent(event)

        def dragMoveEvent(self, event) -> None:  # type: ignore[override]
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return
            super().dragMoveEvent(event)

        def dropEvent(self, event) -> None:  # type: ignore[override]
            if not event.mimeData().hasUrls():
                super().dropEvent(event)
                return

            dropped: List[str] = []
            for url in event.mimeData().urls():
                if not url.isLocalFile():
                    continue
                path = Path(url.toLocalFile()).expanduser().resolve()
                if not path.exists() or not path.is_file():
                    continue
                if path.suffix.lower() not in self.allowed_extensions:
                    continue
                dropped.append(str(path))

            if dropped:
                self.files_dropped.emit(dropped)
            event.acceptProposedAction()

    class ProcessorThread(QThread):
        log_message = pyqtSignal(str)
        finished_result = pyqtSignal(dict)
        failed = pyqtSignal(str)

        def __init__(self, action: str, options: Dict[str, object]) -> None:
            super().__init__()
            self.action = action
            self.options = options

        def run(self) -> None:
            try:
                processor = SubtitleProcessor(log_callback=self.log_message.emit)
                folders = self.options["folders"]
                target_files = self.options.get("target_files", [])
                recursive = bool(self.options.get("recursive", True))
                overwrite = bool(self.options.get("overwrite", False))

                if self.action == "scan":
                    rows = processor.scan_videos(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        only_with_embedded=bool(self.options.get("scan_only_embedded", False)),
                    )
                    payload = {
                        "action": "scan",
                        "count": len(rows),
                        "files": [
                            {
                                "path": r.path,
                                "embedded_subtitle_streams": r.embedded_subtitle_streams,
                                "sidecar_subtitles": r.sidecar_subtitles,
                            }
                            for r in rows
                        ],
                    }
                elif self.action == "remove":
                    summary = processor.remove_embedded_subtitles(
                        folders=folders,
                        recursive=recursive,
                        overwrite=overwrite,
                        output_suffix=str(self.options.get("output_suffix", "_nosubs")),
                        extract_for_restore=bool(self.options.get("extract_for_restore", True)),
                        target_files=target_files,
                    )
                    payload = summary.to_dict()
                elif self.action == "include":
                    summary = processor.include_subtitles(
                        folders=folders,
                        recursive=recursive,
                        overwrite=overwrite,
                        output_suffix=str(self.options.get("output_suffix", "_withsubs")),
                        target_files=target_files,
                        manual_sidecars=dict(self.options.get("manual_sidecars", {})),
                    )
                    payload = summary.to_dict()
                elif self.action == "extract":
                    summary = processor.extract_embedded_subtitles(
                        folders=folders,
                        recursive=recursive,
                        overwrite=overwrite,
                        output_suffix=str(self.options.get("output_suffix", ".embedded_sub")),
                        export_txt=bool(self.options.get("export_txt", True)),
                        target_files=target_files,
                    )
                    payload = summary.to_dict()
                elif self.action == "convert_mkv" or self.action == "convert_mp4":
                    target_format = "mkv" if self.action == "convert_mkv" else "mp4"
                    summary = processor.convert_format(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        target_format=target_format,
                        overwrite=overwrite,
                        output_suffix=str(self.options.get("output_suffix", "_converted")),
                    )
                    payload = summary.to_dict()
                elif self.action == "organize":
                    summary = processor.organize_media(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        organize_movies=bool(self.options.get("organize_movies", True)),
                        organize_tv=bool(self.options.get("organize_tv", True)),
                        organize_config_path=str(self.options.get("organize_config_path", "")).strip() or None,
                    )
                    payload = summary.to_dict()
                elif self.action == "repair":
                    summary = processor.repair_metadata(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        create_backup=bool(self.options.get("create_backup", True)),
                    )
                    payload = summary.to_dict()
                elif self.action == "generate":
                    summary = processor.generate_subtitles(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        model_size=str(self.options.get("model_size", "base")),
                        output_format=str(self.options.get("output_format", "srt")),
                        language=self.options.get("language"),
                    )
                    payload = summary.to_dict()
                else:
                    raise ValueError(f"Unsupported action: {self.action}")

                self.finished_result.emit(payload)
            except Exception as exc:
                self.failed.emit(f"{exc}\n{traceback.format_exc()}")


    class SubtitleToolWindow(QMainWindow):
        def __init__(self, clear_memory: bool = False, use_ai: Optional[bool] = None) -> None:
            super().__init__()
            self.worker: Optional[ProcessorThread] = None
            self.manual_sidecars_by_video: Dict[str, List[str]] = {}
            self._active_manual_video: Optional[str] = None
            self.settings_path = Path(__file__).resolve().parent / SETTINGS_FILE
            self.clear_memory = clear_memory
            self.setWindowTitle("Video Swiss Army Knife - Subtitle & Media Tools")
            self.resize(980, 700)
            
            # Load settings
            settings = self._load_settings()
            self.dark_mode = settings.get("dark_mode", True)
            
            # Handle use_ai setting
            if use_ai is not None:
                # Command-line flag updates the setting
                settings["use_ai"] = use_ai
                self._save_settings(settings)
            
            self.ai_runtime_available, self.ai_missing_dependencies, self.ai_probe_details = probe_ai_runtime()

            # Use saved setting, default to enabled only when runtime dependencies are available.
            default_ai = self.ai_runtime_available
            requested_ai = bool(settings.get("use_ai", default_ai))
            self.ai_requested_but_unavailable = requested_ai and not self.ai_runtime_available

            # Show AI controls only when AI is both requested and importable in this venv.
            self.use_ai = requested_ai and self.ai_runtime_available
            
            self._apply_theme()
            self._build_ui()
            self._log(f"Python executable: {sys.executable}")
            self._log(f"Script path: {Path(__file__).resolve()}")
            if self.ai_runtime_available:
                self._log("AI runtime dependencies detected in current environment.")
            elif self.ai_requested_but_unavailable:
                missing = ", ".join(self.ai_missing_dependencies)
                self._log(f"AI requested but unavailable in current environment (missing: {missing}).")
                if "torch_error" in self.ai_probe_details:
                    self._log(f"torch import error: {self.ai_probe_details['torch_error']}")
                if "whisper_error" in self.ai_probe_details:
                    self._log(f"whisper import error: {self.ai_probe_details['whisper_error']}")
                if "pysubs2_error" in self.ai_probe_details:
                    self._log(f"pysubs2 import error: {self.ai_probe_details['pysubs2_error']}")
            self._check_for_errors()
            self._check_first_run()
            self._load_ui_state()

        def _apply_theme(self) -> None:
            """Apply light or dark theme based on current setting."""
            if self.dark_mode:
                self._apply_dark_theme()
            else:
                self._apply_light_theme()

        def _apply_light_theme(self) -> None:
            """Apply light theme to the application."""
            light_stylesheet = """
                QMainWindow, QDialog, QWidget {
                    background-color: #f0f0f0;
                    color: #000000;
                }
                QGroupBox {
                    border: 1px solid #c0c0c0;
                    border-radius: 4px;
                    margin-top: 10px;
                    padding-top: 10px;
                    font-weight: bold;
                    color: #000000;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 3px;
                }
                QPushButton {
                    background-color: #ffffff;
                    border: 1px solid #c0c0c0;
                    border-radius: 4px;
                    padding: 5px 15px;
                    color: #000000;
                }
                QPushButton:hover {
                    background-color: #e5f3ff;
                }
                QPushButton:pressed {
                    background-color: #cce4ff;
                }
                QPushButton:disabled {
                    background-color: #f5f5f5;
                    color: #a0a0a0;
                }
                QListWidget, QTextEdit, QLineEdit {
                    background-color: #ffffff;
                    border: 1px solid #c0c0c0;
                    border-radius: 3px;
                    padding: 5px;
                    color: #000000;
                }
                QListWidget::item:selected {
                    background-color: #0078d4;
                    color: #ffffff;
                }
                QCheckBox {
                    color: #000000;
                }
                QLabel {
                    color: #000000;
                }
                QProgressBar {
                    border: 1px solid #c0c0c0;
                    border-radius: 3px;
                    text-align: center;
                    background-color: #ffffff;
                    color: #000000;
                }
                QProgressBar::chunk {
                    background-color: #0078d4;
                }
                QScrollBar:vertical {
                    border: none;
                    background: #f0f0f0;
                    width: 12px;
                    margin: 0px;
                }
                QScrollBar::handle:vertical {
                    background: #c0c0c0;
                    min-height: 20px;
                    border-radius: 6px;
                }
                QScrollBar::handle:vertical:hover {
                    background: #a0a0a0;
                }
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                    height: 0px;
                }
                QScrollBar:horizontal {
                    border: none;
                    background: #f0f0f0;
                    height: 12px;
                    margin: 0px;
                }
                QScrollBar::handle:horizontal {
                    background: #c0c0c0;
                    min-width: 20px;
                    border-radius: 6px;
                }
                QScrollBar::handle:horizontal:hover {
                    background: #a0a0a0;
                }
                QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                    width: 0px;
                }
                QMessageBox {
                    background-color: #f0f0f0;
                }
                QMessageBox QLabel {
                    color: #000000;
                }
            """
            self.setStyleSheet(light_stylesheet)

        def _apply_dark_theme(self) -> None:
            """Apply dark theme to the application."""
            dark_stylesheet = """
                QMainWindow, QDialog, QWidget {
                    background-color: #2b2b2b;
                    color: #e0e0e0;
                }
                QGroupBox {
                    border: 1px solid #555;
                    border-radius: 4px;
                    margin-top: 10px;
                    padding-top: 10px;
                    font-weight: bold;
                    color: #e0e0e0;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 3px;
                }
                QPushButton {
                    background-color: #3c3c3c;
                    border: 1px solid #555;
                    border-radius: 4px;
                    padding: 5px 15px;
                    color: #e0e0e0;
                }
                QPushButton:hover {
                    background-color: #4a4a4a;
                }
                QPushButton:pressed {
                    background-color: #2a2a2a;
                }
                QPushButton:disabled {
                    background-color: #2a2a2a;
                    color: #666;
                }
                QListWidget, QTextEdit, QLineEdit {
                    background-color: #1e1e1e;
                    border: 1px solid #555;
                    border-radius: 3px;
                    padding: 5px;
                    color: #e0e0e0;
                }
                QListWidget::item:selected {
                    background-color: #094771;
                }
                QCheckBox {
                    color: #e0e0e0;
                }
                QLabel {
                    color: #e0e0e0;
                }
                QProgressBar {
                    border: 1px solid #555;
                    border-radius: 3px;
                    text-align: center;
                    background-color: #1e1e1e;
                    color: #e0e0e0;
                }
                QProgressBar::chunk {
                    background-color: #0d7d3c;
                }
                QScrollBar:vertical {
                    border: none;
                    background: #2b2b2b;
                    width: 12px;
                    margin: 0px;
                }
                QScrollBar::handle:vertical {
                    background: #555;
                    min-height: 20px;
                    border-radius: 6px;
                }
                QScrollBar::handle:vertical:hover {
                    background: #666;
                }
                QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                    height: 0px;
                }
                QScrollBar:horizontal {
                    border: none;
                    background: #2b2b2b;
                    height: 12px;
                    margin: 0px;
                }
                QScrollBar::handle:horizontal {
                    background: #555;
                    min-width: 20px;
                    border-radius: 6px;
                }
                QScrollBar::handle:horizontal:hover {
                    background: #666;
                }
                QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                    width: 0px;
                }
                QMessageBox {
                    background-color: #2b2b2b;
                }
                QMessageBox QLabel {
                    color: #e0e0e0;
                }
            """
            self.setStyleSheet(dark_stylesheet)

        def _build_ui(self) -> None:
            # Create main scroll area
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            
            container = QWidget()
            self.scroll_area.setWidget(container)
            self.setCentralWidget(self.scroll_area)
            
            root = QVBoxLayout(container)
            root.setSizeConstraint(QVBoxLayout.SizeConstraint.SetMinAndMaxSize)

            folder_box = QGroupBox("Target Folders")
            folder_layout = QVBoxLayout(folder_box)
            self.folder_list = QListWidget()
            self.folder_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
            self.folder_list.setMinimumHeight(100)
            self.folder_list.setMaximumHeight(200)
            folder_layout.addWidget(self.folder_list)

            folder_buttons = QHBoxLayout()
            add_button = QPushButton("Add Folder")
            remove_button = QPushButton("Remove Selected")
            clear_button = QPushButton("Clear")
            folder_buttons.addWidget(add_button)
            folder_buttons.addWidget(remove_button)
            folder_buttons.addWidget(clear_button)
            folder_layout.addLayout(folder_buttons)

            add_button.clicked.connect(self._add_folder)
            remove_button.clicked.connect(self._remove_selected_folders)
            clear_button.clicked.connect(self.folder_list.clear)

            file_box = QGroupBox("Target Video Files (optional)")
            file_layout = QVBoxLayout(file_box)
            self.target_file_list = DragDropPathListWidget(allowed_extensions=VIDEO_EXTENSIONS)
            self.target_file_list.setMinimumHeight(100)
            self.target_file_list.setMaximumHeight(200)
            self.target_file_list.files_dropped.connect(self._add_target_files)
            self.target_file_list.itemSelectionChanged.connect(self._on_target_video_selection_changed)
            file_layout.addWidget(self.target_file_list)

            file_buttons = QHBoxLayout()
            add_files_button = QPushButton("Add Files")
            remove_files_button = QPushButton("Remove Selected")
            clear_files_button = QPushButton("Clear")
            file_buttons.addWidget(add_files_button)
            file_buttons.addWidget(remove_files_button)
            file_buttons.addWidget(clear_files_button)
            file_layout.addLayout(file_buttons)

            file_hint = QLabel("Tip: Drag and drop video files into this list.")
            file_layout.addWidget(file_hint)

            add_files_button.clicked.connect(self._choose_target_files)
            remove_files_button.clicked.connect(self._remove_selected_target_files)
            clear_files_button.clicked.connect(self._clear_target_files)

            subtitle_box = QGroupBox("Manual Subtitle Files for Selected Video (include mode)")
            subtitle_layout = QVBoxLayout(subtitle_box)
            self.manual_subtitle_list = DragDropPathListWidget(allowed_extensions=SUBTITLE_EXTENSIONS)
            self.manual_subtitle_list.setMinimumHeight(80)
            self.manual_subtitle_list.setMaximumHeight(150)
            self.manual_subtitle_list.files_dropped.connect(self._add_manual_subtitles)
            subtitle_layout.addWidget(self.manual_subtitle_list)

            subtitle_buttons = QHBoxLayout()
            add_sub_button = QPushButton("Add Subtitle Files")
            remove_sub_button = QPushButton("Remove Selected")
            clear_sub_button = QPushButton("Clear Current Video List")
            subtitle_buttons.addWidget(add_sub_button)
            subtitle_buttons.addWidget(remove_sub_button)
            subtitle_buttons.addWidget(clear_sub_button)
            subtitle_layout.addLayout(subtitle_buttons)

            subtitle_hint = QLabel(
                "Select one video above, then add subtitle files (or drag/drop) to force include for that video."
            )
            subtitle_layout.addWidget(subtitle_hint)

            add_sub_button.clicked.connect(self._choose_manual_subtitles)
            remove_sub_button.clicked.connect(self._remove_selected_manual_subtitles)
            clear_sub_button.clicked.connect(self._clear_manual_subtitles_for_selected_video)

            options_box = QGroupBox("Options")
            options_layout = QGridLayout(options_box)

            self.recursive_checkbox = QCheckBox("Scan folders recursively")
            self.recursive_checkbox.setChecked(True)
            self.overwrite_checkbox = QCheckBox("Overwrite original files")
            self.extract_checkbox = QCheckBox("Extract embedded subtitles before removal (for restore)")
            self.extract_checkbox.setChecked(True)
            self.export_txt_checkbox = QCheckBox("Export .txt copies for subtitles")
            self.export_txt_checkbox.setChecked(True)
            self.scan_only_embedded_checkbox = QCheckBox("Scan only files with embedded subtitles")
            self.only_selected_targets_checkbox = QCheckBox("Use only selected target video file(s)")

            self.remove_suffix_input = QLineEdit("_nosubs")
            self.include_suffix_input = QLineEdit("_withsubs")
            self.extract_suffix_input = QLineEdit(".embedded_sub")

            options_layout.addWidget(self.recursive_checkbox, 0, 0, 1, 2)
            options_layout.addWidget(self.overwrite_checkbox, 1, 0, 1, 2)
            options_layout.addWidget(self.extract_checkbox, 2, 0, 1, 2)
            options_layout.addWidget(self.export_txt_checkbox, 3, 0, 1, 2)
            options_layout.addWidget(self.scan_only_embedded_checkbox, 4, 0, 1, 2)
            options_layout.addWidget(self.only_selected_targets_checkbox, 5, 0, 1, 2)
            options_layout.addWidget(QLabel("Remove output suffix:"), 6, 0)
            options_layout.addWidget(self.remove_suffix_input, 6, 1)
            options_layout.addWidget(QLabel("Include output suffix:"), 7, 0)
            options_layout.addWidget(self.include_suffix_input, 7, 1)
            options_layout.addWidget(QLabel("Extract output suffix:"), 8, 0)
            options_layout.addWidget(self.extract_suffix_input, 8, 1)
            options_layout.addWidget(QLabel("Conversion output suffix:"), 9, 0)
            self.convert_suffix_input = QLineEdit("_converted")
            options_layout.addWidget(self.convert_suffix_input, 9, 1)

            # Swiss Army Knife section
            tools_box = QGroupBox("Video Tools (Swiss Army Knife)")
            tools_layout = QVBoxLayout(tools_box)
            
            # Organize options
            organize_options = QHBoxLayout()
            self.organize_movies_checkbox = QCheckBox("Organize Movies")
            self.organize_movies_checkbox.setChecked(True)
            self.organize_tv_checkbox = QCheckBox("Organize TV Shows")
            self.organize_tv_checkbox.setChecked(True)
            self.repair_backup_checkbox = QCheckBox("Create Backups when Repairing")
            self.repair_backup_checkbox.setChecked(True)
            organize_options.addWidget(self.organize_movies_checkbox)
            organize_options.addWidget(self.organize_tv_checkbox)
            organize_options.addWidget(self.repair_backup_checkbox)
            organize_options.addStretch()
            tools_layout.addLayout(organize_options)

            organize_config_row = QHBoxLayout()
            organize_config_row.addWidget(QLabel("Organize Rules JSON (optional):"))
            self.organize_rules_input = QLineEdit()
            self.organize_rules_input.setPlaceholderText("e.g. organize_media_rules.example.json")
            self.organize_rules_input.setToolTip(
                "Optional JSON rules for torrent-style cleanup and episode naming. "
                "Leave blank to use built-in behavior."
            )
            self.organize_rules_browse_button = QPushButton("Browse...")
            self.organize_rules_browse_button.clicked.connect(self._choose_organize_rules_file)
            organize_config_row.addWidget(self.organize_rules_input)
            organize_config_row.addWidget(self.organize_rules_browse_button)
            tools_layout.addLayout(organize_config_row)
            
            # Tool buttons
            tools_button_row1 = QHBoxLayout()
            self.convert_mkv_button = QPushButton("Convert to MKV")
            self.convert_mp4_button = QPushButton("Convert to MP4")
            self.organize_button = QPushButton("Organize Media")
            self.repair_button = QPushButton("Repair Metadata")
            tools_button_row1.addWidget(self.convert_mkv_button)
            tools_button_row1.addWidget(self.convert_mp4_button)
            tools_button_row1.addWidget(self.organize_button)
            tools_button_row1.addWidget(self.repair_button)
            tools_layout.addLayout(tools_button_row1)
            
            self.convert_mkv_button.clicked.connect(self._start_convert_mkv)
            self.convert_mp4_button.clicked.connect(self._start_convert_mp4)
            self.organize_button.clicked.connect(self._start_organize)
            self.repair_button.clicked.connect(self._start_repair)
            
            # Subtitle Generation with Whisper AI (only show if use_ai is enabled)
            if self.use_ai:
                tools_layout.addSpacing(10)
                subtitle_gen_label = QLabel("<b>Generate Subtitles (Whisper AI)</b>")
                tools_layout.addWidget(subtitle_gen_label)
                
                whisper_options = QHBoxLayout()
                whisper_options.addWidget(QLabel("Model Size:"))
                self.whisper_model_combo = QComboBox()
                self.whisper_model_combo.addItems(["tiny", "base", "small", "medium", "large", "large-v2", "large-v3"])
                self.whisper_model_combo.setCurrentText("base")
                self.whisper_model_combo.setToolTip(
                    "tiny: Fastest, least accurate (~39M params, ~72MB)\n"
                    "base: Good balance (~74M params, ~140MB)\n"
                    "small: Better accuracy (~244M params, ~460MB)\n"
                    "medium: High accuracy (~769M params, ~1.5GB)\n"
                    "large: Best accuracy (~1550M params, ~2.9GB)\n"
                    "large-v2: Improved large model (~1550M params, ~2.9GB)\n"
                    "large-v3: Latest large model (~1550M params, ~2.9GB)"
                )
                whisper_options.addWidget(self.whisper_model_combo)
                
                whisper_options.addWidget(QLabel("Language (optional):"))
                self.whisper_language_input = QLineEdit()
                self.whisper_language_input.setPlaceholderText("auto-detect")
                self.whisper_language_input.setMaximumWidth(100)
                self.whisper_language_input.setToolTip("Leave blank for auto-detection, or specify: en, es, fr, de, etc.")
                whisper_options.addWidget(self.whisper_language_input)
                
                self.generate_button = QPushButton("Generate Subtitles")
                self.generate_button.setToolTip("Generate subtitles from video audio using Whisper AI")
                whisper_options.addWidget(self.generate_button)
                whisper_options.addStretch()
                
                tools_layout.addLayout(whisper_options)
                self.generate_button.clicked.connect(self._start_generate)
            else:
                # Create dummy attributes for widgets that won't exist
                self.whisper_model_combo = None
                self.whisper_language_input = None
                self.generate_button = None

                if self.ai_requested_but_unavailable:
                    tools_layout.addSpacing(10)
                    missing = ", ".join(self.ai_missing_dependencies)
                    ai_unavailable = QLabel(
                        "AI subtitle generation is enabled in settings but not available in this Python environment.\n"
                        f"Missing: {missing}\n"
                        "Install with: pip install -r requirements_ai.txt"
                    )
                    ai_unavailable.setStyleSheet("color: #d9822b;")
                    tools_layout.addWidget(ai_unavailable)

            button_row = QHBoxLayout()
            self.scan_button = QPushButton("Scan Videos")
            self.remove_button = QPushButton("Remove Embedded Subtitles")
            self.include_button = QPushButton("Include Subtitles Back In")
            self.extract_button = QPushButton("Extract Embedded Subtitles")
            button_row.addWidget(self.scan_button)
            button_row.addWidget(self.remove_button)
            button_row.addWidget(self.include_button)
            button_row.addWidget(self.extract_button)

            self.scan_button.clicked.connect(self._start_scan)
            self.remove_button.clicked.connect(self._start_remove)
            self.include_button.clicked.connect(self._start_include)
            self.extract_button.clicked.connect(self._start_extract)
            
            help_button_row = QHBoxLayout()
            self.theme_toggle_button = QPushButton("Switch to Light Mode" if self.dark_mode else "Switch to Dark Mode")
            self.help_button = QPushButton("Open Help")
            self.tutorial_button = QPushButton("Show Tutorial")
            self.error_history_button = QPushButton("Error History")
            help_button_row.addWidget(self.theme_toggle_button)
            help_button_row.addWidget(self.help_button)
            help_button_row.addWidget(self.tutorial_button)
            help_button_row.addWidget(self.error_history_button)
            help_button_row.addStretch()
            
            self.theme_toggle_button.clicked.connect(self._toggle_theme)
            self.help_button.clicked.connect(self._open_help_dialog)
            self.tutorial_button.clicked.connect(self._show_tutorial)
            self.error_history_button.clicked.connect(self._show_error_history)

            self.progress = QProgressBar()
            self.progress.setRange(0, 1)
            self.progress.setValue(0)

            self.log_box = QTextEdit()
            self.log_box.setReadOnly(True)
            self.log_box.setMinimumHeight(150)
            self.log_box.setMaximumHeight(300)
            self.log_box.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

            root.addWidget(folder_box)
            root.addWidget(file_box)
            root.addWidget(subtitle_box)
            root.addWidget(options_box)
            root.addLayout(button_row)
            root.addWidget(tools_box)
            root.addLayout(help_button_row)
            root.addWidget(self.progress)
            root.addWidget(QLabel("Activity Log"))
            root.addWidget(self.log_box)

            dep_check = SubtitleProcessor().check_dependencies()
            if not dep_check["ffmpeg_found"]:
                self._log(
                    "WARNING: ffmpeg not detected on PATH. "
                    "Install ffmpeg before processing videos."
                )
                self._log_error(
                    "ERR001_FFMPEG_MISSING",
                    "ffmpeg binary not found on system PATH",
                    f"Expected location: {dep_check.get('ffmpeg_path', 'Not found')}"
                )
            if not dep_check["ffprobe_found"]:
                self._log(
                    "WARNING: ffprobe not detected on PATH. "
                    "Install ffmpeg before processing videos."
                )
                self._log_error(
                    "ERR002_FFPROBE_MISSING",
                    "ffprobe binary not found on system PATH",
                    f"Expected location: {dep_check.get('ffprobe_path', 'Not found')}"
                )

            if self.ai_requested_but_unavailable:
                missing = ", ".join(self.ai_missing_dependencies)
                self._log(
                    "WARNING: AI is enabled in settings but unavailable in this environment. "
                    f"Missing: {missing}"
                )
                self._log(f"Current Python environment: {sys.executable}")
                self._log("Install AI deps in this environment with: pip install -r requirements_ai.txt")

        def _log(self, message: str) -> None:
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_box.append(f"[{ts}] {message}")

        def _add_folder(self) -> None:
            folder = QFileDialog.getExistingDirectory(self, "Select Folder")
            if not folder:
                return
            existing = {self.folder_list.item(i).text() for i in range(self.folder_list.count())}
            if folder in existing:
                return
            self.folder_list.addItem(QListWidgetItem(folder))

        def _remove_selected_folders(self) -> None:
            for item in self.folder_list.selectedItems():
                row = self.folder_list.row(item)
                self.folder_list.takeItem(row)

        def _collect_folders(self) -> List[str]:
            return [self.folder_list.item(i).text() for i in range(self.folder_list.count())]

        def _collect_common_options(self) -> Dict[str, object]:
            folders = self._collect_folders()
            target_files = self._collect_target_files()
            if not folders and not target_files:
                raise ValueError("Add at least one folder or target video file before running.")
            return {
                "folders": folders,
                "target_files": target_files,
                "recursive": self.recursive_checkbox.isChecked(),
                "overwrite": self.overwrite_checkbox.isChecked(),
                "extract_for_restore": self.extract_checkbox.isChecked(),
                "export_txt": self.export_txt_checkbox.isChecked(),
                "scan_only_embedded": self.scan_only_embedded_checkbox.isChecked(),
            }

        def _set_running(self, running: bool) -> None:
            self.scan_button.setEnabled(not running)
            self.remove_button.setEnabled(not running)
            self.include_button.setEnabled(not running)
            self.extract_button.setEnabled(not running)
            self.convert_mkv_button.setEnabled(not running)
            self.convert_mp4_button.setEnabled(not running)
            self.organize_button.setEnabled(not running)
            self.repair_button.setEnabled(not running)
            if self.generate_button is not None:
                self.generate_button.setEnabled(not running)
            if running:
                self.progress.setRange(0, 0)
            else:
                self.progress.setRange(0, 1)
                self.progress.setValue(1)

        def _start_worker(self, action: str, options: Dict[str, object]) -> None:
            if self.worker and self.worker.isRunning():
                QMessageBox.warning(self, "Busy", "A task is already running.")
                return

            self.worker = ProcessorThread(action=action, options=options)
            self.worker.log_message.connect(self._log)
            self.worker.finished_result.connect(self._on_result)
            self.worker.failed.connect(self._on_error)
            self.worker.finished.connect(lambda: self._set_running(False))

            self._set_running(True)
            self._log(f"Starting action: {action}")
            self.worker.start()

        def _start_scan(self) -> None:
            try:
                options = self._collect_common_options()
                self._start_worker("scan", options)
            except ValueError as exc:
                self._log_error(
                    "ERR004_VALIDATION_FAILED",
                    "Failed to validate scan options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))

        def _start_remove(self) -> None:
            try:
                options = self._collect_common_options()
                options["output_suffix"] = self.remove_suffix_input.text().strip() or "_nosubs"
                self._start_worker("remove", options)
            except ValueError as exc:
                self._log_error(
                    "ERR005_VALIDATION_FAILED",
                    "Failed to validate remove options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))

        def _start_include(self) -> None:
            try:
                options = self._collect_common_options()
                options["output_suffix"] = self.include_suffix_input.text().strip() or "_withsubs"
                options["manual_sidecars"] = dict(self.manual_sidecars_by_video)
                self._start_worker("include", options)
            except ValueError as exc:
                self._log_error(
                    "ERR006_VALIDATION_FAILED",
                    "Failed to validate include options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))

        def _start_extract(self) -> None:
            try:
                options = self._collect_common_options()
                options["output_suffix"] = self.extract_suffix_input.text().strip() or ".embedded_sub"
                self._start_worker("extract", options)
            except ValueError as exc:
                self._log_error(
                    "ERR007_VALIDATION_FAILED",
                    "Failed to validate extract options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))
        
        def _start_convert_mkv(self) -> None:
            try:
                options = self._collect_common_options()
                options["output_suffix"] = self.convert_suffix_input.text().strip() or "_converted"
                self._start_worker("convert_mkv", options)
            except ValueError as exc:
                self._log_error(
                    "ERR008_VALIDATION_FAILED",
                    "Failed to validate MKV conversion options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))
        
        def _start_convert_mp4(self) -> None:
            try:
                options = self._collect_common_options()
                options["output_suffix"] = self.convert_suffix_input.text().strip() or "_converted"
                self._start_worker("convert_mp4", options)
            except ValueError as exc:
                self._log_error(
                    "ERR009_VALIDATION_FAILED",
                    "Failed to validate MP4 conversion options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))
        
        def _start_organize(self) -> None:
            try:
                options = self._collect_common_options()
                options["organize_movies"] = self.organize_movies_checkbox.isChecked()
                options["organize_tv"] = self.organize_tv_checkbox.isChecked()
                rules_path = self.organize_rules_input.text().strip()
                if rules_path:
                    path = Path(rules_path).expanduser().resolve()
                    if not path.exists() or not path.is_file():
                        QMessageBox.warning(self, "Validation", f"Organize rules JSON not found:\n{path}")
                        return
                    options["organize_config_path"] = str(path)
                else:
                    options["organize_config_path"] = ""
                 
                if not options["organize_movies"] and not options["organize_tv"]:
                    QMessageBox.warning(self, "Validation", "Please select at least one organization option (Movies or TV Shows)")
                    return
                
                reply = QMessageBox.question(
                    self,
                    "Confirm Organization",
                    "This will rename and move files. Are you sure you want to continue?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                
                if reply == QMessageBox.StandardButton.Yes:
                    self._start_worker("organize", options)
            except ValueError as exc:
                self._log_error(
                    "ERR010_VALIDATION_FAILED",
                    "Failed to validate organization options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))

        def _choose_organize_rules_file(self) -> None:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select Organize Rules JSON",
                str(Path(__file__).resolve().parent),
                "JSON Files (*.json)",
            )
            if file_path:
                self.organize_rules_input.setText(file_path)
        
        def _start_repair(self) -> None:
            try:
                options = self._collect_common_options()
                options["create_backup"] = self.repair_backup_checkbox.isChecked()
                
                reply = QMessageBox.question(
                    self,
                    "Confirm Repair",
                    "This will rebuild video containers using FFmpeg. {} Continue?".format(
                        "Backups will be created. " if options["create_backup"] else "NO BACKUPS will be created! "
                    ),
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                
                if reply == QMessageBox.StandardButton.Yes:
                    self._start_worker("repair", options)
            except ValueError as exc:
                self._log_error(
                    "ERR011_VALIDATION_FAILED",
                    "Failed to validate repair options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))
        
        def _start_generate(self) -> None:
            """Start subtitle generation using Whisper AI"""
            try:
                options = self._collect_common_options()
                options["model_size"] = self.whisper_model_combo.currentText()
                options["language"] = self.whisper_language_input.text().strip() or None
                options["output_format"] = "srt"  # Currently only SRT supported
                
                # Check if Whisper is available
                if whisper is None:
                    QMessageBox.warning(
                        self,
                        "Whisper Not Installed",
                        "Whisper AI is not installed. Please install it with:\n\n"
                        "pip install openai-whisper\n\n"
                        "Also requires: pip install pysubs2"
                    )
                    return
                
                if pysubs2 is None:
                    QMessageBox.warning(
                        self,
                        "pysubs2 Not Installed",
                        "pysubs2 library is not installed. Please install it with:\n\n"
                        "pip install pysubs2"
                    )
                    return
                
                # Warn about model download and processing time
                model_size = options["model_size"]
                model_info = {
                    "tiny": "~39M params, ~72MB, fastest",
                    "base": "~74M params, ~140MB, good balance",
                    "small": "~244M params, ~460MB, better accuracy",
                    "medium": "~769M params, ~1.5GB, high accuracy",
                    "large": "~1550M params, ~2.9GB, best accuracy",
                    "large-v2": "~1550M params, ~2.9GB, improved large",
                    "large-v3": "~1550M params, ~2.9GB, latest & best"
                }
                
                reply = QMessageBox.question(
                    self,
                    "Generate Subtitles",
                    f"Generate subtitles using Whisper '{model_size}' model?\n\n"
                    f"Model: {model_info.get(model_size, 'Unknown')}\n\n"
                    f"Note: First run will download the model. "
                    f"Processing may take several minutes per video.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes
                )
                
                if reply == QMessageBox.StandardButton.Yes:
                    self._start_worker("generate", options)
            except ValueError as exc:
                self._log_error(
                    "ERR012_VALIDATION_FAILED",
                    "Failed to validate subtitle generation options",
                    str(exc)
                )
                QMessageBox.warning(self, "Validation", str(exc))

        def _on_result(self, result: Dict[str, object]) -> None:
            action = str(result.get("action", "unknown"))
            if action == "scan":
                files = result.get("files", [])
                count = int(result.get("count", 0))
                self._log(f"Scan complete. Found {count} video file(s).")
                preview_limit = 15
                for item in files[:preview_limit]:
                    sidecars = item.get("sidecar_subtitles", [])
                    self._log(
                        f"- {item.get('path')} | embedded={item.get('embedded_subtitle_streams')} | "
                        f"sidecars={len(sidecars)}"
                    )
                if count > preview_limit:
                    self._log(f"... {count - preview_limit} more file(s) not shown in log.")
            elif action == "generate_subtitles":
                self._log(
                    "Finished {action}: scanned={scanned}, processed={processed}, "
                    "skipped={skipped}, failed={failed}".format(
                        action=action,
                        scanned=result.get("scanned", 0),
                        processed=result.get("processed", 0),
                        skipped=result.get("skipped", 0),
                        failed=result.get("failed", 0),
                    )
                )

                details = result.get("details", [])
                generated_paths: List[str] = []
                if isinstance(details, list):
                    for item in details:
                        if not isinstance(item, dict):
                            continue
                        if item.get("status") == "generated":
                            output_path = item.get("output_path")
                            if isinstance(output_path, str) and output_path:
                                generated_paths.append(output_path)

                if generated_paths:
                    self._log("Subtitle files saved to:")
                    for path in generated_paths:
                        self._log(f"- {path}")
            else:
                self._log(
                    "Finished {action}: scanned={scanned}, processed={processed}, "
                    "skipped={skipped}, failed={failed}".format(
                        action=action,
                        scanned=result.get("scanned", 0),
                        processed=result.get("processed", 0),
                        skipped=result.get("skipped", 0),
                        failed=result.get("failed", 0),
                    )
                )

        def _on_error(self, error_text: str) -> None:
            self._log("Task failed. See error below:")
            self._log(error_text)
            
            # Log error to settings for next startup
            self._log_error(
                "ERR003_OPERATION_FAILED",
                "Subtitle processing operation failed",
                error_text
            )
            
            QMessageBox.critical(self, "Task Failed", "The operation failed. Check the log for details.")

        def _iter_list_values(self, widget: QListWidget) -> List[str]:
            return [widget.item(i).text() for i in range(widget.count())]

        def _add_unique_items(self, widget: QListWidget, values: List[str]) -> int:
            existing = set(self._iter_list_values(widget))
            added = 0
            for value in values:
                if value in existing:
                    continue
                widget.addItem(QListWidgetItem(value))
                existing.add(value)
                added += 1
            return added

        def _choose_target_files(self) -> None:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Select Video Files",
                "",
                "Video Files (*.mp4 *.m4v *.mov *.mkv *.avi *.wmv *.flv *.webm *.mpg *.mpeg *.ts *.m2ts)",
            )
            if not files:
                return
            self._add_target_files(files)

        def _add_target_files(self, files: List[str]) -> None:
            normalized: List[str] = []
            for value in files:
                path = Path(value).expanduser().resolve()
                if not path.exists() or not path.is_file():
                    continue
                if path.suffix.lower() not in VIDEO_EXTENSIONS:
                    continue
                normalized.append(str(path))

            added = self._add_unique_items(self.target_file_list, normalized)
            if added:
                self._log(f"Added {added} target video file(s).")

        def _remove_selected_target_files(self) -> None:
            removed_paths: List[str] = []
            for item in self.target_file_list.selectedItems():
                removed_paths.append(item.text())
                row = self.target_file_list.row(item)
                self.target_file_list.takeItem(row)

            if removed_paths:
                for path in removed_paths:
                    self.manual_sidecars_by_video.pop(path, None)
                self._active_manual_video = None
                self._refresh_manual_subtitle_view()

        def _clear_target_files(self) -> None:
            self.target_file_list.clear()
            self.manual_sidecars_by_video.clear()
            self._active_manual_video = None
            self._refresh_manual_subtitle_view()

        def _collect_target_files(self) -> List[str]:
            if self.only_selected_targets_checkbox.isChecked():
                selected = [item.text() for item in self.target_file_list.selectedItems()]
                if selected:
                    return selected
                if self.target_file_list.count() > 0:
                    raise ValueError("Select at least one target video file or disable 'Use only selected target video file(s)'.")
                return []
            return self._iter_list_values(self.target_file_list)

        def _on_target_video_selection_changed(self) -> None:
            selected = self.target_file_list.selectedItems()
            self._active_manual_video = selected[0].text() if selected else None
            self._refresh_manual_subtitle_view()

        def _refresh_manual_subtitle_view(self) -> None:
            self.manual_subtitle_list.clear()
            if not self._active_manual_video:
                return
            for subtitle in self.manual_sidecars_by_video.get(self._active_manual_video, []):
                self.manual_subtitle_list.addItem(QListWidgetItem(subtitle))

        def _choose_manual_subtitles(self) -> None:
            if not self._active_manual_video:
                QMessageBox.warning(self, "Select Video", "Select a target video file first.")
                return
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Select Subtitle Files",
                "",
                "Subtitle Files (*.srt *.ass *.ssa *.vtt *.sub *.ttml)",
            )
            if not files:
                return
            self._add_manual_subtitles(files)

        def _add_manual_subtitles(self, files: List[str]) -> None:
            if not self._active_manual_video:
                QMessageBox.warning(self, "Select Video", "Select a target video file first.")
                return

            normalized: List[str] = []
            for value in files:
                path = Path(value).expanduser().resolve()
                if not path.exists() or not path.is_file():
                    continue
                if path.suffix.lower() not in SUBTITLE_EXTENSIONS:
                    continue
                normalized.append(str(path))

            existing = self.manual_sidecars_by_video.get(self._active_manual_video, [])
            existing_set = set(existing)
            added = 0
            for subtitle in normalized:
                if subtitle in existing_set:
                    continue
                existing.append(subtitle)
                existing_set.add(subtitle)
                added += 1

            self.manual_sidecars_by_video[self._active_manual_video] = existing
            self._refresh_manual_subtitle_view()
            if added:
                self._log(f"Added {added} subtitle file(s) for {Path(self._active_manual_video).name}.")

        def _remove_selected_manual_subtitles(self) -> None:
            if not self._active_manual_video:
                return
            selected = {item.text() for item in self.manual_subtitle_list.selectedItems()}
            if not selected:
                return

            current = self.manual_sidecars_by_video.get(self._active_manual_video, [])
            updated = [entry for entry in current if entry not in selected]
            if updated:
                self.manual_sidecars_by_video[self._active_manual_video] = updated
            else:
                self.manual_sidecars_by_video.pop(self._active_manual_video, None)
            self._refresh_manual_subtitle_view()

        def _clear_manual_subtitles_for_selected_video(self) -> None:
            if not self._active_manual_video:
                QMessageBox.warning(self, "Select Video", "Select a target video file first.")
                return
            self.manual_sidecars_by_video.pop(self._active_manual_video, None)
            self._refresh_manual_subtitle_view()

        def _open_help_dialog(self) -> None:
            dialog = HelpDialog(self)
            dialog.exec()

        def _show_tutorial(self) -> None:
            tutorial = TutorialDialog(self, self)
            tutorial.exec()

        def _toggle_theme(self) -> None:
            """Toggle between light and dark mode."""
            self.dark_mode = not self.dark_mode
            
            # Update button text
            self.theme_toggle_button.setText("Switch to Light Mode" if self.dark_mode else "Switch to Dark Mode")
            
            # Apply new theme
            self._apply_theme()
            
            # Save preference
            settings = self._load_settings()
            settings["dark_mode"] = self.dark_mode
            self._save_settings(settings)
            
            self._log(f"Switched to {'dark' if self.dark_mode else 'light'} mode")
        
        def _show_error_history(self) -> None:
            """Show dialog with all errors (read and unread)."""
            settings = self._load_settings()
            all_errors = settings.get("errors", [])
            
            if not all_errors:
                QMessageBox.information(self, "Error History", "No errors have been logged.")
                return
            
            dialog = QDialog(self)
            dialog.setWindowTitle("Error History")
            dialog.resize(700, 500)
            
            layout = QVBoxLayout(dialog)
            
            # Create scroll area for errors
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_area.setFrameShape(QFrame.Shape.NoFrame)
            
            content_widget = QWidget()
            content_layout = QVBoxLayout(content_widget)
            
            # Display each error
            for idx, error in enumerate(all_errors, 1):
                error_id = error.get("id", "UNKNOWN")
                message = error.get("message", "No message")
                timestamp = error.get("timestamp", "Unknown time")
                read = error.get("read", False)
                details = error.get("details", "")
                
                error_text = f"[{idx}] {error_id}\n"
                error_text += f"Status: {'Read' if read else 'UNREAD'}\n"
                error_text += f"Time: {timestamp}\n"
                error_text += f"Message: {message}\n"
                if details:
                    error_text += f"Details: {details}"
                
                error_label = QLabel(error_text)
                error_label.setWordWrap(True)
                error_label.setStyleSheet(
                    "QLabel { "
                    "border: 1px solid #666; "
                    "border-radius: 4px; "
                    "padding: 8px; "
                    "margin-bottom: 8px; "
                    "background-color: #ffeeee; "
                    "color: #000000; "
                    "}"
                )
                content_layout.addWidget(error_label)
            
            content_layout.addStretch()
            scroll_area.setWidget(content_widget)
            layout.addWidget(scroll_area)
            
            # Buttons
            button_layout = QHBoxLayout()
            clear_button = QPushButton("Clear All Errors")
            close_button = QPushButton("Close")
            
            clear_button.clicked.connect(lambda: self._clear_errors_and_close(dialog))
            close_button.clicked.connect(dialog.accept)
            
            button_layout.addWidget(clear_button)
            button_layout.addStretch()
            button_layout.addWidget(close_button)
            layout.addLayout(button_layout)
            
            dialog.exec()
        
        def _clear_errors_and_close(self, dialog: QDialog) -> None:
            """Clear all errors and close the dialog."""
            reply = QMessageBox.question(
                dialog,
                "Clear Errors",
                "Are you sure you want to clear all error history?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self._clear_all_errors()
                QMessageBox.information(dialog, "Success", "All errors have been cleared.")
                dialog.accept()

        def _check_for_errors(self) -> None:
            """Check for unread errors and display them to the user."""
            errors = self._get_unread_errors()
            if errors:
                error_messages = []
                for error in errors:
                    error_id = error.get("id", "UNKNOWN")
                    message = error.get("message", "Unknown error")
                    timestamp = error.get("timestamp", "Unknown time")
                    error_messages.append(f"[{error_id}] {message}\nTime: {timestamp}")
                
                full_message = "The following errors were encountered:\n\n" + "\n\n".join(error_messages)
                full_message += "\n\nThese errors have been logged. Click OK to acknowledge."
                
                QMessageBox.warning(self, "Previous Errors Detected", full_message)
                self._mark_errors_read()
        
        def _check_first_run(self) -> None:
            """Check if this is the first run and show tutorial if so."""
            settings = self._load_settings()
            if not settings.get("tutorial_completed", False):
                # Delay tutorial slightly so main window renders first
                QTimer.singleShot(500, self._show_first_run_tutorial)

        def _show_first_run_tutorial(self) -> None:
            reply = QMessageBox.question(
                self,
                "Welcome!",
                "Welcome to Subtitle Tool!\n\nWould you like to take a quick tutorial to learn about the features?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._show_tutorial()
            self._mark_tutorial_completed()

        def _load_settings(self) -> Dict[str, object]:
            if self.settings_path.exists():
                try:
                    content = self.settings_path.read_text(encoding="utf-8")
                    return json.loads(content)
                except (OSError, json.JSONDecodeError):
                    pass
            return {}

        def _save_settings(self, settings: Dict[str, object]) -> None:
            try:
                self.settings_path.write_text(json.dumps(settings, indent=2), encoding="utf-8")
            except OSError:
                pass
        
        def _log_error(self, error_id: str, message: str, details: Optional[str] = None) -> None:
            """Log an error to the settings file with a custom error ID.
            
            Args:
                error_id: Custom error identifier (e.g., 'ERR001', 'FFMPEG_NOT_FOUND')
                message: Human-readable error message
                details: Optional additional details or stack trace
            """
            settings = self._load_settings()
            errors = settings.get("errors", [])
            
            if not isinstance(errors, list):
                errors = []
            
            error_entry = {
                "id": error_id,
                "message": message,
                "timestamp": datetime.now().isoformat(),
                "read": False
            }
            
            if details:
                error_entry["details"] = details
            
            errors.append(error_entry)
            settings["errors"] = errors
            self._save_settings(settings)
        
        def _get_unread_errors(self) -> List[Dict[str, object]]:
            """Get all unread errors from settings."""
            settings = self._load_settings()
            errors = settings.get("errors", [])
            
            if not isinstance(errors, list):
                return []
            
            return [e for e in errors if not e.get("read", False)]
        
        def _mark_errors_read(self) -> None:
            """Mark all errors as read."""
            settings = self._load_settings()
            errors = settings.get("errors", [])
            
            if isinstance(errors, list):
                for error in errors:
                    error["read"] = True
                settings["errors"] = errors
                self._save_settings(settings)
        
        def _clear_all_errors(self) -> None:
            """Clear all errors from settings."""
            settings = self._load_settings()
            settings["errors"] = []
            self._save_settings(settings)

        def _mark_tutorial_completed(self) -> None:
            settings = self._load_settings()
            settings["tutorial_completed"] = True
            self._save_settings(settings)

        def _record_launch_time(self) -> bool:
            """Record app launch time. Returns True if memory should be cleared due to rapid launches."""
            settings = self._load_settings()
            current_time = datetime.now().timestamp()
            launch_times = settings.get("launch_times", [])
            
            # Keep only launches from the last 30 seconds
            recent_launches = [t for t in launch_times if current_time - t < 30]
            recent_launches.append(current_time)
            
            settings["launch_times"] = recent_launches
            self._save_settings(settings)
            
            # If 3 or more launches in 30 seconds, clear memory
            return len(recent_launches) >= 3

        def _save_ui_state(self) -> None:
            """Save current UI state to settings."""
            settings = self._load_settings()
            
            # Collect folder list
            folders = [self.folder_list.item(i).text() for i in range(self.folder_list.count())]
            
            # Collect target files
            target_files = self._collect_target_files()
            
            ui_state = {
                "folders": folders,
                "target_files": target_files,
                "manual_sidecars": self.manual_sidecars_by_video,
                "recursive": self.recursive_checkbox.isChecked(),
                "overwrite": self.overwrite_checkbox.isChecked(),
                "extract": self.extract_checkbox.isChecked(),
                "export_txt": self.export_txt_checkbox.isChecked(),
                "scan_only_embedded": self.scan_only_embedded_checkbox.isChecked(),
                "only_selected_targets": self.only_selected_targets_checkbox.isChecked(),
                "remove_suffix": self.remove_suffix_input.text(),
                "include_suffix": self.include_suffix_input.text(),
                "extract_suffix": self.extract_suffix_input.text(),
                "convert_suffix": self.convert_suffix_input.text(),
                "organize_movies": self.organize_movies_checkbox.isChecked(),
                "organize_tv": self.organize_tv_checkbox.isChecked(),
                "organize_rules_path": self.organize_rules_input.text(),
                "repair_backup": self.repair_backup_checkbox.isChecked(),
                "whisper_model": self.whisper_model_combo.currentText() if self.whisper_model_combo else "base",
                "whisper_language": self.whisper_language_input.text() if self.whisper_language_input else "",
            }
            
            settings["ui_state"] = ui_state
            self._save_settings(settings)

        def _load_ui_state(self) -> None:
            """Load and restore UI state from settings."""
            # Check for rapid launches
            should_clear_rapid = self._record_launch_time()
            
            if self.clear_memory:
                self._log("Memory cleared (--clear flag)")
                return
            
            if should_clear_rapid:
                self._log("Memory cleared (3 rapid launches detected)")
                self._clear_memory()
                return
            
            settings = self._load_settings()
            ui_state = settings.get("ui_state")
            
            if not ui_state:
                return
            
            # Restore folders
            for folder in ui_state.get("folders", []):
                if folder and Path(folder).exists():
                    self.folder_list.addItem(QListWidgetItem(folder))
            
            # Restore target files
            for file_path in ui_state.get("target_files", []):
                if file_path and Path(file_path).exists():
                    self.target_file_list.addItem(QListWidgetItem(file_path))
            
            # Restore manual sidecars
            manual_sidecars = ui_state.get("manual_sidecars", {})
            if manual_sidecars:
                self.manual_sidecars_by_video = manual_sidecars
            
            # Restore checkbox states
            self.recursive_checkbox.setChecked(ui_state.get("recursive", True))
            self.overwrite_checkbox.setChecked(ui_state.get("overwrite", False))
            self.extract_checkbox.setChecked(ui_state.get("extract", True))
            self.export_txt_checkbox.setChecked(ui_state.get("export_txt", True))
            self.scan_only_embedded_checkbox.setChecked(ui_state.get("scan_only_embedded", False))
            self.only_selected_targets_checkbox.setChecked(ui_state.get("only_selected_targets", False))
            
            # Restore text inputs
            self.remove_suffix_input.setText(ui_state.get("remove_suffix", "_nosubs"))
            self.include_suffix_input.setText(ui_state.get("include_suffix", "_withsubs"))
            self.extract_suffix_input.setText(ui_state.get("extract_suffix", ".embedded_sub"))
            self.convert_suffix_input.setText(ui_state.get("convert_suffix", "_converted"))
            
            # Restore Swiss Army Knife tool options
            self.organize_movies_checkbox.setChecked(ui_state.get("organize_movies", True))
            self.organize_tv_checkbox.setChecked(ui_state.get("organize_tv", True))
            self.organize_rules_input.setText(ui_state.get("organize_rules_path", ""))
            self.repair_backup_checkbox.setChecked(ui_state.get("repair_backup", True))
            
            # Restore Whisper AI settings (only if AI is enabled)
            if self.use_ai and self.whisper_model_combo and self.whisper_language_input:
                whisper_model = ui_state.get("whisper_model", "base")
                index = self.whisper_model_combo.findText(whisper_model)
                if index >= 0:
                    self.whisper_model_combo.setCurrentIndex(index)
                self.whisper_language_input.setText(ui_state.get("whisper_language", ""))
            
            self._log("Previous session restored from memory")

        def _clear_memory(self) -> None:
            """Clear saved UI state from settings."""
            settings = self._load_settings()
            if "ui_state" in settings:
                del settings["ui_state"]
            if "launch_times" in settings:
                del settings["launch_times"]
            self._save_settings(settings)

        def closeEvent(self, event) -> None:
            """Save UI state when window closes."""
            self._save_ui_state()
            super().closeEvent(event)



def run_gui(clear_memory: bool = False, use_ai: Optional[bool] = None) -> int:
    if QApplication is None:
        print("PyQt6 is not installed. Install requirements and retry.", file=sys.stderr)
        return 1

    app = QApplication(sys.argv)
    window = SubtitleToolWindow(clear_memory=clear_memory, use_ai=use_ai)
    window.show()
    return app.exec()


def run_api(host: str, port: int) -> int:
    if FastAPI is None or uvicorn is None:
        print("FastAPI/uvicorn are not installed. Install requirements and retry.", file=sys.stderr)
        return 1

    app = create_api_app()
    uvicorn.run(app, host=host, port=port, log_level="info")
    return 0


def cli_print_json(payload: Dict[str, object]) -> None:
    print(json.dumps(payload, indent=2))


def run_cli_action(args: argparse.Namespace) -> int:
    processor = SubtitleProcessor(log_callback=lambda m: print(f"[log] {m}"))

    deps = processor.check_dependencies()
    if not deps["ffmpeg_found"] or not deps["ffprobe_found"]:
        print("ffmpeg/ffprobe not found on PATH. Install ffmpeg first.", file=sys.stderr)
        return 2

    folders = args.folders
    recursive = not args.no_recursive

    if args.mode == "scan":
        rows = processor.scan_videos(
            folders,
            recursive=recursive,
            only_with_embedded=args.only_with_embedded,
        )
        payload = {
            "action": "scan",
            "count": len(rows),
            "files": [
                {
                    "path": r.path,
                    "embedded_subtitle_streams": r.embedded_subtitle_streams,
                    "sidecar_subtitles": r.sidecar_subtitles,
                }
                for r in rows
            ],
        }
        cli_print_json(payload)
        return 0

    if args.mode == "remove":
        summary = processor.remove_embedded_subtitles(
            folders=folders,
            recursive=recursive,
            overwrite=args.overwrite,
            output_suffix=args.suffix,
            extract_for_restore=not args.no_extract,
        )
        cli_print_json(summary.to_dict())
        return 0

    if args.mode == "include":
        summary = processor.include_subtitles(
            folders=folders,
            recursive=recursive,
            overwrite=args.overwrite,
            output_suffix=args.suffix,
        )
        cli_print_json(summary.to_dict())
        return 0

    if args.mode == "extract":
        summary = processor.extract_embedded_subtitles(
            folders=folders,
            recursive=recursive,
            overwrite=args.overwrite,
            output_suffix=args.suffix,
            export_txt=not args.no_txt,
        )
        cli_print_json(summary.to_dict())
        return 0

    print(f"Unsupported mode: {args.mode}", file=sys.stderr)
    return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Subtitle stream utility")
    subparsers = parser.add_subparsers(dest="mode")

    gui_parser = subparsers.add_parser("gui", help="Launch PyQt GUI")
    gui_parser.add_argument("--clear", action="store_true", help="Clear saved UI state/memory")
    gui_parser.add_argument("--no-ai", action="store_true", help="Disable AI subtitle generation (saves 'use_ai=false' setting)")
    gui_parser.add_argument("--use-ai", action="store_true", help="Enable AI subtitle generation (saves 'use_ai=true' setting)")

    api_parser = subparsers.add_parser("api", help="Run FastAPI service")
    api_parser.add_argument("--host", default="127.0.0.1", help="API host")
    api_parser.add_argument("--port", type=int, default=8891, help="API port")

    for mode, default_suffix in (
        ("scan", ""),
        ("remove", "_nosubs"),
        ("include", "_withsubs"),
        ("extract", ".embedded_sub"),
    ):
        cmd = subparsers.add_parser(mode, help=f"Run {mode} operation in CLI mode")
        cmd.add_argument("--folders", nargs="+", required=True, help="One or more folders to process")
        cmd.add_argument("--no-recursive", action="store_true", help="Do not scan subfolders")
        if mode == "scan":
            cmd.add_argument(
                "--only-with-embedded",
                action="store_true",
                help="Only include files that contain embedded subtitle streams",
            )
        if mode in {"remove", "include", "extract"}:
            cmd.add_argument("--overwrite", action="store_true", help="Overwrite original files")
            cmd.add_argument("--suffix", default=default_suffix, help="Output filename suffix")
        if mode == "remove":
            cmd.add_argument(
                "--no-extract",
                action="store_true",
                help="Do not extract embedded subtitle streams before removing",
            )
        if mode == "extract":
            cmd.add_argument(
                "--no-txt",
                action="store_true",
                help="Do not generate .txt versions for extracted subtitles",
            )

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    # Default behavior is GUI when no mode is supplied.
    mode = args.mode or "gui"

    if mode == "gui":
        clear_memory = getattr(args, "clear", False)
        use_ai = None
        if getattr(args, "no_ai", False):
            use_ai = False
        elif getattr(args, "use_ai", False):
            use_ai = True
        return run_gui(clear_memory=clear_memory, use_ai=use_ai)
    if mode == "api":
        return run_api(host=args.host, port=args.port)
    return run_cli_action(args)


if __name__ == "__main__":
    raise SystemExit(main())
