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
from collections import Counter
import importlib
import json
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
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
        QTabWidget,
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
    import numpy as np
except ImportError:
    np = None  # type: ignore[assignment]

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
    available_backends: List[str] = ["text-to-timestamps"]

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
        available_backends.append("openai-whisper")
    except Exception as exc:
        if "openai-whisper / torch" not in missing:
            missing.append("openai-whisper / torch")
        details["whisper_error"] = f"{type(exc).__name__}: {exc}"

    try:
        fw_mod = importlib.import_module("faster_whisper")
        details["faster_whisper"] = str(getattr(fw_mod, "__version__", "installed"))
        available_backends.append("faster-whisper")
    except Exception as exc:
        details["faster_whisper_error"] = f"{type(exc).__name__}: {exc}"

    try:
        wx_mod = importlib.import_module("whisperx")
        details["whisperx"] = str(getattr(wx_mod, "__version__", "installed"))
        available_backends.append("whisperx")
    except Exception as exc:
        details["whisperx_error"] = f"{type(exc).__name__}: {exc}"

    try:
        st_mod = importlib.import_module("stable_whisper")
        details["stable_whisper"] = str(getattr(st_mod, "__version__", "installed"))
        available_backends.append("stable-ts")
    except Exception as exc:
        details["stable_whisper_error"] = f"{type(exc).__name__}: {exc}"

    try:
        wt_mod = importlib.import_module("whisper_timestamped")
        details["whisper_timestamped"] = str(getattr(wt_mod, "__version__", "installed"))
        available_backends.append("whisper-timestamped")
    except Exception as exc:
        details["whisper_timestamped_error"] = f"{type(exc).__name__}: {exc}"

    try:
        sb_mod = importlib.import_module("speechbrain")
        details["speechbrain"] = str(getattr(sb_mod, "__version__", "installed"))
        available_backends.append("speechbrain")
    except Exception as exc:
        details["speechbrain_error"] = f"{type(exc).__name__}: {exc}"

    try:
        vosk_mod = importlib.import_module("vosk")
        details["vosk"] = str(getattr(vosk_mod, "__version__", "installed"))
        available_backends.append("vosk")
    except Exception as exc:
        details["vosk_error"] = f"{type(exc).__name__}: {exc}"

    try:
        aeneas_mod = importlib.import_module("aeneas")
        details["aeneas"] = str(getattr(aeneas_mod, "__version__", "installed"))
    except Exception as exc:
        details["aeneas_error"] = f"{type(exc).__name__}: {exc}"

    try:
        pysubs2_mod = importlib.import_module("pysubs2")
        pysubs2 = pysubs2_mod  # type: ignore[assignment]
        details["pysubs2"] = str(getattr(pysubs2_mod, "VERSION", "installed"))
    except Exception as exc:
        missing.append("pysubs2/pysub2")
        details["pysubs2_error"] = f"{type(exc).__name__}: {exc}"

    if len(available_backends) <= 1:
        missing.append("No supported transcription backend installed")

    details["available_backends"] = ", ".join(sorted(set(available_backends)))
    ai_ready = len(available_backends) > 1 and pysubs2 is not None

    return ai_ready, missing, details

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
LANGUAGE_CODE_MAP = {
    "en": "eng",
    "es": "spa",
    "fr": "fra",
    "de": "deu",
    "it": "ita",
    "pt": "por",
    "ru": "rus",
    "ja": "jpn",
    "ko": "kor",
    "zh": "zho",
    "ar": "ara",
    "hi": "hin",
    "nl": "nld",
    "sv": "swe",
    "no": "nor",
    "da": "dan",
    "fi": "fin",
    "pl": "pol",
    "tr": "tur",
    "el": "ell",
    "he": "heb",
    "uk": "ukr",
    "cs": "ces",
    "ro": "ron",
    "hu": "hun",
    "th": "tha",
    "vi": "vie",
    "id": "ind",
    "ms": "msa",
}
HELP_DOC_NAME = "SUBTITLE_TOOL_HELP.md"
SETTINGS_FILE = ".subtitle_tool_settings.json"
COMMAND_FEEDBACK_LEVELS = {"quiet", "normal", "verbose"}
FFMPEG_LOGLEVELS = {"quiet", "error", "warning", "info", "verbose"}
CONVERSION_BACKENDS = {"auto", "ffmpeg", "mkvtoolnix", "handbrake"}
REPAIR_BACKENDS = {"auto", "ffmpeg", "mkvtoolnix", "makemkv"}
TRANSCRIBE_BACKENDS = {
    "auto",
    "openai-whisper",
    "faster-whisper",
    "whisperx",
    "stable-ts",
    "whisper-timestamped",
    "speechbrain",
    "vosk",
    "text-to-timestamps",
}
SYNC_BACKENDS = {"auto", "whisper-offset", "aeneas"}


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
        mkvmerge_bin: Optional[str] = None,
        handbrake_bin: Optional[str] = None,
        makemkvcon_bin: Optional[str] = None,
        log_callback: Optional[Callable[[str], None]] = None,
        use_hw_accel: bool = False,
        command_feedback: str = "normal",
        ffmpeg_loglevel: str = "warning",
        ffprobe_loglevel: str = "error",
    ) -> None:
        self.ffmpeg_bin = ffmpeg_bin or "ffmpeg"
        self.ffprobe_bin = ffprobe_bin or "ffprobe"
        self.mkvmerge_bin = mkvmerge_bin or "mkvmerge"
        self.handbrake_bin = handbrake_bin or "HandBrakeCLI"
        self.makemkvcon_bin = makemkvcon_bin or "makemkvcon"
        self.log_callback = log_callback
        self.use_hw_accel = use_hw_accel
        self.command_feedback = (
            command_feedback if command_feedback in COMMAND_FEEDBACK_LEVELS else "normal"
        )
        self.ffmpeg_loglevel = ffmpeg_loglevel if ffmpeg_loglevel in FFMPEG_LOGLEVELS else "warning"
        self.ffprobe_loglevel = ffprobe_loglevel if ffprobe_loglevel in FFMPEG_LOGLEVELS else "error"
        # Cache for IMDB episode name lookups - avoids repeated network requests
        # Key: "show_name_lower|season|episode", Value: episode title or None
        self._episode_name_cache: Dict[str, Optional[str]] = {}

    def _log(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)

    def _hw_accel_flags(self) -> List[str]:
        """Return ffmpeg hardware acceleration flags when enabled."""
        return ["-hwaccel", "auto"] if self.use_hw_accel else []

    @staticmethod
    def _ts_input_stability_flags() -> List[str]:
        """Input flags that make ffmpeg probing/remuxing more stable for TS/M2TS sources."""
        return [
            "-analyzeduration",
            "200M",
            "-probesize",
            "200M",
            "-fflags",
            "+genpts+igndts",
        ]

    def _get_temp_workspace_root(self) -> Path:
        custom_dir = os.getenv("SUBTITLE_TOOL_TEMP_DIR", "").strip()
        if custom_dir:
            base = Path(custom_dir).expanduser()
        else:
            local_app_data = os.getenv("LOCALAPPDATA", "").strip()
            if local_app_data:
                base = Path(local_app_data)
            else:
                base = Path(tempfile.gettempdir())

        root = base / "SubtitleTool" / "temp"
        root.mkdir(parents=True, exist_ok=True)
        return root

    @staticmethod
    def _resolve_binary(binary_value: str) -> str:
        if not binary_value:
            return ""
        resolved = shutil.which(binary_value)
        if resolved:
            return resolved
        candidate = Path(binary_value).expanduser()
        if candidate.exists():
            try:
                return str(candidate.resolve())
            except OSError:
                return str(candidate)
        return ""

    @staticmethod
    def _preview_lines(raw_text: str, max_lines: int = 20, max_chars: int = 4000) -> List[str]:
        text = (raw_text or "").strip()
        if not text:
            return []
        lines = text.splitlines()[:max_lines]
        if len("\n".join(lines)) > max_chars:
            trimmed = "\n".join(lines)
            return [trimmed[:max_chars].rstrip() + " ..."]
        return lines

    def _log_command_output(self, stream_name: str, raw_text: str, max_lines: int = 20) -> None:
        for line in self._preview_lines(raw_text, max_lines=max_lines):
            self._log(f"  {stream_name}: {line}")

    def check_dependencies(self) -> Dict[str, object]:
        ffmpeg = self._resolve_binary(self.ffmpeg_bin)
        ffprobe = self._resolve_binary(self.ffprobe_bin)
        mkvmerge = self._resolve_binary(self.mkvmerge_bin)
        handbrake = self._resolve_binary(self.handbrake_bin)
        makemkvcon = self._resolve_binary(self.makemkvcon_bin)
        return {
            "ffmpeg_found": bool(ffmpeg),
            "ffprobe_found": bool(ffprobe),
            "ffmpeg_path": ffmpeg or "",
            "ffprobe_path": ffprobe or "",
            "mkvmerge_found": bool(mkvmerge),
            "mkvmerge_path": mkvmerge or "",
            "mkvtoolnix_found": bool(mkvmerge),
            "handbrake_found": bool(handbrake),
            "handbrake_path": handbrake or "",
            "makemkv_found": bool(makemkvcon),
            "makemkv_path": makemkvcon or "",
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

    def _run_command(
        self,
        args: List[str],
        *,
        description: str = "",
        log_stdout_on_success: bool = False,
        max_feedback_lines: int = 20,
    ) -> subprocess.CompletedProcess[str]:
        self._log("Running: " + " ".join(args))
        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            label = description or Path(args[0]).name
            self._log(f"Command failed ({label}) with exit code {result.returncode}")

        verbose_mode = self.command_feedback == "verbose"
        show_output = verbose_mode or result.returncode != 0
        if show_output:
            self._log_command_output("stderr", result.stderr or "", max_lines=max_feedback_lines)
            if verbose_mode or log_stdout_on_success or result.returncode != 0:
                self._log_command_output("stdout", result.stdout or "", max_lines=max_feedback_lines)
        return result

    def _probe_subtitle_streams(self, video_path: Path) -> List[Dict[str, object]]:
        cmd = [
            self.ffprobe_bin,
            "-v",
            self.ffprobe_loglevel,
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
                self._log(f"Found {len(streams)} subtitle stream(s): {video_path.name}")
                return streams
        except json.JSONDecodeError:
            self._log(f"Failed parsing ffprobe output for {video_path}")
        return []

    def _probe_audio_streams(self, video_path: Path) -> List[Dict[str, object]]:
        cmd = [
            self.ffprobe_bin,
            "-v",
            self.ffprobe_loglevel,
            "-select_streams",
            "a",
            "-show_entries",
            "stream=index,codec_name,bit_rate,channels:stream_tags=language,title",
            "-of",
            "json",
            str(video_path),
        ]
        result = self._run_command(cmd)
        if result.returncode != 0:
            self._log(f"ffprobe audio stream probe failed for {video_path}: {result.stderr.strip()}")
            return []

        try:
            payload = json.loads(result.stdout or "{}")
            streams = payload.get("streams", [])
            if isinstance(streams, list):
                self._log(f"Found {len(streams)} audio stream(s): {video_path.name}")
                return streams
        except json.JSONDecodeError:
            self._log(f"Failed parsing ffprobe audio stream output for {video_path}")
        return []

    def _probe_media_duration_seconds(self, video_path: Path) -> Optional[float]:
        cmd = [
            self.ffprobe_bin,
            "-v",
            self.ffprobe_loglevel,
            "-show_entries",
            "format=duration",
            "-of",
            "default=noprint_wrappers=1:nokey=1",
            str(video_path),
        ]
        result = self._run_command(cmd)
        if result.returncode != 0:
            return None

        raw = (result.stdout or "").strip()
        try:
            value = float(raw)
        except ValueError:
            return None
        if value <= 0:
            return None
        return value

    def _probe_stream_counts(self, video_path: Path) -> Tuple[int, int, int]:
        cmd = [
            self.ffprobe_bin,
            "-v",
            self.ffprobe_loglevel,
            "-show_entries",
            "stream=codec_type",
            "-of",
            "json",
            str(video_path),
        ]
        result = self._run_command(cmd)
        if result.returncode != 0:
            return 0, 0, 0

        video_count = 0
        audio_count = 0
        subtitle_count = 0
        try:
            payload = json.loads(result.stdout or "{}")
            streams = payload.get("streams", [])
            if not isinstance(streams, list):
                return 0, 0, 0
            for stream in streams:
                if not isinstance(stream, dict):
                    continue
                codec_type = str(stream.get("codec_type") or "").strip().lower()
                if codec_type == "video":
                    video_count += 1
                elif codec_type == "audio":
                    audio_count += 1
                elif codec_type == "subtitle":
                    subtitle_count += 1
        except json.JSONDecodeError:
            return 0, 0, 0

        return video_count, audio_count, subtitle_count

    def _validate_muxed_output(
        self,
        source_path: Path,
        output_path: Path,
        *,
        allow_reduced_audio: bool = False,
        allow_reduced_subtitles: bool = False,
    ) -> Tuple[bool, str]:
        if not output_path.exists():
            return False, "output file missing"

        source_ext = source_path.suffix.lower()
        source_is_ts = source_ext in (".ts", ".m2ts")

        source_size = source_path.stat().st_size
        output_size = output_path.stat().st_size
        src_duration = self._probe_media_duration_seconds(source_path)
        out_duration = self._probe_media_duration_seconds(output_path)
        if src_duration and src_duration > 0 and out_duration and out_duration > 0:
            ratio = out_duration / src_duration
            if ratio < 0.98:
                return False, (
                    f"output duration sanity check failed "
                    f"({out_duration:.2f}s vs source {src_duration:.2f}s)"
                )

        src_v, src_a, src_s = self._probe_stream_counts(source_path)
        out_v, out_a, out_s = self._probe_stream_counts(output_path)
        if src_v > 0 and out_v == 0:
            return False, "output has no video streams"
        if src_a > 0 and out_a == 0:
            return False, "output has no audio streams"
        if out_v < src_v:
            return False, f"video stream count reduced ({out_v} vs {src_v})"
        if not allow_reduced_audio and out_a < src_a:
            return False, f"audio stream count reduced ({out_a} vs {src_a})"
        if not allow_reduced_subtitles and out_s < src_s:
            return False, f"subtitle stream count reduced ({out_s} vs {src_s})"

        # TS/M2TS often contains heavy constant-bitrate padding that can disappear
        # after remux; size ratio is not a reliable integrity signal there.
        if not source_is_ts and source_size > 0 and output_size < source_size * 0.90:
            return False, f"output size sanity check failed ({output_size:,} vs {source_size:,} bytes)"

        return True, "ok"

    def _validate_transcoded_output(self, source_path: Path, output_path: Path) -> Tuple[bool, str]:
        if not output_path.exists():
            return False, "output file missing"

        output_size = output_path.stat().st_size
        if output_size < 1_000_000:
            return False, f"output too small ({output_size:,} bytes)"

        source_size = source_path.stat().st_size
        src_duration = self._probe_media_duration_seconds(source_path)
        out_duration = self._probe_media_duration_seconds(output_path)
        if src_duration and out_duration and src_duration > 0:
            ratio = out_duration / src_duration
            if ratio < 0.95:
                return False, (
                    f"output duration sanity check failed "
                    f"({out_duration:.2f}s vs source {src_duration:.2f}s)"
                )

        if source_size >= 5_000_000_000 and output_size < int(source_size * 0.02):
            return False, (
                f"output size suspiciously small for large source "
                f"({output_size:,} vs {source_size:,} bytes)"
            )

        src_v, src_a, _ = self._probe_stream_counts(source_path)
        out_v, out_a, _ = self._probe_stream_counts(output_path)
        if src_v > 0 and out_v == 0:
            return False, "output has no video streams"
        if src_a > 0 and out_a == 0:
            return False, "output has no audio streams"
        return True, "ok"

    def _choose_conversion_backend(self, target_format: str, preferred_backend: str) -> Tuple[Optional[str], str]:
        backend = (preferred_backend or "auto").strip().lower()
        if backend not in CONVERSION_BACKENDS:
            backend = "auto"

        if backend == "auto":
            if target_format == "mkv" and self._resolve_binary(self.mkvmerge_bin):
                return "mkvtoolnix", "auto-selected mkvtoolnix for mkv target"
            if target_format == "mp4" and self._resolve_binary(self.handbrake_bin):
                return "handbrake", "auto-selected handbrake for mp4 target"
            return "ffmpeg", "auto-selected ffmpeg"

        if backend == "mkvtoolnix":
            if target_format != "mkv":
                return None, "mkvtoolnix backend supports mkv output only"
            if not self._resolve_binary(self.mkvmerge_bin):
                return None, "mkvtoolnix selected but mkvmerge was not found"
            return "mkvtoolnix", "using mkvtoolnix backend"

        if backend == "handbrake":
            if not self._resolve_binary(self.handbrake_bin):
                return None, "handbrake backend selected but HandBrakeCLI was not found"
            return "handbrake", "using handbrake backend"

        if backend == "ffmpeg":
            return "ffmpeg", "using ffmpeg backend"

        return None, f"unsupported conversion backend: {backend}"

    def _run_makemkv_repair(self, video: Path, output_path: Path) -> Tuple[bool, str]:
        temp_root = self._get_temp_workspace_root() / f"makemkv_repair_{uuid.uuid4().hex[:8]}"
        temp_root.mkdir(parents=True, exist_ok=True)

        cmd = [
            self.makemkvcon_bin,
            "--robot",
            "--messages=-stdout",
            "--progress=-stdout",
            "mkv",
            f"file:{video}",
            "all",
            str(temp_root),
        ]
        result = self._run_command(
            cmd,
            description="makemkv repair",
            max_feedback_lines=30,
        )
        if result.returncode != 0:
            shutil.rmtree(temp_root, ignore_errors=True)
            return False, (result.stderr or "").strip()[:300] or "makemkv failed"

        candidates = sorted(temp_root.glob("*.mkv"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not candidates:
            shutil.rmtree(temp_root, ignore_errors=True)
            return False, "makemkv completed but no MKV output was generated"

        try:
            shutil.move(str(candidates[0]), str(output_path))
        except OSError as exc:
            shutil.rmtree(temp_root, ignore_errors=True)
            return False, f"failed to move makemkv output: {exc}"

        shutil.rmtree(temp_root, ignore_errors=True)
        return True, "ok"

    @staticmethod
    def _normalize_language_code(code: str) -> str:
        value = (code or "").strip().lower()
        if not value:
            return ""
        if len(value) == 3:
            return value
        return LANGUAGE_CODE_MAP.get(value, value)

    def _extract_audio_sample(
        self,
        video_path: Path,
        stream_index: int,
        output_path: Path,
        start_seconds: Optional[float],
        sample_seconds: Optional[float],
    ) -> bool:
        cmd: List[str] = [
            self.ffmpeg_bin,
            "-y",
            "-loglevel",
            "error",
            "-nostats",
        ]
        cmd.extend(self._hw_accel_flags())
        if start_seconds is not None and start_seconds > 0:
            cmd.extend(["-ss", f"{start_seconds:.3f}"])

        cmd.extend(
            [
                "-i",
                str(video_path),
                "-map",
                f"0:{stream_index}",
                "-vn",
                "-ac",
                "1",
                "-ar",
                "16000",
            ]
        )
        if sample_seconds is not None and sample_seconds > 0:
            cmd.extend(["-t", f"{sample_seconds:.3f}"])
        cmd.extend(["-c:a", "pcm_s16le", str(output_path)])
        result = self._run_command(cmd)
        return result.returncode == 0 and output_path.exists() and output_path.stat().st_size > 0

    def _detect_language_from_audio_sample(self, model: object, audio_path: Path) -> Tuple[str, float]:
        if whisper is None:
            return "", 0.0
        try:
            audio = whisper.load_audio(str(audio_path))
            audio = whisper.pad_or_trim(audio)
            model_n_mels = int(getattr(getattr(model, "dims", None), "n_mels", 80) or 80)
            mel = whisper.log_mel_spectrogram(audio, n_mels=model_n_mels).to(model.device)
            _, probs = model.detect_language(mel)
            top_lang, top_prob = max(probs.items(), key=lambda item: item[1])
            return self._normalize_language_code(str(top_lang)), float(top_prob)
        except Exception as exc:
            # Fallback for mismatched mel dimensions with some model/version combinations.
            try:
                mel = whisper.log_mel_spectrogram(audio, n_mels=80).to(model.device)
                _, probs = model.detect_language(mel)
                top_lang, top_prob = max(probs.items(), key=lambda item: item[1])
                return self._normalize_language_code(str(top_lang)), float(top_prob)
            except Exception:
                self._log(f"Whisper language detection failed for {audio_path.name}: {exc}")
                return "", 0.0

    def _detect_language_from_audio_array(
        self,
        model: object,
        audio_data: Any,
        sample_label: str,
    ) -> Tuple[str, float]:
        if whisper is None:
            return "", 0.0
        try:
            audio = whisper.pad_or_trim(audio_data)
            model_n_mels = int(getattr(getattr(model, "dims", None), "n_mels", 80) or 80)
            mel = whisper.log_mel_spectrogram(audio, n_mels=model_n_mels).to(model.device)
            _, probs = model.detect_language(mel)
            top_lang, top_prob = max(probs.items(), key=lambda item: item[1])
            return self._normalize_language_code(str(top_lang)), float(top_prob)
        except Exception as exc:
            try:
                mel = whisper.log_mel_spectrogram(audio, n_mels=80).to(model.device)
                _, probs = model.detect_language(mel)
                top_lang, top_prob = max(probs.items(), key=lambda item: item[1])
                return self._normalize_language_code(str(top_lang)), float(top_prob)
            except Exception:
                self._log(f"Whisper language detection failed for {sample_label}: {exc}")
                return "", 0.0

    def _detect_language_for_audio_stream(
        self,
        model: object,
        video_path: Path,
        stream_index: int,
        strategy: str,
        snippet_count: int,
        sample_seconds: float,
        duration_seconds: Optional[float] = None,
    ) -> Tuple[str, float, int]:
        duration = duration_seconds if duration_seconds is not None else self._probe_media_duration_seconds(video_path)
        effective_strategy = strategy.lower().strip() or "snippets"
        use_full = effective_strategy == "full"

        starts: List[Optional[float]] = []
        if use_full:
            starts = [None]
        else:
            count = max(1, snippet_count)
            if duration and duration > sample_seconds:
                max_start = max(0.0, duration - sample_seconds)
                starts = [random.uniform(0.0, max_start) for _ in range(count)]
            else:
                starts = [0.0 for _ in range(count)]

        score_by_lang: Counter[str] = Counter()
        successful_samples = 0
        # Minimum raw confidence from Whisper to trust a detection.
        # Below this the audio is likely music, effects, or silence.
        MIN_CONFIDENCE = 0.50

        with tempfile.TemporaryDirectory(prefix="audio_lang_", dir=str(self._get_temp_workspace_root())) as temp_dir:
            temp_root = Path(temp_dir)
            # Snippet mode can be much slower than full-stream mode when each snippet
            # spawns a separate ffmpeg process. Decode once, then sample in-memory.
            can_use_in_memory_snippets = bool(not use_full and len(starts) > 1 and whisper is not None and np is not None)
            if can_use_in_memory_snippets:
                full_stream_path = temp_root / f"stream_{stream_index}_full.wav"
                try:
                    ok = self._extract_audio_sample(
                        video_path=video_path,
                        stream_index=stream_index,
                        output_path=full_stream_path,
                        start_seconds=None,
                        sample_seconds=None,
                    )
                    if ok:
                        full_audio = whisper.load_audio(str(full_stream_path))
                        if len(full_audio) > 0:
                            sr = 16000
                            clip_samples = int(max(1.0, sample_seconds) * sr)
                            for index, start in enumerate(starts):
                                start_s = float(start or 0.0)
                                start_idx = max(0, int(start_s * sr))
                                end_idx = min(len(full_audio), start_idx + clip_samples)
                                clip = full_audio[start_idx:end_idx]
                                if len(clip) == 0:
                                    continue

                                label = f"stream {stream_index} sample {index + 1}/{len(starts)}"
                                lang, confidence = self._detect_language_from_audio_array(model, clip, label)
                                if not lang:
                                    continue

                                if confidence < MIN_CONFIDENCE:
                                    self._log(
                                        f"  {label} -> {lang} ({confidence:.2f}) — skipped (below confidence threshold)"
                                    )
                                    continue

                                successful_samples += 1
                                score_by_lang[lang] += confidence
                                self._log(f"  {label} -> {lang} ({confidence:.2f})")
                finally:
                    try:
                        full_stream_path.unlink(missing_ok=True)
                    except OSError:
                        pass

            if not can_use_in_memory_snippets:
                for index, start in enumerate(starts):
                    sample_path = temp_root / f"stream_{stream_index}_{index + 1:02d}.wav"
                    try:
                        sample_duration = None if use_full else sample_seconds
                        ok = self._extract_audio_sample(
                            video_path=video_path,
                            stream_index=stream_index,
                            output_path=sample_path,
                            start_seconds=start,
                            sample_seconds=sample_duration,
                        )
                        if not ok:
                            continue

                        lang, confidence = self._detect_language_from_audio_sample(model, sample_path)
                        if not lang:
                            continue

                        if confidence < MIN_CONFIDENCE:
                            self._log(
                                f"  stream {stream_index} sample {index + 1}/{len(starts)} -> "
                                f"{lang} ({confidence:.2f}) — skipped (below confidence threshold)"
                            )
                            continue

                        successful_samples += 1
                        score_by_lang[lang] += confidence
                        self._log(
                            f"  stream {stream_index} sample {index + 1}/{len(starts)} -> {lang} ({confidence:.2f})"
                        )
                    finally:
                        try:
                            sample_path.unlink(missing_ok=True)
                        except OSError:
                            pass

        if not score_by_lang:
            return "", 0.0, successful_samples

        detected_lang, detected_score = score_by_lang.most_common(1)[0]
        total_score = float(sum(score_by_lang.values()))
        normalized_confidence = detected_score / total_score if total_score > 0 else 0.0
        return detected_lang, normalized_confidence, successful_samples

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

    def _build_output_paths(
        self,
        source: Path,
        suffix: str,
        overwrite: bool,
        output_root: Optional[str] = None,
    ) -> tuple[Path, Optional[Path]]:
        if overwrite:
            temp_output = source.with_name(f"{source.stem}.tmp_subtitle_tool{source.suffix}")
            return temp_output, source

        destination_parent = source.parent
        if output_root:
            out_dir = Path(output_root).expanduser().resolve()
            out_dir.mkdir(parents=True, exist_ok=True)
            destination_parent = out_dir

        desired = destination_parent / f"{source.stem}{suffix}{source.suffix}"
        if not desired.exists():
            return desired, None

        index = 1
        while True:
            candidate = destination_parent / f"{source.stem}{suffix}_{index}{source.suffix}"
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
                *self._hw_accel_flags(),
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
                    *self._hw_accel_flags(),
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
        output_root: Optional[str] = None,
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

            output_path, replace_target = self._build_output_paths(video, output_suffix, overwrite, output_root)
            cmd = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                "error",
                "-nostats",
                *self._hw_accel_flags(),
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
        output_root: Optional[str] = None,
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

            output_path, replace_target = self._build_output_paths(video, output_suffix, overwrite, output_root)

            cmd: List[str] = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                "error",
                "-nostats",
                *self._hw_accel_flags(),
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
        output_root: Optional[str] = None,
        backend: str = "auto",
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
            
            output_parent = video.parent
            if output_root and not overwrite:
                out_dir = Path(output_root).expanduser().resolve()
                out_dir.mkdir(parents=True, exist_ok=True)
                output_parent = out_dir
            output_path = output_parent / f"{video.stem}{output_suffix}{target_ext}"
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
            
            selected_backend, backend_reason = self._choose_conversion_backend(target_format, backend)
            if not selected_backend:
                summary.failed += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "failed",
                        "reason": backend_reason,
                    }
                )
                continue

            self._log(
                f"Converting {video.name} to {target_format.upper()} "
                f"using {selected_backend} ({backend_reason})..."
            )

            if selected_backend == "ffmpeg":
                cmd = [
                    self.ffmpeg_bin,
                    *self._hw_accel_flags(),
                    "-loglevel",
                    self.ffmpeg_loglevel,
                    "-nostats",
                    "-i",
                    str(video),
                    "-map",
                    "0",
                ]
                if target_format == "mp4":
                    cmd.extend(["-c:v", "copy", "-c:a", "copy", "-c:s", "mov_text"])
                else:
                    cmd.extend(["-c", "copy"])
                cmd.extend(["-y", str(output_path)])
                result = self._run_command(cmd, description="ffmpeg conversion")
            elif selected_backend == "mkvtoolnix":
                cmd = [self.mkvmerge_bin, "-o", str(output_path), str(video)]
                result = self._run_command(cmd, description="mkvmerge conversion", max_feedback_lines=30)
            else:
                hb_format = "av_mp4" if target_format == "mp4" else "av_mkv"
                cmd = [
                    self.handbrake_bin,
                    "-i",
                    str(video),
                    "-o",
                    str(output_path),
                    "--format",
                    hb_format,
                    "--all-audio",
                    "--all-subtitles",
                    "--aencoder",
                    "copy",
                    "--encoder",
                    "x264",
                    "--quality",
                    "20",
                    "--markers",
                ]
                result = self._run_command(cmd, description="HandBrake conversion", max_feedback_lines=30)
            
            if result.returncode != 0:
                summary.failed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "failed",
                    "reason": result.stderr.strip() or "conversion failed"
                })
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass
                continue

            if selected_backend == "handbrake":
                is_valid, invalid_reason = self._validate_transcoded_output(video, output_path)
            else:
                is_valid, invalid_reason = self._validate_muxed_output(
                    video,
                    output_path,
                    allow_reduced_subtitles=(target_format == "mp4"),
                )

            if not is_valid:
                summary.failed += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "failed",
                        "reason": f"conversion validation failed: {invalid_reason}",
                    }
                )
                self._log(f"Validation failed for {video.name}: {invalid_reason}")
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass
                continue
            
            if replace_target:
                output_path.replace(replace_target)
            
            summary.processed += 1
            summary.details.append({
                "file": str(video),
                "status": "converted",
                "reason": f"converted from {current_ext} to {target_ext} using {selected_backend}"
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
        backend: str = "auto",
    ) -> OperationSummary:
        """Repair corrupted video metadata by rebuilding containers with selected backend."""
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
            
            source_ext = video.suffix.lower()
            preferred_backend = (backend or "auto").strip().lower()
            selected_backend = preferred_backend
            if selected_backend not in REPAIR_BACKENDS:
                selected_backend = "auto"

            if selected_backend == "auto":
                if source_ext == ".mkv" and self._resolve_binary(self.mkvmerge_bin):
                    selected_backend = "mkvtoolnix"
                else:
                    selected_backend = "ffmpeg"

            if selected_backend == "mkvtoolnix" and source_ext != ".mkv":
                selected_backend = "ffmpeg"

            if selected_backend == "mkvtoolnix" and not self._resolve_binary(self.mkvmerge_bin):
                selected_backend = "ffmpeg"

            if selected_backend == "makemkv" and not self._resolve_binary(self.makemkvcon_bin):
                summary.failed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "failed",
                    "reason": "makemkv backend selected but makemkvcon was not found"
                })
                continue

            if selected_backend == "makemkv" and source_ext != ".mkv":
                summary.failed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "failed",
                    "reason": "makemkv repair backend currently supports .mkv inputs only"
                })
                continue

            self._log(f"Repairing {video.name} using {selected_backend} backend...")
            
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
                if selected_backend == "mkvtoolnix":
                    cmd = [self.mkvmerge_bin, "-o", str(temp_file), str(video)]
                    result = self._run_command(cmd, description="mkvmerge metadata repair", max_feedback_lines=30)
                elif selected_backend == "makemkv":
                    ok, reason = self._run_makemkv_repair(video, temp_file)
                    result = subprocess.CompletedProcess(
                        args=[self.makemkvcon_bin],
                        returncode=0 if ok else 1,
                        stdout="",
                        stderr=reason if not ok else "",
                    )
                else:
                    cmd = [
                        self.ffmpeg_bin,
                        *self._hw_accel_flags(),
                        "-loglevel",
                        self.ffmpeg_loglevel,
                        "-nostats",
                        "-fflags", "+genpts",  # Generate presentation timestamps
                        "-err_detect", "ignore_err",  # Ignore errors
                        "-i", str(video),
                        "-map", "0",
                        "-c", "copy",  # Copy all streams
                        "-y",
                        str(temp_file)
                    ]
                    result = self._run_command(cmd, description="ffmpeg metadata repair")
                
                # Check if output file was created and has reasonable size
                if temp_file.exists() and temp_file.stat().st_size > 1000:
                    is_valid, invalid_reason = self._validate_muxed_output(video, temp_file)
                    if not is_valid:
                        summary.failed += 1
                        summary.details.append({
                            "file": str(video),
                            "status": "failed",
                            "reason": f"repair validation failed: {invalid_reason}",
                        })
                        try:
                            os.remove(str(temp_file))
                        except OSError:
                            pass
                        continue

                    # Replace original with repaired version
                    os.remove(str(video))
                    os.rename(str(temp_file), str(video))
                    summary.processed += 1
                    summary.details.append({
                        "file": str(video),
                        "status": "repaired",
                        "reason": f"container rebuilt successfully with {selected_backend}"
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

    @staticmethod
    def _safe_float(value: object, default: float = 0.0) -> float:
        try:
            return float(value)
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _extract_transcript_text(segments: List[Dict[str, object]]) -> str:
        pieces: List[str] = []
        for seg in segments:
            txt = str(seg.get("text") or "").strip()
            if txt:
                pieces.append(txt)
        return " ".join(pieces).strip()

    def _text_to_timestamp_segments(self, text: str, total_duration: Optional[float]) -> List[Dict[str, object]]:
        cleaned = re.sub(r"\s+", " ", (text or "").strip())
        if not cleaned:
            return []

        chunks = [p.strip() for p in re.split(r"(?<=[.!?])\s+", cleaned) if p.strip()]
        if not chunks:
            words = cleaned.split()
            step = 10
            chunks = [" ".join(words[i : i + step]) for i in range(0, len(words), step)]

        if not chunks:
            return []

        duration = total_duration if total_duration and total_duration > 0 else max(30.0, len(cleaned.split()) * 0.55)
        weights = [max(1, len(chunk.split())) for chunk in chunks]
        total_weight = sum(weights)

        segments: List[Dict[str, object]] = []
        cursor = 0.0
        for idx, chunk in enumerate(chunks):
            span = duration * (weights[idx] / total_weight)
            start = cursor
            end = min(duration, cursor + span)
            if end <= start:
                end = start + 0.75
            segments.append({"start": start, "end": end, "text": chunk})
            cursor = end
        return segments

    def _normalize_backend_segments(self, segments: List[Dict[str, object]]) -> List[Dict[str, object]]:
        normalized: List[Dict[str, object]] = []
        for seg in segments:
            start = self._safe_float(seg.get("start"), 0.0)
            end = self._safe_float(seg.get("end"), start + 0.75)
            text = str(seg.get("text") or "").strip()
            if not text:
                continue
            if end <= start:
                end = start + 0.75
            normalized.append({"start": start, "end": end, "text": text})
        return normalized

    def _detect_installed_transcription_backends(self) -> set[str]:
        installed: set[str] = {"text-to-timestamps"}

        if whisper is not None:
            installed.add("openai-whisper")
        if importlib.util.find_spec("faster_whisper") is not None:
            installed.add("faster-whisper")
        if importlib.util.find_spec("whisperx") is not None:
            installed.add("whisperx")
        if importlib.util.find_spec("stable_whisper") is not None:
            installed.add("stable-ts")
        if importlib.util.find_spec("whisper_timestamped") is not None:
            installed.add("whisper-timestamped")
        if importlib.util.find_spec("speechbrain") is not None:
            installed.add("speechbrain")
        if importlib.util.find_spec("vosk") is not None:
            installed.add("vosk")
        return installed

    def _resolve_transcription_backend(self, requested_backend: str) -> Tuple[Optional[str], str]:
        backend = (requested_backend or "auto").strip().lower()
        if backend not in TRANSCRIBE_BACKENDS:
            backend = "auto"

        installed = self._detect_installed_transcription_backends()
        preference = [
            "openai-whisper",
            "faster-whisper",
            "whisperx",
            "stable-ts",
            "whisper-timestamped",
            "vosk",
            "speechbrain",
            "text-to-timestamps",
        ]

        if backend == "auto":
            for candidate in preference:
                if candidate in installed:
                    return candidate, f"auto-selected {candidate}"
            return None, "no supported transcription backend is installed"

        if backend == "text-to-timestamps":
            return "text-to-timestamps", "using heuristic text-to-timestamps"

        if backend not in installed:
            return None, f"requested backend '{backend}' is not installed"
        return backend, f"using {backend}"

    def _transcribe_with_openai_whisper(
        self,
        video: Path,
        model_size: str,
        language: Optional[str],
        model_cache: Dict[str, object],
    ) -> Tuple[List[Dict[str, object]], str]:
        if whisper is None:
            raise RuntimeError("openai-whisper is not installed")

        model_key = f"openai:{model_size}"
        model = model_cache.get(model_key)
        if model is None:
            self._log(f"Loading openai-whisper model: {model_size}...")
            model = whisper.load_model(model_size)
            model_cache[model_key] = model

        options: Dict[str, object] = {"task": "transcribe"}
        if language:
            options["language"] = language
        result = model.transcribe(str(video), **options)
        segments = self._normalize_backend_segments(result.get("segments", []))
        detected = str(result.get("language") or "").strip().lower()
        return segments, detected

    def _transcribe_with_faster_whisper(
        self,
        video: Path,
        model_size: str,
        language: Optional[str],
        model_cache: Dict[str, object],
    ) -> Tuple[List[Dict[str, object]], str]:
        module = importlib.import_module("faster_whisper")
        model_key = f"faster:{model_size}"
        model = model_cache.get(model_key)
        if model is None:
            self._log(f"Loading faster-whisper model: {model_size}...")
            model = module.WhisperModel(model_size, device="auto", compute_type="int8")
            model_cache[model_key] = model

        segments_iter, info = model.transcribe(
            str(video),
            language=language or None,
            task="transcribe",
        )
        rows: List[Dict[str, object]] = []
        for seg in segments_iter:
            rows.append({"start": float(seg.start), "end": float(seg.end), "text": str(seg.text).strip()})
        detected = str(getattr(info, "language", "") or "").strip().lower()
        return self._normalize_backend_segments(rows), detected

    def _transcribe_with_whisperx(
        self,
        video: Path,
        model_size: str,
        language: Optional[str],
        model_cache: Dict[str, object],
    ) -> Tuple[List[Dict[str, object]], str]:
        module = importlib.import_module("whisperx")
        model_key = f"whisperx:{model_size}"
        model = model_cache.get(model_key)
        if model is None:
            self._log(f"Loading WhisperX model: {model_size}...")
            model = module.load_model(model_size, device="cpu", compute_type="int8")
            model_cache[model_key] = model

        audio = module.load_audio(str(video))
        result = model.transcribe(audio, batch_size=8, language=language or None)
        segments = result.get("segments", [])
        detected = str(result.get("language") or "").strip().lower()

        try:
            if detected:
                align_model, metadata = module.load_align_model(language_code=detected, device="cpu")
                aligned = module.align(segments, align_model, metadata, audio, "cpu")
                if isinstance(aligned, dict) and isinstance(aligned.get("segments"), list):
                    segments = aligned.get("segments", segments)
        except Exception as exc:
            self._log(f"WhisperX alignment skipped: {exc}")

        return self._normalize_backend_segments(segments), detected

    def _transcribe_with_stable_ts(
        self,
        video: Path,
        model_size: str,
        language: Optional[str],
        model_cache: Dict[str, object],
    ) -> Tuple[List[Dict[str, object]], str]:
        module = importlib.import_module("stable_whisper")
        model_key = f"stable:{model_size}"
        model = model_cache.get(model_key)
        if model is None:
            self._log(f"Loading stable-ts model: {model_size}...")
            model = module.load_model(model_size)
            model_cache[model_key] = model

        result = model.transcribe(str(video), task="transcribe", language=language or None)
        rows: List[Dict[str, object]] = []
        for seg in getattr(result, "segments", []) or []:
            rows.append(
                {
                    "start": float(getattr(seg, "start", 0.0)),
                    "end": float(getattr(seg, "end", 0.0)),
                    "text": str(getattr(seg, "text", "")).strip(),
                }
            )
        detected = str(getattr(result, "language", "") or "").strip().lower()
        return self._normalize_backend_segments(rows), detected

    def _transcribe_with_whisper_timestamped(
        self,
        video: Path,
        model_size: str,
        language: Optional[str],
        model_cache: Dict[str, object],
    ) -> Tuple[List[Dict[str, object]], str]:
        if whisper is None:
            raise RuntimeError("openai-whisper is required for whisper-timestamped")
        wt_module = importlib.import_module("whisper_timestamped")

        model_key = f"wts:{model_size}"
        model = model_cache.get(model_key)
        if model is None:
            self._log(f"Loading whisper-timestamped base model: {model_size}...")
            model = whisper.load_model(model_size)
            model_cache[model_key] = model

        result = wt_module.transcribe(model, str(video), task="transcribe", language=language or None)
        segments = self._normalize_backend_segments(result.get("segments", []))
        detected = str(result.get("language") or "").strip().lower()
        return segments, detected

    def _transcribe_with_vosk(
        self,
        video: Path,
        language: Optional[str],
        model_cache: Dict[str, object],
    ) -> Tuple[List[Dict[str, object]], str]:
        vosk_module = importlib.import_module("vosk")
        import wave

        model_path = os.getenv("VOSK_MODEL_PATH", "").strip()
        if not model_path:
            raise RuntimeError("Set VOSK_MODEL_PATH to a local Vosk model directory")

        model_key = f"vosk:{model_path}"
        vosk_model = model_cache.get(model_key)
        if vosk_model is None:
            self._log(f"Loading Vosk model from: {model_path}")
            vosk_model = vosk_module.Model(model_path)
            model_cache[model_key] = vosk_model

        with tempfile.TemporaryDirectory(prefix="vosk_audio_", dir=str(self._get_temp_workspace_root())) as temp_dir:
            wav_path = Path(temp_dir) / "audio_16k.wav"
            cmd = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                self.ffmpeg_loglevel,
                "-nostats",
                "-i",
                str(video),
                "-ac",
                "1",
                "-ar",
                "16000",
                str(wav_path),
            ]
            result = self._run_command(cmd, description="prepare wav for vosk")
            if result.returncode != 0:
                raise RuntimeError(result.stderr.strip() or "ffmpeg failed preparing Vosk WAV")

            words: List[Dict[str, object]] = []
            with wave.open(str(wav_path), "rb") as wf:
                recognizer = vosk_module.KaldiRecognizer(vosk_model, wf.getframerate())
                recognizer.SetWords(True)
                while True:
                    data = wf.readframes(4000)
                    if not data:
                        break
                    if recognizer.AcceptWaveform(data):
                        payload = json.loads(recognizer.Result() or "{}")
                        words.extend(payload.get("result", []))

                payload = json.loads(recognizer.FinalResult() or "{}")
                words.extend(payload.get("result", []))

        if not words:
            raise RuntimeError("Vosk produced no timed words")

        segments: List[Dict[str, object]] = []
        chunk: List[Dict[str, object]] = []
        for word in words:
            chunk.append(word)
            if len(chunk) >= 10:
                start = self._safe_float(chunk[0].get("start"), 0.0)
                end = self._safe_float(chunk[-1].get("end"), start + 0.8)
                text = " ".join(str(w.get("word") or "").strip() for w in chunk).strip()
                if text:
                    segments.append({"start": start, "end": end, "text": text})
                chunk = []
        if chunk:
            start = self._safe_float(chunk[0].get("start"), 0.0)
            end = self._safe_float(chunk[-1].get("end"), start + 0.8)
            text = " ".join(str(w.get("word") or "").strip() for w in chunk).strip()
            if text:
                segments.append({"start": start, "end": end, "text": text})

        detected = (language or "").strip().lower()
        return self._normalize_backend_segments(segments), detected

    def _transcribe_with_speechbrain_text(self, video: Path, model_cache: Dict[str, object]) -> str:
        savedir = str(self._get_temp_workspace_root() / "speechbrain_asr")
        load_errors: List[str] = []

        # SpeechBrain is more reliable with explicit PCM WAV input than direct MP4/MKV input.
        # Some runtime environments lack torchaudio/ffmpeg backend wiring for container decode.
        with tempfile.TemporaryDirectory(prefix="speechbrain_audio_", dir=str(self._get_temp_workspace_root())) as temp_dir:
            wav_path = Path(temp_dir) / "speechbrain_16k.wav"
            if not self._extract_audio_sample(
                video_path=video,
                stream_index=0,
                output_path=wav_path,
                start_seconds=None,
                sample_seconds=None,
            ):
                raise RuntimeError("SpeechBrain input audio extraction failed (ffmpeg could not produce 16k WAV)")

            asr_variants = [
                ("speechbrain.inference.ASR", "ASR", "speechbrain_asr"),
                ("speechbrain.inference", "EncoderDecoderASR", "speechbrain_enc"),
                ("speechbrain.pretrained", "EncoderDecoderASR", "speechbrain_pretrained_enc"),
            ]

            for module_name, class_name, cache_key in asr_variants:
                try:
                    module = importlib.import_module(module_name)
                    asr_cls = getattr(module, class_name, None)
                    if asr_cls is None:
                        load_errors.append(f"{module_name}.{class_name} unavailable")
                        continue

                    model = model_cache.get(cache_key)
                    if model is None:
                        model = asr_cls.from_hparams(
                            source="speechbrain/asr-crdnn-rnnlm-librispeech",
                            savedir=savedir,
                        )
                        model_cache[cache_key] = model

                    text = str(model.transcribe_file(str(wav_path))).strip()
                    if text:
                        return text
                    load_errors.append(f"{module_name}.{class_name} returned empty transcript")
                except Exception as exc:
                    load_errors.append(f"{module_name}.{class_name}: {type(exc).__name__}: {exc}")

        if load_errors:
            raise RuntimeError("SpeechBrain ASR failed: " + " | ".join(load_errors[-3:]))
        raise RuntimeError("SpeechBrain ASR model could not be loaded")

    def _transcribe_with_backend(
        self,
        video: Path,
        backend: str,
        model_size: str,
        language: Optional[str],
        model_cache: Dict[str, object],
    ) -> Tuple[List[Dict[str, object]], str]:
        if backend == "openai-whisper":
            return self._transcribe_with_openai_whisper(video, model_size, language, model_cache)
        if backend == "faster-whisper":
            return self._transcribe_with_faster_whisper(video, model_size, language, model_cache)
        if backend == "whisperx":
            return self._transcribe_with_whisperx(video, model_size, language, model_cache)
        if backend == "stable-ts":
            return self._transcribe_with_stable_ts(video, model_size, language, model_cache)
        if backend == "whisper-timestamped":
            return self._transcribe_with_whisper_timestamped(video, model_size, language, model_cache)
        if backend == "vosk":
            return self._transcribe_with_vosk(video, language, model_cache)
        if backend == "speechbrain":
            transcript = self._transcribe_with_speechbrain_text(video, model_cache)
            duration = self._probe_media_duration_seconds(video)
            return self._text_to_timestamp_segments(transcript, duration), (language or "eng")
        if backend == "text-to-timestamps":
            transcript = ""
            for candidate in ["speechbrain", "openai-whisper", "faster-whisper", "vosk"]:
                try:
                    if candidate == "speechbrain":
                        transcript = self._transcribe_with_speechbrain_text(video, model_cache)
                    else:
                        candidate_segments, _ = self._transcribe_with_backend(
                            video,
                            candidate,
                            model_size,
                            language,
                            model_cache,
                        )
                        transcript = self._extract_transcript_text(candidate_segments)
                    if transcript:
                        break
                except Exception:
                    continue
            if not transcript:
                raise RuntimeError("No transcript source was available for text-to-timestamps")
            duration = self._probe_media_duration_seconds(video)
            return self._text_to_timestamp_segments(transcript, duration), (language or "")

        raise RuntimeError(f"Unsupported transcription backend: {backend}")
    
    def generate_subtitles(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        model_size: str = "base",
        output_format: str = "srt",
        language: Optional[str] = None,
        backend: str = "auto",
    ) -> OperationSummary:
        """Generate subtitles from video audio using selectable AI backends."""
        summary = OperationSummary(action="generate_subtitles")

        if pysubs2 is None:
            self._log("ERROR: pysubs2 (aka pysub2) not installed. Run: pip install pysubs2")
            summary.failed = 1
            summary.details.append({
                "file": "N/A",
                "status": "failed",
                "reason": "pysubs2 library not installed"
            })
            return summary

        selected_backend, backend_reason = self._resolve_transcription_backend(backend)
        if not selected_backend:
            summary.failed = 1
            summary.details.append(
                {
                    "file": "N/A",
                    "status": "failed",
                    "reason": backend_reason,
                }
            )
            self._log(f"ERROR: {backend_reason}")
            return summary

        self._log(f"Subtitle generation backend: {selected_backend} ({backend_reason})")
        
        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)
        
        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)
        
        if not videos:
            self._log("No video files found to generate subtitles")
            return summary

        model_cache: Dict[str, object] = {}
        
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
                self._log(f"  Transcribing audio with {selected_backend}...")
                segments, detected_lang = self._transcribe_with_backend(
                    video=video,
                    backend=selected_backend,
                    model_size=model_size,
                    language=language,
                    model_cache=model_cache,
                )
                if not segments:
                    raise RuntimeError("transcription backend returned no segments")
                
                # Create subtitle file using pysubs2
                subs = pysubs2.SSAFile()
                for segment in segments:
                    event = pysubs2.SSAEvent(
                        start=int(segment["start"] * 1000),  # Convert to milliseconds
                        end=int(segment["end"] * 1000),
                        text=segment["text"].strip()
                    )
                    subs.append(event)
                
                # Save subtitle file
                subs.save(str(output_path))

                detected_lang = detected_lang or "unknown"
                self._log(f"  Generated {output_path.name} (language: {detected_lang})")
                self._log(f"  Saved subtitle to: {output_path}")
                
                summary.processed += 1
                summary.details.append({
                    "file": str(video),
                    "status": "generated",
                    "reason": (
                        f"created {output_path.name} with {len(segments)} segments "
                        f"using {selected_backend}"
                    ),
                    "output_path": str(output_path),
                    "backend": selected_backend,
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

    def detect_and_tag_audio_languages(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        model_size: str = "base",
        strategy: str = "snippets",
        snippet_count: int = 3,
        sample_seconds: float = 25.0,
        overwrite: bool = False,
        output_suffix: str = "_langtagged",
        overwrite_existing_tags: bool = False,
        detect_only: bool = False,
        output_root: Optional[str] = None,
    ) -> OperationSummary:
        """Detect audio stream languages with Whisper and optionally write language metadata tags."""
        summary = OperationSummary(action="tag_audio_language")

        if whisper is None:
            self._log("ERROR: openai-whisper not installed. Run: pip install openai-whisper")
            summary.failed = 1
            summary.details.append(
                {
                    "file": "N/A",
                    "status": "failed",
                    "reason": "Whisper library not installed",
                }
            )
            return summary

        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)

        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)

        if not videos:
            self._log("No video files found to detect audio language")
            return summary

        try:
            self._log(f"Loading Whisper model for audio language tagging: {model_size}...")
            model = whisper.load_model(model_size)
        except Exception as exc:
            summary.failed = len(videos)
            for video in videos:
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "failed",
                        "reason": f"Failed to load model: {exc}",
                    }
                )
            return summary

        if detect_only:
            self._log("Detect-only mode enabled: metadata tags will NOT be written to media files.")

        for video in videos:
            self._log(f"Detecting audio language: {video.name}")
            audio_streams = self._probe_audio_streams(video)
            if not audio_streams:
                summary.skipped += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "skipped",
                        "reason": "no audio streams",
                    }
                )
                continue

            # Probe duration once per file so we can estimate per-stream sizes.
            video_duration = self._probe_media_duration_seconds(video)

            updates: List[Tuple[int, str]] = []
            detected_streams: List[Dict[str, object]] = []
            changed_streams = 0
            attempted_streams = 0
            for stream_order, stream in enumerate(audio_streams):
                stream_index = int(stream.get("index") or -1)
                if stream_index < 0:
                    continue

                tags = stream.get("tags") or {}
                existing_lang = ""
                if isinstance(tags, dict):
                    existing_lang = self._normalize_language_code(str(tags.get("language") or ""))

                if existing_lang and not overwrite_existing_tags:
                    continue

                attempted_streams += 1
                # Estimate stream size from per-stream bitrate × duration.
                _br_raw = stream.get("bit_rate")
                _br_bps: Optional[int] = None
                try:
                    _br_bps = int(_br_raw) if _br_raw else None
                except (TypeError, ValueError):
                    pass
                _estimated_bytes: Optional[int] = (
                    int(_br_bps / 8 * video_duration)
                    if _br_bps and video_duration
                    else None
                )
                detected_lang, confidence, sample_hits = self._detect_language_for_audio_stream(
                    model=model,
                    video_path=video,
                    stream_index=stream_index,
                    strategy=strategy,
                    snippet_count=snippet_count,
                    sample_seconds=sample_seconds,
                    duration_seconds=video_duration,
                )
                if not detected_lang:
                    self._log(
                        f"  Could not detect language for stream {stream_order + 1} ({stream_index}); "
                        f"successful samples: {sample_hits}"
                    )
                    # Still track in detected_streams so the stream appears in the
                    # language selection dialog and the user can keep or prune it.
                    _entry: Dict[str, object] = {
                        "stream_order": stream_order,
                        "stream_index": stream_index,
                        "language": "unknown",
                        "confidence": 0.0,
                    }
                    if _estimated_bytes is not None:
                        _entry["estimated_bytes"] = _estimated_bytes
                    detected_streams.append(_entry)
                    continue

                _det_entry: Dict[str, object] = {
                    "stream_order": stream_order,
                    "stream_index": stream_index,
                    "language": detected_lang,
                    "confidence": round(float(confidence), 4),
                }
                if _estimated_bytes is not None:
                    _det_entry["estimated_bytes"] = _estimated_bytes
                detected_streams.append(_det_entry)

                updates.append((stream_order, detected_lang))
                changed_streams += 1
                self._log(
                    f"  Stream {stream_order + 1} ({stream_index}) -> {detected_lang} "
                    f"(confidence {confidence:.2f})"
                )

            if not updates:
                summary.skipped += 1
                reason = "no detectable stream languages"
                if attempted_streams == 0 and not overwrite_existing_tags:
                    reason = "language tags already present"
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "skipped",
                        "reason": reason,
                        "detected_streams": detected_streams,
                    }
                )
                continue

            if detect_only:
                summary.processed += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "detected",
                        "reason": f"detected {changed_streams} audio stream language(s); no file changes",
                        "detected_streams": detected_streams,
                    }
                )
                continue

            output_path, replace_target = self._build_output_paths(video, output_suffix, overwrite, output_root)
            src_ext = video.suffix.lower()
            is_ts_source = src_ext in (".m2ts", ".ts")

            cmd: List[str] = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                "warning",
                "-nostats",
            ]
            if is_ts_source:
                # Large/complex Blu-ray transport streams need a bigger muxing queue
                cmd.extend(self._ts_input_stability_flags())
            cmd.extend([
                "-i",
                str(video),
            ])
            if is_ts_source:
                cmd.extend(["-max_muxing_queue_size", "1024"])
            cmd.extend(["-map", "0", "-c", "copy"])
            for stream_order, language_code in updates:
                cmd.extend([f"-metadata:s:a:{stream_order}", f"language={language_code}"])
            cmd.append(str(output_path))

            self._log("  Remuxing full source file with -map 0 (all streams); detection snippets do not limit output length.")

            result = self._run_command(cmd)
            ffmpeg_msgs = (result.stderr or "").strip()
            if ffmpeg_msgs:
                for _line in ffmpeg_msgs.splitlines()[:15]:
                    self._log(f"  ffmpeg: {_line}")

            if result.returncode != 0:
                summary.failed += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "failed",
                        "reason": ffmpeg_msgs[:300] or "ffmpeg failed",
                        "detected_streams": detected_streams,
                    }
                )
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass
                continue

            is_valid, invalid_reason = self._validate_muxed_output(video, output_path)
            if not is_valid:
                self._log(f"  Output validation failed: {invalid_reason}")
                src_v, src_a, src_s = self._probe_stream_counts(video)
                out_v, out_a, out_s = self._probe_stream_counts(output_path)
                self._log(
                    f"  Stream counts source(v/a/s)={src_v}/{src_a}/{src_s} -> "
                    f"output(v/a/s)={out_v}/{out_a}/{out_s}"
                )
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass
                # The m2ts muxer can silently truncate large/complex Blu-ray streams
                # (TrueHD/Atmos, DTS-HD MA, etc.) and still exit with code 0.
                # Fall back to MKV, which handles all Blu-ray codec types reliably.
                if is_ts_source:
                    self._log("  Retrying as MKV to work around m2ts muxer limitation...")
                    mkv_output = output_path.with_suffix(".mkv")
                    mkv_cmd: List[str] = [
                        self.ffmpeg_bin,
                        "-y",
                        "-loglevel",
                        "warning",
                        "-nostats",
                        *self._hw_accel_flags(),
                    ]
                    if is_ts_source:
                        mkv_cmd.extend(self._ts_input_stability_flags())
                    mkv_cmd.extend([
                        "-i",
                        str(video),
                        "-max_muxing_queue_size",
                        "1024",
                        "-map",
                        "0",
                        "-c",
                        "copy",
                    ])
                    for stream_order, language_code in updates:
                        mkv_cmd.extend([f"-metadata:s:a:{stream_order}", f"language={language_code}"])
                    mkv_cmd.append(str(mkv_output))
                    mkv_result = self._run_command(mkv_cmd)
                    mkv_msgs = (mkv_result.stderr or "").strip()
                    if mkv_msgs:
                        for _line in mkv_msgs.splitlines()[:15]:
                            self._log(f"  ffmpeg: {_line}")
                    if mkv_result.returncode == 0:
                        mkv_valid, mkv_reason = self._validate_muxed_output(video, mkv_output)
                        if mkv_valid:
                            self._log(f"  MKV fallback succeeded -> {mkv_output.name}")
                            output_path = mkv_output
                            replace_target = None  # cannot overwrite .m2ts with .mkv in-place
                            is_valid = True
                        else:
                            self._log(f"  MKV fallback validation failed: {mkv_reason}")
                            src_v, src_a, src_s = self._probe_stream_counts(video)
                            out_v, out_a, out_s = self._probe_stream_counts(mkv_output)
                            self._log(
                                f"  Stream counts source(v/a/s)={src_v}/{src_a}/{src_s} -> "
                                f"output(v/a/s)={out_v}/{out_a}/{out_s}"
                            )
                            try:
                                mkv_output.unlink(missing_ok=True)
                            except OSError:
                                pass
                    else:
                        try:
                            mkv_output.unlink(missing_ok=True)
                        except OSError:
                            pass

                if not is_valid:
                    summary.failed += 1
                    summary.details.append(
                        {
                            "file": str(video),
                            "status": "failed",
                            "reason": invalid_reason,
                            "detected_streams": detected_streams,
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
                    "reason": f"tagged {changed_streams} audio stream(s)",
                    "detected_streams": detected_streams,
                }
            )

        return summary

    def prune_audio_streams(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        keep_audio_orders_by_file: Dict[str, List[int]],
        overwrite: bool = False,
        output_suffix: str = "_audiopruned",
        output_root: Optional[str] = None,
    ) -> OperationSummary:
        """Keep only selected audio stream orders for each file using stream copy."""
        summary = OperationSummary(action="prune_audio_streams")

        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)

        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)
        if not videos:
            self._log("No video files found for audio pruning")
            return summary

        normalized_keep: Dict[str, List[int]] = {}
        for key, values in (keep_audio_orders_by_file or {}).items():
            try:
                orders = sorted({int(v) for v in values if int(v) >= 0})
            except Exception:
                orders = []
            normalized_keep[str(Path(key).resolve())] = orders

        for video in videos:
            video_key = str(video.resolve())
            keep_orders = normalized_keep.get(video_key, [])
            is_ts_source = video.suffix.lower() in (".ts", ".m2ts")
            audio_streams = self._probe_audio_streams(video)
            audio_count = len(audio_streams)

            if audio_count == 0:
                summary.skipped += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "skipped",
                        "reason": "no audio streams",
                    }
                )
                continue

            valid_keep_orders = [o for o in keep_orders if 0 <= o < audio_count]
            if not valid_keep_orders:
                summary.skipped += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "skipped",
                        "reason": "no selected audio streams to keep",
                    }
                )
                continue

            output_path, replace_target = self._build_output_paths(video, output_suffix, overwrite, output_root)
            cmd: List[str] = [
                self.ffmpeg_bin,
                "-y",
                "-loglevel",
                "error",
                "-nostats",
                *self._hw_accel_flags(),
            ]
            if is_ts_source:
                cmd.extend(self._ts_input_stability_flags())
            cmd.extend([
                "-i",
                str(video),
                "-map",
                "0:v?",
                "-map",
                "0:s?",
                "-map",
                "0:d?",
                "-map",
                "0:t?",
            ])
            for order in valid_keep_orders:
                cmd.extend(["-map", f"0:a:{order}"])
            if is_ts_source:
                cmd.extend(["-max_muxing_queue_size", "1024"])
            cmd.extend(["-c", "copy", str(output_path)])

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
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass
                continue

            is_valid, invalid_reason = self._validate_muxed_output(
                video,
                output_path,
                allow_reduced_audio=True,
            )
            if not is_valid:
                self._log(f"  Output validation failed: {invalid_reason}")
                src_v, src_a, src_s = self._probe_stream_counts(video)
                out_v, out_a, out_s = self._probe_stream_counts(output_path)
                self._log(
                    f"  Stream counts source(v/a/s)={src_v}/{src_a}/{src_s} -> "
                    f"output(v/a/s)={out_v}/{out_a}/{out_s}"
                )
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass

                if is_ts_source:
                    self._log("  Retrying prune as MKV to work around m2ts muxer limitation...")
                    mkv_output = output_path.with_suffix(".mkv")
                    mkv_cmd: List[str] = [
                        self.ffmpeg_bin,
                        "-y",
                        "-loglevel",
                        "warning",
                        "-nostats",
                        *self._hw_accel_flags(),
                    ]
                    mkv_cmd.extend(self._ts_input_stability_flags())
                    mkv_cmd.extend([
                        "-i",
                        str(video),
                        "-map",
                        "0:v?",
                        "-map",
                        "0:s?",
                        "-map",
                        "0:d?",
                        "-map",
                        "0:t?",
                    ])
                    for order in valid_keep_orders:
                        mkv_cmd.extend(["-map", f"0:a:{order}"])
                    mkv_cmd.extend(["-max_muxing_queue_size", "1024", "-c", "copy", str(mkv_output)])

                    mkv_result = self._run_command(mkv_cmd)
                    if mkv_result.returncode == 0:
                        mkv_valid, mkv_reason = self._validate_muxed_output(
                            video,
                            mkv_output,
                            allow_reduced_audio=True,
                        )
                        if mkv_valid:
                            self._log(f"  MKV prune fallback succeeded -> {mkv_output.name}")
                            output_path = mkv_output
                            replace_target = None
                            is_valid = True
                        else:
                            self._log(f"  MKV prune fallback validation failed: {mkv_reason}")
                            src_v, src_a, src_s = self._probe_stream_counts(video)
                            out_v, out_a, out_s = self._probe_stream_counts(mkv_output)
                            self._log(
                                f"  Stream counts source(v/a/s)={src_v}/{src_a}/{src_s} -> "
                                f"output(v/a/s)={out_v}/{out_a}/{out_s}"
                            )
                            try:
                                mkv_output.unlink(missing_ok=True)
                            except OSError:
                                pass
                    else:
                        mkv_msgs = (mkv_result.stderr or "").strip()
                        if mkv_msgs:
                            for _line in mkv_msgs.splitlines()[:15]:
                                self._log(f"  ffmpeg: {_line}")
                        try:
                            mkv_output.unlink(missing_ok=True)
                        except OSError:
                            pass

                if not is_valid:
                    summary.failed += 1
                    summary.details.append(
                        {
                            "file": str(video),
                            "status": "failed",
                            "reason": invalid_reason,
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
                    "reason": (
                        f"kept {len(valid_keep_orders)}/{audio_count} audio stream(s) "
                        f"(removed {max(0, audio_count - len(valid_keep_orders))})"
                    ),
                }
            )

        return summary

    # ------------------------------------------------------------------
    # Subtitle sync helpers
    # ------------------------------------------------------------------

    def _compute_subtitle_offset(
        self,
        whisper_segments: List[Dict[str, object]],
        sub_events: List[object],
        max_offset_seconds: float = 300.0,
    ) -> Tuple[float, int]:
        """Vote-accumulate the best global time offset (seconds) to shift sub_events onto whisper_segments.

        A positive offset means subtitles need to be moved *later* in time.
        Returns (offset_seconds, vote_count).
        """
        w_starts = sorted(
            float(seg["start"])
            for seg in whisper_segments
            if str(seg.get("text", "")).strip()
        )
        s_starts = sorted(
            float(ev.start) / 1000.0
            for ev in sub_events
            if str(getattr(ev, "text", "") or "").strip()
        )
        if not w_starts or not s_starts:
            return 0.0, 0

        # Sample at most 300 items each to keep O(n^2) cost bounded
        sample_w = w_starts[:: max(1, len(w_starts) // 300)]
        sample_s = s_starts[:: max(1, len(s_starts) // 300)]

        # Hough-style vote accumulation: each (w, s) pair votes for offset = w - s
        # Quantize to 0.1 s resolution (QUANT = 10 → 1 unit = 0.1 s)
        QUANT = 10
        votes: Counter[int] = Counter()
        for w in sample_w:
            for s in sample_s:
                d = w - s
                if -max_offset_seconds <= d <= max_offset_seconds:
                    votes[round(d * QUANT)] += 1

        if not votes:
            return 0.0, 0

        # Find the 1-second window (±5 bins) with the highest total votes
        best_bin = max(votes, key=lambda b: sum(votes.get(b + delta, 0) for delta in range(-5, 6)))
        window_votes = sum(votes.get(best_bin + delta, 0) for delta in range(-5, 6))

        # Weighted-average the bins inside that window for sub-0.1 s precision
        total_weight = 0.0
        total_count = 0
        for delta in range(-5, 6):
            v = votes.get(best_bin + delta, 0)
            total_weight += (best_bin + delta) * v
            total_count += v
        precise_bin = total_weight / total_count if total_count else best_bin

        return precise_bin / QUANT, window_votes

    def _verify_subtitle_sync(
        self,
        whisper_segments: List[Dict[str, object]],
        sub_events: List[object],
        tolerance_seconds: float = 2.0,
    ) -> Tuple[float, int, int]:
        """Check what fraction of Whisper speech segments have a subtitle covering them.

        Returns (coverage_ratio, matched_count, total_speech_count).
        """
        speech = [
            (float(seg["start"]), float(seg["end"]))
            for seg in whisper_segments
            if str(seg.get("text", "")).strip()
        ]
        if not speech:
            return 1.0, 0, 0

        sub_intervals = sorted(
            (float(ev.start) / 1000.0, float(ev.end) / 1000.0)
            for ev in sub_events
            if str(getattr(ev, "text", "") or "").strip()
        )
        if not sub_intervals:
            return 0.0, 0, len(speech)

        matched = 0
        for ws, we in speech:
            lo = ws - tolerance_seconds
            hi = we + tolerance_seconds
            for ss, se in sub_intervals:
                if ss <= hi and se >= lo:
                    matched += 1
                    break

        return matched / len(speech), matched, len(speech)

    def _resolve_sync_backend(self, requested_backend: str) -> Tuple[Optional[str], str]:
        backend = (requested_backend or "auto").strip().lower()
        if backend not in SYNC_BACKENDS:
            backend = "auto"

        if backend == "auto":
            if whisper is not None:
                return "whisper-offset", "auto-selected whisper-offset"
            if importlib.util.find_spec("aeneas") is not None:
                return "aeneas", "auto-selected aeneas"
            return None, "neither openai-whisper nor aeneas is installed"

        if backend == "whisper-offset" and whisper is None:
            return None, "whisper-offset selected but openai-whisper is not installed"
        if backend == "aeneas" and importlib.util.find_spec("aeneas") is None:
            return None, "aeneas backend selected but aeneas is not installed"
        return backend, f"using {backend}"

    def _sync_subtitles_with_aeneas(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        language: Optional[str],
        overwrite: bool,
        output_suffix: str,
        output_root: Optional[str],
    ) -> OperationSummary:
        summary = OperationSummary(action="sync_subtitles")

        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)
        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)

        if not videos:
            self._log("No video files found to sync subtitles")
            return summary

        task_module = importlib.import_module("aeneas.task")
        exec_module = importlib.import_module("aeneas.executetask")
        Task = getattr(task_module, "Task")
        ExecuteTask = getattr(exec_module, "ExecuteTask")

        lang_code = self._normalize_language_code(language or "eng") or "eng"
        if len(lang_code) != 3:
            lang_code = "eng"

        for video in videos:
            self._log(f"Syncing subtitles with Aeneas: {video.name}")
            sidecars = [
                p for p in self._find_sidecar_subtitles(video)
                if p.suffix.lower() in {".srt", ".ass", ".ssa", ".vtt", ".sub", ".ttml"}
            ]
            if not sidecars:
                summary.skipped += 1
                summary.details.append(
                    {
                        "file": str(video),
                        "status": "skipped",
                        "reason": "no text subtitle sidecar found",
                    }
                )
                continue

            subtitle_path = sidecars[0]
            try:
                subs = pysubs2.load(str(subtitle_path))
            except Exception as exc:
                summary.failed += 1
                summary.details.append({"file": str(video), "status": "failed", "reason": f"Could not load subtitle: {exc}"})
                continue

            non_empty_indices: List[int] = []
            text_lines: List[str] = []
            for idx, ev in enumerate(subs.events):
                cleaned = self._strip_subtitle_tags(str(getattr(ev, "text", "") or "")).strip()
                if cleaned:
                    non_empty_indices.append(idx)
                    text_lines.append(cleaned)

            if not text_lines:
                summary.skipped += 1
                summary.details.append({"file": str(video), "status": "skipped", "reason": "subtitle file has no text events"})
                continue

            with tempfile.TemporaryDirectory(prefix="aeneas_sync_", dir=str(self._get_temp_workspace_root())) as temp_dir:
                text_file = Path(temp_dir) / "subtitle_text.txt"
                map_file = Path(temp_dir) / "sync_map.json"
                text_file.write_text("\n".join(text_lines), encoding="utf-8")

                task_config = (
                    f"task_language={lang_code}|"
                    "is_text_type=plain|"
                    "os_task_file_format=json"
                )

                try:
                    task = Task(config_string=task_config)
                    task.audio_file_path_absolute = str(video.resolve())
                    task.text_file_path_absolute = str(text_file.resolve())
                    task.sync_map_file_path_absolute = str(map_file.resolve())
                    ExecuteTask(task).execute()
                    task.output_sync_map_file()
                except Exception as exc:
                    summary.failed += 1
                    summary.details.append(
                        {
                            "file": str(video),
                            "status": "failed",
                            "reason": f"Aeneas alignment failed: {exc}",
                        }
                    )
                    continue

                try:
                    payload = json.loads(map_file.read_text(encoding="utf-8"))
                    fragments = payload.get("fragments", [])
                except Exception as exc:
                    summary.failed += 1
                    summary.details.append(
                        {
                            "file": str(video),
                            "status": "failed",
                            "reason": f"Could not parse Aeneas output: {exc}",
                        }
                    )
                    continue

                if not isinstance(fragments, list) or not fragments:
                    summary.failed += 1
                    summary.details.append(
                        {
                            "file": str(video),
                            "status": "failed",
                            "reason": "Aeneas returned no alignment fragments",
                        }
                    )
                    continue

                shifted_subs = pysubs2.SSAFile()
                shifted_subs.info.update(subs.info)
                shifted_subs.styles.update(subs.styles)
                shifted_subs.events = [ev.copy() for ev in subs.events]

                mapped = 0
                for idx, fragment in enumerate(fragments):
                    if idx >= len(non_empty_indices):
                        break
                    ev_index = non_empty_indices[idx]
                    begin = self._safe_float(fragment.get("begin"), 0.0)
                    end = self._safe_float(fragment.get("end"), begin + 0.75)
                    if end <= begin:
                        end = begin + 0.75
                    shifted_subs.events[ev_index].start = max(0, int(begin * 1000))
                    shifted_subs.events[ev_index].end = max(
                        shifted_subs.events[ev_index].start + 100,
                        int(end * 1000),
                    )
                    mapped += 1

            if overwrite:
                out_path = subtitle_path
            else:
                output_parent = subtitle_path.parent
                if output_root:
                    output_parent = Path(output_root).expanduser().resolve()
                    output_parent.mkdir(parents=True, exist_ok=True)
                out_path = output_parent / f"{subtitle_path.stem}{output_suffix}{subtitle_path.suffix}"
                idx = 1
                while out_path.exists():
                    out_path = output_parent / f"{subtitle_path.stem}{output_suffix}_{idx}{subtitle_path.suffix}"
                    idx += 1

            try:
                shifted_subs.save(str(out_path))
            except Exception as exc:
                summary.failed += 1
                summary.details.append({"file": str(video), "status": "failed", "reason": f"Could not save synced subtitle: {exc}"})
                continue

            mapped_ratio = mapped / max(1, len(non_empty_indices))
            quality = "good" if mapped_ratio >= 0.85 else ("fair" if mapped_ratio >= 0.5 else "low")
            summary.processed += 1
            summary.details.append(
                {
                    "file": str(video),
                    "status": "synced",
                    "reason": f"aeneas mapped {mapped}/{len(non_empty_indices)} events ({int(mapped_ratio * 100)}%) quality={quality}",
                    "output_path": str(out_path),
                    "quality": quality,
                    "backend": "aeneas",
                }
            )

        return summary

    def sync_subtitles(
        self,
        folders: List[str],
        recursive: bool,
        target_files: List[str],
        model_size: str = "base",
        language: Optional[str] = None,
        overwrite: bool = False,
        output_suffix: str = "_synced",
        max_offset_seconds: float = 300.0,
        verification_tolerance_seconds: float = 2.0,
        output_root: Optional[str] = None,
        sync_backend: str = "auto",
    ) -> OperationSummary:
        """Re-sync existing subtitle files to match audio timing using selectable sync backends.

        This shifts existing subtitle timestamps — it does NOT generate new subtitles.
        After syncing each file it verifies subtitle coverage against Whisper speech segments
        and reports a quality rating.
        """
        summary = OperationSummary(action="sync_subtitles")

        selected_sync_backend, sync_reason = self._resolve_sync_backend(sync_backend)
        if not selected_sync_backend:
            summary.failed = 1
            summary.details.append({"file": "N/A", "status": "failed", "reason": sync_reason})
            self._log(f"ERROR: {sync_reason}")
            return summary

        self._log(f"Subtitle sync backend: {selected_sync_backend} ({sync_reason})")

        if selected_sync_backend == "aeneas":
            return self._sync_subtitles_with_aeneas(
                folders=folders,
                recursive=recursive,
                target_files=target_files,
                language=language,
                overwrite=overwrite,
                output_suffix=output_suffix,
                output_root=output_root,
            )

        if whisper is None:
            self._log("ERROR: openai-whisper not installed. Run: pip install openai-whisper")
            summary.failed = 1
            summary.details.append({"file": "N/A", "status": "failed", "reason": "Whisper library not installed"})
            return summary

        if pysubs2 is None:
            self._log("ERROR: pysubs2 not installed. Run: pip install pysubs2")
            summary.failed = 1
            summary.details.append({"file": "N/A", "status": "failed", "reason": "pysubs2 library not installed"})
            return summary

        videos = [Path(f) for f in target_files if Path(f).exists()]
        for video in self._iter_video_files(folders, recursive):
            videos.append(video)
        videos = list({str(v): v for v in videos}.values())
        summary.scanned = len(videos)

        if not videos:
            self._log("No video files found to sync subtitles")
            return summary

        try:
            self._log(f"Loading Whisper model for subtitle sync: {model_size}...")
            model = whisper.load_model(model_size)
            self._log("Model loaded.")
        except Exception as exc:
            summary.failed = len(videos)
            for video in videos:
                summary.details.append({"file": str(video), "status": "failed", "reason": f"Failed to load model: {exc}"})
            return summary

        for video in videos:
            self._log(f"Syncing subtitles: {video.name}")

            # Find existing sidecar subtitle files (skip image-based .sup)
            sidecars = [
                p for p in self._find_sidecar_subtitles(video)
                if p.suffix.lower() in {".srt", ".ass", ".ssa", ".vtt", ".sub", ".ttml"}
            ]
            if not sidecars:
                summary.skipped += 1
                summary.details.append({
                    "file": str(video),
                    "status": "skipped",
                    "reason": "no text subtitle sidecar found (will not generate new subtitles)",
                })
                continue

            subtitle_path = sidecars[0]
            self._log(f"  Subtitle file: {subtitle_path.name}")

            try:
                subs = pysubs2.load(str(subtitle_path))
            except Exception as exc:
                summary.failed += 1
                summary.details.append({"file": str(video), "status": "failed", "reason": f"Could not load subtitle: {exc}"})
                continue

            if not subs.events:
                summary.skipped += 1
                summary.details.append({"file": str(video), "status": "skipped", "reason": "subtitle file has no events"})
                continue

            # Transcribe with Whisper to get speech segment timestamps
            try:
                self._log(f"  Transcribing audio (model={model_size})…")
                transcribe_opts: Dict[str, object] = {"task": "transcribe"}
                if language:
                    transcribe_opts["language"] = language
                result = model.transcribe(str(video), **transcribe_opts)
            except Exception as exc:
                summary.failed += 1
                summary.details.append({"file": str(video), "status": "failed", "reason": f"Whisper transcription failed: {exc}"})
                continue

            segments: List[Dict[str, object]] = result.get("segments", [])
            if not segments:
                summary.skipped += 1
                summary.details.append({"file": str(video), "status": "skipped", "reason": "Whisper found no speech — cannot compute offset"})
                continue

            offset_seconds, votes = self._compute_subtitle_offset(segments, subs.events, max_offset_seconds)
            offset_ms = int(round(offset_seconds * 1000))

            if votes < 2:
                summary.skipped += 1
                self._log(f"  Offset confidence too low (votes={votes}); skipping")
                summary.details.append({"file": str(video), "status": "skipped", "reason": f"Low offset confidence (votes={votes})"})
                continue

            self._log(f"  Computed offset: {offset_seconds:+.3f}s  (votes={votes})")

            # Apply offset — clamp to non-negative times
            shifted_subs = pysubs2.SSAFile()
            shifted_subs.info.update(subs.info)
            shifted_subs.styles.update(subs.styles)
            for ev in subs.events:
                new_ev = ev.copy()
                new_ev.start = max(0, ev.start + offset_ms)
                new_ev.end = max(new_ev.start + 100, ev.end + offset_ms)
                shifted_subs.events.append(new_ev)

            # Determine output path
            if overwrite:
                out_path = subtitle_path
            else:
                output_parent = subtitle_path.parent
                if output_root:
                    output_parent = Path(output_root).expanduser().resolve()
                    output_parent.mkdir(parents=True, exist_ok=True)
                out_path = output_parent / f"{subtitle_path.stem}{output_suffix}{subtitle_path.suffix}"
                idx = 1
                while out_path.exists():
                    out_path = output_parent / f"{subtitle_path.stem}{output_suffix}_{idx}{subtitle_path.suffix}"
                    idx += 1

            try:
                shifted_subs.save(str(out_path))
                self._log(f"  Saved: {out_path.name}")
            except Exception as exc:
                summary.failed += 1
                summary.details.append({"file": str(video), "status": "failed", "reason": f"Could not save synced subtitle: {exc}"})
                continue

            # --- Verification pass ---
            coverage, matched, total = self._verify_subtitle_sync(
                segments, shifted_subs.events, verification_tolerance_seconds
            )
            coverage_pct = int(round(coverage * 100))
            quality = "good" if coverage >= 0.70 else ("fair" if coverage >= 0.40 else "low")
            self._log(
                f"  Verification: {matched}/{total} speech segments covered "
                f"({coverage_pct}%)  →  quality={quality}"
            )

            summary.processed += 1
            summary.details.append({
                "file": str(video),
                "status": "synced",
                "reason": (
                    f"offset={offset_seconds:+.3f}s  "
                    f"coverage={coverage_pct}% ({matched}/{total} segs)  "
                    f"quality={quality}"
                ),
                "output_path": str(out_path),
                "offset_seconds": offset_seconds,
                "coverage_pct": coverage_pct,
                "quality": quality,
            })

        return summary


class JobPayload(BaseModel):
    folders: List[str] = Field(default_factory=list)
    target_files: List[str] = Field(default_factory=list)
    manual_sidecars: Dict[str, List[str]] = Field(default_factory=dict)
    recursive: bool = True
    overwrite: bool = False
    output_root: str = ""
    output_suffix: str = ""
    extract_for_restore: bool = True
    export_txt: bool = True
    scan_only_embedded: bool = False
    ai_backend: str = "auto"
    model_size: str = "base"
    language_strategy: str = "snippets"
    snippet_count: int = 3
    sample_seconds: float = 25.0
    overwrite_existing_tags: bool = False
    detect_only_audio_tagging: bool = False
    keep_audio_orders_by_file: Dict[str, List[int]] = Field(default_factory=dict)
    prune_audio_suffix: str = "_audiopruned"
    sync_backend: str = "auto"
    sync_language: str = ""
    sync_max_offset_seconds: float = 300.0
    sync_verification_tolerance: float = 2.0


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
                    output_root=payload.output_root or None,
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
                    output_root=payload.output_root or None,
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
            elif action == "tag_audio_language":
                summary = self.processor.detect_and_tag_audio_languages(
                    folders=payload.folders,
                    recursive=payload.recursive,
                    target_files=payload.target_files,
                    model_size=payload.model_size or "base",
                    strategy=payload.language_strategy or "snippets",
                    snippet_count=max(1, int(payload.snippet_count or 3)),
                    sample_seconds=max(5.0, float(payload.sample_seconds or 25.0)),
                    overwrite=payload.overwrite,
                    output_suffix=payload.output_suffix or "_langtagged",
                    overwrite_existing_tags=bool(payload.overwrite_existing_tags),
                    detect_only=bool(payload.detect_only_audio_tagging),
                    output_root=payload.output_root or None,
                )
                result = summary.to_dict()
            elif action == "prune_audio_streams":
                summary = self.processor.prune_audio_streams(
                    folders=payload.folders,
                    recursive=payload.recursive,
                    target_files=payload.target_files,
                    keep_audio_orders_by_file=dict(payload.keep_audio_orders_by_file or {}),
                    overwrite=payload.overwrite,
                    output_suffix=payload.prune_audio_suffix or "_audiopruned",
                    output_root=payload.output_root or None,
                )
                result = summary.to_dict()
            elif action == "sync_subtitles":
                summary = self.processor.sync_subtitles(
                    folders=payload.folders,
                    recursive=payload.recursive,
                    target_files=payload.target_files,
                    model_size=payload.model_size or "base",
                    language=payload.sync_language or None,
                    overwrite=payload.overwrite,
                    output_suffix=payload.output_suffix or "_synced",
                    max_offset_seconds=max(10.0, float(payload.sync_max_offset_seconds or 300.0)),
                    verification_tolerance_seconds=max(0.5, float(payload.sync_verification_tolerance or 2.0)),
                    output_root=payload.output_root or None,
                    sync_backend=payload.sync_backend or "auto",
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

    @app.post("/jobs/tag-audio-language")
    def start_tag_audio_language(payload: JobPayload) -> Dict[str, object]:
        try:
            job = manager.submit("tag_audio_language", payload)
            return {"job_id": job.id, "status": job.status}
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/jobs/sync-subtitles")
    def start_sync_subtitles(payload: JobPayload) -> Dict[str, object]:
        try:
            job = manager.submit("sync_subtitles", payload)
            return {"job_id": job.id, "status": job.status}
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

    @app.post("/jobs/prune-audio-streams")
    def start_prune_audio_streams(payload: JobPayload) -> Dict[str, object]:
        try:
            job = manager.submit("prune_audio_streams", payload)
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
                    "description": "Additional video tools for conversion, organization, and repair. Use the Tooling & Diagnostics tab to configure FFmpeg/FFprobe, MKVToolNix, HandBrakeCLI, and MakeMKV paths.",
                    "widget": None,
                },
                {
                    "title": "Convert to MKV/MP4",
                    "description": "Convert videos between MKV and MP4 using selectable backends: FFmpeg, MKVToolNix, or HandBrakeCLI. Auto mode picks the best available backend.",
                    "widget": self.main_window.convert_mkv_button,
                },
                {
                    "title": "Organize Media",
                    "description": "Automatically organize media files: move movies up one level and rename TV show episodes to S##E## format. Configure options with checkboxes.",
                    "widget": self.main_window.organize_button,
                },
                {
                    "title": "Repair Metadata",
                    "description": "Rebuild corrupted containers with backend selection (FFmpeg, MKVToolNix, or MakeMKV for MKV files). Useful for fixing truncated/corrupt media issues.",
                    "widget": self.main_window.repair_button,
                },
                {
                    "title": "Generate Subtitles with AI Backends",
                    "description": "Generate subtitles using selectable backends: OpenAI Whisper, faster-whisper, WhisperX, stable-ts, whisper-timestamped, SpeechBrain, Vosk, or Text-to-Timestamps. Auto mode picks the best available backend.",
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
                processor = SubtitleProcessor(
                    ffmpeg_bin=str(self.options.get("ffmpeg_bin", "")).strip() or None,
                    ffprobe_bin=str(self.options.get("ffprobe_bin", "")).strip() or None,
                    mkvmerge_bin=str(self.options.get("mkvmerge_bin", "")).strip() or None,
                    handbrake_bin=str(self.options.get("handbrake_bin", "")).strip() or None,
                    makemkvcon_bin=str(self.options.get("makemkvcon_bin", "")).strip() or None,
                    log_callback=self.log_message.emit,
                    use_hw_accel=bool(self.options.get("use_hw_accel", False)),
                    command_feedback=str(self.options.get("command_feedback", "normal")),
                    ffmpeg_loglevel=str(self.options.get("ffmpeg_loglevel", "warning")),
                    ffprobe_loglevel=str(self.options.get("ffprobe_loglevel", "error")),
                )
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
                        output_root=str(self.options.get("output_root", "")).strip() or None,
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
                        output_root=str(self.options.get("output_root", "")).strip() or None,
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
                        output_root=str(self.options.get("output_root", "")).strip() or None,
                        backend=str(self.options.get("convert_backend", "auto")),
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
                        backend=str(self.options.get("repair_backend", "auto")),
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
                        backend=str(self.options.get("ai_backend", "auto")),
                    )
                    payload = summary.to_dict()
                elif self.action == "tag_audio_language":
                    summary = processor.detect_and_tag_audio_languages(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        model_size=str(self.options.get("model_size", "base")),
                        strategy=str(self.options.get("language_strategy", "snippets")),
                        snippet_count=max(1, int(self.options.get("snippet_count", 3))),
                        sample_seconds=max(5.0, float(self.options.get("sample_seconds", 25.0))),
                        overwrite=overwrite,
                        output_suffix=str(self.options.get("output_suffix", "_langtagged")),
                        overwrite_existing_tags=bool(self.options.get("overwrite_existing_tags", False)),
                        detect_only=bool(self.options.get("detect_only_audio_tagging", False)),
                        output_root=str(self.options.get("output_root", "")).strip() or None,
                    )
                    payload = summary.to_dict()
                elif self.action == "prune_audio_streams":
                    summary = processor.prune_audio_streams(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        keep_audio_orders_by_file=dict(self.options.get("keep_audio_orders_by_file", {})),
                        overwrite=overwrite,
                        output_suffix=str(self.options.get("prune_audio_suffix", "_audiopruned")),
                        output_root=str(self.options.get("output_root", "")).strip() or None,
                    )
                    payload = summary.to_dict()
                elif self.action == "sync_subtitles":
                    summary = processor.sync_subtitles(
                        folders=folders,
                        recursive=recursive,
                        target_files=target_files,
                        model_size=str(self.options.get("model_size", "base")),
                        language=self.options.get("sync_language") or None,
                        overwrite=overwrite,
                        output_suffix=str(self.options.get("output_suffix", "_synced")),
                        max_offset_seconds=max(10.0, float(self.options.get("sync_max_offset_seconds", 300.0))),
                        verification_tolerance_seconds=max(0.5, float(self.options.get("sync_verification_tolerance", 2.0))),
                        output_root=str(self.options.get("output_root", "")).strip() or None,
                        sync_backend=str(self.options.get("sync_backend", "auto")),
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
            self.hw_accel_checkbox = QCheckBox("Use hardware acceleration (GPU)")
            self.hw_accel_checkbox.setToolTip(
                "Pass -hwaccel auto to ffmpeg. Enables GPU-assisted video demuxing and decoding\n"
                "where supported (NVIDIA NVDEC, Intel QuickSync, AMD VCE, etc.)."
            )
            self.save_next_to_source_checkbox = QCheckBox("Save outputs next to source files")
            self.save_next_to_source_checkbox.setChecked(True)
            self.custom_output_dir_input = QLineEdit()
            self.custom_output_dir_input.setPlaceholderText("Custom output folder (optional)")
            self.custom_output_dir_input.setEnabled(False)
            self.custom_output_dir_browse_button = QPushButton("Browse...")
            self.custom_output_dir_browse_button.setEnabled(False)
            self.custom_output_dir_browse_button.clicked.connect(self._choose_output_directory)

            def _toggle_output_controls(checked: bool) -> None:
                self.custom_output_dir_input.setEnabled(not checked)
                self.custom_output_dir_browse_button.setEnabled(not checked)

            self.save_next_to_source_checkbox.toggled.connect(_toggle_output_controls)

            self.remove_suffix_input = QLineEdit("_nosubs")
            self.include_suffix_input = QLineEdit("_withsubs")
            self.extract_suffix_input = QLineEdit(".embedded_sub")

            options_layout.addWidget(self.recursive_checkbox, 0, 0, 1, 2)
            options_layout.addWidget(self.overwrite_checkbox, 1, 0, 1, 2)
            options_layout.addWidget(self.extract_checkbox, 2, 0, 1, 2)
            options_layout.addWidget(self.export_txt_checkbox, 3, 0, 1, 2)
            options_layout.addWidget(self.scan_only_embedded_checkbox, 4, 0, 1, 2)
            options_layout.addWidget(self.only_selected_targets_checkbox, 5, 0, 1, 2)
            options_layout.addWidget(self.hw_accel_checkbox, 6, 0, 1, 2)
            options_layout.addWidget(self.save_next_to_source_checkbox, 7, 0, 1, 2)
            options_layout.addWidget(QLabel("Custom output folder:"), 8, 0)
            custom_out_row = QHBoxLayout()
            custom_out_row.setContentsMargins(0, 0, 0, 0)
            custom_out_row.addWidget(self.custom_output_dir_input)
            custom_out_row.addWidget(self.custom_output_dir_browse_button)
            options_layout.addLayout(custom_out_row, 8, 1)
            options_layout.addWidget(QLabel("Remove output suffix:"), 9, 0)
            options_layout.addWidget(self.remove_suffix_input, 9, 1)
            options_layout.addWidget(QLabel("Include output suffix:"), 10, 0)
            options_layout.addWidget(self.include_suffix_input, 10, 1)
            options_layout.addWidget(QLabel("Extract output suffix:"), 11, 0)
            options_layout.addWidget(self.extract_suffix_input, 11, 1)
            options_layout.addWidget(QLabel("Conversion output suffix:"), 12, 0)
            self.convert_suffix_input = QLineEdit("_converted")
            options_layout.addWidget(self.convert_suffix_input, 12, 1)

            # Swiss Army Knife section
            tools_box = QGroupBox("Video Tools (Swiss Army Knife)")
            tools_root_layout = QVBoxLayout(tools_box)
            self.tools_tabs = QTabWidget()
            tools_root_layout.addWidget(self.tools_tabs)

            operations_widget = QWidget()
            tools_layout = QVBoxLayout(operations_widget)
            self.tools_tabs.addTab(operations_widget, "Operations")

            diagnostics_widget = QWidget()
            diagnostics_layout = QGridLayout(diagnostics_widget)
            self.tools_tabs.addTab(diagnostics_widget, "Tooling && Diagnostics")
            
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

            backend_row = QHBoxLayout()
            backend_row.addWidget(QLabel("Convert backend:"))
            self.convert_backend_combo = QComboBox()
            self.convert_backend_combo.addItem("Auto (recommended)", "auto")
            self.convert_backend_combo.addItem("FFmpeg", "ffmpeg")
            self.convert_backend_combo.addItem("MKVToolNix (mkvmerge)", "mkvtoolnix")
            self.convert_backend_combo.addItem("HandBrakeCLI", "handbrake")
            backend_row.addWidget(self.convert_backend_combo)

            backend_row.addWidget(QLabel("Repair backend:"))
            self.repair_backend_combo = QComboBox()
            self.repair_backend_combo.addItem("Auto (recommended)", "auto")
            self.repair_backend_combo.addItem("FFmpeg", "ffmpeg")
            self.repair_backend_combo.addItem("MKVToolNix (mkvmerge)", "mkvtoolnix")
            self.repair_backend_combo.addItem("MakeMKV (makemkvcon)", "makemkv")
            backend_row.addWidget(self.repair_backend_combo)
            backend_row.addStretch()
            tools_layout.addLayout(backend_row)
            
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
            
            # AI tools section (only show if use_ai is enabled)
            if self.use_ai:
                tools_layout.addSpacing(10)
                subtitle_gen_label = QLabel("<b>AI Tools</b>")
                tools_layout.addWidget(subtitle_gen_label)
                
                whisper_options = QHBoxLayout()
                whisper_options.addWidget(QLabel("Backend:"))
                self.ai_backend_combo = QComboBox()
                self.ai_backend_combo.addItem("Auto (best available)", "auto")
                self.ai_backend_combo.addItem("OpenAI Whisper", "openai-whisper")
                self.ai_backend_combo.addItem("faster-whisper", "faster-whisper")
                self.ai_backend_combo.addItem("WhisperX", "whisperx")
                self.ai_backend_combo.addItem("stable-ts", "stable-ts")
                self.ai_backend_combo.addItem("whisper-timestamped", "whisper-timestamped")
                self.ai_backend_combo.addItem("SpeechBrain", "speechbrain")
                self.ai_backend_combo.addItem("Vosk", "vosk")
                self.ai_backend_combo.addItem("Text to Timestamps (heuristic)", "text-to-timestamps")
                whisper_options.addWidget(self.ai_backend_combo)

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
                self.generate_button.setToolTip("Generate subtitles from video audio using the selected AI model")
                whisper_options.addWidget(self.generate_button)
                whisper_options.addStretch()
                
                tools_layout.addLayout(whisper_options)
                self.generate_button.clicked.connect(self._start_generate)

                language_tag_options = QHBoxLayout()
                language_tag_options.addWidget(QLabel("Lang Analysis:"))
                self.audio_lang_strategy_combo = QComboBox()
                self.audio_lang_strategy_combo.addItem("Random snippets (faster)", "snippets")
                self.audio_lang_strategy_combo.addItem("Whole stream (slower, deeper)", "full")
                language_tag_options.addWidget(self.audio_lang_strategy_combo)

                language_tag_options.addWidget(QLabel("Snippets:"))
                self.audio_lang_snippets_input = QLineEdit("3")
                self.audio_lang_snippets_input.setMaximumWidth(60)
                self.audio_lang_snippets_input.setToolTip("Used for random-snippet mode")
                language_tag_options.addWidget(self.audio_lang_snippets_input)

                language_tag_options.addWidget(QLabel("Seconds/sample:"))
                self.audio_lang_seconds_input = QLineEdit("25")
                self.audio_lang_seconds_input.setMaximumWidth(60)
                language_tag_options.addWidget(self.audio_lang_seconds_input)

                self.audio_lang_overwrite_checkbox = QCheckBox("Overwrite existing language tags")
                language_tag_options.addWidget(self.audio_lang_overwrite_checkbox)

                self.audio_lang_detect_only_checkbox = QCheckBox("Detect only (no file changes)")
                self.audio_lang_detect_only_checkbox.setToolTip(
                    "Only analyze and report detected languages; do not write metadata tags or remux files."
                )
                language_tag_options.addWidget(self.audio_lang_detect_only_checkbox)

                self.tag_audio_language_button = QPushButton("Detect + Tag Audio Language")
                self.tag_audio_language_button.setToolTip("Use AI language detection per audio stream and write metadata tags")
                language_tag_options.addWidget(self.tag_audio_language_button)
                language_tag_options.addStretch()
                tools_layout.addLayout(language_tag_options)
                self.tag_audio_language_button.clicked.connect(self._start_tag_audio_language)

                # Subtitle Sync row
                tools_layout.addSpacing(4)
                sync_options_row = QHBoxLayout()
                sync_options_row.addWidget(QLabel("Subtitle Sync:"))

                sync_options_row.addWidget(QLabel("Backend:"))
                self.sync_backend_combo = QComboBox()
                self.sync_backend_combo.addItem("Auto", "auto")
                self.sync_backend_combo.addItem("Whisper Offset", "whisper-offset")
                self.sync_backend_combo.addItem("Aeneas", "aeneas")
                sync_options_row.addWidget(self.sync_backend_combo)

                sync_options_row.addWidget(QLabel("Language hint:"))
                self.sync_language_input = QLineEdit()
                self.sync_language_input.setPlaceholderText("auto")
                self.sync_language_input.setMaximumWidth(70)
                self.sync_language_input.setToolTip("Optional ISO language code (en, es, fr…) for AI transcription")
                sync_options_row.addWidget(self.sync_language_input)

                self.sync_overwrite_checkbox = QCheckBox("Overwrite original subtitle")
                self.sync_overwrite_checkbox.setToolTip("Replace the original subtitle file in-place instead of saving <name>_synced")
                sync_options_row.addWidget(self.sync_overwrite_checkbox)

                self.sync_button = QPushButton("Sync Subtitles to Audio")
                self.sync_button.setToolTip(
                    "Shift existing sidecar subtitle timings to match audio.\n"
                    "Choose Whisper Offset or Aeneas backend in the sync row.\n"
                    "No new subtitles are generated; existing subtitle events are re-timed."
                )
                sync_options_row.addWidget(self.sync_button)
                sync_options_row.addStretch()
                tools_layout.addLayout(sync_options_row)
                self.sync_button.clicked.connect(self._start_sync_subtitles)
            else:
                # Create dummy attributes for widgets that won't exist
                self.ai_backend_combo = None
                self.whisper_model_combo = None
                self.whisper_language_input = None
                self.generate_button = None
                self.audio_lang_strategy_combo = None
                self.audio_lang_snippets_input = None
                self.audio_lang_seconds_input = None
                self.audio_lang_overwrite_checkbox = None
                self.audio_lang_detect_only_checkbox = None
                self.tag_audio_language_button = None
                self.sync_backend_combo = None
                self.sync_language_input = None
                self.sync_overwrite_checkbox = None
                self.sync_button = None

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

            # Tooling & diagnostics tab
            self.ffmpeg_bin_input = QLineEdit()
            self.ffmpeg_bin_input.setPlaceholderText("ffmpeg")
            self.ffmpeg_bin_input.setText("ffmpeg")
            self.ffprobe_bin_input = QLineEdit()
            self.ffprobe_bin_input.setPlaceholderText("ffprobe")
            self.ffprobe_bin_input.setText("ffprobe")
            self.mkvmerge_bin_input = QLineEdit()
            self.mkvmerge_bin_input.setPlaceholderText("mkvmerge")
            self.mkvmerge_bin_input.setText("mkvmerge")
            self.handbrake_bin_input = QLineEdit()
            self.handbrake_bin_input.setPlaceholderText("HandBrakeCLI")
            self.handbrake_bin_input.setText("HandBrakeCLI")
            self.makemkvcon_bin_input = QLineEdit()
            self.makemkvcon_bin_input.setPlaceholderText("makemkvcon")
            self.makemkvcon_bin_input.setText("makemkvcon")

            def _tool_row(row: int, label: str, line_edit: QLineEdit) -> None:
                diagnostics_layout.addWidget(QLabel(label), row, 0)
                diagnostics_layout.addWidget(line_edit, row, 1)
                browse_button = QPushButton("Browse...")
                browse_button.clicked.connect(lambda _=False, target=line_edit: self._choose_executable_path(target))
                diagnostics_layout.addWidget(browse_button, row, 2)

            _tool_row(0, "FFmpeg binary:", self.ffmpeg_bin_input)
            _tool_row(1, "FFprobe binary:", self.ffprobe_bin_input)
            _tool_row(2, "MKVToolNix (mkvmerge):", self.mkvmerge_bin_input)
            _tool_row(3, "HandBrakeCLI:", self.handbrake_bin_input)
            _tool_row(4, "MakeMKV (makemkvcon):", self.makemkvcon_bin_input)

            diagnostics_layout.addWidget(QLabel("Command feedback:"), 5, 0)
            self.command_feedback_combo = QComboBox()
            self.command_feedback_combo.addItem("Quiet", "quiet")
            self.command_feedback_combo.addItem("Normal", "normal")
            self.command_feedback_combo.addItem("Verbose", "verbose")
            self.command_feedback_combo.setCurrentIndex(1)
            diagnostics_layout.addWidget(self.command_feedback_combo, 5, 1, 1, 2)

            diagnostics_layout.addWidget(QLabel("FFmpeg loglevel:"), 6, 0)
            self.ffmpeg_loglevel_combo = QComboBox()
            for level in ["quiet", "error", "warning", "info", "verbose"]:
                self.ffmpeg_loglevel_combo.addItem(level, level)
            ffmpeg_idx = self.ffmpeg_loglevel_combo.findData("warning")
            if ffmpeg_idx >= 0:
                self.ffmpeg_loglevel_combo.setCurrentIndex(ffmpeg_idx)
            diagnostics_layout.addWidget(self.ffmpeg_loglevel_combo, 6, 1, 1, 2)

            diagnostics_layout.addWidget(QLabel("FFprobe loglevel:"), 7, 0)
            self.ffprobe_loglevel_combo = QComboBox()
            for level in ["quiet", "error", "warning", "info", "verbose"]:
                self.ffprobe_loglevel_combo.addItem(level, level)
            ffprobe_idx = self.ffprobe_loglevel_combo.findData("error")
            if ffprobe_idx >= 0:
                self.ffprobe_loglevel_combo.setCurrentIndex(ffprobe_idx)
            diagnostics_layout.addWidget(self.ffprobe_loglevel_combo, 7, 1, 1, 2)

            self.tool_status_refresh_button = QPushButton("Refresh Tool Status")
            self.tool_status_refresh_button.clicked.connect(
                lambda: self._refresh_dependency_status(log_missing=False, announce=True)
            )
            diagnostics_layout.addWidget(self.tool_status_refresh_button, 8, 0, 1, 3)

            self.tool_status_box = QTextEdit()
            self.tool_status_box.setReadOnly(True)
            self.tool_status_box.setMinimumHeight(140)
            self.tool_status_box.setMaximumHeight(220)
            diagnostics_layout.addWidget(self.tool_status_box, 9, 0, 1, 3)

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

            self._refresh_dependency_status(log_missing=True, announce=False)

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
            output_root = ""
            if not self.save_next_to_source_checkbox.isChecked():
                output_root = self.custom_output_dir_input.text().strip()
                if not output_root:
                    raise ValueError("Choose a custom output folder or enable 'Save outputs next to source files'.")
                out_path = Path(output_root).expanduser().resolve()
                out_path.mkdir(parents=True, exist_ok=True)
                output_root = str(out_path)
            return {
                "folders": folders,
                "target_files": target_files,
                "recursive": self.recursive_checkbox.isChecked(),
                "overwrite": self.overwrite_checkbox.isChecked(),
                "extract_for_restore": self.extract_checkbox.isChecked(),
                "export_txt": self.export_txt_checkbox.isChecked(),
                "scan_only_embedded": self.scan_only_embedded_checkbox.isChecked(),
                "use_hw_accel": self.hw_accel_checkbox.isChecked(),
                "output_root": output_root,
                "ffmpeg_bin": self.ffmpeg_bin_input.text().strip() or "ffmpeg",
                "ffprobe_bin": self.ffprobe_bin_input.text().strip() or "ffprobe",
                "mkvmerge_bin": self.mkvmerge_bin_input.text().strip() or "mkvmerge",
                "handbrake_bin": self.handbrake_bin_input.text().strip() or "HandBrakeCLI",
                "makemkvcon_bin": self.makemkvcon_bin_input.text().strip() or "makemkvcon",
                "command_feedback": str(self.command_feedback_combo.currentData() or "normal"),
                "ffmpeg_loglevel": str(self.ffmpeg_loglevel_combo.currentData() or "warning"),
                "ffprobe_loglevel": str(self.ffprobe_loglevel_combo.currentData() or "error"),
                "convert_backend": str(self.convert_backend_combo.currentData() or "auto"),
                "repair_backend": str(self.repair_backend_combo.currentData() or "auto"),
            }

        def _choose_output_directory(self) -> None:
            folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
            if folder:
                self.custom_output_dir_input.setText(folder)

        def _choose_executable_path(self, target_input: QLineEdit) -> None:
            file_path, _ = QFileDialog.getOpenFileName(self, "Select Executable")
            if file_path:
                target_input.setText(file_path)

        def _refresh_dependency_status(self, log_missing: bool = False, announce: bool = True) -> Dict[str, object]:
            ffmpeg_bin = self.ffmpeg_bin_input.text().strip() or "ffmpeg"
            ffprobe_bin = self.ffprobe_bin_input.text().strip() or "ffprobe"
            mkvmerge_bin = self.mkvmerge_bin_input.text().strip() or "mkvmerge"
            handbrake_bin = self.handbrake_bin_input.text().strip() or "HandBrakeCLI"
            makemkv_bin = self.makemkvcon_bin_input.text().strip() or "makemkvcon"

            probe = SubtitleProcessor(
                ffmpeg_bin=ffmpeg_bin,
                ffprobe_bin=ffprobe_bin,
                mkvmerge_bin=mkvmerge_bin,
                handbrake_bin=handbrake_bin,
                makemkvcon_bin=makemkv_bin,
            )
            dep_check = probe.check_dependencies()

            lines = [
                f"FFmpeg: {'FOUND' if dep_check.get('ffmpeg_found') else 'MISSING'}  ({dep_check.get('ffmpeg_path') or ffmpeg_bin})",
                f"FFprobe: {'FOUND' if dep_check.get('ffprobe_found') else 'MISSING'}  ({dep_check.get('ffprobe_path') or ffprobe_bin})",
                f"MKVToolNix (mkvmerge): {'FOUND' if dep_check.get('mkvmerge_found') else 'MISSING'}  ({dep_check.get('mkvmerge_path') or mkvmerge_bin})",
                f"HandBrakeCLI: {'FOUND' if dep_check.get('handbrake_found') else 'MISSING'}  ({dep_check.get('handbrake_path') or handbrake_bin})",
                f"MakeMKV (makemkvcon): {'FOUND' if dep_check.get('makemkv_found') else 'MISSING'}  ({dep_check.get('makemkv_path') or makemkv_bin})",
            ]
            self.tool_status_box.setPlainText("\n".join(lines))

            if log_missing:
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
            elif announce:
                self._log("Tool status refreshed.")
            return dep_check

        def _set_running(self, running: bool) -> None:
            self.scan_button.setEnabled(not running)
            self.remove_button.setEnabled(not running)
            self.include_button.setEnabled(not running)
            self.extract_button.setEnabled(not running)
            self.convert_mkv_button.setEnabled(not running)
            self.convert_mp4_button.setEnabled(not running)
            self.organize_button.setEnabled(not running)
            self.repair_button.setEnabled(not running)
            self.tool_status_refresh_button.setEnabled(not running)
            if self.generate_button is not None:
                self.generate_button.setEnabled(not running)
            if self.tag_audio_language_button is not None:
                self.tag_audio_language_button.setEnabled(not running)
            if self.sync_button is not None:
                self.sync_button.setEnabled(not running)
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
                selected_backend = str(self.repair_backend_combo.currentData() or "auto")
                
                reply = QMessageBox.question(
                    self,
                    "Confirm Repair",
                    "This will rebuild video containers using backend '{}'. {} Continue?".format(
                        selected_backend,
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
            """Start subtitle generation using the selected AI backend."""
            try:
                options = self._collect_common_options()
                options["ai_backend"] = (
                    str(self.ai_backend_combo.currentData() or "auto") if self.ai_backend_combo else "auto"
                )
                options["model_size"] = self.whisper_model_combo.currentText()
                options["language"] = self.whisper_language_input.text().strip() or None
                options["output_format"] = "srt"  # Currently only SRT supported

                if pysubs2 is None:
                    QMessageBox.warning(
                        self,
                        "pysubs2/pysub2 Not Installed",
                        "pysubs2 library (sometimes searched as pysub2) is not installed.\n\n"
                        "pip install pysubs2"
                    )
                    return

                backend_probe = SubtitleProcessor()
                selected_backend, backend_reason = backend_probe._resolve_transcription_backend(
                    str(options.get("ai_backend", "auto"))
                )
                if not selected_backend:
                    QMessageBox.warning(
                        self,
                        "AI Backend Not Available",
                        backend_reason,
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
                    f"Generate subtitles using backend '{selected_backend}' and model '{model_size}'?\n\n"
                    f"Backend selection: {backend_reason}\n"
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

        def _start_tag_audio_language(self) -> None:
            try:
                options = self._collect_common_options()
                options["model_size"] = self.whisper_model_combo.currentText() if self.whisper_model_combo else "base"
                strategy = "snippets"
                if self.audio_lang_strategy_combo is not None:
                    strategy = str(self.audio_lang_strategy_combo.currentData() or "snippets")
                options["language_strategy"] = strategy

                snippet_count = 3
                if self.audio_lang_snippets_input is not None:
                    snippet_count = int((self.audio_lang_snippets_input.text() or "3").strip())
                sample_seconds = 25.0
                if self.audio_lang_seconds_input is not None:
                    sample_seconds = float((self.audio_lang_seconds_input.text() or "25").strip())
                options["snippet_count"] = max(1, snippet_count)
                options["sample_seconds"] = max(5.0, sample_seconds)
                options["overwrite_existing_tags"] = bool(
                    self.audio_lang_overwrite_checkbox.isChecked()
                ) if self.audio_lang_overwrite_checkbox is not None else False
                options["detect_only_audio_tagging"] = bool(
                    self.audio_lang_detect_only_checkbox.isChecked()
                ) if self.audio_lang_detect_only_checkbox is not None else False
                options["output_suffix"] = "_langtagged"

                if whisper is None:
                    QMessageBox.warning(
                        self,
                        "Whisper Not Installed",
                        "Whisper AI is not installed. Please install it with:\n\n"
                        "pip install openai-whisper",
                    )
                    return

                reply = QMessageBox.question(
                    self,
                    "Tag Audio Language",
                    "Detect language from audio streams?\n\n"
                    "If 'Detect only' is enabled, no media files will be modified.\n"
                    "Otherwise, metadata tags will be written using stream copy (no re-encode).\n\n"
                    "Tip: Random snippets is faster. Whole stream can improve confidence on difficult content.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if reply == QMessageBox.StandardButton.Yes:
                    self._start_worker("tag_audio_language", options)
            except ValueError as exc:
                self._log_error(
                    "ERR013_VALIDATION_FAILED",
                    "Failed to validate audio language tagging options",
                    str(exc),
                )
                QMessageBox.warning(self, "Validation", str(exc))

        def _start_sync_subtitles(self) -> None:
            try:
                options = self._collect_common_options()
                options["model_size"] = self.whisper_model_combo.currentText() if self.whisper_model_combo else "base"
                options["sync_backend"] = (
                    str(self.sync_backend_combo.currentData() or "auto") if self.sync_backend_combo else "auto"
                )
                options["sync_language"] = self.sync_language_input.text().strip() if self.sync_language_input else ""
                options["output_suffix"] = "_synced"
                options["sync_max_offset_seconds"] = 300.0
                options["sync_verification_tolerance"] = 2.0
                if self.sync_overwrite_checkbox is not None:
                    options["overwrite"] = self.sync_overwrite_checkbox.isChecked()

                if pysubs2 is None:
                    QMessageBox.warning(
                        self,
                        "pysubs2/pysub2 Not Installed",
                        "pysubs2 (aka pysub2) is not installed. Please install it with:\n\n"
                        "pip install pysubs2",
                    )
                    return

                sync_probe = SubtitleProcessor()
                selected_sync_backend, sync_reason = sync_probe._resolve_sync_backend(
                    str(options.get("sync_backend", "auto"))
                )
                if not selected_sync_backend:
                    QMessageBox.warning(
                        self,
                        "Sync Backend Not Available",
                        sync_reason,
                    )
                    return

                reply = QMessageBox.question(
                    self,
                    "Sync Subtitles to Audio",
                    f"Sync backend: {selected_sync_backend} ({sync_reason})\n\n"
                    "This will align existing sidecar subtitles to the video audio timeline.\n"
                    "No new subtitles will be generated.\n\n"
                    "Continue?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if reply == QMessageBox.StandardButton.Yes:
                    self._start_worker("sync_subtitles", options)
            except ValueError as exc:
                self._log_error(
                    "ERR014_VALIDATION_FAILED",
                    "Failed to validate subtitle sync options",
                    str(exc),
                )
                QMessageBox.warning(self, "Validation", str(exc))

        def _prompt_audio_languages_to_keep(
            self,
            details: List[Dict[str, object]],
        ) -> Optional[Tuple[Dict[str, List[int]], List[str]]]:
            language_counts: Counter[str] = Counter()
            language_bytes: Dict[str, int] = {}
            for item in details:
                detected = item.get("detected_streams", [])
                if not isinstance(detected, list):
                    continue
                for stream in detected:
                    if not isinstance(stream, dict):
                        continue
                    lang = str(stream.get("language") or "").strip().lower()
                    if lang:
                        language_counts[lang] += 1
                        eb = stream.get("estimated_bytes")
                        if isinstance(eb, (int, float)) and eb > 0:
                            language_bytes[lang] = language_bytes.get(lang, 0) + int(eb)

            if not language_counts:
                return None

            def _fmt_bytes(total: int) -> str:
                mb = total / (1024 * 1024)
                if mb >= 1024:
                    return f"~{mb / 1024:.1f} GB"
                return f"~{mb:.0f} MB"

            dialog = QDialog(self)
            dialog.setWindowTitle("Choose Audio Languages to Keep")
            dialog.resize(520, 420)

            layout = QVBoxLayout(dialog)
            layout.addWidget(
                QLabel(
                    "Select the detected audio languages to keep.\n"
                    "All other detected audio streams will be removed."
                )
            )

            list_widget = QListWidget()
            for lang, count in sorted(language_counts.items(), key=lambda x: (-x[1], x[0])):
                size_str = f",  {_fmt_bytes(language_bytes[lang])}" if lang in language_bytes else ""
                label = f"{lang}  ({count} stream{'s' if count != 1 else ''}{size_str})"
                item = QListWidgetItem(label)
                item.setData(Qt.ItemDataRole.UserRole, lang)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Checked)
                list_widget.addItem(item)
            layout.addWidget(list_widget)

            helper_row = QHBoxLayout()
            check_all_button = QPushButton("Check All")
            uncheck_all_button = QPushButton("Uncheck All")
            helper_row.addWidget(check_all_button)
            helper_row.addWidget(uncheck_all_button)
            helper_row.addStretch()
            layout.addLayout(helper_row)

            def _set_all(state: Qt.CheckState) -> None:
                for i in range(list_widget.count()):
                    list_widget.item(i).setCheckState(state)

            check_all_button.clicked.connect(lambda: _set_all(Qt.CheckState.Checked))
            uncheck_all_button.clicked.connect(lambda: _set_all(Qt.CheckState.Unchecked))

            buttons = QDialogButtonBox(
                QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
            )
            ok_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
            if ok_button is not None:
                ok_button.setText("Trim Audio Streams")
            cancel_button = buttons.button(QDialogButtonBox.StandardButton.Cancel)
            if cancel_button is not None:
                cancel_button.setText("Skip")
            layout.addWidget(buttons)

            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)

            if dialog.exec() != int(QDialog.DialogCode.Accepted):
                return None

            selected_languages: set[str] = set()
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                if item.checkState() == Qt.CheckState.Checked:
                    lang = str(item.data(Qt.ItemDataRole.UserRole) or "").strip().lower()
                    if lang:
                        selected_languages.add(lang)

            if not selected_languages:
                QMessageBox.warning(self, "No Languages Selected", "Please select at least one language to keep.")
                return None

            keep_map: Dict[str, List[int]] = {}
            for item in details:
                file_path = str(item.get("file") or "").strip()
                if not file_path:
                    continue
                detected = item.get("detected_streams", [])
                if not isinstance(detected, list):
                    continue

                keep_orders: List[int] = []
                for stream in detected:
                    if not isinstance(stream, dict):
                        continue
                    lang = str(stream.get("language") or "").strip().lower()
                    if lang not in selected_languages:
                        continue
                    try:
                        order = int(stream.get("stream_order"))
                    except Exception:
                        continue
                    if order >= 0:
                        keep_orders.append(order)

                if keep_orders:
                    keep_map[str(Path(file_path).resolve())] = sorted(set(keep_orders))

            if not keep_map:
                QMessageBox.warning(
                    self,
                    "No Matching Streams",
                    "None of the selected languages matched detected stream data for this run.",
                )
                return None

            return keep_map, sorted(selected_languages)

        def _prompt_and_start_audio_prune(self, result: Dict[str, object]) -> None:
            details = result.get("details", [])
            if not isinstance(details, list):
                return

            promptable_details: List[Dict[str, object]] = []
            for item in details:
                if not isinstance(item, dict):
                    continue
                detected = item.get("detected_streams", [])
                if isinstance(detected, list) and detected:
                    promptable_details.append(item)

            if not promptable_details:
                return

            reply = QMessageBox.question(
                self,
                "Audio Stream Cleanup",
                "Detection is complete. Do you want to choose languages to keep\n"
                "and remove other detected audio streams to reduce file size?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

            selection = self._prompt_audio_languages_to_keep(promptable_details)
            if selection is None:
                return

            keep_map, selected_languages = selection
            self._log("Selected languages to keep: " + ", ".join(selected_languages))
            options: Dict[str, object] = {
                "folders": [],
                "target_files": list(keep_map.keys()),
                "recursive": False,
                "overwrite": self.overwrite_checkbox.isChecked(),
                "prune_audio_suffix": "_audiopruned",
                "keep_audio_orders_by_file": keep_map,
            }
            QTimer.singleShot(50, lambda: self._start_worker("prune_audio_streams", options))

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
            elif action == "tag_audio_language" or action == "tag_audio_languages":
                self._log(
                    "Audio language detection complete: scanned={scanned}, processed={processed}, "
                    "skipped={skipped}, failed={failed}".format(
                        scanned=result.get("scanned", 0),
                        processed=result.get("processed", 0),
                        skipped=result.get("skipped", 0),
                        failed=result.get("failed", 0),
                    )
                )
                if int(result.get("processed", 0)) > 0:
                    self._prompt_and_start_audio_prune(result)
            elif action == "prune_audio_streams":
                self._log(
                    "Audio pruning complete: scanned={scanned}, processed={processed}, "
                    "skipped={skipped}, failed={failed}".format(
                        scanned=result.get("scanned", 0),
                        processed=result.get("processed", 0),
                        skipped=result.get("skipped", 0),
                        failed=result.get("failed", 0),
                    )
                )
            elif action == "sync_subtitles":
                self._log(
                    "Subtitle sync complete: scanned={scanned}, synced={processed}, "
                    "skipped={skipped}, failed={failed}".format(
                        scanned=result.get("scanned", 0),
                        processed=result.get("processed", 0),
                        skipped=result.get("skipped", 0),
                        failed=result.get("failed", 0),
                    )
                )
                details = result.get("details", [])
                if isinstance(details, list):
                    for item in details:
                        if not isinstance(item, dict):
                            continue
                        status = item.get("status", "")
                        reason = item.get("reason", "")
                        fname = Path(item.get("output_path", "") or item.get("file", "")).name or "?"
                        if status == "synced":
                            self._log(f"  \u2713 {fname} \u2014 {reason}")
                        elif status == "skipped":
                            self._log(f"  ~ Skipped: {reason}")
                        elif status == "failed":
                            self._log(f"  \u2717 Failed: {reason}")
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
                "selected_folders": [item.text() for item in self.folder_list.selectedItems()],
                "selected_targets": [item.text() for item in self.target_file_list.selectedItems()],
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
                "convert_backend": self.convert_backend_combo.currentData(),
                "repair_backend": self.repair_backend_combo.currentData(),
                "ffmpeg_bin": self.ffmpeg_bin_input.text(),
                "ffprobe_bin": self.ffprobe_bin_input.text(),
                "mkvmerge_bin": self.mkvmerge_bin_input.text(),
                "handbrake_bin": self.handbrake_bin_input.text(),
                "makemkvcon_bin": self.makemkvcon_bin_input.text(),
                "command_feedback": self.command_feedback_combo.currentData(),
                "ffmpeg_loglevel": self.ffmpeg_loglevel_combo.currentData(),
                "ffprobe_loglevel": self.ffprobe_loglevel_combo.currentData(),
                "ai_backend": self.ai_backend_combo.currentData() if self.ai_backend_combo else "auto",
                "whisper_model": self.whisper_model_combo.currentText() if self.whisper_model_combo else "base",
                "whisper_language": self.whisper_language_input.text() if self.whisper_language_input else "",
                "audio_lang_strategy": self.audio_lang_strategy_combo.currentData() if self.audio_lang_strategy_combo else "snippets",
                "audio_lang_snippets": self.audio_lang_snippets_input.text() if self.audio_lang_snippets_input else "3",
                "audio_lang_seconds": self.audio_lang_seconds_input.text() if self.audio_lang_seconds_input else "25",
                "audio_lang_overwrite": self.audio_lang_overwrite_checkbox.isChecked() if self.audio_lang_overwrite_checkbox else False,
                "audio_lang_detect_only": self.audio_lang_detect_only_checkbox.isChecked() if self.audio_lang_detect_only_checkbox else False,
                "sync_backend": self.sync_backend_combo.currentData() if self.sync_backend_combo else "auto",
                "sync_language": self.sync_language_input.text() if self.sync_language_input else "",
                "sync_overwrite": self.sync_overwrite_checkbox.isChecked() if self.sync_overwrite_checkbox else False,
                "hw_accel": self.hw_accel_checkbox.isChecked(),
                "save_next_to_source": self.save_next_to_source_checkbox.isChecked(),
                "custom_output_dir": self.custom_output_dir_input.text(),
                "tools_tab_index": self.tools_tabs.currentIndex(),
                "window_x": self.x(),
                "window_y": self.y(),
                "window_width": self.width(),
                "window_height": self.height(),
                "window_maximized": self.isMaximized(),
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

            selected_folders = set(ui_state.get("selected_folders", []))
            if selected_folders:
                for i in range(self.folder_list.count()):
                    item = self.folder_list.item(i)
                    if item.text() in selected_folders:
                        item.setSelected(True)

            selected_targets = set(ui_state.get("selected_targets", []))
            if selected_targets:
                for i in range(self.target_file_list.count()):
                    item = self.target_file_list.item(i)
                    if item.text() in selected_targets:
                        item.setSelected(True)
            
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
            self.save_next_to_source_checkbox.setChecked(ui_state.get("save_next_to_source", True))
            self.custom_output_dir_input.setText(str(ui_state.get("custom_output_dir", "")))
            self.custom_output_dir_input.setEnabled(not self.save_next_to_source_checkbox.isChecked())
            self.custom_output_dir_browse_button.setEnabled(not self.save_next_to_source_checkbox.isChecked())
            
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

            convert_backend = ui_state.get("convert_backend", "auto")
            convert_index = self.convert_backend_combo.findData(convert_backend)
            if convert_index >= 0:
                self.convert_backend_combo.setCurrentIndex(convert_index)

            repair_backend = ui_state.get("repair_backend", "auto")
            repair_index = self.repair_backend_combo.findData(repair_backend)
            if repair_index >= 0:
                self.repair_backend_combo.setCurrentIndex(repair_index)

            self.ffmpeg_bin_input.setText(str(ui_state.get("ffmpeg_bin", self.ffmpeg_bin_input.text())))
            self.ffprobe_bin_input.setText(str(ui_state.get("ffprobe_bin", self.ffprobe_bin_input.text())))
            self.mkvmerge_bin_input.setText(str(ui_state.get("mkvmerge_bin", self.mkvmerge_bin_input.text())))
            self.handbrake_bin_input.setText(str(ui_state.get("handbrake_bin", self.handbrake_bin_input.text())))
            self.makemkvcon_bin_input.setText(str(ui_state.get("makemkvcon_bin", self.makemkvcon_bin_input.text())))

            command_feedback = ui_state.get("command_feedback", "normal")
            feedback_index = self.command_feedback_combo.findData(command_feedback)
            if feedback_index >= 0:
                self.command_feedback_combo.setCurrentIndex(feedback_index)

            ffmpeg_level = ui_state.get("ffmpeg_loglevel", "warning")
            ffmpeg_level_index = self.ffmpeg_loglevel_combo.findData(ffmpeg_level)
            if ffmpeg_level_index >= 0:
                self.ffmpeg_loglevel_combo.setCurrentIndex(ffmpeg_level_index)

            ffprobe_level = ui_state.get("ffprobe_loglevel", "error")
            ffprobe_level_index = self.ffprobe_loglevel_combo.findData(ffprobe_level)
            if ffprobe_level_index >= 0:
                self.ffprobe_loglevel_combo.setCurrentIndex(ffprobe_level_index)
            
            # Restore Whisper AI settings (only if AI is enabled)
            if self.use_ai and self.whisper_model_combo and self.whisper_language_input:
                ai_backend = ui_state.get("ai_backend", "auto")
                if self.ai_backend_combo is not None:
                    ai_backend_index = self.ai_backend_combo.findData(ai_backend)
                    if ai_backend_index >= 0:
                        self.ai_backend_combo.setCurrentIndex(ai_backend_index)
                whisper_model = ui_state.get("whisper_model", "base")
                index = self.whisper_model_combo.findText(whisper_model)
                if index >= 0:
                    self.whisper_model_combo.setCurrentIndex(index)
                self.whisper_language_input.setText(ui_state.get("whisper_language", ""))

            if self.use_ai and self.audio_lang_strategy_combo and self.audio_lang_snippets_input and self.audio_lang_seconds_input:
                strategy = ui_state.get("audio_lang_strategy", "snippets")
                index = self.audio_lang_strategy_combo.findData(strategy)
                if index >= 0:
                    self.audio_lang_strategy_combo.setCurrentIndex(index)
                self.audio_lang_snippets_input.setText(str(ui_state.get("audio_lang_snippets", "3")))
                self.audio_lang_seconds_input.setText(str(ui_state.get("audio_lang_seconds", "25")))
                if self.audio_lang_overwrite_checkbox:
                    self.audio_lang_overwrite_checkbox.setChecked(bool(ui_state.get("audio_lang_overwrite", False)))
                if self.audio_lang_detect_only_checkbox:
                    self.audio_lang_detect_only_checkbox.setChecked(bool(ui_state.get("audio_lang_detect_only", False)))

            if self.use_ai and self.sync_language_input:
                if self.sync_backend_combo is not None:
                    sync_backend = ui_state.get("sync_backend", "auto")
                    sync_backend_index = self.sync_backend_combo.findData(sync_backend)
                    if sync_backend_index >= 0:
                        self.sync_backend_combo.setCurrentIndex(sync_backend_index)
                self.sync_language_input.setText(str(ui_state.get("sync_language", "")))
                if self.sync_overwrite_checkbox:
                    self.sync_overwrite_checkbox.setChecked(bool(ui_state.get("sync_overwrite", False)))

            self.hw_accel_checkbox.setChecked(bool(ui_state.get("hw_accel", False)))

            tools_tab_index = int(ui_state.get("tools_tab_index", 0))
            if 0 <= tools_tab_index < self.tools_tabs.count():
                self.tools_tabs.setCurrentIndex(tools_tab_index)

            if bool(ui_state.get("window_maximized", False)):
                self.setWindowState(self.windowState() | Qt.WindowState.WindowMaximized)
            else:
                x = int(ui_state.get("window_x", self.x()))
                y = int(ui_state.get("window_y", self.y()))
                w = int(ui_state.get("window_width", self.width()))
                h = int(ui_state.get("window_height", self.height()))
                if w > 200 and h > 150:
                    self.setGeometry(x, y, w, h)

            self._refresh_dependency_status(log_missing=False, announce=False)
            
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
    processor = SubtitleProcessor(
        ffmpeg_bin=(args.ffmpeg_bin or None),
        ffprobe_bin=(args.ffprobe_bin or None),
        mkvmerge_bin=(args.mkvmerge_bin or None),
        handbrake_bin=(args.handbrake_bin or None),
        makemkvcon_bin=(args.makemkvcon_bin or None),
        command_feedback=args.command_feedback,
        ffmpeg_loglevel=args.ffmpeg_loglevel,
        ffprobe_loglevel=args.ffprobe_loglevel,
        log_callback=lambda m: print(f"[log] {m}"),
    )

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
            output_root=args.output_root or None,
        )
        cli_print_json(summary.to_dict())
        return 0

    if args.mode == "include":
        summary = processor.include_subtitles(
            folders=folders,
            recursive=recursive,
            overwrite=args.overwrite,
            output_suffix=args.suffix,
            output_root=args.output_root or None,
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

    if args.mode == "tag-audio-language":
        summary = processor.detect_and_tag_audio_languages(
            folders=folders,
            recursive=recursive,
            target_files=[],
            model_size=args.model_size,
            strategy=args.strategy,
            snippet_count=max(1, args.snippets),
            sample_seconds=max(5.0, args.sample_seconds),
            overwrite=args.overwrite,
            output_suffix=args.suffix,
            overwrite_existing_tags=args.overwrite_existing_tags,
            detect_only=bool(args.detect_only),
            output_root=args.output_root or None,
        )
        cli_print_json(summary.to_dict())
        return 0

    if args.mode == "sync-subtitles":
        summary = processor.sync_subtitles(
            folders=folders,
            recursive=recursive,
            target_files=[],
            model_size=args.model_size,
            language=args.language or None,
            overwrite=args.overwrite,
            output_suffix=args.suffix,
            max_offset_seconds=max(10.0, args.max_offset),
            verification_tolerance_seconds=max(0.5, args.tolerance),
            output_root=args.output_root or None,
            sync_backend=args.sync_backend,
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
        ("tag-audio-language", "_langtagged"),
        ("sync-subtitles", "_synced"),
    ):
        cmd = subparsers.add_parser(mode, help=f"Run {mode} operation in CLI mode")
        cmd.add_argument("--folders", nargs="+", required=True, help="One or more folders to process")
        cmd.add_argument("--no-recursive", action="store_true", help="Do not scan subfolders")
        cmd.add_argument("--ffmpeg-bin", default="", help="Path or command name for ffmpeg")
        cmd.add_argument("--ffprobe-bin", default="", help="Path or command name for ffprobe")
        cmd.add_argument("--mkvmerge-bin", default="", help="Path or command name for mkvmerge (MKVToolNix)")
        cmd.add_argument("--handbrake-bin", default="", help="Path or command name for HandBrakeCLI")
        cmd.add_argument("--makemkvcon-bin", default="", help="Path or command name for makemkvcon")
        cmd.add_argument(
            "--command-feedback",
            choices=["quiet", "normal", "verbose"],
            default="normal",
            help="How much command stderr/stdout feedback to print",
        )
        cmd.add_argument(
            "--ffmpeg-loglevel",
            choices=["quiet", "error", "warning", "info", "verbose"],
            default="warning",
            help="FFmpeg loglevel passed to ffmpeg commands",
        )
        cmd.add_argument(
            "--ffprobe-loglevel",
            choices=["quiet", "error", "warning", "info", "verbose"],
            default="error",
            help="FFprobe loglevel passed to ffprobe commands",
        )
        if mode == "scan":
            cmd.add_argument(
                "--only-with-embedded",
                action="store_true",
                help="Only include files that contain embedded subtitle streams",
            )
        if mode in {"remove", "include", "extract"}:
            cmd.add_argument("--overwrite", action="store_true", help="Overwrite original files")
            cmd.add_argument("--suffix", default=default_suffix, help="Output filename suffix")
            if mode in {"remove", "include"}:
                cmd.add_argument("--output-root", default="", help="Custom output folder (default: next to source)")
        if mode == "tag-audio-language":
            cmd.add_argument("--overwrite", action="store_true", help="Overwrite original files")
            cmd.add_argument("--suffix", default=default_suffix, help="Output filename suffix")
            cmd.add_argument("--output-root", default="", help="Custom output folder (default: next to source)")
            cmd.add_argument("--model-size", default="base", help="Whisper model size")
            cmd.add_argument(
                "--strategy",
                choices=["snippets", "full"],
                default="snippets",
                help="Language analysis strategy",
            )
            cmd.add_argument("--snippets", type=int, default=3, help="Number of random snippets (snippets strategy)")
            cmd.add_argument("--sample-seconds", type=float, default=25.0, help="Duration per snippet in seconds")
            cmd.add_argument(
                "--overwrite-existing-tags",
                action="store_true",
                help="Re-detect and overwrite existing audio language tags",
            )
            cmd.add_argument(
                "--detect-only",
                action="store_true",
                help="Detect/report stream languages only; do not write metadata or create output files",
            )
        if mode == "sync-subtitles":
            cmd.add_argument("--overwrite", action="store_true", help="Overwrite original subtitle file in-place")
            cmd.add_argument("--suffix", default=default_suffix, help="Output subtitle filename suffix")
            cmd.add_argument("--output-root", default="", help="Custom output folder (default: next to source)")
            cmd.add_argument("--model-size", default="base", help="Whisper model size")
            cmd.add_argument("--language", default="", help="Language hint for Whisper (e.g. en, es); blank = auto-detect")
            cmd.add_argument(
                "--sync-backend",
                choices=["auto", "whisper-offset", "aeneas"],
                default="auto",
                help="Backend for subtitle sync",
            )
            cmd.add_argument("--max-offset", type=float, default=300.0, help="Maximum offset in seconds to consider (default: 300)")
            cmd.add_argument("--tolerance", type=float, default=2.0, help="Verification coverage tolerance in seconds (default: 2.0)")
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
