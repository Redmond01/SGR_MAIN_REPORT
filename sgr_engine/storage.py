"""Disk-backed, per-session workbook storage for low-memory hosting.

Streamlit Community Cloud gives each app roughly 1 GB of RAM.  Holding every
uploaded workbook, every processed workbook, and a ZIP archive as ``bytes`` in
``st.session_state`` multiplies memory use several times over.  Instead, each
browser session gets its own temporary folder::

    <tmp>/sgr_sessions/<session_id>/
        raw/     original uploaded workbooks (read-only source for processing)
        ready/   validated, processed workbooks queued for download

``st.session_state`` then only needs the small :class:`StationState` objects.
The "Download All" ZIP is streamed from ``ready/`` to a file on disk using
``ZIP_STORED`` (``.xlsx`` files are already compressed), so packaging 60
workbooks needs almost no extra RAM.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from typing import Dict, Iterable, List, Optional

__all__ = ["SessionStore", "cleanup_stale_sessions"]

_ROOT = os.path.join(tempfile.gettempdir(), "sgr_sessions")
_SAFE_RE = re.compile(r"[^A-Za-z0-9._ ()\-]+")


def _safe_name(file_name: str) -> str:
    """Strip any path components / unsafe characters from an uploaded file name."""
    base = os.path.basename(str(file_name).replace("\\", "/"))
    base = _SAFE_RE.sub("_", base).strip() or "station.xlsx"
    return base


def cleanup_stale_sessions(max_age_hours: float = 24.0) -> int:
    """Delete session folders untouched for ``max_age_hours`` (abandoned browser tabs)."""
    if not os.path.isdir(_ROOT):
        return 0
    cutoff = time.time() - max_age_hours * 3600
    removed = 0
    for name in os.listdir(_ROOT):
        path = os.path.join(_ROOT, name)
        try:
            if os.path.isdir(path) and os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)
                removed += 1
        except OSError:
            continue
    return removed


class SessionStore:
    """Per-session folder holding raw uploads and validated output workbooks."""

    def __init__(self, session_id: Optional[str] = None):
        self.session_id = session_id or uuid.uuid4().hex
        self.root = os.path.join(_ROOT, self.session_id)
        self.raw_dir = os.path.join(self.root, "raw")
        self.ready_dir = os.path.join(self.root, "ready")
        self.export_dir = os.path.join(self.root, "export")
        for d in (self.raw_dir, self.ready_dir, self.export_dir):
            os.makedirs(d, exist_ok=True)

    # ------------------------------------------------------------ helpers
    def touch(self):
        try:
            os.utime(self.root, None)
        except OSError:
            pass

    @staticmethod
    def _atomic_write(path: str, data: bytes):
        """Write to a temp file then rename, so a crash never leaves a half-written workbook."""
        tmp = f"{path}.part"
        with open(tmp, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)

    # ------------------------------------------------------------ raw uploads
    def raw_path(self, file_name: str) -> str:
        return os.path.join(self.raw_dir, _safe_name(file_name))

    def save_raw(self, file_name: str, data: bytes) -> str:
        path = self.raw_path(file_name)
        self._atomic_write(path, data)
        self.touch()
        return path

    def load_raw(self, file_name: str) -> bytes:
        with open(self.raw_path(file_name), "rb") as f:
            return f.read()

    def has_raw(self, file_name: str) -> bool:
        return os.path.isfile(self.raw_path(file_name))

    # ------------------------------------------------------------ ready queue
    def ready_path(self, file_name: str) -> str:
        return os.path.join(self.ready_dir, _safe_name(file_name))

    def save_ready(self, file_name: str, data: bytes) -> str:
        path = self.ready_path(file_name)
        self._atomic_write(path, data)
        self.touch()
        return path

    def has_ready(self, file_name: str) -> bool:
        return os.path.isfile(self.ready_path(file_name))

    def load_ready(self, file_name: str) -> bytes:
        with open(self.ready_path(file_name), "rb") as f:
            return f.read()

    def remove_ready(self, file_name: str):
        try:
            os.remove(self.ready_path(file_name))
        except FileNotFoundError:
            pass

    def ready_size(self, file_name: str) -> int:
        try:
            return os.path.getsize(self.ready_path(file_name))
        except OSError:
            return 0

    # ------------------------------------------------------------ removal
    def remove_station(self, file_name: str):
        for p in (self.raw_path(file_name), self.ready_path(file_name)):
            try:
                os.remove(p)
            except FileNotFoundError:
                pass

    def clear(self):
        shutil.rmtree(self.root, ignore_errors=True)
        for d in (self.raw_dir, self.ready_dir, self.export_dir):
            os.makedirs(d, exist_ok=True)

    # ------------------------------------------------------------ ZIP export
    def build_zip(
        self,
        file_names: Iterable[str],
        zip_name: str,
        extra_files: Optional[Dict[str, bytes]] = None,
    ) -> str:
        """Stream the given ready workbooks into ``export/<zip_name>`` and return its path.

        Uses ``ZIP_STORED`` so each workbook is copied in chunks from disk without
        being decompressed or held in RAM.  Every member is CRC-checked afterwards.
        """
        zip_path = os.path.join(self.export_dir, _safe_name(zip_name))
        tmp_path = f"{zip_path}.part"
        names: List[str] = []
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_STORED, allowZip64=True) as zf:
            for fn in file_names:
                src = self.ready_path(fn)
                if not os.path.isfile(src):
                    raise FileNotFoundError(f"Queued workbook missing on disk: {fn}")
                arc = _safe_name(fn)
                zf.write(src, arcname=arc)
                names.append(arc)
            for arc, data in (extra_files or {}).items():
                zf.writestr(_safe_name(arc), data)
                names.append(_safe_name(arc))

        # Integrity check: every member must read back with a valid CRC
        with zipfile.ZipFile(tmp_path, "r") as zf:
            bad = zf.testzip()
            if bad is not None:
                raise zipfile.BadZipFile(f"Corrupted member in archive: {bad}")
            if sorted(zf.namelist()) != sorted(names):
                raise zipfile.BadZipFile("Archive member list does not match queued workbooks")

        os.replace(tmp_path, zip_path)
        self.touch()
        return zip_path
