"""Helpers for resolving resource and data paths across packaging targets."""

from __future__ import annotations

import os
import shutil
import sys
from functools import lru_cache
from pathlib import Path
from typing import Iterable

APP_NAME = "Filehopper"


def is_frozen() -> bool:
    """Return ``True`` when running from a frozen/packaged executable."""

    return bool(getattr(sys, "frozen", False))


@lru_cache()
def bundle_root() -> Path:
    """Return the directory that contains bundled application resources."""

    if is_frozen() and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent


@lru_cache()
def storage_dir() -> Path:
    """Return the directory that should contain mutable application data."""

    if is_frozen():
        return _user_data_dir()
    return Path.cwd()


def _user_data_dir() -> Path:
    home = Path.home()
    if sys.platform.startswith("win"):
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if base:
            return Path(base) / APP_NAME
        return home / "AppData" / "Local" / APP_NAME
    if sys.platform == "darwin":
        return home / "Library" / "Application Support" / APP_NAME
    xdg_home = os.environ.get("XDG_DATA_HOME")
    if xdg_home:
        return Path(xdg_home) / APP_NAME
    return home / ".local" / "share" / APP_NAME


def data_file(name: str) -> str:
    """Return a writable path for the data file ``name``.

    When running from source the file remains relative so tests can redirect
    storage by changing the current working directory. For frozen builds the
    path resolves to the per-user storage directory.
    """

    if is_frozen():
        path = storage_dir() / name
        path.parent.mkdir(parents=True, exist_ok=True)
        return str(path)
    return name


def ensure_runtime_files(filenames: Iterable[str]) -> None:
    """Ensure default data files exist in the runtime storage directory."""

    if not is_frozen():
        return

    src_root = bundle_root()
    dst_root = storage_dir()
    dst_root.mkdir(parents=True, exist_ok=True)

    for name in filenames:
        dest = dst_root / name
        if dest.exists():
            continue
        source = src_root / name
        if source.exists():
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, dest)
        else:
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.touch()
