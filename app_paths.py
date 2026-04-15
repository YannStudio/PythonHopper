"""Helpers for resolving resource and data paths across packaging targets."""

from __future__ import annotations

import os
import shutil
import sys
from functools import lru_cache
from pathlib import Path
from typing import Iterable, Optional

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


def runtime_asset_root() -> Path:
    """Return a stable writable root for mutable assets such as uploaded logos."""

    if is_frozen():
        return storage_dir()
    return bundle_root()


def runtime_asset_dir(name: str) -> Path:
    """Return a writable asset subdirectory below :func:`runtime_asset_root`."""

    path = runtime_asset_root() / name
    path.mkdir(parents=True, exist_ok=True)
    return path


def to_runtime_relative_path(path: Path) -> str:
    """Return ``path`` relative to :func:`runtime_asset_root` when possible."""

    try:
        return path.resolve().relative_to(runtime_asset_root().resolve()).as_posix()
    except Exception:
        return str(path)


def resolve_runtime_path(
    path_str: str,
    *,
    extra_roots: Optional[Iterable[Path]] = None,
) -> Optional[Path]:
    """Resolve ``path_str`` against known runtime roots.

    Relative paths are first looked up below :func:`runtime_asset_root`, then
    fall back to the mutable storage directory, the application bundle root,
    and finally the current working directory for backward compatibility.
    """

    if not path_str:
        return None

    path = Path(path_str)
    if path.is_absolute():
        return path

    roots = [runtime_asset_root(), storage_dir(), bundle_root(), Path.cwd()]
    if extra_roots:
        roots.extend(Path(root) for root in extra_roots)

    seen: set[str] = set()
    fallback: Optional[Path] = None
    for root in roots:
        try:
            candidate = (root / path).resolve()
        except Exception:
            candidate = root / path
        key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)
        if fallback is None:
            fallback = candidate
        if candidate.exists():
            return candidate
    return fallback


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
