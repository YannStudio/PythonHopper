"""Safe helpers for writing mutable JSON data files."""

from __future__ import annotations

import json
import os
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Optional

BACKUP_DIR_NAME = ".filehopper_backups"
DISABLE_BACKUPS_ENV = "FILEHOPPER_DISABLE_BACKUPS"
DEFAULT_BACKUP_KEEP = 25


def backups_enabled() -> bool:
    """Return whether automatic data-file backups should be written."""

    value = os.environ.get(DISABLE_BACKUPS_ENV, "").strip().lower()
    return value not in {"1", "true", "yes", "on"}


def backup_existing_file(path: str | os.PathLike[str], *, keep: int = DEFAULT_BACKUP_KEEP) -> Optional[Path]:
    """Copy an existing file to the backup folder and return the backup path."""

    source = Path(path)
    if not backups_enabled() or not source.exists() or not source.is_file():
        return None

    backup_dir = source.parent / BACKUP_DIR_NAME / source.stem
    backup_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = backup_dir / f"{source.stem}_{timestamp}{source.suffix}"
    counter = 2
    while candidate.exists():
        candidate = backup_dir / f"{source.stem}_{timestamp}_{counter}{source.suffix}"
        counter += 1

    shutil.copy2(source, candidate)
    _prune_backups(backup_dir, keep=keep)
    return candidate


def write_json_with_backup(
    path: str | os.PathLike[str],
    payload: Any,
    *,
    indent: int = 2,
    ensure_ascii: bool = False,
    keep_backups: int = DEFAULT_BACKUP_KEEP,
) -> None:
    """Write JSON atomically after backing up the previous version if needed."""

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, indent=indent, ensure_ascii=ensure_ascii)

    if target.exists():
        try:
            if target.read_text(encoding="utf-8") == text:
                return
        except OSError:
            pass
        backup_existing_file(target, keep=keep_backups)

    with NamedTemporaryFile(
        "w",
        encoding="utf-8",
        dir=str(target.parent),
        delete=False,
        prefix=f".{target.name}.",
        suffix=".tmp",
    ) as handle:
        temp_path = Path(handle.name)
        handle.write(text)

    try:
        os.replace(temp_path, target)
    except Exception:
        try:
            temp_path.unlink(missing_ok=True)
        finally:
            raise


def _prune_backups(backup_dir: Path, *, keep: int) -> None:
    if keep <= 0:
        return

    backups = [path for path in backup_dir.iterdir() if path.is_file()]
    backups.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    for old_backup in backups[keep:]:
        try:
            old_backup.unlink()
        except OSError:
            pass
