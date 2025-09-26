"""Helpers for resolving stored client logo paths across the application."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable, Optional, Union

from clients_db import CLIENTS_DB_FILE

PathLike = Union[str, os.PathLike[str]]

MODULE_DIR = Path(__file__).resolve().parent
CLIENTS_DB_PATH = (MODULE_DIR / CLIENTS_DB_FILE).resolve()
CLIENT_LOGO_DIR = CLIENTS_DB_PATH.parent / "client_logos"


def _iter_logo_candidates(path: Path) -> Iterable[Path]:
    """Yield candidate absolute paths for a stored logo reference."""

    seen: set[Path] = set()

    def _yield_once(candidate: Path) -> Iterable[Path]:
        if candidate not in seen:
            seen.add(candidate)
            yield candidate

    if path.is_absolute():
        yield from _yield_once(path)
        rel = Path(path.name) if path.name else Path()
    else:
        rel = path

    rel_variants = [rel]
    if rel and rel.name:
        rel_variants.append(Path(rel.name))
    if rel.parts and rel.parts[0] == CLIENT_LOGO_DIR.name:
        stripped = Path(*rel.parts[1:]) if len(rel.parts) > 1 else Path(rel.name)
        rel_variants.append(stripped)

    bases = [Path.cwd(), MODULE_DIR, CLIENT_LOGO_DIR]
    for base in bases:
        for rel_candidate in rel_variants:
            candidate = base / rel_candidate if rel_candidate else base
            yield from _yield_once(candidate)


def resolve_logo_path(path: Optional[PathLike]) -> Optional[Path]:
    """Resolve ``path`` against common search locations.

    ``path`` may be absolute or relative. Relative paths are attempted against the
    current working directory, the application directory, and the ``client_logos``
    directory next to ``clients_db.json``. The first existing file is returned.
    """

    if not path:
        return None

    stored = Path(path)
    for candidate in _iter_logo_candidates(stored):
        if candidate.exists():
            return candidate
    return None
