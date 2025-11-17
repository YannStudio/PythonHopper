from __future__ import annotations

import datetime
import locale
import os
import re
from collections import defaultdict
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List

from export_bundle import create_export_bundle as _create_export_bundle


def _to_str(x: Any) -> str:
    return "" if x is None else str(x)


@lru_cache()
def favorite_marker() -> str:
    """Return the preferred marker for favorites (★ when supported)."""

    return "★" if _supports_unicode_star() else "*"


@lru_cache()
def favorite_prefix() -> str:
    """Return the marker followed by a trailing space for display lists."""

    marker = favorite_marker()
    return f"{marker} " if marker else ""


def strip_favorite_marker(text: str) -> str:
    """Remove a leading favorite marker (★/*) from ``text`` if present."""

    value = _to_str(text)
    for marker in ("★ ", "* "):
        if marker and marker in value:
            return value.replace(marker, "", 1)
    return value


def _supports_unicode_star() -> bool:
    """Detect whether the runtime can encode the ★ character safely."""

    candidate = "★"
    encodings = []
    pref = locale.getpreferredencoding(False)
    if pref:
        encodings.append(pref)
    if os.name == "nt":
        encodings.append("mbcs")
    for enc in encodings:
        try:
            candidate.encode(enc)
        except (UnicodeEncodeError, LookupError):
            continue
        else:
            return True
    return False


_VAT_RE = re.compile(r"^[A-Z]{2}[A-Z0-9]{2,12}$")


def validate_vat(vat: str) -> str:
    """Validate and normalize a VAT number.

    Returns the normalized VAT number (uppercase) when the input matches the
    basic pattern of two letters followed by 2-12 alphanumeric characters.
    Otherwise an empty string is returned.
    """
    v = _to_str(vat).strip().upper()
    if not _VAT_RE.fullmatch(v):
        return ""
    if not any(ch.isdigit() for ch in v[2:]):
        return ""
    return v


def _num_to_2dec(val: Any) -> str:
    """Parse '1,23'/'1.23' -> '1.23' met 2 dec; anders string teruggeven."""
    s = _to_str(val).strip()
    if not s:
        return ""
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        f = float(s)
        return f"{f:.2f}"
    except Exception:
        return _to_str(val)


def _pn_wrap_25(pn: str) -> str:
    """Harde regelbreuk NA 25 tekens; tot 25 op één regel, rest op volgende."""
    pn = _to_str(pn)
    if len(pn) <= 25:
        return pn
    return pn[:25] + "<br/>" + pn[25:]


def _material_nowrap(s: str) -> str:
    """Voorkom wrapping in Paragraph: vervang spaties door &nbsp;."""
    return _to_str(s).replace(" ", "&nbsp;")


_EXPORT_BUNDLE_DIR_RE = re.compile(r"^\d{4}-\d{2}-\d{2}_.+")


def _index_sort_key(path: str) -> tuple[int, int, str]:
    parts = Path(path).parts
    is_bundle = any(_EXPORT_BUNDLE_DIR_RE.match(part) for part in parts)
    is_latest = any(part.lower() == "latest" for part in parts)
    priority_group = 0 if (is_latest or is_bundle) else 1
    depth_score = -len(parts)
    return priority_group, depth_score, path.lower()


def _build_file_index(source_folder: str, selected_exts: List[str]) -> Dict[str, List[str]]:
    idx = defaultdict(list)
    sel = set(e.lower() for e in selected_exts)
    for rootdir, _, files in os.walk(source_folder):
        for f in files:
            name, ext = os.path.splitext(f)
            if ext.lower() in sel:
                idx[name].append(os.path.join(rootdir, f))
    for key, paths in idx.items():
        paths.sort(key=_index_sort_key)
    return idx


def _unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while True:
        candidate = f"{base} ({i}){ext}"
        if not os.path.exists(candidate):
            return candidate
        i += 1


@dataclass(slots=True)
class ExportBundleResult:
    """Result metadata for :func:`create_export_bundle`."""

    root_dir: str
    bundle_dir: str
    folder_name: str
    dry_run: bool
    used_fallback: bool
    warnings: List[str]
    latest_symlink: str | None = None


_INVALID_BUNDLE_CHARS = set('<>:\\"/|?*')


def _sanitize_bundle_component(value: object) -> str:
    """Return a filesystem-friendly representation of ``value``."""

    text = _to_str(value).strip()
    if not text:
        return ""
    # Collapse whitespace and drop control characters
    text = " ".join(text.split())
    cleaned = []
    for ch in text:
        if ch in _INVALID_BUNDLE_CHARS or ord(ch) < 32:
            cleaned.append("_")
            continue
        if ch == os.sep or (os.altsep and ch == os.altsep):
            cleaned.append("-")
            continue
        cleaned.append(ch)
    result = "".join(cleaned).strip(" .-_")
    return result


def create_export_bundle(
    base_dir: str,
    project_number: str | None,
    project_name: str | None,
    *,
    latest_symlink: bool | str = False,
    dry_run: bool = False,
    timestamp: datetime.datetime | datetime.date | None = None,
) -> ExportBundleResult:
    """Create or determine an export bundle directory within ``base_dir``."""

    root_dir = os.path.abspath(base_dir)
    warnings: List[str] = []
    used_fallback = False

    pn_raw = _to_str(project_number).strip()
    pn_clean = _sanitize_bundle_component(pn_raw)
    if pn_raw and not pn_clean:
        warnings.append("Projectnummer bevat geen geldige tekens en is overgeslagen.")
    if not pn_clean:
        pn_clean = "project"
        used_fallback = True

    name_raw = _to_str(project_name).strip()
    name_for_slug = name_raw
    if name_raw and not _sanitize_bundle_component(name_raw):
        warnings.append("Projectnaam bevat geen geldige tekens en is overgeslagen.")
        name_for_slug = ""
    if not name_for_slug:
        used_fallback = True

    if not pn_raw and not name_raw:
        warnings.append(
            "Projectnummer of -naam ontbreekt; er wordt een fallbackmap gebruikt."
        )

    bundle_date: datetime.date | None
    if timestamp is None:
        bundle_date = None
    elif isinstance(timestamp, datetime.datetime):
        bundle_date = timestamp.date()
    elif isinstance(timestamp, datetime.date):
        bundle_date = timestamp
    else:  # pragma: no cover - defensive
        raise TypeError("timestamp must be a date, datetime, or None")

    if not os.path.exists(root_dir):
        os.makedirs(root_dir, exist_ok=True)
    elif not os.path.isdir(root_dir):  # pragma: no cover - defensive
        raise NotADirectoryError(f"Exportbasis is geen map: {root_dir}")

    bundle_path = _create_export_bundle(
        root_dir,
        pn_clean,
        name_for_slug or pn_clean,
        date=bundle_date,
        dry_run=dry_run,
        create_latest_symlink=False,
    )

    folder_name = bundle_path.name
    bundle_dir = str(bundle_path)

    latest_path: str | None = None
    link_requested = bool(latest_symlink)
    link_name = "latest"
    if isinstance(latest_symlink, str):
        custom = _sanitize_bundle_component(latest_symlink)
        if custom:
            link_name = custom
        elif latest_symlink.strip():
            warnings.append(
                "Naam voor 'latest'-symlink bevat geen geldige tekens; standaardnaam gebruikt."
            )
    if link_requested:
        latest_path = os.path.abspath(os.path.join(root_dir, link_name))
        if not dry_run:
            try:
                if os.path.lexists(latest_path):
                    if os.path.islink(latest_path) or not os.path.isdir(latest_path):
                        os.unlink(latest_path)
                    else:
                        warnings.append(
                            f"Kan symlink '{latest_path}' niet maken: pad bestaat al en is geen symlink."
                        )
                        latest_path = None
                if latest_path is not None:
                    os.symlink(bundle_dir, latest_path, target_is_directory=True)
            except (OSError, NotImplementedError) as exc:
                warnings.append(f"Kon symlink '{latest_path}' niet maken: {exc}")
                latest_path = None

    return ExportBundleResult(
        root_dir=root_dir,
        bundle_dir=bundle_dir,
        folder_name=folder_name,
        dry_run=dry_run,
        used_fallback=used_fallback,
        warnings=warnings,
        latest_symlink=latest_path,
    )
