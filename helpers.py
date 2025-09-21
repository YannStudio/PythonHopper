from __future__ import annotations

import datetime
import os
import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Dict, List


def _to_str(x: Any) -> str:
    return "" if x is None else str(x)


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


def _build_file_index(source_folder: str, selected_exts: List[str]) -> Dict[str, List[str]]:
    idx = defaultdict(list)
    sel = set(e.lower() for e in selected_exts)
    for rootdir, _, files in os.walk(source_folder):
        for f in files:
            name, ext = os.path.splitext(f)
            if ext.lower() in sel:
                idx[name].append(os.path.join(rootdir, f))
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
    timestamp: datetime.datetime | None = None,
) -> ExportBundleResult:
    """Create or determine an export bundle directory within ``base_dir``.

    The directory name is derived from ``project_number`` and ``project_name``.
    Missing or unusable values fall back to ``bundle-YYYYMMDD-HHMMSS``. When
    ``dry_run`` is ``True`` the directory and optional symlink are not created.

    ``latest_symlink`` enables creation of a symlink inside ``base_dir`` that
    points to the bundle directory. When a string is provided it is used as the
    symlink name; otherwise ``"latest"`` is used. Existing non-symlink paths
    with the same name are left untouched and reported via ``warnings``.
    """

    root_dir = os.path.abspath(base_dir)
    components: List[str] = []
    warnings: List[str] = []

    pn = _sanitize_bundle_component(project_number)
    if pn:
        components.append(pn)
    elif _to_str(project_number).strip():
        warnings.append("Projectnummer bevat geen geldige tekens en is overgeslagen.")

    pname = _sanitize_bundle_component(project_name)
    if pname:
        components.append(pname)
    elif _to_str(project_name).strip():
        warnings.append("Projectnaam bevat geen geldige tekens en is overgeslagen.")

    used_fallback = False
    if components:
        folder_name = " - ".join(components)
    else:
        used_fallback = True
        if not (project_number or project_name):
            warnings.append(
                "Projectnummer of -naam ontbreekt; er wordt een fallbackmap gebruikt."
            )
        timestamp = timestamp or datetime.datetime.now()
        folder_name = timestamp.strftime("bundle-%Y%m%d-%H%M%S")

    bundle_dir = os.path.abspath(os.path.join(root_dir, folder_name))

    if not dry_run:
        os.makedirs(bundle_dir, exist_ok=True)

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
