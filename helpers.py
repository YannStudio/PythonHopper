from typing import Any, List, Dict
import datetime
import errno
import os
import re
import shutil
from collections import defaultdict


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


def _sanitize_bundle_label(label: str | None) -> str:
    """Return a filesystem friendly label for export bundles."""

    text = _to_str(label).strip()
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    # Replace all characters that may cause issues on filesystems.
    sanitized = re.sub(r"[^0-9A-Za-z._-]+", "_", text)
    sanitized = sanitized.strip("._-")
    return sanitized


def create_export_bundle(
    root: str,
    label: str | None = None,
    *,
    timestamp: datetime.datetime | None = None,
    latest_name: str = "latest",
    max_attempts: int = 50,
) -> Dict[str, str]:
    """Create a timestamped export bundle directory inside ``root``.

    Parameters
    ----------
    root:
        Base directory that will contain per-run export bundles.
    label:
        Optional descriptor that becomes part of the directory name after
        sanitising. When empty, only the timestamp will be used.
    timestamp:
        Optional datetime used for deterministic naming in tests.
    latest_name:
        Name of the symlink that should point to the most recent bundle.
    max_attempts:
        Number of attempts when incrementing suffixes for existing bundles.

    Returns
    -------
    dict
        Mapping with ``root``, ``name``, ``path`` and ``latest`` keys.
    """

    if max_attempts < 1:
        raise ValueError("max_attempts must be >= 1")

    ts = timestamp or datetime.datetime.now()
    base_root = os.path.abspath(root)
    os.makedirs(base_root, exist_ok=True)

    label_clean = _sanitize_bundle_label(label)
    base_name = ts.strftime("%Y%m%d-%H%M%S")
    if label_clean:
        base_name = f"{base_name}_{label_clean}"

    attempt = 0
    bundle_path = ""
    bundle_name = ""
    while attempt < max_attempts:
        suffix = "" if attempt == 0 else f"-{attempt + 1}"
        candidate_name = f"{base_name}{suffix}"
        candidate_path = os.path.join(base_root, candidate_name)
        try:
            os.makedirs(candidate_path)
            bundle_path = candidate_path
            bundle_name = candidate_name
            break
        except FileExistsError:
            attempt += 1
            continue
        except OSError as exc:
            if exc.errno == errno.EEXIST:
                attempt += 1
                continue
            raise
    else:
        raise RuntimeError(
            f"Kon geen exportbundel aanmaken in {base_root} na {max_attempts} pogingen"
        )

    latest_path = os.path.join(base_root, latest_name)
    try:
        if os.path.lexists(latest_path):
            if os.path.islink(latest_path) or not os.path.isdir(latest_path):
                os.unlink(latest_path)
            else:
                shutil.rmtree(latest_path)
        os.symlink(bundle_name, latest_path)
    except OSError:
        # Symlinks might not be supported; expose that to callers via None.
        latest_path = None

    return {
        "root": base_root,
        "name": bundle_name,
        "path": bundle_path,
        "latest": latest_path,
    }
