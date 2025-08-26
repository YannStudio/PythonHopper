from typing import Any, List, Dict
import os
from collections import defaultdict


def _to_str(x: Any) -> str:
    return "" if x is None else str(x)


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
