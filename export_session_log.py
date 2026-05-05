from __future__ import annotations

import datetime as _dt
import hashlib
import json
import os
from pathlib import Path
from typing import Any, Dict, Mapping, MutableMapping


EXPORT_SESSION_LOG_FILENAME = "Filehopper-exportlog.json"
EXPORT_SESSION_LOG_SCHEMA_VERSION = 1

_STATE_KEYS = (
    "selections",
    "groups",
    "doc_types",
    "doc_numbers",
    "remarks",
    "deliveries",
    "exports",
    "en1090",
    "pricing",
)


def _to_str(value: Any) -> str:
    return "" if value is None else str(value)


def _clean_mapping(value: Any) -> Dict[str, Any]:
    if not isinstance(value, Mapping):
        return {}
    cleaned: Dict[str, Any] = {}
    for raw_key, raw_value in value.items():
        key = _to_str(raw_key).strip()
        if not key:
            continue
        cleaned[key] = raw_value
    return cleaned


def normalize_state_dict(value: Any) -> Dict[str, Any]:
    """Return a resilient plain-dict representation of a supplier selection state."""

    if isinstance(value, Mapping):
        source = value
    else:
        source = {
            key: getattr(value, key, {})
            for key in _STATE_KEYS
            if hasattr(value, key)
        }
        source["remember"] = bool(getattr(value, "remember", True))

    state: Dict[str, Any] = {}
    for key in _STATE_KEYS:
        state[key] = _clean_mapping(source.get(key, {}))
    state["remember"] = bool(source.get("remember", True))
    return state


def convert_offers_to_orders(state: Mapping[str, Any]) -> Dict[str, Any]:
    """Convert Offerteaanvraag document types to Bestelbon and clear OFF numbers."""

    converted = normalize_state_dict(state)
    doc_types = dict(converted.get("doc_types", {}))
    doc_numbers = dict(converted.get("doc_numbers", {}))
    for key, doc_type in list(doc_types.items()):
        if _to_str(doc_type).strip().lower().startswith("offerte"):
            doc_types[key] = "Bestelbon"
            if _to_str(doc_numbers.get(key)).strip().upper().startswith("OFF"):
                doc_numbers[key] = ""
    converted["doc_types"] = doc_types
    converted["doc_numbers"] = doc_numbers
    return converted


def _bom_fingerprint(bom_df: Any) -> Dict[str, Any]:
    if bom_df is None:
        return {"row_count": 0, "columns": [], "sha256": ""}
    try:
        columns = [_to_str(col) for col in list(getattr(bom_df, "columns", []))]
        row_count = int(len(bom_df))
    except Exception:
        return {"row_count": 0, "columns": [], "sha256": ""}

    preferred = [
        "PartNumber",
        "Description",
        "Production",
        "Finish",
        "RAL color",
        "Materiaal",
        "Aantal",
        "Oppervlakte",
        "Gewicht",
    ]
    present = [column for column in preferred if column in columns]
    if not present:
        present = columns

    digest = hashlib.sha256()
    digest.update("|".join(columns).encode("utf-8", "replace"))
    digest.update(f"\nrows={row_count}\n".encode("ascii"))
    try:
        for row in bom_df[present].fillna("").astype(str).itertuples(index=False, name=None):
            digest.update("\t".join(row).encode("utf-8", "replace"))
            digest.update(b"\n")
    except Exception:
        digest.update(repr((columns, row_count)).encode("utf-8", "replace"))
    return {
        "row_count": row_count,
        "columns": columns,
        "fingerprint_columns": present,
        "sha256": digest.hexdigest(),
    }


def build_export_session_log(
    *,
    project_number: str = "",
    project_name: str = "",
    client_name: str = "",
    bom_source_path: str = "",
    bom_df: Any = None,
    state: Any = None,
    app_version: str = "",
) -> Dict[str, Any]:
    source_path = _to_str(bom_source_path).strip()
    return {
        "schema_version": EXPORT_SESSION_LOG_SCHEMA_VERSION,
        "app": {
            "name": "Filehopper",
            "version": _to_str(app_version).strip(),
        },
        "created_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "project": {
            "number": _to_str(project_number).strip(),
            "name": _to_str(project_name).strip(),
            "client": _to_str(client_name).strip(),
        },
        "bom": {
            "source_path": source_path,
            "filename": os.path.basename(source_path) if source_path else "",
            **_bom_fingerprint(bom_df),
        },
        "order_state": normalize_state_dict(state or {}),
    }


def write_export_session_log(export_dir: str | os.PathLike[str], payload: Mapping[str, Any]) -> str:
    path = Path(export_dir) / EXPORT_SESSION_LOG_FILENAME
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(dict(payload), handle, indent=2, ensure_ascii=False)
    return str(path)


def find_export_session_logs(
    root_dir: str | os.PathLike[str],
    *,
    limit: int = 20,
) -> list[str]:
    """Return export session logs below ``root_dir``, newest first."""

    root = Path(root_dir)
    if not root.exists() or not root.is_dir():
        return []
    matches: list[tuple[float, str]] = []
    for path in root.rglob(EXPORT_SESSION_LOG_FILENAME):
        try:
            mtime = path.stat().st_mtime
        except OSError:
            continue
        matches.append((mtime, str(path)))
    matches.sort(key=lambda item: (item[0], item[1].lower()), reverse=True)
    if limit <= 0:
        return [path for _mtime, path in matches]
    return [path for _mtime, path in matches[:limit]]


def load_export_session_log(path: str | os.PathLike[str]) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, MutableMapping):
        raise ValueError("Exportlog heeft geen geldig JSON-object.")
    version = data.get("schema_version")
    if version != EXPORT_SESSION_LOG_SCHEMA_VERSION:
        raise ValueError(f"Niet-ondersteunde exportlog versie: {version!r}.")
    state = data.get("order_state", {})
    data["order_state"] = normalize_state_dict(state)
    return dict(data)
