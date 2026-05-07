from __future__ import annotations

import datetime as _dt
import hashlib
import json
import os
from copy import deepcopy
from collections.abc import Iterable as IterableABC
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, MutableMapping


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

_IMPORT_SECTION_TO_STATE_KEYS = {
    "suppliers": ("selections",),
    "groups": ("groups",),
    "documents": ("doc_types", "doc_numbers"),
    "remarks": ("remarks",),
    "deliveries": ("deliveries",),
    "exports": ("exports",),
    "en1090": ("en1090",),
    "pricing": ("pricing",),
}


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


def _clean_string_list(values: Any) -> list[str]:
    if not isinstance(values, IterableABC) or isinstance(values, (str, bytes, Mapping)):
        return []
    cleaned: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = _to_str(value).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        cleaned.append(text)
    return cleaned


def _clean_generated_documents(values: Any) -> list[Dict[str, Any]]:
    if not isinstance(values, IterableABC) or isinstance(values, (str, bytes, Mapping)):
        return []
    cleaned: list[Dict[str, Any]] = []
    for value in values:
        if not isinstance(value, Mapping):
            continue
        path = _to_str(value.get("path")).strip()
        if not path:
            continue
        record: Dict[str, Any] = {"path": path}
        for key in (
            "kind",
            "format",
            "selection_key",
            "context_kind",
            "context_label",
            "doc_type",
            "doc_number",
            "supplier",
        ):
            text = _to_str(value.get(key)).strip()
            if text:
                record[key] = text
        selection_keys = _clean_string_list(value.get("selection_keys"))
        if selection_keys:
            record["selection_keys"] = selection_keys
        cleaned.append(record)
    return cleaned


def normalize_export_info(value: Any) -> Dict[str, Any]:
    source = value if isinstance(value, Mapping) else {}
    return {
        "generated_documents": _clean_generated_documents(
            source.get("generated_documents", [])
        ),
        "status_messages": _clean_string_list(source.get("status_messages", [])),
        "path_limit_warnings": _clean_string_list(
            source.get("path_limit_warnings", [])
        ),
    }


def resolve_export_document_path(
    export_log_path: str | os.PathLike[str],
    record: Mapping[str, Any],
) -> str:
    """Resolve a generated-document record path relative to its exportlog."""

    if not isinstance(record, Mapping):
        return ""
    raw_path = _to_str(record.get("path")).strip()
    if not raw_path:
        return ""
    candidate = Path(raw_path)
    if candidate.is_absolute():
        return str(candidate)

    base_dir = Path(export_log_path).resolve().parent
    resolved = (base_dir / raw_path).resolve()
    try:
        resolved.relative_to(base_dir)
    except ValueError:
        return ""
    return str(resolved)


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


def state_keys_for_import_sections(sections: Iterable[str]) -> set[str]:
    """Return normalized order-state keys covered by import section names."""

    keys: set[str] = set()
    for section in sections or []:
        name = _to_str(section).strip()
        if name in _STATE_KEYS:
            keys.add(name)
            continue
        keys.update(_IMPORT_SECTION_TO_STATE_KEYS.get(name, ()))
    return keys


def merge_order_state_sections(
    current_state: Any,
    incoming_state: Any,
    sections: Iterable[str],
) -> Dict[str, Any]:
    """Merge selected exportlog sections into the current order state."""

    current = normalize_state_dict(current_state)
    incoming = normalize_state_dict(incoming_state)
    selected_keys = state_keys_for_import_sections(sections)
    merged: Dict[str, Any] = {}
    for key in _STATE_KEYS:
        source = incoming if key in selected_keys else current
        merged[key] = deepcopy(source.get(key, {}))
    merged["remember"] = bool(current.get("remember", True))
    return merged


def state_selection_keys(state: Any) -> set[str]:
    """Return all selection keys referenced by a normalized or raw order state."""

    normalized = normalize_state_dict(state)
    keys: set[str] = set()
    for name in _STATE_KEYS:
        value = normalized.get(name, {})
        if not isinstance(value, Mapping):
            continue
        keys.update(_to_str(key).strip() for key in value.keys() if _to_str(key).strip())
    return keys


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


def _format_selection_key(key: str) -> str:
    prefix, sep, identifier = key.partition("::")
    if not sep:
        return key
    labels = {
        "production": "Productie",
        "finish": "Afwerking",
        "opticutter": "Brutemateriaal",
    }
    label = labels.get(prefix, prefix or "Selectie")
    return f"{label}: {identifier}"


def summarize_export_log_compatibility(
    payload: Mapping[str, Any],
    current_selection_keys: Iterable[str],
    *,
    current_bom_df: Any = None,
) -> Dict[str, Any]:
    """Compare an exportlog against the currently visible order rows and BOM."""

    state = payload.get("order_state", {}) if isinstance(payload, Mapping) else {}
    incoming_keys = state_selection_keys(state)
    current_keys = {
        _to_str(key).strip()
        for key in current_selection_keys
        if _to_str(key).strip()
    }
    matched = incoming_keys & current_keys
    missing = incoming_keys - current_keys
    new = current_keys - incoming_keys

    log_bom = payload.get("bom", {}) if isinstance(payload, Mapping) else {}
    current_bom = _bom_fingerprint(current_bom_df) if current_bom_df is not None else {}
    bom_changed = False
    if isinstance(log_bom, Mapping) and current_bom:
        log_sha = _to_str(log_bom.get("sha256")).strip()
        current_sha = _to_str(current_bom.get("sha256")).strip()
        if log_sha and current_sha:
            bom_changed = log_sha != current_sha
        else:
            try:
                bom_changed = int(log_bom.get("row_count", -1)) != int(
                    current_bom.get("row_count", -1)
                )
            except Exception:
                bom_changed = False

    return {
        "incoming_keys": sorted(incoming_keys, key=str.lower),
        "current_keys": sorted(current_keys, key=str.lower),
        "matched_keys": sorted(matched, key=str.lower),
        "missing_keys": sorted(missing, key=str.lower),
        "new_keys": sorted(new, key=str.lower),
        "bom_changed": bom_changed,
        "log_bom": dict(log_bom) if isinstance(log_bom, Mapping) else {},
        "current_bom": current_bom,
    }


def format_export_log_compatibility_message(
    summary: Mapping[str, Any],
    *,
    max_items: int = 8,
) -> str:
    """Return a concise Dutch review message for compatibility differences."""

    lines: list[str] = []
    log_bom = summary.get("log_bom", {})
    current_bom = summary.get("current_bom", {})
    if summary.get("bom_changed"):
        log_rows = (
            _to_str(log_bom.get("row_count")).strip()
            if isinstance(log_bom, Mapping)
            else ""
        )
        current_rows = (
            _to_str(current_bom.get("row_count")).strip()
            if isinstance(current_bom, Mapping)
            else ""
        )
        if log_rows or current_rows:
            lines.append(
                "De huidige BOM lijkt te verschillen van de BOM in de exportlog "
                f"(log: {log_rows or '?'} rijen, huidig: {current_rows or '?'} rijen)."
            )
        else:
            lines.append("De huidige BOM lijkt te verschillen van de BOM in de exportlog.")

    def _append_key_block(title: str, keys: object) -> None:
        if not isinstance(keys, list) or not keys:
            return
        lines.append(f"{title} ({len(keys)}):")
        shown = keys[:max_items]
        lines.extend(f"- {_format_selection_key(_to_str(key))}" for key in shown)
        remaining = len(keys) - len(shown)
        if remaining > 0:
            lines.append(f"- ... en {remaining} meer")

    _append_key_block(
        "Regels uit de exportlog die niet op deze bestelbonpagina staan",
        summary.get("missing_keys"),
    )
    _append_key_block(
        "Nieuwe regels op deze bestelbonpagina zonder exportlogwaarden",
        summary.get("new_keys"),
    )

    if not lines:
        return ""

    matched_count = len(summary.get("matched_keys", []) or [])
    incoming_count = len(summary.get("incoming_keys", []) or [])
    if incoming_count:
        lines.append(f"Gevonden matches: {matched_count} van {incoming_count} exportlogregel(s).")
    return "\n".join(lines)


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
    generated_documents: Any = None,
    status_messages: Any = None,
    path_limit_warnings: Any = None,
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
        "export": normalize_export_info(
            {
                "generated_documents": generated_documents,
                "status_messages": status_messages,
                "path_limit_warnings": path_limit_warnings,
            }
        ),
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
    data["export"] = normalize_export_info(data.get("export", {}))
    return dict(data)
