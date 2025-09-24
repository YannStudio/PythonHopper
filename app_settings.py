from __future__ import annotations

import json
from copy import deepcopy
from dataclasses import dataclass, field, asdict, fields
from pathlib import Path
from typing import Any, Iterable, List, Optional

from suppliers_db import SUPPLIERS_DB_FILE

SETTINGS_FILE = Path(SUPPLIERS_DB_FILE).with_name("app_settings.json")


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        val = value.strip().lower()
        if val in {"1", "true", "yes", "on"}:
            return True
        if val in {"0", "false", "no", "off"}:
            return False
    return False


def _as_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _normalize_patterns(patterns: Iterable[Any]) -> List[str]:
    cleaned: List[str] = []
    seen = set()
    for raw in patterns:
        if not isinstance(raw, str):
            continue
        pat = raw.strip()
        if not pat:
            continue
        if not pat.startswith("."):
            pat = "." + pat.lstrip(".")
        pat = pat.lower()
        if pat not in seen:
            cleaned.append(pat)
            seen.add(pat)
    return cleaned


def _normalize_key(value: str) -> str:
    base = [
        (ch.lower() if ch.isalnum() else "_")
        for ch in value.strip()
    ]
    key = "".join(base).strip("_")
    return key or "ext"


@dataclass
class FileExtensionSetting:
    key: str
    label: str
    patterns: List[str] = field(default_factory=list)
    enabled: bool = False

    @classmethod
    def from_any(cls, data: Any) -> "FileExtensionSetting":
        if isinstance(data, FileExtensionSetting):
            return cls(
                key=data.key,
                label=data.label,
                patterns=list(data.patterns),
                enabled=bool(data.enabled),
            )
        if not isinstance(data, dict):
            raise ValueError("file extension data must be a mapping")

        key = _as_str(data.get("key", "")).strip().lower()
        label = _as_str(data.get("label", "")).strip()
        patterns_raw = data.get("patterns", [])
        if isinstance(patterns_raw, str):
            patterns_iter = [patterns_raw]
        else:
            patterns_iter = patterns_raw
        patterns = _normalize_patterns(patterns_iter if patterns_iter is not None else [])

        if not key:
            if patterns:
                key = patterns[0].lstrip(".")
            elif label:
                key = _normalize_key(label)
        if not key:
            key = "ext"
        key = _normalize_key(key)

        if not label:
            if patterns:
                label = ", ".join(patterns)
            else:
                label = key.upper()

        if not patterns:
            patterns = ["." + key]

        enabled = _as_bool(data.get("enabled", False))

        return cls(key=key, label=label, patterns=patterns, enabled=enabled)

    @classmethod
    def from_user_input(
        cls,
        label: str,
        patterns_text: str,
        enabled: bool,
        *,
        key: Optional[str] = None,
    ) -> "FileExtensionSetting":
        label = _as_str(label).strip()
        raw = patterns_text.replace(";", ",")
        parts: List[str] = []
        for chunk in raw.split(","):
            chunk = chunk.strip()
            if not chunk:
                continue
            parts.extend(chunk.split())
        patterns = _normalize_patterns(parts)
        if not patterns:
            raise ValueError("Geef minstens één bestandsextensie op")
        if not label:
            label = ", ".join(patterns)
        if key:
            norm_key = _normalize_key(key)
        else:
            norm_key = _normalize_key(patterns[0].lstrip("."))
        return cls(key=norm_key, label=label, patterns=patterns, enabled=bool(enabled))


DEFAULT_FILE_EXTENSIONS: List[FileExtensionSetting] = [
    FileExtensionSetting(key="pdf", label="PDF (.pdf)", patterns=[".pdf"], enabled=False),
    FileExtensionSetting(
        key="step",
        label="STEP (.step, .stp)",
        patterns=[".step", ".stp"],
        enabled=False,
    ),
    FileExtensionSetting(key="dxf", label="DXF (.dxf)", patterns=[".dxf"], enabled=False),
    FileExtensionSetting(key="dwg", label="DWG (.dwg)", patterns=[".dwg"], enabled=False),
]


SUGGESTED_FILE_EXTENSION_GROUPS: List[tuple[str, List[FileExtensionSetting]]] = [
    (
        "Documenten",
        [
            FileExtensionSetting(
                key="word",
                label="Word (.doc, .docx)",
                patterns=[".doc", ".docx"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="excel",
                label="Excel (.xls, .xlsx, .xlsm)",
                patterns=[".xls", ".xlsx", ".xlsm"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="powerpoint",
                label="PowerPoint (.ppt, .pptx)",
                patterns=[".ppt", ".pptx"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="text",
                label="Tekstbestanden (.txt, .rtf)",
                patterns=[".txt", ".rtf"],
                enabled=False,
            ),
        ],
    ),
    (
        "Afbeeldingen",
        [
            FileExtensionSetting(
                key="fotos",
                label="Foto's (.jpg, .jpeg, .png)",
                patterns=[".jpg", ".jpeg", ".png"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="bitmap",
                label="Bitmap (.bmp, .tif, .tiff)",
                patterns=[".bmp", ".tif", ".tiff"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="vector",
                label="Vector (.svg, .ai)",
                patterns=[".svg", ".ai"],
                enabled=False,
            ),
        ],
    ),
    (
        "3D CAD – SolidWorks",
        [
            FileExtensionSetting(
                key="solidworks_part",
                label="SolidWorks Part (.sldprt)",
                patterns=[".sldprt"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="solidworks_assembly",
                label="SolidWorks Assembly (.sldasm)",
                patterns=[".sldasm"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="solidworks_drawing",
                label="SolidWorks Drawing (.slddrw)",
                patterns=[".slddrw"],
                enabled=False,
            ),
        ],
    ),
    (
        "3D CAD – Autodesk Inventor",
        [
            FileExtensionSetting(
                key="inventor_part",
                label="Inventor Part (.ipt)",
                patterns=[".ipt"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="inventor_assembly",
                label="Inventor Assembly (.iam)",
                patterns=[".iam"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="inventor_drawing",
                label="Inventor Drawing (.idw)",
                patterns=[".idw"],
                enabled=False,
            ),
        ],
    ),
    (
        "Archief",
        [
            FileExtensionSetting(
                key="zip",
                label="ZIP-archief (.zip)",
                patterns=[".zip"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="seven_zip",
                label="7-Zip (.7z)",
                patterns=[".7z"],
                enabled=False,
            ),
            FileExtensionSetting(
                key="rar",
                label="RAR (.rar)",
                patterns=[".rar"],
                enabled=False,
            ),
        ],
    ),
]


@dataclass
class AppSettings:
    source_folder: str = ""
    dest_folder: str = ""
    project_number: str = ""
    project_name: str = ""
    file_extensions: List[FileExtensionSetting] = field(
        default_factory=lambda: deepcopy(DEFAULT_FILE_EXTENSIONS)
    )
    zip_per_production: bool = True
    export_date_prefix: bool = False
    export_date_suffix: bool = False
    custom_prefix_enabled: bool = False
    custom_prefix_text: str = ""
    custom_suffix_enabled: bool = False
    custom_suffix_text: str = ""
    bundle_latest: bool = False
    bundle_dry_run: bool = False
    _path: Path = field(default=SETTINGS_FILE, repr=False, compare=False)

    @classmethod
    def load(cls, path: Optional[Any] = None) -> "AppSettings":
        settings_path = Path(path) if path is not None else SETTINGS_FILE
        try:
            with open(settings_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except FileNotFoundError:
            inst = cls()
            inst._path = settings_path
            return inst
        except (json.JSONDecodeError, OSError, TypeError, ValueError):
            inst = cls()
            inst._path = settings_path
            return inst

        inst = cls.from_dict(data)
        inst._path = settings_path
        return inst

    @classmethod
    def from_dict(cls, data: Any) -> "AppSettings":
        inst = cls()
        if not isinstance(data, dict):
            return inst
        legacy_flags = {}
        for key in ("pdf", "step", "dxf", "dwg"):
            if key in data:
                legacy_flags[key] = _as_bool(data.get(key))

        for field_info in fields(cls):
            name = field_info.name
            if name == "_path":
                continue
            if name not in data:
                if name == "file_extensions" and legacy_flags:
                    setattr(
                        inst,
                        name,
                        [
                            FileExtensionSetting(
                                key=ext.key,
                                label=ext.label,
                                patterns=list(ext.patterns),
                                enabled=legacy_flags.get(ext.key, ext.enabled),
                            )
                            for ext in DEFAULT_FILE_EXTENSIONS
                        ],
                    )
                continue
            cur_val = getattr(inst, name)
            raw = data.get(name)
            if isinstance(cur_val, bool):
                setattr(inst, name, _as_bool(raw))
            elif isinstance(cur_val, list) and name == "file_extensions":
                extensions: List[FileExtensionSetting] = []
                if isinstance(raw, list):
                    for item in raw:
                        try:
                            ext = FileExtensionSetting.from_any(item)
                        except ValueError:
                            continue
                        existing_keys = {e.key for e in extensions}
                        base_key = ext.key
                        suffix = 2
                        key_candidate = base_key
                        while key_candidate in existing_keys:
                            key_candidate = f"{base_key}_{suffix}"
                            suffix += 1
                        if key_candidate != ext.key:
                            ext = FileExtensionSetting(
                                key=key_candidate,
                                label=ext.label,
                                patterns=list(ext.patterns),
                                enabled=ext.enabled,
                            )
                        extensions.append(ext)
                if not extensions:
                    extensions = [
                        FileExtensionSetting(
                            key=ext.key,
                            label=ext.label,
                            patterns=list(ext.patterns),
                            enabled=legacy_flags.get(ext.key, ext.enabled),
                        )
                        for ext in DEFAULT_FILE_EXTENSIONS
                    ]
                else:
                    if legacy_flags:
                        for ext in extensions:
                            if ext.key in legacy_flags:
                                ext.enabled = legacy_flags[ext.key]
                setattr(inst, name, extensions)
            else:
                setattr(inst, name, _as_str(raw))
        return inst

    def save(self, path: Optional[Any] = None) -> None:
        settings_path = Path(path) if path is not None else getattr(self, "_path", SETTINGS_FILE)
        payload = {k: v for k, v in asdict(self).items() if k != "_path"}
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        with open(settings_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2, ensure_ascii=False)
        self._path = settings_path

