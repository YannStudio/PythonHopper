from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict, fields
from pathlib import Path
from typing import Any, Optional

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


@dataclass
class AppSettings:
    source_folder: str = ""
    dest_folder: str = ""
    project_number: str = ""
    project_name: str = ""
    pdf: bool = False
    step: bool = False
    dxf: bool = False
    dwg: bool = False
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
        for field_info in fields(cls):
            name = field_info.name
            if name == "_path":
                continue
            if name not in data:
                continue
            cur_val = getattr(inst, name)
            raw = data.get(name)
            if isinstance(cur_val, bool):
                setattr(inst, name, _as_bool(raw))
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

