"""Preset storage for PDF work dossier ordering."""

from __future__ import annotations

import json
import os
import re
from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional

from app_paths import data_file
from data_storage import write_json_with_backup

PDF_WORKDOSSIER_PRESETS_DB_FILE = data_file("pdf_workdossier_presets.json")


def _as_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _as_bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        text = value.strip().lower()
        if text in {"1", "true", "yes", "y", "ja", "aan", "on"}:
            return True
        if text in {"0", "false", "no", "n", "nee", "uit", "off"}:
            return False
    return default


def _normalize_text(value: Any) -> str:
    return _as_str(value).strip()


def _normalize_priority(value: Any, default: int = 100) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return default


def _normalize_identifiers(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, str):
        parts = re.split(r"[\r\n,;]+", value)
    elif isinstance(value, (list, tuple, set)):
        parts = list(value)
    else:
        parts = [_as_str(value)]

    cleaned: List[str] = []
    seen: set[str] = set()
    for raw in parts:
        text = _normalize_text(raw)
        if not text:
            continue
        key = text.casefold()
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(text)
    return cleaned


@dataclass
class PdfWorkDossierSection:
    name: str
    identifiers: List[str] = field(default_factory=list)
    include_bom_pdf: bool = False
    enabled: bool = True

    @classmethod
    def from_any(cls, data: Any) -> "PdfWorkDossierSection":
        if isinstance(data, PdfWorkDossierSection):
            return cls(
                name=data.name,
                identifiers=list(data.identifiers),
                include_bom_pdf=bool(data.include_bom_pdf),
                enabled=bool(data.enabled),
            )
        if not isinstance(data, dict):
            raise ValueError("PDF section must be a mapping")

        name = _normalize_text(data.get("name"))
        if not name:
            raise ValueError("PDF section name is required")

        return cls(
            name=name,
            identifiers=_normalize_identifiers(data.get("identifiers")),
            include_bom_pdf=_as_bool(data.get("include_bom_pdf"), False),
            enabled=_as_bool(data.get("enabled"), True),
        )


@dataclass
class PdfWorkDossierPreset:
    name: str
    enabled: bool = True
    priority: int = 100
    sections: List[PdfWorkDossierSection] = field(default_factory=list)
    include_unmatched: bool = True
    unmatched_section_name: str = "Overige"

    @classmethod
    def from_any(cls, data: Any) -> "PdfWorkDossierPreset":
        if isinstance(data, PdfWorkDossierPreset):
            return cls(
                name=data.name,
                enabled=bool(data.enabled),
                priority=int(data.priority),
                sections=[PdfWorkDossierSection.from_any(section) for section in data.sections],
                include_unmatched=bool(data.include_unmatched),
                unmatched_section_name=data.unmatched_section_name or "Overige",
            )
        if not isinstance(data, dict):
            raise ValueError("PDF preset must be a mapping")

        name = _normalize_text(data.get("name"))
        if not name:
            raise ValueError("PDF preset name is required")

        raw_sections = data.get("sections", [])
        sections: List[PdfWorkDossierSection] = []
        if isinstance(raw_sections, list):
            for raw_section in raw_sections:
                try:
                    sections.append(PdfWorkDossierSection.from_any(raw_section))
                except Exception:
                    continue

        return cls(
            name=name,
            enabled=_as_bool(data.get("enabled"), True),
            priority=_normalize_priority(data.get("priority"), 100),
            sections=sections,
            include_unmatched=_as_bool(data.get("include_unmatched"), True),
            unmatched_section_name=_normalize_text(data.get("unmatched_section_name"))
            or "Overige",
        )


def default_pdf_workdossier_preset() -> PdfWorkDossierPreset:
    """Return the built-in production order preset for work dossiers."""

    return PdfWorkDossierPreset(
        name="Werkdossier standaard",
        sections=[
            PdfWorkDossierSection("Hoofdassembly", include_bom_pdf=True),
            PdfWorkDossierSection(
                "Assembly tekeningen",
                identifiers=["Assembly", "Assemblage", "Montage"],
            ),
            PdfWorkDossierSection(
                "Laserwerk",
                identifiers=["Laserwerk", "Laser cutting", "Lasersnijden", "Laser"],
            ),
            PdfWorkDossierSection(
                "Tube laserwerk",
                identifiers=["Tube laserwerk", "Tube laser", "Buislaser"],
            ),
            PdfWorkDossierSection(
                "Spare parts",
                identifiers=["Spare parts", "Reserveonderdelen", "Onderdelen"],
            ),
        ],
    )


class PdfWorkDossierPresetsDB:
    def __init__(self, presets: Optional[List[PdfWorkDossierPreset]] = None):
        self.presets: List[PdfWorkDossierPreset] = presets or []

    @staticmethod
    def load(path: str = PDF_WORKDOSSIER_PRESETS_DB_FILE) -> "PdfWorkDossierPresetsDB":
        if not os.path.exists(path):
            return PdfWorkDossierPresetsDB()
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            return PdfWorkDossierPresetsDB()

        raw_presets = data if isinstance(data, list) else data.get("presets", [])
        presets: List[PdfWorkDossierPreset] = []
        if isinstance(raw_presets, list):
            for raw_preset in raw_presets:
                try:
                    presets.append(PdfWorkDossierPreset.from_any(raw_preset))
                except Exception:
                    continue
        return PdfWorkDossierPresetsDB(presets)

    def save(self, path: str = PDF_WORKDOSSIER_PRESETS_DB_FILE) -> None:
        write_json_with_backup(path, self.to_dict())

    def to_dict(self) -> Dict[str, Any]:
        return {"presets": [asdict(preset) for preset in self.presets]}

    def presets_sorted(self) -> List[PdfWorkDossierPreset]:
        return sorted(
            [preset for preset in self.presets if preset.enabled],
            key=lambda preset: (-int(preset.priority), preset.name.casefold()),
        )

    def get(self, name: str) -> Optional[PdfWorkDossierPreset]:
        key = _normalize_text(name).casefold()
        for preset in self.presets:
            if preset.name.casefold() == key:
                return preset
        return None

    def upsert(
        self,
        preset: PdfWorkDossierPreset,
        old_name: Optional[str] = None,
    ) -> None:
        cloned = PdfWorkDossierPreset.from_any(preset)
        key = _normalize_text(old_name or cloned.name).casefold()
        for index, existing in enumerate(self.presets):
            if existing.name.casefold() == key:
                self.presets[index] = cloned
                return
        self.presets.append(cloned)
