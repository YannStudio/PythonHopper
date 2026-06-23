"""Preset rules for assigning spare-part rows to order groups."""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from typing import Any, Mapping, Sequence

from app_paths import data_file
from data_storage import write_json_with_backup
from helpers import _to_str


SPARE_PART_PRESETS_DB_FILE = data_file("spare_part_presets.json")
SPARE_PART_PRESET_FIELDS = {
    "supplier": "supplier",
    "supplier code": "supplier_code",
    "supplier_code": "supplier_code",
    "manufacturer": "manufacturer",
    "fabrikant": "manufacturer",
    "manufacturer code": "manufacturer_code",
    "manufacturer_code": "manufacturer_code",
    "fabrikant code": "manufacturer_code",
}
SPARE_PART_PRESET_MATCH_TYPES = {"exact", "contains", "startswith"}


def _clean(value: object) -> str:
    return _to_str(value).strip()


def _key(value: object) -> str:
    return _clean(value).casefold()


def _as_bool(value: object, default: bool = True) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        text = value.strip().casefold()
        if text in {"1", "true", "yes", "y", "ja", "aan", "on"}:
            return True
        if text in {"0", "false", "no", "n", "nee", "uit", "off"}:
            return False
    return default


def _as_int(value: object, default: int = 100) -> int:
    try:
        return int(_clean(value))
    except Exception:
        return default


def normalize_spare_part_preset_field(value: object) -> str:
    return SPARE_PART_PRESET_FIELDS.get(_key(value), "manufacturer")


def normalize_spare_part_match_type(value: object) -> str:
    text = _key(value)
    if text in {"begint met", "starts with", "startwith", "prefix"}:
        return "startswith"
    if text in {"bevat", "contains", "contain"}:
        return "contains"
    if text in SPARE_PART_PRESET_MATCH_TYPES:
        return text
    return "exact"


@dataclass
class SparePartPresetRule:
    name: str
    enabled: bool = True
    priority: int = 100
    match_field: str = "manufacturer"
    match_type: str = "exact"
    pattern: str = ""
    target_group: str = ""

    @classmethod
    def from_any(cls, data: Any) -> "SparePartPresetRule":
        if isinstance(data, SparePartPresetRule):
            return cls(
                name=data.name,
                enabled=bool(data.enabled),
                priority=int(data.priority),
                match_field=normalize_spare_part_preset_field(data.match_field),
                match_type=normalize_spare_part_match_type(data.match_type),
                pattern=data.pattern,
                target_group=data.target_group,
            )
        if not isinstance(data, Mapping):
            raise ValueError("spare-part preset rule must be a mapping")

        match_data = data.get("match", {})
        apply_data = data.get("apply", {})
        if not isinstance(match_data, Mapping):
            match_data = {}
        if not isinstance(apply_data, Mapping):
            apply_data = {}

        name = _clean(data.get("name"))
        if not name:
            raise ValueError("spare-part preset rule name is required")
        target_group = _clean(data.get("target_group", apply_data.get("target_group")))
        if not target_group:
            raise ValueError("spare-part preset target_group is required")

        return cls(
            name=name,
            enabled=_as_bool(data.get("enabled"), True),
            priority=_as_int(data.get("priority"), 100),
            match_field=normalize_spare_part_preset_field(
                data.get("match_field", match_data.get("field"))
            ),
            match_type=normalize_spare_part_match_type(
                data.get("match_type", match_data.get("type"))
            ),
            pattern=_clean(data.get("pattern", match_data.get("pattern"))),
            target_group=target_group,
        )

    def field_value(self, item: object) -> str:
        return _clean(getattr(item, self.match_field, ""))

    def matches(self, item: object) -> bool:
        if not self.enabled or not self.pattern:
            return False
        value_key = _key(self.field_value(item))
        pattern_key = _key(self.pattern)
        if self.match_type == "contains":
            return pattern_key in value_key
        if self.match_type == "startswith":
            return value_key.startswith(pattern_key)
        return value_key == pattern_key

    def summary(self) -> str:
        field_label = {
            "supplier": "Supplier",
            "supplier_code": "Supplier code",
            "manufacturer": "Manufacturer",
            "manufacturer_code": "Manufacturer code",
        }.get(self.match_field, self.match_field)
        match_label = {
            "exact": "is",
            "contains": "bevat",
            "startswith": "begint met",
        }.get(self.match_type, self.match_type)
        return f"{field_label} {match_label} {self.pattern} -> {self.target_group}"


def example_spare_part_preset_rules() -> list[SparePartPresetRule]:
    """Return disabled example rules that demonstrate common spare-part presets."""

    return [
        SparePartPresetRule(
            name="Voorbeeld - RS supplier naar Electro",
            enabled=False,
            priority=90,
            match_field="supplier",
            match_type="contains",
            pattern="RS",
            target_group="Electro",
        ),
        SparePartPresetRule(
            name="Voorbeeld - Festo fabrikant naar Pneumatica",
            enabled=False,
            priority=80,
            match_field="manufacturer",
            match_type="exact",
            pattern="Festo",
            target_group="Pneumatica",
        ),
        SparePartPresetRule(
            name="Voorbeeld - SM fabrikantcode naar Mechanisch",
            enabled=False,
            priority=70,
            match_field="manufacturer_code",
            match_type="startswith",
            pattern="SM-",
            target_group="Mechanisch",
        ),
        SparePartPresetRule(
            name="Voorbeeld - ND suppliercode naar Herbaroof",
            enabled=False,
            priority=60,
            match_field="supplier_code",
            match_type="startswith",
            pattern="ND",
            target_group="Herbaroof",
        ),
    ]


class SparePartPresetsDB:
    def __init__(self, rules: Sequence[SparePartPresetRule] | None = None):
        self.rules: list[SparePartPresetRule] = list(rules or [])

    @staticmethod
    def load(path: str = SPARE_PART_PRESETS_DB_FILE) -> "SparePartPresetsDB":
        if not os.path.exists(path):
            return SparePartPresetsDB()
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            return SparePartPresetsDB()
        raw_rules = data if isinstance(data, list) else data.get("rules", [])
        rules: list[SparePartPresetRule] = []
        for raw_rule in raw_rules:
            try:
                rules.append(SparePartPresetRule.from_any(raw_rule))
            except Exception:
                continue
        return SparePartPresetsDB(rules)

    def save(self, path: str = SPARE_PART_PRESETS_DB_FILE) -> None:
        write_json_with_backup(path, self.to_dict())

    def to_dict(self) -> dict[str, Any]:
        return {"rules": [asdict(rule) for rule in self.rules]}

    def rules_sorted(self) -> list[SparePartPresetRule]:
        return sorted(
            self.rules,
            key=lambda rule: (-int(rule.priority), _key(rule.name)),
        )

    def _idx_by_name(self, name: str) -> int:
        name_key = _key(name)
        for index, rule in enumerate(self.rules):
            if _key(rule.name) == name_key:
                return index
        return -1

    def get(self, name: str) -> SparePartPresetRule | None:
        idx = self._idx_by_name(name)
        return self.rules[idx] if idx >= 0 else None

    def upsert(
        self,
        rule: SparePartPresetRule,
        old_name: str | None = None,
    ) -> None:
        idx = self._idx_by_name(old_name or rule.name)
        cloned = SparePartPresetRule.from_any(rule)
        if idx >= 0:
            self.rules[idx] = cloned
        else:
            self.rules.append(cloned)

    def remove(self, name: str) -> bool:
        idx = self._idx_by_name(name)
        if idx < 0:
            return False
        self.rules.pop(idx)
        return True

    def overrides_for_items(self, items: Sequence[object]) -> dict[str, str]:
        overrides: dict[str, str] = {}
        for item in items:
            identity_key = _clean(getattr(item, "identity_key", ""))
            if not identity_key:
                continue
            for rule in self.rules_sorted():
                if rule.matches(item):
                    overrides[identity_key] = rule.target_group
                    break
        return overrides
