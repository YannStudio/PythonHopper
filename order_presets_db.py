import json
import os
import re
from dataclasses import asdict, dataclass, field
from typing import Any, Dict, List, Optional

from app_paths import data_file
from data_storage import write_json_with_backup

ORDER_PRESETS_DB_FILE = data_file("order_presets.json")
ORDER_PRESET_KINDS = {"production", "finish", "opticutter"}
ORDER_PRESET_DOC_TYPES = {"", "Geen", "Bestelbon", "Standaard bon", "Offerteaanvraag"}


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


def _normalize_key(value: Any) -> str:
    return _normalize_text(value).casefold()


def _normalize_selection_kind(value: Any) -> str:
    text = _normalize_key(value)
    if text in ORDER_PRESET_KINDS:
        return text
    return "production"


def _normalize_doc_type(value: Any) -> str:
    text = _normalize_text(value)
    if text in ORDER_PRESET_DOC_TYPES:
        return text
    return ""


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
        key = _normalize_key(text)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(text)
    return cleaned


def _normalize_optional_bool(value: Any) -> Optional[bool]:
    if value in (None, ""):
        return None
    if isinstance(value, str):
        text = value.strip().lower()
        if text in {"ongewijzigd", "leave", "default", "none"}:
            return None
    return _as_bool(value)


@dataclass(frozen=True)
class OrderPresetContext:
    client: str = ""
    selection_kind: str = "production"
    identifier: str = ""

    @classmethod
    def from_any(cls, data: Any) -> "OrderPresetContext":
        if isinstance(data, OrderPresetContext):
            return data
        if not isinstance(data, dict):
            raise ValueError("preset context must be a mapping")
        return cls(
            client=_normalize_text(data.get("client")),
            selection_kind=_normalize_selection_kind(data.get("selection_kind")),
            identifier=_normalize_text(data.get("identifier")),
        )


@dataclass
class OrderPresetEvaluation:
    supplier: str = ""
    doc_type: str = ""
    delivery: str = ""
    remark: str = ""
    en1090: Optional[bool] = None
    matched_rule_names: List[str] = field(default_factory=list)
    applied_rule_names: List[str] = field(default_factory=list)

    def has_changes(self) -> bool:
        return bool(
            self.supplier
            or self.doc_type
            or self.delivery
            or self.remark
            or self.en1090 is not None
        )


@dataclass
class OrderPresetRule:
    name: str
    enabled: bool = True
    priority: int = 100
    auto_apply: bool = True
    client: str = ""
    selection_kind: str = "production"
    identifiers: List[str] = field(default_factory=list)
    supplier: str = ""
    doc_type: str = ""
    delivery: str = ""
    remark: str = ""
    en1090: Optional[bool] = None

    @classmethod
    def from_any(cls, data: Any) -> "OrderPresetRule":
        if isinstance(data, OrderPresetRule):
            return cls(
                name=data.name,
                enabled=bool(data.enabled),
                priority=int(data.priority),
                auto_apply=bool(data.auto_apply),
                client=data.client,
                selection_kind=_normalize_selection_kind(data.selection_kind),
                identifiers=list(data.identifiers),
                supplier=data.supplier,
                doc_type=_normalize_doc_type(data.doc_type),
                delivery=data.delivery,
                remark=data.remark,
                en1090=data.en1090,
            )
        if not isinstance(data, dict):
            raise ValueError("preset rule must be a mapping")

        match_data = data.get("match", {})
        apply_data = data.get("apply", {})
        if not isinstance(match_data, dict):
            match_data = {}
        if not isinstance(apply_data, dict):
            apply_data = {}

        name = _normalize_text(data.get("name"))
        if not name:
            raise ValueError("preset rule name is required")

        selection_kind = _normalize_selection_kind(
            data.get("selection_kind") or match_data.get("selection_kind")
        )
        identifiers = _normalize_identifiers(
            data.get("identifiers", match_data.get("identifiers"))
        )

        return cls(
            name=name,
            enabled=_as_bool(data.get("enabled"), True),
            priority=_normalize_priority(data.get("priority"), 100),
            auto_apply=_as_bool(data.get("auto_apply"), True),
            client=_normalize_text(data.get("client", match_data.get("client"))),
            selection_kind=selection_kind,
            identifiers=identifiers,
            supplier=_normalize_text(data.get("supplier", apply_data.get("supplier"))),
            doc_type=_normalize_doc_type(data.get("doc_type", apply_data.get("doc_type"))),
            delivery=_normalize_text(data.get("delivery", apply_data.get("delivery"))),
            remark=_normalize_text(data.get("remark", apply_data.get("remark"))),
            en1090=_normalize_optional_bool(data.get("en1090", apply_data.get("en1090"))),
        )

    def specificity_score(self) -> int:
        score = 0
        if self.client:
            score += 10
        if self.identifiers:
            score += 5
            if len(self.identifiers) == 1:
                score += 1
        return score

    def matches(self, context: OrderPresetContext) -> bool:
        if not self.enabled:
            return False
        if _normalize_selection_kind(context.selection_kind) != self.selection_kind:
            return False
        if self.client and _normalize_key(self.client) != _normalize_key(context.client):
            return False
        if self.identifiers:
            identifier_key = _normalize_key(context.identifier)
            allowed = {_normalize_key(item) for item in self.identifiers}
            if identifier_key not in allowed:
                return False
        return True

    def action_summary(self) -> str:
        parts: List[str] = []
        if self.supplier:
            parts.append(f"leverancier={self.supplier}")
        if self.doc_type:
            parts.append(f"document={self.doc_type}")
        if self.delivery:
            parts.append(f"leveradres={self.delivery}")
        if self.remark:
            parts.append(f"opmerking={self.remark}")
        if self.en1090 is not None:
            parts.append(f"EN1090={'aan' if self.en1090 else 'uit'}")
        return ", ".join(parts)

    def selection_summary(self) -> str:
        if not self.identifiers:
            return "(alle)"
        return ", ".join(self.identifiers)


class OrderPresetsDB:
    def __init__(self, rules: Optional[List[OrderPresetRule]] = None):
        self.rules: List[OrderPresetRule] = rules or []

    @staticmethod
    def load(path: str = ORDER_PRESETS_DB_FILE) -> "OrderPresetsDB":
        if not os.path.exists(path):
            return OrderPresetsDB()
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            return OrderPresetsDB()

        if isinstance(data, list):
            raw_rules = data
        else:
            raw_rules = data.get("rules", [])

        rules: List[OrderPresetRule] = []
        for raw_rule in raw_rules:
            try:
                rules.append(OrderPresetRule.from_any(raw_rule))
            except Exception:
                continue
        return OrderPresetsDB(rules)

    def save(self, path: str = ORDER_PRESETS_DB_FILE) -> None:
        write_json_with_backup(path, self.to_dict())

    def to_dict(self) -> Dict[str, Any]:
        return {"rules": [asdict(rule) for rule in self.rules]}

    def rules_sorted(self) -> List[OrderPresetRule]:
        return sorted(
            self.rules,
            key=lambda rule: (
                -int(rule.priority),
                -rule.specificity_score(),
                _normalize_key(rule.name),
            ),
        )

    def _idx_by_name(self, name: str) -> int:
        key = _normalize_key(name)
        for index, rule in enumerate(self.rules):
            if _normalize_key(rule.name) == key:
                return index
        return -1

    def get(self, name: str) -> Optional[OrderPresetRule]:
        idx = self._idx_by_name(name)
        return self.rules[idx] if idx >= 0 else None

    def upsert(self, rule: OrderPresetRule, old_name: Optional[str] = None) -> None:
        key = old_name or rule.name
        idx = self._idx_by_name(key)
        cloned = OrderPresetRule.from_any(rule)
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

    def toggle_enabled(self, name: str) -> bool:
        idx = self._idx_by_name(name)
        if idx < 0:
            return False
        current = self.rules[idx]
        self.rules[idx] = OrderPresetRule.from_any(
            {**asdict(current), "enabled": not current.enabled}
        )
        return True

    def evaluate(
        self,
        context: OrderPresetContext | Dict[str, Any],
        *,
        auto_apply_only: bool = False,
    ) -> OrderPresetEvaluation:
        context_obj = OrderPresetContext.from_any(context)
        evaluation = OrderPresetEvaluation()
        applied_fields: set[str] = set()

        for rule in self.rules_sorted():
            if auto_apply_only and not rule.auto_apply:
                continue
            if not rule.matches(context_obj):
                continue

            evaluation.matched_rule_names.append(rule.name)
            rule_contributed = False

            if rule.supplier and "supplier" not in applied_fields:
                evaluation.supplier = rule.supplier
                applied_fields.add("supplier")
                rule_contributed = True
            if rule.doc_type and "doc_type" not in applied_fields:
                evaluation.doc_type = rule.doc_type
                applied_fields.add("doc_type")
                rule_contributed = True
            if rule.delivery and "delivery" not in applied_fields:
                evaluation.delivery = rule.delivery
                applied_fields.add("delivery")
                rule_contributed = True
            if rule.remark and "remark" not in applied_fields:
                evaluation.remark = rule.remark
                applied_fields.add("remark")
                rule_contributed = True
            if rule.en1090 is not None and "en1090" not in applied_fields:
                evaluation.en1090 = bool(rule.en1090)
                applied_fields.add("en1090")
                rule_contributed = True

            if rule_contributed:
                evaluation.applied_rule_names.append(rule.name)

            if {
                "supplier",
                "doc_type",
                "delivery",
                "remark",
                "en1090",
            }.issubset(applied_fields):
                break

        return evaluation
