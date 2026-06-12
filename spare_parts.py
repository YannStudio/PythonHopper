"""Helpers for detecting and grouping BOM spare-part rows."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from itertools import combinations
from typing import Iterable, Mapping, Sequence

from helpers import _to_str


SPARE_PARTS_PRODUCTION_LABEL = "Spare Parts"
SPARE_PARTS_FULL_LIST_KEY = "full"
SPARE_PARTS_FULL_LIST_LABEL = "Volledige lijst"
SPARE_PARTS_UNASSIGNED_KEY = "unassigned"
SPARE_PARTS_UNASSIGNED_LABEL = "Nog toe te wijzen"
SPARE_PARTS_CUSTOM_SOURCE = "custom"

SUPPLIER_FIELDS = ("Supplier", "Leverancier")
SUPPLIER_CODE_FIELDS = (
    "Supplier code",
    "Supplier Code",
    "SupplierCode",
    "Leverancier code",
    "Leverancierscode",
)
MANUFACTURER_FIELDS = ("Manufacturer", "Fabrikant")
MANUFACTURER_CODE_FIELDS = (
    "Manufacturer code",
    "Manufacturer Code",
    "ManufacturerCode",
    "Fabrikant code",
    "Fabrikantcode",
)
QUANTITY_FIELDS = ("Aantal", "Quantity", "Qty", "St.", "Stuks", "St")


@dataclass(slots=True, frozen=True)
class SparePartItem:
    row_index: object
    part_number: str = ""
    description: str = ""
    quantity: str = ""
    material: str = ""
    supplier: str = ""
    supplier_code: str = ""
    manufacturer: str = ""
    manufacturer_code: str = ""

    @property
    def status(self) -> str:
        if not (self.supplier or self.manufacturer):
            return "Mist leverancier/fabrikant"
        if not (self.supplier_code or self.manufacturer_code):
            return "Mist code"
        return "OK"

    @property
    def identity_key(self) -> str:
        key = "sparepart:" + "|".join(
            _identity_part(part)
            for part in (
                self.row_index,
                self.part_number,
                self.description,
                self.manufacturer_code,
                self.supplier_code,
            )
        )
        return key

    @property
    def match_key(self) -> str:
        return _spare_part_match_key(
            self.part_number,
            self.description,
            self.manufacturer_code,
            self.supplier_code,
        )

    def to_order_item(
        self,
        *,
        group_label: str = "",
        group_key: str = "",
    ) -> dict[str, object]:
        label = " - ".join(
            part
            for part in (self.part_number, self.description, self.manufacturer_code)
            if part
        )
        return {
            "PartNumber": self.part_number,
            "Description": self.description,
            "Aantal": self.quantity,
            "Materiaal": self.material,
            "Supplier": self.supplier,
            "Supplier code": self.supplier_code,
            "Manufacturer": self.manufacturer,
            "Manufacturer code": self.manufacturer_code,
            "Bestelgroep": group_label,
            "Bestelgroep key": group_key,
            "Status": self.status,
            "key": self.identity_key,
            "label": label or self.identity_key,
            "quantity": self.quantity,
        }


@dataclass(slots=True)
class SparePartGroup:
    key: str
    label: str
    route_name: str
    route_source: str
    items: list[SparePartItem] = field(default_factory=list)
    is_full_list: bool = False
    default_supplier: str = ""
    default_doc_type: str = "Bestelbon"
    item_group_labels: dict[str, str] = field(default_factory=dict)
    item_group_keys: dict[str, str] = field(default_factory=dict)

    @property
    def display_label(self) -> str:
        return f"{SPARE_PARTS_PRODUCTION_LABEL} - {self.label}"

    @property
    def item_count(self) -> int:
        return len(self.items)

    @property
    def missing_count(self) -> int:
        return sum(1 for item in self.items if item.status != "OK")

    def to_mapping(self) -> dict[str, object]:
        mapped_items = []
        for item in self.items:
            item_key = item.identity_key
            mapped_items.append(
                item.to_order_item(
                    group_label=self.item_group_labels.get(item_key, self.label),
                    group_key=self.item_group_keys.get(item_key, self.key),
                )
            )
        return {
            "key": self.key,
            "label": self.label,
            "display_label": self.display_label,
            "route_name": self.route_name,
            "route_source": self.route_source,
            "is_full_list": self.is_full_list,
            "default_supplier": self.default_supplier,
            "default_doc_type": self.default_doc_type,
            "item_count": self.item_count,
            "missing_count": self.missing_count,
            "items": mapped_items,
        }


def _normalize_label(value: object) -> str:
    text = _to_str(value).strip()
    normalized = unicodedata.normalize("NFKD", text)
    normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    normalized = normalized.casefold()
    normalized = re.sub(r"[^0-9a-z]+", " ", normalized)
    return " ".join(normalized.split())


def _slug(value: object) -> str:
    normalized = _normalize_label(value)
    slug = re.sub(r"[^0-9a-z]+", "-", normalized).strip("-")
    return slug or "onbekend"


def _identity_part(value: object) -> str:
    return _to_str(value).strip().replace("|", "/")


def make_custom_spare_part_group_key(label: object) -> str:
    return f"{SPARE_PARTS_CUSTOM_SOURCE}--{_slug(label)}"


def _spare_part_match_key(
    part_number: object,
    description: object,
    manufacturer_code: object,
    supplier_code: object,
) -> str:
    parts = [
        _normalize_label(part_number),
        _normalize_label(description),
        _normalize_label(manufacturer_code),
        _normalize_label(supplier_code),
    ]
    if not any(parts):
        return ""
    return "|".join(parts)


def spare_part_identity_match_key(identity_key: object) -> str:
    candidates = _spare_part_identity_match_keys(identity_key)
    return sorted(candidates)[0] if candidates else ""


def _spare_part_identity_match_keys(identity_key: object) -> set[str]:
    text = _to_str(identity_key).strip()
    prefix = "sparepart:"
    if not text.startswith(prefix):
        return set()
    parts = text[len(prefix) :].split("|")
    if parts:
        parts = parts[1:]
    if not parts:
        return set()
    if len(parts) >= 4:
        match_key = _spare_part_match_key(parts[0], parts[1], parts[2], parts[3])
        return {match_key} if match_key else set()

    # Legacy exportlogs skipped empty fields in identity_key. Rebuild all ordered
    # position candidates and let the caller require a unique current-row match.
    candidates: set[str] = set()
    for positions in combinations(range(4), len(parts)):
        values = ["", "", "", ""]
        for position, value in zip(positions, parts):
            values[position] = value
        match_key = _spare_part_match_key(values[0], values[1], values[2], values[3])
        if match_key:
            candidates.add(match_key)
    return candidates


def match_spare_part_group_overrides(
    items: Sequence[SparePartItem],
    group_overrides: Mapping[str, str] | None,
) -> dict[str, str]:
    """Return exportlog group overrides matched to the current BOM spare parts."""

    if not items or not group_overrides:
        return {}

    current_keys = {item.identity_key for item in items}
    matched: dict[str, str] = {}
    fallback_overrides: list[tuple[str, str]] = []
    for raw_key, raw_label in group_overrides.items():
        key = _to_str(raw_key).strip()
        label = _to_str(raw_label).strip()
        if not key or not label:
            continue
        if key in current_keys:
            matched[key] = label
        else:
            fallback_overrides.append((key, label))

    if not fallback_overrides:
        return matched

    unique_by_match_key: dict[str, str] = {}
    duplicate_match_keys: set[str] = set()
    for item in items:
        match_key = item.match_key
        if not match_key:
            continue
        if match_key in unique_by_match_key:
            duplicate_match_keys.add(match_key)
            continue
        unique_by_match_key[match_key] = item.identity_key
    for match_key in duplicate_match_keys:
        unique_by_match_key.pop(match_key, None)

    for old_key, label in fallback_overrides:
        candidate_current_keys = {
            unique_by_match_key[match_key]
            for match_key in _spare_part_identity_match_keys(old_key)
            if match_key in unique_by_match_key
        }
        if len(candidate_current_keys) != 1:
            continue
        current_key = next(iter(candidate_current_keys))
        if current_key not in matched:
            matched[current_key] = label
    return matched


def is_spare_parts_production(value: object) -> bool:
    normalized = _normalize_label(value)
    return normalized in {"spare parts", "spare part", "spareparts"}


def _field_value(row: Mapping[str, object], names: Sequence[str]) -> str:
    for name in names:
        value = row.get(name)
        text = _to_str(value).strip()
        if text:
            return text

    normalized_lookup = {
        _normalize_label(key).replace(" ", ""): key for key in row.keys()
    }
    for name in names:
        key = normalized_lookup.get(_normalize_label(name).replace(" ", ""))
        if key is None:
            continue
        text = _to_str(row.get(key)).strip()
        if text:
            return text
    return ""


def _iter_rows(source: object) -> Iterable[tuple[object, Mapping[str, object]]]:
    iterrows = getattr(source, "iterrows", None)
    if callable(iterrows):
        for index, row in iterrows():
            yield index, row
        return
    if isinstance(source, Iterable):
        for index, row in enumerate(source):
            if isinstance(row, Mapping):
                yield index, row


def collect_spare_part_items(source: object) -> list[SparePartItem]:
    items: list[SparePartItem] = []
    for index, row in _iter_rows(source):
        if not is_spare_parts_production(row.get("Production")):
            continue
        item = SparePartItem(
            row_index=index,
            part_number=_field_value(row, ("PartNumber", "Part number", "Artikel nr.")),
            description=_field_value(row, ("Description", "Omschrijving")),
            quantity=_field_value(row, QUANTITY_FIELDS),
            material=_field_value(row, ("Materiaal", "Material")),
            supplier=_field_value(row, SUPPLIER_FIELDS),
            supplier_code=_field_value(row, SUPPLIER_CODE_FIELDS),
            manufacturer=_field_value(row, MANUFACTURER_FIELDS),
            manufacturer_code=_field_value(row, MANUFACTURER_CODE_FIELDS),
        )
        items.append(item)
    return items


def build_spare_part_groups(
    items: Sequence[SparePartItem],
    *,
    include_full_list: bool = True,
    group_overrides: Mapping[str, str] | None = None,
) -> list[SparePartGroup]:
    if not items:
        return []

    groups_by_key: dict[str, SparePartGroup] = {}
    item_group_labels: dict[str, str] = {}
    item_group_keys: dict[str, str] = {}
    group_overrides = group_overrides or {}
    for item in items:
        item_key = item.identity_key
        override_label = _to_str(group_overrides.get(item_key)).strip()
        if override_label:
            if _normalize_label(override_label) == _normalize_label(
                SPARE_PARTS_UNASSIGNED_LABEL
            ):
                source = "unassigned"
                route_name = SPARE_PARTS_UNASSIGNED_LABEL
                key = SPARE_PARTS_UNASSIGNED_KEY
            else:
                source = SPARE_PARTS_CUSTOM_SOURCE
                route_name = override_label
                key = make_custom_spare_part_group_key(route_name)
            default_supplier = ""
        elif item.supplier:
            source = "supplier"
            route_name = item.supplier
            key = f"supplier--{_slug(route_name)}"
            default_supplier = item.supplier
        elif item.manufacturer:
            source = "manufacturer"
            route_name = item.manufacturer
            key = f"manufacturer--{_slug(route_name)}"
            default_supplier = ""
        else:
            source = "unassigned"
            route_name = SPARE_PARTS_UNASSIGNED_LABEL
            key = SPARE_PARTS_UNASSIGNED_KEY
            default_supplier = ""

        item_group_labels[item_key] = route_name
        item_group_keys[item_key] = key
        group = groups_by_key.get(key)
        if group is None:
            group = SparePartGroup(
                key=key,
                label=route_name,
                route_name=route_name,
                route_source=source,
                default_supplier=default_supplier,
            )
            groups_by_key[key] = group
        group.items.append(item)

    groups = sorted(
        groups_by_key.values(),
        key=lambda group: (
            {
                "supplier": 0,
                SPARE_PARTS_CUSTOM_SOURCE: 1,
                "manufacturer": 2,
                "unassigned": 3,
            }.get(group.route_source, 4),
            group.label.casefold(),
        ),
    )
    if include_full_list:
        groups.insert(
            0,
            SparePartGroup(
                key=SPARE_PARTS_FULL_LIST_KEY,
                label=SPARE_PARTS_FULL_LIST_LABEL,
                route_name=SPARE_PARTS_FULL_LIST_LABEL,
                route_source="full",
                items=list(items),
                is_full_list=True,
                default_doc_type="Standaard bon",
                item_group_labels=dict(item_group_labels),
                item_group_keys=dict(item_group_keys),
            ),
        )
    return groups


def summarize_spare_part_warnings(groups: Sequence[SparePartGroup]) -> list[str]:
    if not groups:
        return []

    full_group = next((group for group in groups if group.is_full_list), None)
    items = list(full_group.items if full_group is not None else [])
    if not items:
        seen: set[str] = set()
        for group in groups:
            for item in group.items:
                key = item.identity_key
                if key in seen:
                    continue
                seen.add(key)
                items.append(item)

    warnings: list[str] = []
    without_route = sum(1 for item in items if not (item.supplier or item.manufacturer))
    without_code = sum(
        1 for item in items if not (item.supplier_code or item.manufacturer_code)
    )
    unassigned = next(
        (group.item_count for group in groups if group.key == SPARE_PARTS_UNASSIGNED_KEY),
        0,
    )
    groups_without_supplier = sum(
        1
        for group in groups
        if not group.is_full_list
        and group.route_source != "supplier"
        and not group.default_supplier
    )

    if unassigned:
        warnings.append(f"{unassigned} nog toe te wijzen")
    if without_route:
        warnings.append(f"{without_route} zonder leverancier/fabrikant")
    if without_code:
        warnings.append(f"{without_code} zonder supplier/fabrikantcode")
    if groups_without_supplier:
        warnings.append(f"{groups_without_supplier} groep(en) zonder standaardleverancier")
    return warnings


def collect_spare_part_groups(
    source: object,
    *,
    group_overrides: Mapping[str, str] | None = None,
) -> list[SparePartGroup]:
    return build_spare_part_groups(
        collect_spare_part_items(source),
        group_overrides=group_overrides,
    )
