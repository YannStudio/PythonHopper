"""Helpers for detecting and grouping BOM spare-part rows."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
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
            _to_str(part).strip()
            for part in (
                self.row_index,
                self.part_number,
                self.description,
                self.manufacturer_code,
                self.supplier_code,
            )
            if _to_str(part).strip()
        )
        return key

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


def make_custom_spare_part_group_key(label: object) -> str:
    return f"{SPARE_PARTS_CUSTOM_SOURCE}--{_slug(label)}"


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


def collect_spare_part_groups(
    source: object,
    *,
    group_overrides: Mapping[str, str] | None = None,
) -> list[SparePartGroup]:
    return build_spare_part_groups(
        collect_spare_part_items(source),
        group_overrides=group_overrides,
    )
