import pandas as pd
from openpyxl import load_workbook

from models import Supplier
from orders import (
    copy_per_production_and_orders,
    make_spare_part_default_key,
    make_spare_part_selection_key,
    parse_selection_key,
)
from spare_parts import (
    SPARE_PARTS_CUSTOM_SOURCE,
    SPARE_PARTS_FULL_LIST_KEY,
    SPARE_PARTS_UNASSIGNED_KEY,
    build_spare_part_groups,
    collect_spare_part_groups,
    collect_spare_part_items,
    is_spare_parts_production,
    make_custom_spare_part_group_key,
)
from suppliers_db import SuppliersDB


def test_spare_parts_detection_is_tolerant():
    assert is_spare_parts_production("Spare Parts") is True
    assert is_spare_parts_production(" spare-parts ") is True
    assert is_spare_parts_production("Laser") is False


def test_collect_spare_parts_reads_supplier_and_manufacturer_fields():
    df = pd.DataFrame(
        [
            {
                "PartNumber": "458",
                "Description": "Hangkastje",
                "Production": "Spare Parts",
                "Aantal": 2,
                "Supplier": "Herbaroof",
                "Supplier code": "ND SM-25",
                "Manufacturer": "Herbaroof",
                "Manufacturer code": "ND SM-25",
            },
            {
                "PartNumber": "999",
                "Production": "Laser",
                "Supplier": "Ignored",
            },
        ]
    )

    items = collect_spare_part_items(df)

    assert len(items) == 1
    assert items[0].part_number == "458"
    assert items[0].supplier == "Herbaroof"
    assert items[0].manufacturer_code == "ND SM-25"
    assert items[0].status == "OK"


def test_spare_part_groups_prefer_supplier_and_keep_full_list():
    items = collect_spare_part_items(
        [
            {
                "PartNumber": "A",
                "Production": "Spare Parts",
                "Supplier": "Herbaroof",
                "Supplier code": "H-1",
                "Manufacturer": "Maker A",
                "Manufacturer code": "M-1",
            },
            {
                "PartNumber": "B",
                "Production": "Spare Parts",
                "Manufacturer": "Maker B",
                "Manufacturer code": "M-2",
            },
        ]
    )

    groups = build_spare_part_groups(items)

    assert groups[0].key == SPARE_PARTS_FULL_LIST_KEY
    assert groups[0].is_full_list is True
    assert groups[0].default_doc_type == "Standaard bon"
    assert groups[0].item_count == 2
    route_groups = {group.key: group for group in groups[1:]}
    assert route_groups["supplier--herbaroof"].default_supplier == "Herbaroof"
    assert route_groups["manufacturer--maker-b"].default_supplier == ""


def test_spare_part_groups_track_unassigned_and_missing_codes():
    items = collect_spare_part_items(
        [
            {
                "PartNumber": "A",
                "Production": "Spare Parts",
                "Supplier": "Electro",
            },
            {"PartNumber": "B", "Production": "Spare Parts"},
        ]
    )

    groups = build_spare_part_groups(items)
    by_key = {group.key: group for group in groups}

    assert by_key["supplier--electro"].missing_count == 1
    assert by_key[SPARE_PARTS_UNASSIGNED_KEY].label == "Nog toe te wijzen"
    assert by_key[SPARE_PARTS_UNASSIGNED_KEY].missing_count == 1


def test_spare_part_groups_accept_manual_overrides_without_mutating_bom():
    df = pd.DataFrame(
        [
            {
                "PartNumber": "A",
                "Production": "Spare Parts",
                "Supplier": "Herbaroof",
                "Supplier code": "H-1",
            },
            {
                "PartNumber": "B",
                "Production": "Spare Parts",
                "Manufacturer": "Maker B",
                "Manufacturer code": "M-2",
            },
        ]
    )
    items = collect_spare_part_items(df)
    custom_key = make_custom_spare_part_group_key("Electro")
    overrides = {
        items[0].identity_key: "Electro",
        items[1].identity_key: "Nog toe te wijzen",
    }

    groups = build_spare_part_groups(items, group_overrides=overrides)
    by_key = {group.key: group for group in groups}

    assert by_key[custom_key].route_source == SPARE_PARTS_CUSTOM_SOURCE
    assert by_key[custom_key].label == "Electro"
    assert by_key[custom_key].items == [items[0]]
    assert by_key[SPARE_PARTS_UNASSIGNED_KEY].items == [items[1]]
    full_items = by_key[SPARE_PARTS_FULL_LIST_KEY].to_mapping()["items"]
    assert [item["Bestelgroep"] for item in full_items] == [
        "Electro",
        "Nog toe te wijzen",
    ]
    assert list(df["Production"]) == ["Spare Parts", "Spare Parts"]


def test_spare_part_selection_key_roundtrip():
    key = make_spare_part_selection_key("supplier--herbaroof")

    assert parse_selection_key(key) == ("sparepart", "supplier--herbaroof")
    assert make_spare_part_default_key("supplier--herbaroof").endswith(
        "supplier--herbaroof"
    )


def test_spare_part_groups_export_full_list_and_supplier_order(tmp_path):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()
    (src / "PN1.pdf").write_text("one", encoding="utf-8")
    (src / "PN2.pdf").write_text("two", encoding="utf-8")
    db = SuppliersDB(storage_path=tmp_path / "suppliers_db.json")
    db.upsert(Supplier.from_any({"supplier": "Herbaroof"}))
    bom_df = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Description": "Spare rail",
                "Production": "Spare Parts",
                "Aantal": 2,
                "Supplier": "Herbaroof",
                "Supplier code": "ND SM-25",
                "Manufacturer": "Herbaroof",
                "Manufacturer code": "MF-25",
            },
            {
                "PartNumber": "PN2",
                "Description": "Loose cover",
                "Production": "Spare Parts",
                "Aantal": 1,
                "Supplier": "Herbaroof",
                "Supplier code": "ND SM-30",
                "Manufacturer": "Maker",
                "Manufacturer code": "MF-30",
            },
        ]
    )
    groups = [group.to_mapping() for group in collect_spare_part_groups(bom_df)]

    copied, chosen = copy_per_production_and_orders(
        str(src),
        str(dest),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        {},
        True,
        spare_part_groups=groups,
        spare_part_override_map={"supplier--herbaroof": "Herbaroof"},
    )

    assert copied == 2
    assert chosen[make_spare_part_selection_key(SPARE_PARTS_FULL_LIST_KEY)] == ""
    assert chosen[make_spare_part_selection_key("supplier--herbaroof")] == "Herbaroof"
    assert (
        db.get_default(make_spare_part_default_key("supplier--herbaroof"))
        == "Herbaroof"
    )
    full_docs = list((dest / "Spare Parts").glob("Standaard*Spare*Parts*.xlsx"))
    supplier_docs = list(
        (dest / "Spare Parts-Herbaroof").glob("Bestelbon*Spare*Parts-Herbaroof*.xlsx")
    )
    assert full_docs
    assert supplier_docs

    workbook = load_workbook(full_docs[0])
    values = [
        value
        for row in workbook.active.iter_rows(values_only=True)
        for value in row
        if value is not None
    ]
    assert "Supplier code" in values
    assert "Fabrikant code" in values
    assert "Bestelgroep" in values
    assert "Status" in values
    assert "ND SM-25" in values
    assert "MF-30" in values


def test_spare_part_manual_group_is_used_for_export_without_mutating_bom(tmp_path):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()
    (src / "PN1.pdf").write_text("one", encoding="utf-8")
    db = SuppliersDB(storage_path=tmp_path / "suppliers_db.json")
    db.upsert(Supplier.from_any({"supplier": "RS Components"}))
    bom_df = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Description": "Spare cable",
                "Production": "Spare Parts",
                "Aantal": 3,
                "Supplier": "RS Components",
                "Supplier code": "RS-1",
                "Manufacturer": "Maker",
                "Manufacturer code": "M-1",
            }
        ]
    )
    items = collect_spare_part_items(bom_df)
    custom_group_key = make_custom_spare_part_group_key("Electro")
    groups = [
        group.to_mapping()
        for group in collect_spare_part_groups(
            bom_df,
            group_overrides={items[0].identity_key: "Electro"},
        )
    ]

    _copied, chosen = copy_per_production_and_orders(
        str(src),
        str(dest),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        {},
        True,
        spare_part_groups=groups,
        spare_part_override_map={custom_group_key: "RS Components"},
    )

    assert chosen[make_spare_part_selection_key(custom_group_key)] == "RS Components"
    assert list(bom_df["Production"]) == ["Spare Parts"]
    custom_docs = list(
        (dest / "Spare Parts-Electro").glob("Bestelbon*Spare*Parts-Electro*.xlsx")
    )
    assert custom_docs
    workbook = load_workbook(custom_docs[0])
    values = [
        value
        for row in workbook.active.iter_rows(values_only=True)
        for value in row
        if value is not None
    ]
    assert "Bestelgroep" in values
    assert "Electro" in values
