from spare_part_presets import (
    SparePartPresetRule,
    SparePartPresetsDB,
    normalize_spare_part_match_type,
    normalize_spare_part_preset_field,
)
from spare_parts import collect_spare_part_items


def test_spare_part_preset_rule_matches_supported_fields_and_types():
    item = collect_spare_part_items(
        [
            {
                "PartNumber": "A",
                "Production": "Spare Parts",
                "Supplier": "RS Components",
                "Supplier code": "RS-204",
                "Manufacturer": "Festo",
                "Manufacturer code": "XYZ-100",
            }
        ]
    )[0]

    assert SparePartPresetRule(
        name="Supplier",
        match_field="supplier",
        match_type="exact",
        pattern="rs components",
        target_group="Electro",
    ).matches(item)
    assert SparePartPresetRule(
        name="Manufacturer code prefix",
        match_field="manufacturer_code",
        match_type="startswith",
        pattern="XYZ",
        target_group="Pneumatica",
    ).matches(item)
    assert SparePartPresetRule(
        name="Supplier code contains",
        match_field="supplier_code",
        match_type="contains",
        pattern="20",
        target_group="Electro",
    ).matches(item)


def test_spare_part_presets_db_builds_overrides_by_priority():
    items = collect_spare_part_items(
        [
            {
                "PartNumber": "A",
                "Production": "Spare Parts",
                "Supplier": "RS Components",
                "Supplier code": "RS-204",
                "Manufacturer": "Festo",
            }
        ]
    )
    db = SparePartPresetsDB(
        [
            SparePartPresetRule(
                name="General supplier",
                priority=10,
                match_field="supplier",
                match_type="exact",
                pattern="RS Components",
                target_group="Electro",
            ),
            SparePartPresetRule(
                name="Specific code",
                priority=50,
                match_field="supplier_code",
                match_type="startswith",
                pattern="RS-",
                target_group="RS",
            ),
        ]
    )

    overrides = db.overrides_for_items(items)

    assert overrides == {items[0].identity_key: "RS"}


def test_spare_part_presets_db_save_and_load(tmp_path):
    path = tmp_path / "spare_part_presets.json"
    db = SparePartPresetsDB(
        [
            SparePartPresetRule(
                name="Maker",
                enabled=False,
                priority=25,
                match_field="fabrikant",
                match_type="begint met",
                pattern="Her",
                target_group="Herbaroof",
            )
        ]
    )

    db.save(str(path))
    loaded = SparePartPresetsDB.load(str(path))

    assert len(loaded.rules) == 1
    assert loaded.rules[0].name == "Maker"
    assert loaded.rules[0].enabled is False
    assert loaded.rules[0].match_field == "manufacturer"
    assert loaded.rules[0].match_type == "startswith"


def test_spare_part_preset_normalizers_have_stable_defaults():
    assert normalize_spare_part_preset_field("Supplier code") == "supplier_code"
    assert normalize_spare_part_preset_field("unknown") == "manufacturer"
    assert normalize_spare_part_match_type("bevat") == "contains"
    assert normalize_spare_part_match_type("unknown") == "exact"
