import json

from app_diagnostics import (
    build_diagnostic_report,
    create_data_file_backups,
    format_report_for_clipboard,
)
from app_settings import AppSettings
from clients_db import ClientsDB
from data_storage import write_json_with_backup
from models import Client, Supplier
from order_presets_db import OrderPresetRule, OrderPresetsDB
from spare_part_presets import SparePartPresetRule, SparePartPresetsDB
from suppliers_db import SuppliersDB


def test_diagnostic_report_counts_data_files_and_backups(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    SuppliersDB([Supplier.from_any({"supplier": "ACME"})]).save("suppliers_db.json")
    ClientsDB([Client.from_any({"name": "Client A"})]).save("clients_db.json")
    OrderPresetsDB([OrderPresetRule(name="Rule A", supplier="ACME")]).save(
        "order_presets.json"
    )
    AppSettings().save("app_settings.json")
    write_json_with_backup("delivery_addresses_db.json", {"addresses": []})

    SuppliersDB(
        [
            Supplier.from_any({"supplier": "ACME"}),
            Supplier.from_any({"supplier": "BETA"}),
        ]
    ).save("suppliers_db.json")

    report = build_diagnostic_report(AppSettings())
    by_name = {item.filename: item for item in report.data_files}

    assert by_name["suppliers_db.json"].status == "OK"
    assert by_name["suppliers_db.json"].count_label == "2"
    assert by_name["suppliers_db.json"].backup_count == 1
    assert by_name["clients_db.json"].count_label == "1"
    assert not report.warnings

    text = format_report_for_clipboard(report)
    assert "Filehopper" in text
    assert "suppliers_db.json" in text


def test_diagnostic_report_flags_invalid_json_and_missing_preset_supplier(
    tmp_path, monkeypatch
):
    monkeypatch.chdir(tmp_path)

    (tmp_path / "suppliers_db.json").write_text("{not json", encoding="utf-8")
    ClientsDB([Client.from_any({"name": "Client A"})]).save("clients_db.json")
    OrderPresetsDB([OrderPresetRule(name="Rule A", supplier="Missing Supplier")]).save(
        "order_presets.json"
    )

    report = build_diagnostic_report(AppSettings())
    by_name = {item.filename: item for item in report.data_files}

    assert by_name["suppliers_db.json"].status == "Ongeldige JSON"
    assert any("onbekende leverancier" in warning for warning in report.warnings)


def test_diagnostic_report_flags_spare_part_preset_issues(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    SparePartPresetsDB(
        [
            SparePartPresetRule(
                name="RS",
                match_field="supplier",
                match_type="exact",
                pattern="",
                target_group="Electro",
            ),
            SparePartPresetRule(
                name="RS duplicate",
                match_field="supplier",
                match_type="exact",
                pattern="RS Components",
                target_group="Electro",
            ),
            SparePartPresetRule(
                name="RS overlap",
                match_field="supplier",
                match_type="exact",
                pattern="RS Components",
                target_group="Pneumatica",
            ),
        ]
    ).save("spare_part_presets.json")

    report = build_diagnostic_report(AppSettings())

    assert any("geen matchwaarde" in warning for warning in report.warnings)
    assert any("Overlappende actieve spare-part presets" in warning for warning in report.warnings)


def test_create_data_file_backups_uses_existing_files(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "suppliers_db.json").write_text(
        json.dumps({"suppliers": []}), encoding="utf-8"
    )

    backups = create_data_file_backups()

    assert len(backups) == 1
    assert backups[0].exists()
