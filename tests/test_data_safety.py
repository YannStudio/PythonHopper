import json

import pandas as pd

from models import Supplier
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def _basic_bom_export(tmp_path, db):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()
    (src / "PN1.pdf").write_text("dummy", encoding="utf-8")
    bom_df = pd.DataFrame(
        [{"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}]
    )

    return copy_per_production_and_orders(
        str(src),
        str(dest),
        bom_df,
        [".pdf"],
        db,
        {"Laser": "ACME"},
        {},
        {"Laser": "BB-1"},
        True,
        client=None,
        delivery_map={},
        export_bom=False,
        export_related_files=False,
    )


def test_in_memory_supplier_db_does_not_overwrite_real_supplier_file(tmp_path):
    real_path = tmp_path / "suppliers_db.json"
    SuppliersDB(
        [
            Supplier.from_any({"supplier": "ACME"}),
            Supplier.from_any({"supplier": "BETA"}),
        ],
        storage_path=real_path,
    ).save()

    transient_db = SuppliersDB([Supplier.from_any({"supplier": "ACME"})])
    _basic_bom_export(tmp_path, transient_db)

    reloaded = SuppliersDB.load(real_path)
    assert [supplier.supplier for supplier in reloaded.suppliers] == ["ACME", "BETA"]
    assert reloaded.defaults_by_production == {}


def test_loaded_supplier_db_persists_defaults_without_dropping_suppliers(tmp_path):
    real_path = tmp_path / "suppliers_db.json"
    SuppliersDB(
        [
            Supplier.from_any({"supplier": "ACME"}),
            Supplier.from_any({"supplier": "BETA"}),
        ],
        storage_path=real_path,
    ).save()

    loaded_db = SuppliersDB.load(real_path)
    _basic_bom_export(tmp_path, loaded_db)

    reloaded = SuppliersDB.load(real_path)
    assert [supplier.supplier for supplier in reloaded.suppliers] == ["ACME", "BETA"]
    assert reloaded.defaults_by_production == {"Laser": "ACME"}


def test_json_saves_create_backup_before_overwrite(tmp_path):
    path = tmp_path / "suppliers_db.json"
    db = SuppliersDB([Supplier.from_any({"supplier": "ACME"})], storage_path=path)
    db.save()

    db.upsert(Supplier.from_any({"supplier": "BETA"}))
    db.save()

    backup_dir = tmp_path / ".filehopper_backups" / "suppliers_db"
    backups = list(backup_dir.glob("suppliers_db_*.json"))
    assert len(backups) == 1

    backup_data = json.loads(backups[0].read_text(encoding="utf-8"))
    assert [supplier["supplier"] for supplier in backup_data["suppliers"]] == ["ACME"]
