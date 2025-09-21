import os
import pandas as pd
from helpers import create_export_bundle
from models import Supplier
from suppliers_db import SuppliersDB
from orders import copy_per_production_and_orders


def test_defaults_persist(tmp_path, monkeypatch):
    # Ensure suppliers DB file is created inside temporary directory
    monkeypatch.chdir(tmp_path)

    # Setup supplier database with two suppliers
    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))
    db.upsert(Supplier.from_any({"supplier": "BETA"}))

    # Create source and destination directories
    src = tmp_path / "src"
    dst = tmp_path / "dst"
    src.mkdir(); dst.mkdir()

    # Create dummy files for each part
    (src / "PN1.pdf").write_text("dummy")
    (src / "PN2.pdf").write_text("dummy")

    # BOM with two productions
    bom_df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1},
        {"PartNumber": "PN2", "Description": "", "Production": "Plasma", "Aantal": 1},
    ])

    overrides = {"Laser": "ACME", "Plasma": "BETA"}

    # Run the copy and order generation, remembering defaults
    bundle = create_export_bundle(dst, "Defaults")
    cnt, chosen, bundle_info = copy_per_production_and_orders(
        str(src),
        str(bundle["path"]),
        bom_df,
        [".pdf"],
        db,
        overrides,
        {},
        {"Laser": "1", "Plasma": "2"},
        True,
        client=None,
        delivery_map={},
        bundle=bundle,
    )

    assert cnt == 2
    assert chosen == overrides
    if bundle_info.get("latest"):
        assert (tmp_path / "dst" / "latest").is_symlink()

    # Defaults should be updated in memory
    assert db.defaults_by_production == overrides

    # Reload from disk and ensure defaults persisted
    db2 = SuppliersDB.load()
    assert db2.defaults_by_production == overrides
