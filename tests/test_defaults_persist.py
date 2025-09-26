import os
import pandas as pd
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
    cnt, chosen, warnings = copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        overrides,
        {},
        {"Laser": "1", "Plasma": "2"},
        True,
        client=None,
        delivery_map={},
    )

    assert cnt == 2
    assert chosen == overrides
    assert not warnings

    # Defaults should be updated in memory
    assert db.defaults_by_production == overrides

    # Reload from disk and ensure defaults persisted
    db2 = SuppliersDB.load()
    assert db2.defaults_by_production == overrides
