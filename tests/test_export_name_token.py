import zipfile

import pandas as pd

import orders
from models import Supplier
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    db = SuppliersDB([
        Supplier.from_any({"supplier": "ACME"}),
    ])
    return db


def _build_bom() -> pd.DataFrame:
    return pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}
    ])


def test_export_token_applied_to_files_and_zip(tmp_path, monkeypatch):
    monkeypatch.setattr(orders, "SUPPLIERS_DB_FILE", str(tmp_path / "suppliers.json"))

    src = tmp_path / "src"
    dest = tmp_path / "dest"
    dest_zip = tmp_path / "dest_zip"
    src.mkdir()
    dest.mkdir()
    dest_zip.mkdir()

    (src / "PN1.pdf").write_text("dummy")

    db = _make_db()
    bom_df = _build_bom()

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        bom_df,
        [".pdf"],
        db,
        {"Laser": ""},
        {},
        {},
        False,
        export_name_token="REV-A",
    )
    assert cnt == 1
    exported = dest / "Laser" / "PN1-REV-A.pdf"
    assert exported.exists()

    cnt_zip, _ = copy_per_production_and_orders(
        str(src),
        str(dest_zip),
        bom_df,
        [".pdf"],
        db,
        {"Laser": ""},
        {},
        {},
        False,
        zip_parts=True,
        export_name_token="REV-A",
    )
    assert cnt_zip == 1
    zip_path = dest_zip / "Laser" / "Laser.zip"
    assert zip_path.exists()
    with zipfile.ZipFile(zip_path) as zf:
        assert "PN1-REV-A.pdf" in zf.namelist()
