import zipfile

import pandas as pd

from models import Supplier
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    return SuppliersDB([
        Supplier.from_any({"supplier": "ACME"}),
    ])


def _build_bom() -> pd.DataFrame:
    return pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1},
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1},
    ])


def test_duplicate_parts_single_export(tmp_path):
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
        {"Laser": "ACME"},
        {},
        {},
        False,
    )
    assert cnt == 1
    exported_files = list((dest / "Laser").glob("PN1*.pdf"))
    assert len(exported_files) == 1

    cnt_zip, _ = copy_per_production_and_orders(
        str(src),
        str(dest_zip),
        bom_df,
        [".pdf"],
        db,
        {"Laser": "ACME"},
        {},
        {},
        False,
        zip_parts=True,
    )
    assert cnt_zip == 1
    zip_path = dest_zip / "Laser" / "Laser.zip"
    assert zip_path.exists()
    with zipfile.ZipFile(zip_path) as zf:
        names = zf.namelist()
        assert names.count("PN1.pdf") == 1
