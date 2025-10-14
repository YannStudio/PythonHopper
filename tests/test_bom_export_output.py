import datetime

import pandas as pd

from models import Supplier
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    return SuppliersDB([Supplier.from_any({"supplier": "ACME"})])


def _basic_bom() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "PartNumber": "PN-001",
                "Description": "Behuizing",
                "Production": "Laser",
                "Aantal": 2,
                "Status": "✅",
            }
        ]
    )


def test_bom_export_written_with_iso_date(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")

    df = _basic_bom()

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=True,
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"BOM-FileHopper-Export-{today}.xlsx"
    assert export_path.is_file()

    exported = pd.read_excel(export_path)
    pd.testing.assert_frame_equal(
        df.reset_index(drop=True), exported, check_dtype=False
    )


def test_bom_export_can_be_disabled(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")

    df = _basic_bom()

    copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        export_bom=False,
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"BOM-FileHopper-Export-{today}.xlsx"
    assert not export_path.exists()
