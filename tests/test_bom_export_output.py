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
                "Bestanden gevonden": "pdf",
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
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

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
        bom_source_path=str(bom_path),
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"ProjectX-BOM-{today}.xlsx"
    assert export_path.is_file()

    exported = pd.read_excel(export_path)
    expected = df.drop(columns=["Bestanden gevonden", "Status"], errors="ignore")
    pd.testing.assert_frame_equal(
        expected.reset_index(drop=True), exported, check_dtype=False
    )
    assert "Bestanden gevonden" not in exported.columns
    assert "Status" not in exported.columns


def test_bom_export_can_be_disabled(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "ProjectX-BOM.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

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
        bom_source_path=str(bom_path),
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"ProjectX-BOM-{today}.xlsx"
    assert not export_path.exists()


def test_bom_export_strips_suffix_after_bom(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    bom_path = src / "154215-a21-BOM-partsonly.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

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
        bom_source_path=str(bom_path),
    )

    today = datetime.date.today().strftime("%Y-%m-%d")
    export_path = dest / f"154215-a21-BOM-{today}.xlsx"
    assert export_path.exists()


def test_related_exports_copied_next_to_bom(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir(); dest.mkdir()

    (src / "PN-001.pdf").write_text("dummy", encoding="utf-8")
    assembly_stem = "20230171-v2-a00"
    (src / f"{assembly_stem}.pdf").write_text("top", encoding="utf-8")
    bom_path = src / f"{assembly_stem}-BOM-PartsOnly.xlsx"
    bom_path.write_text("bom", encoding="utf-8")

    df = _basic_bom()

    cnt, _ = copy_per_production_and_orders(
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
        bom_source_path=str(bom_path),
    )

    assert cnt == 2
    assert (dest / f"{assembly_stem}.pdf").is_file()
