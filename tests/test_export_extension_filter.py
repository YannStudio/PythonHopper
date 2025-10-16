import zipfile

import pandas as pd

from orders import copy_per_production_and_orders
from models import Supplier
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    return SuppliersDB([Supplier.from_any({"supplier": "ACME"})])


def _basic_bom() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "PartNumber": "PN-001",
                "Production": "Laser",
            }
        ]
    )


def test_export_skips_unselected_extensions(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    (src / "PN-001.dxf").write_text("dxf", encoding="utf-8")
    (src / "PN-001.ai").write_text("ai", encoding="utf-8")

    df = _basic_bom()

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".dxf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        remember_defaults=False,
        export_bom=False,
    )

    prod_dir = dest / "Laser"
    exported_dxf = sorted(p.name for p in prod_dir.glob("*.dxf"))
    assert exported_dxf == ["PN-001.dxf"]
    assert not any(p.suffix.lower() == ".ai" for p in prod_dir.iterdir())
    assert cnt == 1


def test_zip_export_skips_unselected_extensions(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    (src / "PN-001.dxf").write_text("dxf", encoding="utf-8")
    (src / "PN-001.ai").write_text("ai", encoding="utf-8")

    df = _basic_bom()

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".dxf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        remember_defaults=False,
        export_bom=False,
        zip_parts=True,
    )

    prod_dir = dest / "Laser"
    zip_path = prod_dir / "Laser.zip"
    assert zip_path.is_file()
    with zipfile.ZipFile(zip_path) as zf:
        assert zf.namelist() == ["PN-001.dxf"]
    assert cnt == 1
