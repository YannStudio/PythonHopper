import pandas as pd

from orders import copy_per_production_and_orders, make_production_selection_key
from models import Supplier
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    return SuppliersDB([Supplier.from_any({"supplier": "ACME"})])


def _bom_two_productions() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"PartNumber": "PN-001", "Production": "Laser"},
            {"PartNumber": "PN-002", "Production": "Assembly"},
        ]
    )


def test_production_export_filter_skips_disabled(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    (src / "PN-001.dxf").write_text("laser", encoding="utf-8")
    (src / "PN-002.dxf").write_text("assembly", encoding="utf-8")

    df = _bom_two_productions()

    cnt, chosen = copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".dxf"],
        _make_db(),
        {"Laser": "ACME", "Assembly": "ACME"},
        {},
        {},
        remember_defaults=False,
        export_bom=False,
        production_export_filter={"Laser": True, "Assembly": False},
    )

    laser_dir = dest / "Laser"
    assembly_dir = dest / "Assembly"

    assert laser_dir.is_dir()
    assert not assembly_dir.exists()
    assert sorted(p.name for p in laser_dir.glob("*.dxf")) == ["PN-001.dxf"]
    assert cnt == 1
    assert make_production_selection_key("Laser") in chosen
    assert make_production_selection_key("Assembly") not in chosen
