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


def test_order_documents_can_be_generated_without_copying_source_files(
    tmp_path, monkeypatch
):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()
    (src / "PN-001.pdf").write_text("drawing", encoding="utf-8")
    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN-001",
                "Production": "Laser",
                "Description": "Part",
                "Aantal": 1,
            }
        ]
    )

    def fake_pdf(path, *_args, **_kwargs):
        with open(path, "wb") as fh:
            fh.write(b"%PDF-1.4\n%%EOF\n")

    monkeypatch.setattr("orders.generate_pdf_order_platypus", fake_pdf)
    generated = []

    cnt, chosen = copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {"Laser": "BB-1"},
        remember_defaults=False,
        export_bom=False,
        export_related_files=False,
        generated_documents=generated,
    )

    assert cnt == 0
    assert chosen[make_production_selection_key("Laser")] == "ACME"
    assert not (dest / "Laser" / "PN-001.pdf").exists()
    pdf_records = [
        record
        for record in generated
        if record.get("format") == "pdf" and record.get("kind") == "order"
    ]
    assert pdf_records
    assert all((dest / record["path"]).exists() for record in pdf_records)
    assert all(
        str(record["path"]).replace("\\", "/").startswith("Laser/")
        for record in pdf_records
    )
