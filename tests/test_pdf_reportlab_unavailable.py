import datetime

import pandas as pd
import pytest

from models import Supplier
from suppliers_db import SuppliersDB
import orders


def test_pdf_export_skipped_when_reportlab_missing(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))

    source = tmp_path / "src"
    dest = tmp_path / "dest"
    source.mkdir()
    dest.mkdir()

    (source / "PN1.pdf").write_text("dummy")

    bom_df = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Description": "onderdeel",
                "Production": "Laser",
                "Aantal": 1,
            }
        ]
    )

    monkeypatch.setattr(orders, "REPORTLAB_OK", False, raising=False)

    count, chosen, warnings = orders.copy_per_production_and_orders(
        str(source),
        str(dest),
        bom_df,
        [".pdf"],
        db,
        {"Laser": "ACME"},
        doc_type_map={},
        doc_num_map={},
        remember_defaults=False,
        client=None,
        delivery_map={},
    )

    assert count == 1
    assert chosen == {"Laser": "ACME"}
    assert any("ReportLab" in warn for warn in warnings)

    today = datetime.date.today().strftime("%Y-%m-%d")
    prod_dir = dest / "Laser"
    excel_path = prod_dir / f"Bestelbon_Laser_{today}.xlsx"
    pdf_path = prod_dir / f"Bestelbon_Laser_{today}.pdf"

    assert excel_path.exists()
    assert not pdf_path.exists()

    exported_files = {p.name for p in prod_dir.iterdir()}
    assert "PN1.pdf" in exported_files


def test_generate_pdf_order_platypus_reports_missing_dependency(tmp_path, monkeypatch):
    monkeypatch.setattr(orders, "REPORTLAB_OK", False, raising=False)

    supplier = Supplier.from_any({"supplier": "Test"})

    with pytest.raises(orders.PDFGenerationUnavailableError) as excinfo:
        orders.generate_pdf_order_platypus(
            str(tmp_path / "out.pdf"),
            {"name": "Test", "address": ""},
            supplier,
            "Laser",
            [],
        )

    assert "ReportLab" in str(excinfo.value)
