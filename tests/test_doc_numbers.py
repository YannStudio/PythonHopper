import datetime
import os

import pandas as pd
import openpyxl
import pytest
from PyPDF2 import PdfReader

from models import Supplier
from suppliers_db import SuppliersDB
from orders import copy_per_production_and_orders


def test_doc_number_in_name_and_header(tmp_path):
    reportlab = pytest.importorskip("reportlab")

    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))
    src = tmp_path / "src"
    src.mkdir()
    (src / "PN1.pdf").write_text("dummy")
    bom_df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}
    ])
    dst = tmp_path / "dst"
    dst.mkdir()

    doc_num_map = {"Laser": "123"}
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        doc_num_map,
        False,
        client=None,
        delivery_map={},
    )

    prod_folder = dst / "Laser"
    today = datetime.date.today().strftime("%Y-%m-%d")

    xlsx_path = prod_folder / f"Bestelbon_123_Laser_{today}.xlsx"
    assert xlsx_path.exists()
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    assert ws["A1"].value == "Bestelbon nr."
    assert ws["B1"].value == "123"

    pdf_path = prod_folder / f"Bestelbon_123_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Bestelbon nr. 123" in text

