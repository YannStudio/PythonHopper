import datetime
import pandas as pd
import openpyxl
import pytest
from PyPDF2 import PdfReader

from models import Supplier
from suppliers_db import SuppliersDB
from orders import copy_per_production_and_orders


def test_project_info_in_documents(tmp_path):
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

    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        {},
        False,
        client=None,
        delivery_map={},
        project_number="PRJ123",
        project_name="New Project",
    )

    prod_folder = dst / "Laser"
    today = datetime.date.today().strftime("%Y-%m-%d")

    xlsx_path = prod_folder / f"Bestelbon_Laser_{today}.xlsx"
    assert xlsx_path.exists()
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    col_a = [ws[f"A{i}"].value for i in range(1, 20)]
    assert "Projectnr." in col_a
    assert "Projectnaam" in col_a
    row_num = col_a.index("Projectnr.") + 1
    row_name = col_a.index("Projectnaam") + 1
    assert ws[f"B{row_num}"].value == "PRJ123"
    assert ws[f"B{row_name}"].value == "New Project"

    pdf_path = prod_folder / f"Bestelbon_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Projectnr.: PRJ123" in text
    assert "Projectnaam: New Project" in text
