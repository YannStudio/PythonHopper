import datetime
import os

import pandas as pd
import openpyxl
import pytest
from PyPDF2 import PdfReader

from models import Supplier
from suppliers_db import SuppliersDB
from orders import copy_per_production_and_orders, _prefix_for_doc_type


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

    xlsx_path = prod_folder / f"Bestelbon_BB-123_Laser_{today}.xlsx"
    assert xlsx_path.exists()
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    assert ws["A1"].value == "Nummer"
    assert ws["B1"].value == "BB-123"
    assert ws["A2"].value == "Datum"
    assert ws["B2"].value == today

    pdf_path = prod_folder / f"Bestelbon_BB-123_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    lines = text.splitlines()
    assert f"Nummer: BB-123" in text
    assert f"Datum: {today}" in text
    assert "BB-123" not in lines[0]


def test_prefix_helper():
    assert _prefix_for_doc_type("Bestelbon") == "BB-"
    assert _prefix_for_doc_type("Offerteaanvraag") == "OFF-"
    assert _prefix_for_doc_type("Onbekend") == ""


def test_offerte_prefix_in_output(tmp_path):
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

    doc_num_map = {"Laser": "42"}
    doc_type_map = {"Laser": "Offerteaanvraag"}
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        doc_type_map,
        doc_num_map,
        False,
        client=None,
        delivery_map={},
    )

    prod_folder = dst / "Laser"
    today = datetime.date.today().strftime("%Y-%m-%d")

    xlsx_path = prod_folder / f"Offerteaanvraag_OFF-42_Laser_{today}.xlsx"
    assert xlsx_path.exists()
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    assert ws["A1"].value == "Nummer"
    assert ws["B1"].value == "OFF-42"
    assert ws["A2"].value == "Datum"
    assert ws["B2"].value == today

    pdf_path = prod_folder / f"Offerteaanvraag_OFF-42_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    lines = text.splitlines()
    assert f"Nummer: OFF-42" in text
    assert f"Datum: {today}" in text
    assert "OFF-42" not in lines[0]


def test_missing_doc_number_omits_prefix_and_header(tmp_path):
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
    )

    prod_folder = dst / "Laser"
    today = datetime.date.today().strftime("%Y-%m-%d")

    xlsx_path = prod_folder / f"Bestelbon_Laser_{today}.xlsx"
    assert xlsx_path.exists()
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    # Without document number the first header line should be the supplier
    assert ws["A1"].value == "Leverancier"
    rows = list(ws.iter_rows(min_row=1, max_row=10, max_col=2, values_only=True))
    assert ("Datum", today) in rows

    pdf_path = prod_folder / f"Bestelbon_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Nummer:" not in text
    assert f"Datum: {today}" in text

