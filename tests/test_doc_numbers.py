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
    assert ws["A1"].value == "Bestelbon nr."
    assert ws["B1"].value == "BB-123"

    pdf_path = prod_folder / f"Bestelbon_BB-123_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Bestelbon nr. BB-123" in text


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
    assert ws["A1"].value == "Offerteaanvraag nr."
    assert ws["B1"].value == "OFF-42"

    pdf_path = prod_folder / f"Offerteaanvraag_OFF-42_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Offerteaanvraag nr. OFF-42" in text

