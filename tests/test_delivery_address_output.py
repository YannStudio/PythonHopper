import io
import os
import time
from pathlib import Path
import pandas as pd
import pytest
openpyxl = pytest.importorskip("openpyxl")
from PyPDF2 import PdfReader

from models import Supplier, DeliveryAddress
from suppliers_db import SuppliersDB
from orders import copy_per_production_and_orders


def _setup_basic(tmp_path):
    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))
    src = tmp_path / "src"
    src.mkdir()
    (src / "PN1.pdf").write_text("dummy")
    bom_df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}
    ])
    return db, src, bom_df


def _read_pdf_text(path, retries=5, delay=0.05):
    last_error = None
    for _ in range(retries):
        try:
            data = Path(path).read_bytes()
            reader = PdfReader(io.BytesIO(data))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as exc:  # pragma: no cover - best effort retry
            last_error = exc
            try:
                text = data.decode("latin-1", errors="ignore")
                return text
            except Exception:
                pass
            time.sleep(delay)
    raise last_error


def test_delivery_address_present_absent(tmp_path):
    reportlab = pytest.importorskip("reportlab")
    db, src, bom_df = _setup_basic(tmp_path)

    delivery = DeliveryAddress(name="Magazijn", address="Straat 1", remarks="achterdeur")

    # With delivery address
    dst1 = tmp_path / "dst1"
    dst1.mkdir()
    copy_per_production_and_orders(
        str(src),
        str(dst1),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        {"Laser": "1"},
        False,
        client=None,
        delivery_map={"Laser": delivery},
    )
    prod_folder = dst1 / "Laser"
    xlsx = next(f for f in os.listdir(prod_folder) if f.endswith(".xlsx"))
    wb = openpyxl.load_workbook(prod_folder / xlsx)
    ws = wb.active
    col_a = [ws[f"A{i}"].value for i in range(1, 20)]
    assert "Leveradres" in col_a
    row = col_a.index("Leveradres") + 1
    # Name is on the next line with an empty label
    assert ws[f"B{row}"].value in (None, "")
    assert ws[f"B{row+1}"].value == "Magazijn"
    assert ws[f"B{row+2}"].value == "Straat 1"
    assert ws[f"B{row+3}"].value == "achterdeur"
    pdf = next(f for f in os.listdir(prod_folder) if f.endswith(".pdf") and "_" in f)
    text = _read_pdf_text(prod_folder / pdf)
    assert "Leveradres:\nMagazijn" in text

    # Without delivery address
    dst2 = tmp_path / "dst2"
    dst2.mkdir()
    copy_per_production_and_orders(
        str(src),
        str(dst2),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        {"Laser": "2"},
        False,
        client=None,
        delivery_map={},
    )
    prod_folder2 = dst2 / "Laser"
    xlsx2 = next(f for f in os.listdir(prod_folder2) if f.endswith(".xlsx"))
    wb2 = openpyxl.load_workbook(prod_folder2 / xlsx2)
    ws2 = wb2.active
    col_a2 = [ws2[f"A{i}"].value for i in range(1, 20)]
    assert "Leveradres" not in col_a2
    pdf2 = next(f for f in os.listdir(prod_folder2) if f.endswith(".pdf") and "_" in f)
    text2 = _read_pdf_text(prod_folder2 / pdf2)
    assert "Leveradres" not in text2


def test_delivery_address_per_production(tmp_path):
    reportlab = pytest.importorskip("reportlab")
    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))
    src = tmp_path / "src"
    src.mkdir()
    (src / "PN1.pdf").write_text("dummy")
    (src / "PN2.pdf").write_text("dummy")
    bom_df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1},
        {"PartNumber": "PN2", "Description": "", "Production": "Plasma", "Aantal": 1},
    ])

    d1 = DeliveryAddress(name="Magazijn", address="Straat 1")
    d2 = DeliveryAddress(name="Depot", address="Weg 2")
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
        {"Laser": "10", "Plasma": "20"},
        False,
        client=None,
        delivery_map={"Laser": d1, "Plasma": d2},
    )

    # Laser folder checks
    laser = dst / "Laser"
    xlsx1 = next(f for f in os.listdir(laser) if f.endswith(".xlsx"))
    wb1 = openpyxl.load_workbook(laser / xlsx1)
    ws1 = wb1.active
    col_a1 = [ws1[f"A{i}"].value for i in range(1, 20)]
    row1 = col_a1.index("Leveradres") + 1
    assert ws1[f"B{row1}"].value in (None, "")
    assert ws1[f"B{row1+1}"].value == "Magazijn"
    pdf1 = next(f for f in os.listdir(laser) if f.endswith(".pdf") and "_" in f)
    text1 = _read_pdf_text(laser / pdf1)
    assert "Leveradres:\nMagazijn" in text1

    # Plasma folder checks
    plasma = dst / "Plasma"
    xlsx2 = next(f for f in os.listdir(plasma) if f.endswith(".xlsx"))
    wb2 = openpyxl.load_workbook(plasma / xlsx2)
    ws2 = wb2.active
    col_a2 = [ws2[f"A{i}"].value for i in range(1, 20)]
    row2 = col_a2.index("Leveradres") + 1
    assert ws2[f"B{row2}"].value in (None, "")
    assert ws2[f"B{row2+1}"].value == "Depot"
    pdf2 = next(f for f in os.listdir(plasma) if f.endswith(".pdf") and "_" in f)
    text2 = _read_pdf_text(plasma / pdf2)
    assert "Leveradres:\nDepot" in text2


def test_delivery_address_placeholder_prints(tmp_path):
    reportlab = pytest.importorskip("reportlab")
    db, src, bom_df = _setup_basic(tmp_path)

    # Placeholder address should still render in the delivery block
    delivery = DeliveryAddress(name="Magazijn", address="Bestelling wordt opgehaald")

    dst = tmp_path / "dst_placeholder"
    dst.mkdir()
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        {"Laser": "3"},
        False,
        client=None,
        delivery_map={"Laser": delivery},
    )

    prod_folder = dst / "Laser"
    pdf = next(f for f in os.listdir(prod_folder) if f.endswith(".pdf") and "_" in f)
    text = _read_pdf_text(prod_folder / pdf)
    assert "Bestelling wordt opgehaald" in text


def test_export_remark_rendered_under_delivery(tmp_path):
    reportlab = pytest.importorskip("reportlab")
    db, src, bom_df = _setup_basic(tmp_path)

    delivery = DeliveryAddress(
        name="Magazijn", address="Straat 1", remarks="Afspraak balie"
    )

    dst = tmp_path / "dst_export"
    dst.mkdir()
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        {"Laser": "Exportbon"},
        {"Laser": "5"},
        False,
        client=None,
        delivery_map={"Laser": delivery},
        remarks_map={"Laser": "Export aanwijzing"},
    )

    prod_folder = dst / "Laser"
    pdf = next(f for f in os.listdir(prod_folder) if f.endswith(".pdf") and "_" in f)
    text = _read_pdf_text(prod_folder / pdf)

    expected = (
        "Leveradres:\nMagazijn\nStraat 1\nAfspraak balie\nOpmerking:\nExport aanwijzing"
    )
    assert expected in text
    assert "Opmerking: Export aanwijzing" not in text

    xlsx = next(f for f in os.listdir(prod_folder) if f.endswith(".xlsx"))
    wb = openpyxl.load_workbook(prod_folder / xlsx)
    ws = wb.active
    col_a = [ws[f"A{i}"].value for i in range(1, 40)]
    col_b = [ws[f"B{i}"].value for i in range(1, 40)]
    lever_row = col_a.index("Leveradres") + 1
    assert "Export aanwijzing" not in col_b[: lever_row - 1]
    assert any(
        ws[f"A{i}"].value == "Opmerking" and ws[f"B{i}"].value == "Export aanwijzing"
        for i in range(lever_row, lever_row + 8)
    )


def test_export_remark_rendered_under_delivery_with_spaced_doc_type(tmp_path):
    reportlab = pytest.importorskip("reportlab")
    db, src, bom_df = _setup_basic(tmp_path)

    delivery = DeliveryAddress(
        name="Magazijn", address="Straat 1", remarks="Afspraak balie"
    )

    dst = tmp_path / "dst_export_spaced"
    dst.mkdir()
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        {"Laser": "Export bon"},
        {"Laser": "5"},
        False,
        client=None,
        delivery_map={"Laser": delivery},
        remarks_map={"Laser": "Export aanwijzing"},
    )

    prod_folder = dst / "Laser"
    pdf = next(f for f in os.listdir(prod_folder) if f.endswith(".pdf") and "_" in f)
    text = _read_pdf_text(prod_folder / pdf)

    expected = (
        "Leveradres:\nMagazijn\nStraat 1\nAfspraak balie\nOpmerking:\nExport aanwijzing"
    )
    assert expected in text
    assert "Opmerking: Export aanwijzing" not in text

    xlsx = next(f for f in os.listdir(prod_folder) if f.endswith(".xlsx"))
    wb = openpyxl.load_workbook(prod_folder / xlsx)
    ws = wb.active
    col_a = [ws[f"A{i}"].value for i in range(1, 40)]
    col_b = [ws[f"B{i}"].value for i in range(1, 40)]
    lever_row = col_a.index("Leveradres") + 1
    assert "Export aanwijzing" not in col_b[: lever_row - 1]
    assert any(
        ws[f"A{i}"].value == "Opmerking" and ws[f"B{i}"].value == "Export aanwijzing"
        for i in range(lever_row, lever_row + 8)
    )


def test_export_remark_rendered_under_delivery_with_abbreviated_doc_type(tmp_path):
    reportlab = pytest.importorskip("reportlab")
    db, src, bom_df = _setup_basic(tmp_path)

    delivery = DeliveryAddress(
        name="Magazijn", address="Straat 1", remarks="Afspraak balie"
    )

    dst = tmp_path / "dst_export_abbrev"
    dst.mkdir()
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        {"Laser": "Exp bon"},
        {"Laser": "7"},
        False,
        client=None,
        delivery_map={"Laser": delivery},
        remarks_map={"Laser": "Export aanwijzing"},
    )

    prod_folder = dst / "Laser"
    pdf = next(f for f in os.listdir(prod_folder) if f.endswith(".pdf") and "_" in f)
    text = _read_pdf_text(prod_folder / pdf)

    expected = (
        "Leveradres:\nMagazijn\nStraat 1\nAfspraak balie\nOpmerking:\nExport aanwijzing"
    )
    assert expected in text
    assert "Opmerking: Export aanwijzing" not in text

    xlsx = next(f for f in os.listdir(prod_folder) if f.endswith(".xlsx"))
    wb = openpyxl.load_workbook(prod_folder / xlsx)
    ws = wb.active
    col_a = [ws[f"A{i}"].value for i in range(1, 40)]
    col_b = [ws[f"B{i}"].value for i in range(1, 40)]
    lever_row = col_a.index("Leveradres") + 1
    assert "Export aanwijzing" not in col_b[: lever_row - 1]
    assert any(
        ws[f"A{i}"].value == "Opmerking" and ws[f"B{i}"].value == "Export aanwijzing"
        for i in range(lever_row, lever_row + 8)
    )


def test_export_remark_rendered_without_delivery(tmp_path):
    reportlab = pytest.importorskip("reportlab")
    db, src, bom_df = _setup_basic(tmp_path)

    dst = tmp_path / "dst_export_no_delivery"
    dst.mkdir()
    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        {},
        {"Laser": "Export"},
        {"Laser": "5"},
        False,
        client=None,
        delivery_map={},
        remarks_map={"Laser": "Export aanwijzing"},
    )

    prod_folder = dst / "Laser"
    pdf = next(f for f in os.listdir(prod_folder) if f.endswith(".pdf"))
    reader = PdfReader(prod_folder / pdf)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)

    assert "Opmerking:\nExport aanwijzing" in text
    assert "Opmerking: Export aanwijzing" not in text
    assert "Leveradres" not in text

    xlsx = next(f for f in os.listdir(prod_folder) if f.endswith(".xlsx"))
    wb = openpyxl.load_workbook(prod_folder / xlsx)
    ws = wb.active
    opmerking_rows = [
        (ws[f"A{i}"].value, ws[f"B{i}"].value)
        for i in range(1, 30)
        if ws[f"A{i}"].value == "Opmerking"
    ]
    assert ("Opmerking", "Export aanwijzing") in opmerking_rows
