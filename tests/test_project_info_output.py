import datetime

import pandas as pd
import pytest

openpyxl = pytest.importorskip("openpyxl")
from PyPDF2 import PdfReader

import cli
from cli import build_parser, cli_copy_per_prod
from orders import copy_per_production_and_orders
from models import Supplier
from suppliers_db import SuppliersDB
from clients_db import ClientsDB
from delivery_addresses_db import DeliveryAddressesDB


def test_project_info_in_documents(tmp_path, monkeypatch):
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

    parser = build_parser()
    args = parser.parse_args([
        "copy-per-prod",
        "--source",
        str(src),
        "--dest",
        str(dst),
        "--bom",
        str(tmp_path / "bom.xlsx"),
        "--exts",
        "pdf",
        "--project-number",
        "PRJ123",
        "--project-name",
        "New Project",
    ])

    monkeypatch.setattr(cli, "load_bom", lambda path: bom_df)
    monkeypatch.setattr(SuppliersDB, "load", classmethod(lambda cls, path: db))
    monkeypatch.setattr(ClientsDB, "load", classmethod(lambda cls, path: ClientsDB([])))
    monkeypatch.setattr(
        DeliveryAddressesDB, "load", classmethod(lambda cls, path: DeliveryAddressesDB([]))
    )

    cli_copy_per_prod(args)

    today = datetime.date.today().strftime("%Y-%m-%d")
    slug = "new-project"
    expected_prefix = f"{today}_PRJ123_{slug}"
    bundle_dirs = [p for p in dst.iterdir() if p.is_dir() and p.name.startswith(expected_prefix)]
    assert bundle_dirs, "Geen bundelmap aangemaakt"
    assert len(bundle_dirs) == 1, "Meerdere bundelmappen aangetroffen"
    prod_folder = bundle_dirs[0] / "Laser"

    xlsx_path = prod_folder / f"Bestelbon_Laser_{today}.xlsx"
    assert xlsx_path.exists()
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    col_a = [ws[f"A{i}"].value for i in range(1, 20)]
    assert "Projectnummer" in col_a
    assert "Projectnaam" in col_a
    row_num = col_a.index("Projectnummer") + 1
    row_name = col_a.index("Projectnaam") + 1
    assert ws[f"B{row_num}"].value == "PRJ123"
    assert ws[f"B{row_name}"].value == "New Project"

    pdf_path = prod_folder / f"Bestelbon_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Projectnummer: PRJ123" in text
    assert "Projectnaam: New Project" in text


def test_documents_created_without_suppliers(tmp_path, monkeypatch):
    pytest.importorskip("reportlab")

    monkeypatch.chdir(tmp_path)
    db = SuppliersDB()

    src = tmp_path / "src"
    src.mkdir()
    (src / "PN1.pdf").write_text("dummy")

    dest = tmp_path / "dest"
    dest.mkdir()

    bom_df = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Description": "",
                "Production": "Laser",
                "Aantal": 1,
            }
        ]
    )

    cnt, chosen, warnings = copy_per_production_and_orders(
        str(src),
        str(dest),
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

    assert cnt == 1
    assert chosen == {"Laser": ""}
    assert not warnings

    today = datetime.date.today().strftime("%Y-%m-%d")
    prod_folder = dest / "Laser"

    xlsx_path = prod_folder / f"Bestelbon_Laser_{today}.xlsx"
    assert xlsx_path.exists()
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    header_values = {ws[f"A{i}"].value: ws[f"B{i}"].value for i in range(1, ws.max_row + 1)}
    assert "Leverancier" in header_values
    assert header_values["Leverancier"] in ("", None)

    pdf_path = prod_folder / f"Bestelbon_Laser_{today}.pdf"
    assert pdf_path.exists()
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    assert "Besteld bij:" in text
