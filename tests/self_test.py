from __future__ import annotations

import os
import tempfile
import datetime

import pandas as pd
import pytest

openpyxl = pytest.importorskip("openpyxl")

from models import Supplier, Client
from suppliers_db import SuppliersDB
from clients_db import ClientsDB
from bom import load_bom
from orders import (
    copy_per_production_and_orders,
    DEFAULT_FOOTER_NOTE,
    write_order_excel,
    make_production_selection_key,
)


def run_tests() -> int:
    print("Running self-tests...")
    db = SuppliersDB()
    db.upsert(Supplier.from_any({
        "supplier": "ACME",
        "description": "Snijwerk",
        "btw": "BE123",
        "adress_1": "Teststraat 1 bus 2",
        "address_2": "BE-2000 Antwerpen",
        "land": "BE",
        "e-mail sales": "x@y.z",
        "tel. sales": "+32 123",
    }))
    assert db.suppliers and db.suppliers[0].adres_2 == "BE-2000 Antwerpen"
    db.toggle_fav("ACME")
    assert db.suppliers[0].favorite
    db.set_default("Laser", "ACME")
    assert db.get_default("Laser") == "ACME"

    cdb = ClientsDB()
    cdb.upsert(Client.from_any({
        "name": "TestClient",
        "address": "Straat 1, 1000 Brussel",
        "vat": "BE000",
        "email": "test@example.com",
    }))
    client = cdb.clients[0]

    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "src")
        dst = os.path.join(td, "dst")
        os.makedirs(src)
        os.makedirs(dst)
        open(os.path.join(src, "PN1.pdf"), "wb").write(b"%PDF-1.4")
        open(os.path.join(src, "PN1.stp"), "wb").write(b"step")
        bom = os.path.join(td, "bom.xlsx")
        df = pd.DataFrame([
            {
                "PartNumber": "PN1-THIS-IS-A-VERY-LONG-CODE-OVER-25CHARS",
                "Description": "Lange omschrijving die netjes moet wrappen.",
                "Production": "Laser",
                "Aantal": 2,
                "Materiaal": "S235JR",
                "Oppervlakte (m²)": "1,23",
                "Gewicht (kg)": "4,56",
            },
            {
                "PartNumber": "PN2",
                "Description": "Geen files",
                "Production": "Laser",
                "Aantal": 1000,
                "Material": "Alu 5754",
                "Area": "0.50",
                "Weight": "1.00",
            },
        ])
        df.to_excel(bom, index=False, engine="openpyxl")
        ldf = load_bom(bom)
        assert ldf["Aantal"].max() <= 999  # capped
        cnt, chosen = copy_per_production_and_orders(
            src,
            dst,
            ldf,
            [".pdf", ".stp"],
            db,
            {},
            {},
            {"Laser": "1"},
            True,
            client=client,
            delivery_map={},
            footer_note=DEFAULT_FOOTER_NOTE,
        )
        assert cnt == 2
        assert chosen.get(make_production_selection_key("Laser")) == "ACME"
        prod_folder = os.path.join(dst, "Laser")
        assert os.path.exists(os.path.join(prod_folder, "PN1.pdf"))
        assert os.path.exists(os.path.join(prod_folder, "PN1.stp"))
        xlsx = [f for f in os.listdir(prod_folder) if f.lower().endswith(".xlsx")]
        assert xlsx, "Excel bestelbon niet aangemaakt"
        wb = openpyxl.load_workbook(os.path.join(prod_folder, xlsx[0]))
        ws = wb.active
        today_display = datetime.date.today().strftime("%Y-%m-%d")
        date_token = datetime.date.today().strftime("%Y%m%d")
        assert ws["A1"].value == "Nummer" and ws["B1"].value == "BB-1"
        assert ws["A2"].value == "Datum" and ws["B2"].value == today_display
        assert ws["A4"].value == "Bedrijf" and ws["B4"].value == client.name
        assert ws["A9"].value == "Leverancier" and ws["B9"].value == "ACME"
        assert ws["A10"].value == "Adres"
        assert (
            ws["B10"].value == "Teststraat 1 bus 2, BE-2000 Antwerpen, BE"
        )
        assert ws["A11"].value == "BTW" and ws["B11"].value == "BE123"
        assert ws["A12"].value == "E-mail" and ws["B12"].value == "x@y.z"
        assert ws["A13"].value == "Tel" and ws["B13"].value == "+32 123"
        header_row = None
        for row in range(1, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == "PartNumber":
                header_row = row
                break
        assert header_row is not None, "PartNumber header niet gevonden"
        pn_cell = ws.cell(row=header_row + 1, column=1)
        assert pn_cell.alignment is not None
        assert pn_cell.alignment.wrap_text, "PartNumber wrap_text niet geactiveerd"
        pdfs = [f for f in os.listdir(prod_folder) if f.lower().endswith(".pdf")]
        assert pdfs, "PDF bestelbon niet aangemaakt"

        dst_dates = os.path.join(td, "dst_dates")
        os.makedirs(dst_dates)
        cnt_dates, _ = copy_per_production_and_orders(
            src,
            dst_dates,
            ldf,
            [".pdf", ".stp"],
            db,
            {},
            {},
            {"Laser": "1"},
            True,
            client=client,
            delivery_map={},
            footer_note=DEFAULT_FOOTER_NOTE,
            date_suffix_exports=True,
        )
        assert cnt_dates == 2
        prod_folder_dates = os.path.join(dst_dates, "Laser")
        suffix_pdf_name = f"PN1-{date_token}.pdf"
        suffix_stp_name = f"PN1-{date_token}.stp"
        assert "_" not in suffix_pdf_name and "_" not in suffix_stp_name
        assert os.path.exists(os.path.join(prod_folder_dates, suffix_pdf_name))
        assert os.path.exists(os.path.join(prod_folder_dates, suffix_stp_name))

        dst_prefix = os.path.join(td, "dst_prefix")
        os.makedirs(dst_prefix)
        cnt_prefix, _ = copy_per_production_and_orders(
            src,
            dst_prefix,
            ldf,
            [".pdf", ".stp"],
            db,
            {},
            {},
            {"Laser": "1"},
            True,
            client=client,
            delivery_map={},
            footer_note=DEFAULT_FOOTER_NOTE,
            date_prefix_exports=True,
        )
        assert cnt_prefix == 2
        prod_folder_prefix = os.path.join(dst_prefix, "Laser")
        prefix_pdf_name = f"{date_token}-PN1.pdf"
        prefix_stp_name = f"{date_token}-PN1.stp"
        assert os.path.exists(os.path.join(prod_folder_prefix, prefix_pdf_name))
        assert os.path.exists(os.path.join(prod_folder_prefix, prefix_stp_name))

        dst_prefix_suffix = os.path.join(td, "dst_prefix_suffix")
        os.makedirs(dst_prefix_suffix)
        cnt_prefix_suffix, _ = copy_per_production_and_orders(
            src,
            dst_prefix_suffix,
            ldf,
            [".pdf", ".stp"],
            db,
            {},
            {},
            {"Laser": "1"},
            True,
            client=client,
            delivery_map={},
            footer_note=DEFAULT_FOOTER_NOTE,
            date_prefix_exports=True,
            date_suffix_exports=True,
        )
        assert cnt_prefix_suffix == 2
        prod_folder_prefix_suffix = os.path.join(dst_prefix_suffix, "Laser")
        prefix_suffix_pdf = f"{date_token}-PN1-{date_token}.pdf"
        prefix_suffix_stp = f"{date_token}-PN1-{date_token}.stp"
        assert "_" not in prefix_suffix_pdf and "_" not in prefix_suffix_stp
        assert os.path.exists(
            os.path.join(prod_folder_prefix_suffix, prefix_suffix_pdf)
        )
        assert os.path.exists(
            os.path.join(prod_folder_prefix_suffix, prefix_suffix_stp)
        )
    print("All tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(run_tests())


def test_order_excel_partnumber_wrap_text(tmp_path):
    path = tmp_path / "order.xlsx"
    items = [
        {
            "PartNumber": "PN1-THIS-IS-A-VERY-LONG-CODE-OVER-25CHARS",
            "Description": "Lange omschrijving",
            "Materiaal": "S235JR",
            "Aantal": 1,
            "Oppervlakte": "1,23",
            "Gewicht": "4,56",
        }
    ]
    write_order_excel(str(path), items, doc_number="BB-1")
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    header_row = None
    for row in range(1, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == "PartNumber":
            header_row = row
            break
    assert header_row is not None, "PartNumber header niet gevonden"
    pn_cell = ws.cell(row=header_row + 1, column=1)
    assert pn_cell.alignment is not None
    assert pn_cell.alignment.wrap_text
