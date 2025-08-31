import pandas as pd
from openpyxl import load_workbook

from models import Supplier, Client
from suppliers_db import SuppliersDB
from orders import copy_per_production_and_orders


def test_delivery_address_used_in_order(tmp_path, monkeypatch):
    """The selected delivery address should appear in the order document."""
    # operate within temporary directory to avoid side effects
    monkeypatch.chdir(tmp_path)

    # supplier database with one supplier
    db = SuppliersDB([Supplier.from_any({"supplier": "ACME"})])

    # create source and destination directories
    src = tmp_path / "src"
    dst = tmp_path / "dst"
    src.mkdir()
    dst.mkdir()

    # dummy source file
    (src / "PN1.pdf").write_text("dummy")

    # BOM with a single production entry
    bom_df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1},
    ])

    supplier_map = {"Laser": "ACME"}
    delivery_map = {"Laser": "Custom Street 5"}

    client = Client.from_any({"name": "Client", "address": "Base Addr"})

    cnt, chosen = copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        supplier_map,
        delivery_map,
        False,
        client=client,
    )

    assert cnt == 1

    # verify that the generated Excel order contains the chosen delivery address
    prod_dir = dst / "Laser"
    excel_files = list(prod_dir.glob("Bestelbon_Laser_*.xlsx"))
    assert excel_files, "Order Excel file not created"

    wb = load_workbook(excel_files[0])
    ws = wb.active
    # Address is the second header row, second column
    assert ws.cell(row=2, column=2).value == "Custom Street 5"

