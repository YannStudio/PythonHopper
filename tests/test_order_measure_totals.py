import pandas as pd
import pytest

import orders
from models import Supplier
from suppliers_db import SuppliersDB


def test_write_order_excel_adds_surface_and_weight_total_row(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")

    path = tmp_path / "bestelbon.xlsx"
    orders.write_order_excel(
        str(path),
        [
            {
                "PartNumber": "PN-1",
                "Description": "Onderdeel",
                "Materiaal": "S235",
                "Aantal": 4,
                "Oppervlakte": "0,50",
                "Gewicht": "7,80",
            }
        ],
        context_label="Laser",
        total_surface_m2=2.0,
        total_weight_kg=31.2,
    )

    ws = openpyxl.load_workbook(path).active
    header_row = next(
        row for row in range(1, ws.max_row + 1) if ws.cell(row, 1).value == "PartNumber"
    )
    assert ws.cell(header_row, 5).value == "m\u00b2/st"
    assert ws.cell(header_row, 6).value == "kg/st"
    total_row = next(
        row for row in range(1, ws.max_row + 1) if ws.cell(row, 1).value == "Totaal"
    )

    assert str(ws.cell(total_row, 5).value).replace(",", ".") == "2.00"
    assert str(ws.cell(total_row, 6).value).replace(",", ".") == "31.20"


def test_write_order_excel_drops_empty_standard_material_column(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")

    path = tmp_path / "assembly.xlsx"
    orders.write_order_excel(
        str(path),
        [
            {
                "PartNumber": "ASM-1",
                "Description": "Assembly item",
                "Materiaal": "",
                "Aantal": 1,
                "Oppervlakte": "1.23",
                "Gewicht": "4.56",
            }
        ],
        context_label="Assembly",
        total_surface_m2=1.23,
        total_weight_kg=4.56,
    )

    ws = openpyxl.load_workbook(path).active
    header_row = next(
        row for row in range(1, ws.max_row + 1) if ws.cell(row, 1).value == "PartNumber"
    )
    headers = [
        ws.cell(header_row, column).value
        for column in range(1, ws.max_column + 1)
        if ws.cell(header_row, column).value
    ]

    assert headers == ["PartNumber", "Description", "Aantal", "m\u00b2/st", "kg/st"]
    assert "Materiaal" not in headers


def test_standard_order_route_calculates_surface_and_weight_from_quantity(
    tmp_path, monkeypatch
):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN-1",
                "Production": "Laser",
                "Description": "Onderdeel",
                "Materiaal": "S235",
                "Length profile": "2500",
                "Aantal": 4,
                "Oppervlakte": "0,50",
                "Gewicht": "7,80",
            }
        ]
    )
    captured = {}

    def fake_excel(_path, _items, *_args, **kwargs):
        captured["excel_sections"] = kwargs.get("sections")

    def fake_pdf(_path, _company, _supplier, _context, _items, **kwargs):
        captured["pdf_sections"] = kwargs.get("sections")

    monkeypatch.setattr(orders, "write_order_excel", fake_excel)
    monkeypatch.setattr(orders, "generate_pdf_order_platypus", fake_pdf)

    orders.copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [],
        SuppliersDB([Supplier.from_any({"supplier": "ACME"})]),
        {"Laser": "ACME"},
        {},
        {},
        remember_defaults=False,
        export_bom=False,
        export_related_files=False,
    )

    excel_section = captured["excel_sections"][0]
    pdf_section = captured["pdf_sections"][0]
    assert excel_section.total_surface_m2 == pytest.approx(2.0)
    assert excel_section.total_weight_kg == pytest.approx(31.2)
    assert pdf_section.total_surface_m2 == pytest.approx(2.0)
    assert pdf_section.total_weight_kg == pytest.approx(31.2)
    assert excel_section.items[0]["Lengte"] == "2500"
    assert pdf_section.items[0]["Lengte"] == "2500"
