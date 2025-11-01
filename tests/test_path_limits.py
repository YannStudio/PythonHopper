import os

import pandas as pd
import pytest

import orders
from models import Supplier
from suppliers_db import SuppliersDB


def test_fit_filename_shortens_when_path_too_long(tmp_path):
    long_dir = tmp_path
    # Create nested directories to make the absolute path relatively long.
    for idx in range(4):
        long_dir = long_dir / ("segment" + str(idx))
        long_dir.mkdir()

    filename = "Standard bon_BOM-07_Powder coating _ Git zwart - RAL 9005_2025-10-26.xlsx"
    limit = len(os.path.abspath(long_dir)) + 40
    safe_name = orders._fit_filename_within_path(str(long_dir), filename, max_path=limit)

    assert safe_name.endswith(".xlsx")
    assert len(os.path.join(os.path.abspath(long_dir), safe_name)) <= limit
    assert safe_name != filename


def test_fit_filename_returns_original_when_within_limit(tmp_path):
    filename = "Bestelbon_PN1_2025-10-26.pdf"
    limit = len(os.path.abspath(tmp_path)) + len(filename) + 10
    safe_name = orders._fit_filename_within_path(str(tmp_path), filename, max_path=limit)

    assert safe_name == filename


def test_fit_filename_raises_when_directory_too_long(tmp_path):
    # Choose a limit that makes the directory itself exceed the limit.
    limit = len(os.path.abspath(tmp_path))
    with pytest.raises(OSError):
        orders._fit_filename_within_path(str(tmp_path), "example.pdf", max_path=limit)


def test_copy_per_production_warns_about_path_limit(tmp_path, monkeypatch):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    long_prod = "Laser-" + "X" * 80
    long_finish = "Poedercoating " + "Y" * 60

    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Description": "Omschrijving",
                "Materiaal": "Staal",
                "Aantal": 1,
                "Oppervlakte": "",
                "Gewicht": "",
                "Production": long_prod,
                "Finish": long_finish,
                "RAL color": "RAL 9005",
            }
        ]
    )

    db = SuppliersDB([Supplier.from_any({"supplier": "ACME"})])

    warnings: list[str] = []
    monkeypatch.setattr(orders, "_WINDOWS_MAX_PATH", len(os.path.abspath(dest)) + 80)
    monkeypatch.setattr(orders, "write_order_excel", lambda *args, **kwargs: None)
    monkeypatch.setattr(orders, "generate_pdf_order_platypus", lambda *args, **kwargs: None)

    orders.copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [],
        db,
        {long_prod: "ACME"},
        {},
        {},
        False,
        path_limit_warnings=warnings,
    )

    assert warnings, "er worden pad-waarschuwingen verwacht"
    assert any("Productie" in msg for msg in warnings)
    assert any("Afwerking" in msg for msg in warnings)
