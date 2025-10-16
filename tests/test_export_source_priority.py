from __future__ import annotations

import pandas as pd

from helpers import _build_file_index
from models import Supplier
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    return SuppliersDB([Supplier.from_any({"supplier": "ACME"})])


def test_file_index_prioritizes_non_bundle_sources(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    bundle_dir = src / "2024-01-01_Project-Alpha"
    bundle_dir.mkdir()

    real = src / "PN1.pdf"
    bundle_copy = bundle_dir / "PN1.pdf"
    real.write_text("new")
    bundle_copy.write_text("old")

    idx = _build_file_index(str(src), [".pdf"])
    assert idx["PN1"][-1] == str(real)


def test_copy_prefers_files_outside_bundle_directories(tmp_path):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    bundle_dir = src / "2024-01-01_Project-Alpha"
    bundle_dir.mkdir()

    real = src / "PN1.pdf"
    bundle_copy = bundle_dir / "PN1.pdf"
    real.write_text("new")
    bundle_copy.write_text("old")

    bom_df = pd.DataFrame(
        [{"PartNumber": "PN1", "Description": "", "Production": "Laser"}]
    )

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        bom_df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
    )

    assert cnt == 2
    exported = next((dest / "Laser").glob("*.pdf"))
    assert exported.read_text() == "new"
