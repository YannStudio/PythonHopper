import os
import pandas as pd

from models import Supplier
from suppliers_db import SuppliersDB
from orders import copy_per_production_and_orders


def test_offerteaanvraag_prefix(tmp_path):
    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))

    src = tmp_path / "src"
    dst = tmp_path / "dst"
    src.mkdir(); dst.mkdir()

    (src / "PN1.pdf").write_text("dummy")

    bom_df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}
    ])

    overrides = {"Laser": "ACME"}

    copy_per_production_and_orders(
        str(src),
        str(dst),
        bom_df,
        [".pdf"],
        db,
        overrides,
        False,
        doc_type="offerteaanvraag",
    )

    prod_folder = dst / "Laser"
    files = list(os.listdir(prod_folder))
    assert any(f.startswith("Offerteaanvraag_") and f.endswith(".pdf") for f in files)
    assert any(f.startswith("Offerteaanvraag_") and f.endswith(".xlsx") for f in files)
