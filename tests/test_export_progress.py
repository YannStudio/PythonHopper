import pandas as pd

import orders
from models import Supplier
from progress import ProgressEvent
from suppliers_db import SuppliersDB


def test_copy_per_production_reports_progress_events(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()
    (src / "PN-001.dxf").write_text("dxf", encoding="utf-8")

    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN-001",
                "Production": "Laser",
                "Description": "Part",
                "Aantal": 1,
            }
        ]
    )
    db = SuppliersDB([Supplier.from_any({"supplier": "ACME"})])
    events: list[ProgressEvent] = []

    monkeypatch.setattr(orders, "write_order_excel", lambda *args, **kwargs: None)
    monkeypatch.setattr(
        orders,
        "generate_pdf_order_platypus",
        lambda *args, **kwargs: None,
    )

    count, _chosen = orders.copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".dxf"],
        db,
        {"Laser": "ACME"},
        {},
        {},
        remember_defaults=False,
        export_bom=False,
        export_related_files=False,
        progress_callback=events.append,
    )

    phases = [event.phase for event in events]
    assert count == 1
    assert phases[0] == "prepare"
    assert phases[-1] == "done"
    assert {"scan", "productions", "documents", "save"}.issubset(phases)
    assert events[-1].normalized_percent() == 100
    assert all(event.normalized_percent() < 100 for event in events[:-1])
