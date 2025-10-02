import pandas as pd
import pytest

import orders
import step_previews
from models import Supplier
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def _make_db() -> SuppliersDB:
    return SuppliersDB([Supplier.from_any({"supplier": "ACME"})])


def _build_bom() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Description": "",
                "Production": "Laser",
                "Aantal": 1,
            }
        ]
    )


def test_packlist_skipped_without_step_files(tmp_path):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    (src / "PN1.pdf").write_text("dummy")

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        _build_bom(),
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
    )
    assert cnt == 1
    packlists = list((dest / "Laser").glob("Paklijst_*.pdf"))
    assert packlists == []


@pytest.mark.skipif(not orders.REPORTLAB_OK, reason="ReportLab not available")
def test_packlist_skipped_when_renderer_missing(tmp_path, monkeypatch):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    (src / "PN1.step").write_text("dummy")

    captured: dict[str, object] = {}

    def fake_render(items, out_dir, size=step_previews.DEFAULT_SIZE):
        captured["items"] = list(items)
        return []

    monkeypatch.setattr(orders.step_previews, "render_step_files", fake_render)

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        _build_bom(),
        [".step"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
    )
    assert cnt == 1
    assert "items" in captured
    packlists = list((dest / "Laser").glob("Paklijst_*.pdf"))
    assert packlists == []
