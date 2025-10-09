from models import Supplier
from orders import pick_supplier_for_production
from suppliers_db import SuppliersDB


def test_pick_supplier_with_prefetched_list(tmp_path, monkeypatch):
    """Providing a cached supplier list keeps the selection identical."""

    monkeypatch.chdir(tmp_path)
    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))
    db.upsert(Supplier.from_any({"supplier": "BETA"}))

    overrides = {"Laser": "BETA"}
    prefetched = db.suppliers_sorted()

    direct = pick_supplier_for_production("Laser", db, overrides)
    cached = pick_supplier_for_production(
        "Laser", db, overrides, suppliers_sorted=prefetched
    )

    assert cached == direct


def test_pick_supplier_skips_sparepart(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))

    result = pick_supplier_for_production("SparePart", db, {})

    assert result.supplier == ""
