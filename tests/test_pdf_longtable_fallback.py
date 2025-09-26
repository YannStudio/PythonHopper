import importlib

import pytest

from models import DeliveryAddress, Supplier


def test_generate_pdf_order_platypus_longtable_fallback(tmp_path):
    platypus = pytest.importorskip("reportlab.platypus")

    had_longtable = hasattr(platypus, "LongTable")
    original_longtable = getattr(platypus, "LongTable", None)
    if had_longtable:
        delattr(platypus, "LongTable")

    orders_module = importlib.import_module("orders")
    orders_module = importlib.reload(orders_module)

    try:
        assert orders_module.REPORTLAB_OK
        assert orders_module.LongTable is orders_module.Table

        out_path = tmp_path / "order.pdf"
        orders_module.generate_pdf_order_platypus(
            str(out_path),
            {
                "name": "ACME Corp",
                "address": "Main Street 1",
                "vat": "BE0123456789",
                "email": "info@example.com",
            },
            Supplier(supplier="Supplier", adres_1="Street 12", postcode="1000", gemeente="Brussels"),
            "PROD-1",
            [
                {
                    "PartNumber": "PN-1",
                    "Description": "Panel",
                    "Materiaal": "Steel",
                    "Aantal": "2",
                    "Oppervlakte": "1.5",
                    "Gewicht": "3.0",
                }
            ],
            delivery=DeliveryAddress(name="Site", address="Warehouse", remarks="Dock 5"),
        )

        assert out_path.exists()
        assert out_path.stat().st_size > 0
    finally:
        if had_longtable:
            platypus.LongTable = original_longtable
        importlib.reload(orders_module)
