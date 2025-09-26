import importlib

import pytest

from models import DeliveryAddress, Supplier


def test_generate_pdf_order_requires_reportlab(monkeypatch, tmp_path):
    orders_module = importlib.import_module("orders")
    orders_module = importlib.reload(orders_module)

    monkeypatch.setattr(orders_module, "REPORTLAB_OK", False, raising=False)

    out_path = tmp_path / "order.pdf"

    with pytest.raises(orders_module.ReportLabUnavailableError) as excinfo:
        orders_module.generate_pdf_order_platypus(
            str(out_path),
            {
                "name": "ACME Corp",
                "address": "Main Street 1",
                "vat": "BE0123456789",
                "email": "info@example.com",
            },
            Supplier(
                supplier="Supplier",
                adres_1="Street 12",
                postcode="1000",
                gemeente="Brussels",
            ),
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

    assert "ReportLab" in str(excinfo.value)
    assert not out_path.exists()

    importlib.reload(orders_module)
