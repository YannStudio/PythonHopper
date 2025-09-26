import base64
import importlib

import pytest

from models import DeliveryAddress, Supplier


PNG_RED_DOT = (
    "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAI0lEQVR4nGNgGAWjYBSMglEwCkb9/0foBAWjYBQAAAbBBf8HY2xwAAAAAElFTkSuQmCC"
)


def _write_logo(path):
    path.write_bytes(base64.b64decode(PNG_RED_DOT))


def _make_company_info(tmp_path, with_logo: bool = True):
    info = {
        "name": "ACME Corp",
        "address": "Main Street 1",
        "vat": "BE0123456789",
        "email": "info@example.com",
    }
    if with_logo:
        logo_path = tmp_path / "logo.png"
        _write_logo(logo_path)
        info["logo_path"] = str(logo_path)
    return info


def _generate_sample_order(orders_module, path, company_info):
    orders_module.generate_pdf_order_platypus(
        str(path),
        company_info,
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


def test_generate_pdf_order_platypus_keep_together_fallback(tmp_path):
    platypus = pytest.importorskip("reportlab.platypus")
    pytest.importorskip("PyPDF2")

    had_keeptogether = hasattr(platypus, "KeepTogether")
    original_keeptogether = getattr(platypus, "KeepTogether", None)
    if had_keeptogether:
        delattr(platypus, "KeepTogether")

    orders_module = importlib.import_module("orders")
    orders_module = importlib.reload(orders_module)

    try:
        assert orders_module.REPORTLAB_OK
        assert orders_module.KeepTogether is None

        out_path = tmp_path / "order_with_logo.pdf"
        _generate_sample_order(orders_module, out_path, _make_company_info(tmp_path))

        assert out_path.exists() and out_path.stat().st_size > 0

        import PyPDF2

        reader = PyPDF2.PdfReader(str(out_path))
        page_text = "".join((page.extract_text() or "") for page in reader.pages)
        assert "ACME Corp" in page_text

        out_text_only = tmp_path / "order_text_only.pdf"
        _generate_sample_order(orders_module, out_text_only, _make_company_info(tmp_path, False))
        assert out_text_only.exists() and out_text_only.stat().st_size > 0

        reader_no_logo = PyPDF2.PdfReader(str(out_text_only))
        page_text_no_logo = "".join((page.extract_text() or "") for page in reader_no_logo.pages)
        assert "ACME Corp" in page_text_no_logo

    finally:
        if had_keeptogether:
            platypus.KeepTogether = original_keeptogether
        importlib.reload(orders_module)
