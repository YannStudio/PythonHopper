import pytest

from models import Supplier
from orders import REPORTLAB_OK, generate_pdf_order_platypus
from orders.core import format_supplier_address


def _one_line_items() -> list[dict[str, object]]:
    return [
        {
            "PartNumber": "PN-1",
            "Description": "Onderdeel",
            "Materiaal": "",
            "Aantal": 1,
            "Oppervlakte": "",
            "Gewicht": "",
        }
    ]


def test_supplier_address_skips_repeated_postcode_city():
    supplier = Supplier(
        supplier="ML Coating",
        adres_1="Boutersemdreef 30/1",
        adres_2="B-2240 Zandhoven",
        postcode="2240",
        gemeente="Zandhoven",
        land="Belgie",
    )

    assert (
        format_supplier_address(supplier)
        == "Boutersemdreef 30/1, B-2240 Zandhoven, Belgie"
    )


def test_supplier_address_keeps_postcode_city_when_missing_from_address_2():
    supplier = Supplier(
        supplier="Supplier BV",
        adres_1="Industrieweg 1",
        adres_2="Unit 2",
        postcode="2000",
        gemeente="Antwerpen",
        land="Belgie",
    )

    assert (
        format_supplier_address(supplier)
        == "Industrieweg 1, Unit 2, 2000 Antwerpen, Belgie"
    )


@pytest.mark.skipif(not REPORTLAB_OK, reason="ReportLab niet beschikbaar")
def test_pdf_supplier_address_does_not_repeat_postcode_city(tmp_path):
    supplier = Supplier(
        supplier="ML Coating",
        adres_1="Boutersemdreef 30/1",
        adres_2="B-2240 Zandhoven",
        postcode="2240",
        gemeente="Zandhoven",
        land="Belgie",
        sales_email="info@mlcoating.be",
        phone="+32 476 71 13 45",
    )
    pdf_path = tmp_path / "order.pdf"

    generate_pdf_order_platypus(
        str(pdf_path),
        {},
        supplier,
        production="Bar Chair",
        items=_one_line_items(),
        doc_type="Offerteaanvraag",
        doc_number="OFF001",
    )

    from PyPDF2 import PdfReader

    text = "\n".join(page.extract_text() or "" for page in PdfReader(str(pdf_path)).pages)
    assert text.count("2240 Zandhoven") == 1
