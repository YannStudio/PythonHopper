from models import Supplier
from suppliers_db import SuppliersDB


def test_supplier_manual_replace_clears_address_fields():
    db = SuppliersDB(
        [
            Supplier(
                supplier="ACME",
                adres_1="Street 1",
                postcode="2000",
                gemeente="Antwerpen",
                land="Belgie",
            )
        ]
    )

    edited = Supplier.from_any(
        {
            "supplier": "ACME",
            "adres_1": "Street 1",
            "postcode": "",
            "gemeente": "",
            "land": "",
        }
    )
    db.upsert(edited, replace=True)

    supplier = db.suppliers[0]
    assert supplier.postcode is None
    assert supplier.gemeente is None
    assert supplier.land is None


def test_supplier_merge_keeps_existing_address_fields_when_source_is_empty():
    db = SuppliersDB(
        [
            Supplier(
                supplier="ACME",
                postcode="2000",
                gemeente="Antwerpen",
                land="Belgie",
            )
        ]
    )

    csv_record = Supplier.from_any(
        {
            "supplier": "ACME",
            "postcode": "",
            "gemeente": "",
            "land": "",
        }
    )
    db.upsert(csv_record)

    supplier = db.suppliers[0]
    assert supplier.postcode == "2000"
    assert supplier.gemeente == "Antwerpen"
    assert supplier.land == "Belgie"


def test_supplier_from_any_keeps_product_description_separate():
    supplier = Supplier.from_any(
        {
            "supplier": "ACME",
            "description": "General supplier note",
            "product_description": "Steel profiles",
        }
    )

    assert supplier.description == "General supplier note"
    assert supplier.product_description == "Steel profiles"
