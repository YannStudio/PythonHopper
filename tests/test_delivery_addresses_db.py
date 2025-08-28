from models import DeliveryAddress
from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_ADDRESSES_DB_FILE


def test_delivery_addresses_db_crud(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    db = DeliveryAddressesDB()
    addr = DeliveryAddress.from_any({
        "name": "Depot",
        "address": "Straat 1",
        "contact": "Jan",
        "phone": "0000",
        "email": "jan@example.com",
    })
    db.upsert(addr)
    db.save(DELIVERY_ADDRESSES_DB_FILE)
    assert db.get("Depot").contact == "Jan"

    db2 = DeliveryAddressesDB.load(DELIVERY_ADDRESSES_DB_FILE)
    assert db2.get("Depot").email == "jan@example.com"

    db2.toggle_fav("Depot")
    db2.remove("Depot")
    assert db2.get("Depot") is None
