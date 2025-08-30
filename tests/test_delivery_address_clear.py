from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE
from models import DeliveryAddress


def test_clearing_remark_persists(tmp_path, monkeypatch):
    """Removing the remark should overwrite previous data with None."""
    # work inside temporary directory so we don't touch real files
    monkeypatch.chdir(tmp_path)

    # start with an address that has a remark
    db = DeliveryAddressesDB([
        DeliveryAddress(name="Test", address="Street", remarks="to clear"),
    ])
    db.save()  # create initial file

    # simulate editing the address and clearing the remark
    cleared = DeliveryAddress.from_any({"name": "Test", "address": "Street", "remarks": ""})
    db.upsert(cleared)
    db.save()

    # reload and ensure the remark is gone
    db2 = DeliveryAddressesDB.load(DELIVERY_DB_FILE)
    addr = db2.get("Test")
    assert addr is not None
    assert addr.remarks is None

