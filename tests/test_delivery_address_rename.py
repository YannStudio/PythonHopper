from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE
from models import DeliveryAddress
from cli import build_parser, cli_delivery_addresses


def test_upsert_allows_rename(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    db = DeliveryAddressesDB([
        DeliveryAddress(name="Old", address="Street 1"),
    ])
    db.save()
    db.upsert(DeliveryAddress(name="New", address="Street 1"), old_name="Old")
    db.save()
    db2 = DeliveryAddressesDB.load(DELIVERY_DB_FILE)
    assert db2.get("Old") is None
    renamed = db2.get("New")
    assert renamed is not None
    assert renamed.address == "Street 1"


def test_cli_rename(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    db = DeliveryAddressesDB([
        DeliveryAddress(name="Old", address="Street 1"),
    ])
    db.save()
    parser = build_parser()
    args = parser.parse_args(["delivery-addresses", "rename", "Old", "New"])
    cli_delivery_addresses(args)
    db2 = DeliveryAddressesDB.load(DELIVERY_DB_FILE)
    assert db2.get("Old") is None
    renamed = db2.get("New")
    assert renamed is not None
    assert renamed.address == "Street 1"
