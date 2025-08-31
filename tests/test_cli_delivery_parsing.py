import pandas as pd
from models import Supplier, DeliveryAddress
from suppliers_db import SuppliersDB
from clients_db import ClientsDB
from delivery_addresses_db import DeliveryAddressesDB
import cli
from cli import build_parser, cli_copy_per_prod


def test_cli_delivery_parsing(monkeypatch, tmp_path):
    parser = build_parser()
    args = parser.parse_args([
        "copy-per-prod",
        "--source", str(tmp_path / "src"),
        "--dest", str(tmp_path / "dst"),
        "--bom", str(tmp_path / "bom.xlsx"),
        "--exts", "pdf",
        "--delivery", "Laser=Addr1",
        "--delivery", "Plasma=Addr2",
    ])

    # minimal environment
    (tmp_path / "src").mkdir()
    (tmp_path / "dst").mkdir()
    df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}
    ])
    monkeypatch.setattr(cli, "load_bom", lambda path: df)

    sdb = SuppliersDB([Supplier.from_any({"supplier": "ACME"})])
    monkeypatch.setattr(SuppliersDB, "load", classmethod(lambda cls, path: sdb))
    cdb = ClientsDB([])
    monkeypatch.setattr(ClientsDB, "load", classmethod(lambda cls, path: cdb))
    ddb = DeliveryAddressesDB([
        DeliveryAddress(name="Addr1", address="A"),
        DeliveryAddress(name="Addr2", address="B"),
    ])
    monkeypatch.setattr(DeliveryAddressesDB, "load", classmethod(lambda cls, path: ddb))

    captured = {}

    def fake_copy(*args, **kwargs):
        captured.update(kwargs)
        return 0, {}

    monkeypatch.setattr(cli, "copy_per_production_and_orders", fake_copy)
    cli_copy_per_prod(args)
    assert set(captured["delivery_map"]) == {"Laser", "Plasma"}
    assert captured["delivery_map"]["Laser"].name == "Addr1"
    assert captured["delivery_map"]["Plasma"].name == "Addr2"


def test_cli_delivery_special_tokens(monkeypatch, tmp_path):
    parser = build_parser()
    args = parser.parse_args([
        "copy-per-prod",
        "--source",
        str(tmp_path / "src"),
        "--dest",
        str(tmp_path / "dst"),
        "--bom",
        str(tmp_path / "bom.xlsx"),
        "--exts",
        "pdf",
        "--delivery",
        "Laser=Bestelling wordt opgehaald",
        "--delivery",
        "Plasma=Leveradres wordt nog meegedeeld",
    ])

    (tmp_path / "src").mkdir()
    (tmp_path / "dst").mkdir()
    df = pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}
    ])
    monkeypatch.setattr(cli, "load_bom", lambda path: df)

    sdb = SuppliersDB([Supplier.from_any({"supplier": "ACME"})])
    monkeypatch.setattr(SuppliersDB, "load", classmethod(lambda cls, path: sdb))
    cdb = ClientsDB([])
    monkeypatch.setattr(ClientsDB, "load", classmethod(lambda cls, path: cdb))
    ddb = DeliveryAddressesDB([])
    monkeypatch.setattr(DeliveryAddressesDB, "load", classmethod(lambda cls, path: ddb))

    captured = {}

    def fake_copy(*args, **kwargs):
        captured.update(kwargs)
        return 0, {}

    monkeypatch.setattr(cli, "copy_per_production_and_orders", fake_copy)
    cli_copy_per_prod(args)
    assert set(captured["delivery_map"]) == {"Laser", "Plasma"}
    assert (
        captured["delivery_map"]["Laser"].name == "Bestelling wordt opgehaald"
    )
    assert (
        captured["delivery_map"]["Plasma"].name
        == "Leveradres wordt nog meegedeeld"
    )
