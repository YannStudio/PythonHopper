import cli
from cli import build_parser, cli_suppliers, cli_clients
from suppliers_db import SuppliersDB
from clients_db import ClientsDB


def test_cli_supplier_add_valid_vat(monkeypatch, capsys):
    parser = build_parser()
    args = parser.parse_args(["suppliers", "add", "ACME", "--btw", "BE0123456789"])
    db = SuppliersDB([])
    monkeypatch.setattr(SuppliersDB, "load", staticmethod(lambda *a, **kw: db))
    monkeypatch.setattr(db, "save", lambda *a, **kw: None)
    rc = cli_suppliers(args)
    assert rc == 0
    assert db.suppliers[0].btw == "BE0123456789"


def test_cli_supplier_add_invalid_vat(monkeypatch, capsys):
    parser = build_parser()
    args = parser.parse_args(["suppliers", "add", "ACME", "--btw", "INVALID"])
    db = SuppliersDB([])
    monkeypatch.setattr(SuppliersDB, "load", staticmethod(lambda *a, **kw: db))
    monkeypatch.setattr(db, "save", lambda *a, **kw: None)
    rc = cli_suppliers(args)
    captured = capsys.readouterr()
    assert rc == 2
    assert "Ongeldig BTW-nummer" in captured.out
    assert db.suppliers == []


def test_cli_client_add_valid_vat(monkeypatch, capsys):
    parser = build_parser()
    args = parser.parse_args(["clients", "add", "Foo", "--vat", "BE0123456789"])
    db = ClientsDB([])
    monkeypatch.setattr(ClientsDB, "load", staticmethod(lambda *a, **kw: db))
    monkeypatch.setattr(db, "save", lambda *a, **kw: None)
    rc = cli_clients(args)
    assert rc == 0
    assert db.clients[0].vat == "BE0123456789"


def test_cli_client_add_invalid_vat(monkeypatch, capsys):
    parser = build_parser()
    args = parser.parse_args(["clients", "add", "Foo", "--vat", "BADVAT"])
    db = ClientsDB([])
    monkeypatch.setattr(ClientsDB, "load", staticmethod(lambda *a, **kw: db))
    monkeypatch.setattr(db, "save", lambda *a, **kw: None)
    rc = cli_clients(args)
    captured = capsys.readouterr()
    assert rc == 2
    assert "Ongeldig BTW-nummer" in captured.out
    assert db.clients == []
