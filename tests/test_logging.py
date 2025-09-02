import logging

from cli import build_parser, cli_suppliers
from main import main
from suppliers_db import SuppliersDB


def test_cli_suppliers_logs(monkeypatch, caplog):
    parser = build_parser()
    args = parser.parse_args(["suppliers", "list"])
    monkeypatch.setattr(SuppliersDB, "load", classmethod(lambda cls, path: SuppliersDB([])))
    with caplog.at_level(logging.INFO):
        cli_suppliers(args)
    assert "(geen leveranciers)" in caplog.text


def test_verbose_sets_debug_level(monkeypatch):
    monkeypatch.setattr(SuppliersDB, "load", classmethod(lambda cls, path: SuppliersDB([])))
    logging.getLogger().handlers.clear()
    main(["suppliers", "list"])
    assert logging.getLogger().getEffectiveLevel() == logging.INFO
    logging.getLogger().handlers.clear()
    main(["--verbose", "suppliers", "list"])
    assert logging.getLogger().getEffectiveLevel() == logging.DEBUG

