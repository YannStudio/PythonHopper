import argparse
import logging

import cli
import suppliers_db
from helpers import _to_str as original_to_str


def test_import_csv_logs_error_for_missing_supplier(tmp_path, caplog, monkeypatch, capsys):
    csv_path = tmp_path / "suppliers.csv"
    csv_path.write_text("Description\nOnly desc\n", encoding="latin1")

    db_path = tmp_path / "suppliers_db.json"
    monkeypatch.setattr(cli, "SUPPLIERS_DB_FILE", str(db_path))
    monkeypatch.setattr(suppliers_db, "SUPPLIERS_DB_FILE", str(db_path))

    def fake_to_str(x):
        if x is None:
            return "missing"
        return original_to_str(x)

    monkeypatch.setattr(cli, "_to_str", fake_to_str)

    args = argparse.Namespace(
        action="import-csv",
        csv=str(csv_path),
        btw=None,
        adres_1=None,
        adres_2=None,
        email=None,
        phone=None,
    )

    caplog.set_level(logging.ERROR)
    cli.cli_suppliers(args)

    out, err = capsys.readouterr()
    assert "Overgeslagen: 1" in out
    assert "Record 1" in caplog.text
    assert "Supplier name is missing in record." in caplog.text
