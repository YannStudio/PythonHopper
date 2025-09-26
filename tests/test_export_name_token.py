import re
import zipfile

import pandas as pd
import pytest

import cli
import orders
from cli import build_parser, cli_copy_per_prod
from clients_db import ClientsDB
from delivery_addresses_db import DeliveryAddressesDB
from models import Supplier
from orders import copy_per_production_and_orders, REPORTLAB_OK
from suppliers_db import SuppliersDB

requires_reportlab = pytest.mark.skipif(
    not REPORTLAB_OK, reason="ReportLab is vereist voor PDF-exporttests"
)


def _make_db() -> SuppliersDB:
    return SuppliersDB([
        Supplier.from_any({"supplier": "ACME"}),
    ])


def _build_bom() -> pd.DataFrame:
    return pd.DataFrame([
        {"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}
    ])


@pytest.mark.parametrize(
    ("prefix", "suffix", "expected"),
    [
        (False, True, "PN1-REV-A.pdf"),
        (True, False, "REV-A-PN1.pdf"),
        (True, True, "REV-A-PN1-REV-A.pdf"),
    ],
)
@requires_reportlab
def test_export_token_positions(tmp_path, monkeypatch, prefix, suffix, expected):
    monkeypatch.setattr(orders, "SUPPLIERS_DB_FILE", str(tmp_path / "suppliers.json"))

    src = tmp_path / "src"
    dest = tmp_path / f"dest_{prefix}_{suffix}"
    dest_zip = tmp_path / f"dest_zip_{prefix}_{suffix}"
    src.mkdir()
    dest.mkdir()
    dest_zip.mkdir()

    (src / "PN1.pdf").write_text("dummy")

    db = _make_db()
    bom_df = _build_bom()

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        bom_df,
        [".pdf"],
        db,
        {"Laser": ""},
        {},
        {},
        False,
        export_name_prefix_text="REV-A",
        export_name_prefix_enabled=prefix,
        export_name_suffix_text="REV-A",
        export_name_suffix_enabled=suffix,
    )
    assert cnt == 1
    exported = dest / "Laser" / expected
    assert exported.exists()

    cnt_zip, _ = copy_per_production_and_orders(
        str(src),
        str(dest_zip),
        bom_df,
        [".pdf"],
        db,
        {"Laser": ""},
        {},
        {},
        False,
        zip_parts=True,
        export_name_prefix_text="REV-A",
        export_name_prefix_enabled=prefix,
        export_name_suffix_text="REV-A",
        export_name_suffix_enabled=suffix,
    )
    assert cnt_zip == 1
    zip_dir = dest_zip / "Laser"
    zip_files = sorted(zip_dir.glob("Laser*.zip"))
    assert len(zip_files) == 1
    zip_path = zip_files[0]
    assert re.fullmatch(r"Laser(?:_.+)?", zip_path.stem)
    with zipfile.ZipFile(zip_path) as zf:
        assert expected in zf.namelist()
        info = zf.getinfo(expected)
        if getattr(zipfile, "zlib", None):
            assert info.compress_type == zipfile.ZIP_DEFLATED
        else:
            assert info.compress_type == zipfile.ZIP_STORED


@requires_reportlab
def test_export_token_disabled(tmp_path, monkeypatch):
    monkeypatch.setattr(orders, "SUPPLIERS_DB_FILE", str(tmp_path / "suppliers.json"))

    src = tmp_path / "src"
    dest = tmp_path / "dest_disabled"
    src.mkdir()
    dest.mkdir()
    (src / "PN1.pdf").write_text("dummy")

    db = _make_db()
    bom_df = _build_bom()

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        bom_df,
        [".pdf"],
        db,
        {"Laser": ""},
        {},
        {},
        False,
        export_name_prefix_text="REV-A",
        export_name_prefix_enabled=False,
        export_name_suffix_text="REV-A",
        export_name_suffix_enabled=False,
    )
    assert cnt == 1
    exported = dest / "Laser" / "PN1.pdf"
    assert exported.exists()


def test_cli_export_token_flags(monkeypatch, tmp_path):
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
        "--export-prefix-text",
        "REV-A",
        "--export-prefix-enabled",
        "--export-suffix-text",
        "REV-A",
        "--no-export-suffix-enabled",
    ])

    (tmp_path / "src").mkdir()
    (tmp_path / "dst").mkdir()
    monkeypatch.setattr(cli, "load_bom", lambda path: _build_bom())

    sdb = _make_db()
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

    assert captured["export_name_prefix_text"] == "REV-A"
    assert captured["export_name_prefix_enabled"] is True
    assert captured["export_name_suffix_text"] == "REV-A"
    assert captured["export_name_suffix_enabled"] is False
