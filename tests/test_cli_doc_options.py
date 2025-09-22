import pandas as pd
from models import Supplier
from suppliers_db import SuppliersDB
from clients_db import ClientsDB
from delivery_addresses_db import DeliveryAddressesDB
import cli
from cli import build_parser, cli_copy_per_prod


def test_cli_doc_options_parsing(monkeypatch, tmp_path):
    parser = build_parser()
    args = parser.parse_args([
        "copy-per-prod",
        "--source", str(tmp_path / "src"),
        "--dest", str(tmp_path / "dst"),
        "--bom", str(tmp_path / "bom.xlsx"),
        "--exts", "pdf",
        "--doc-type", "Laser=Offerteaanvraag",
        "--doc-number", "Laser=123",
        "--doc-number", "Plasma=456",
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

    assert captured["doc_type_map"] == {"Laser": "Offerteaanvraag"}
    assert captured["doc_num_map"] == {"Laser": "123", "Plasma": "456"}
    assert captured["export_name_prefix_text"] == ""
    assert captured["export_name_suffix_text"] == ""
