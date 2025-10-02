import pandas as pd
import pytest

import cli
from app_settings import AppSettings
from bom import load_bom
from cli import build_parser, cli_copy_per_prod
from models import Supplier
from orders import copy_per_production_and_orders, _normalize_finish_folder
from suppliers_db import SuppliersDB
from clients_db import ClientsDB
from delivery_addresses_db import DeliveryAddressesDB


def _make_db() -> SuppliersDB:
    return SuppliersDB([Supplier.from_any({"supplier": "ACME"})])


def test_load_bom_finish_columns_and_copy(tmp_path):
    bom_path = tmp_path / "bom.csv"
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Production": "Laser",
                "finish": "Poedercoating",
                "RAL COLOR": "RAL 9005",
                "Aantal": 2,
            },
            {
                "PartNumber": "PN2",
                "Production": "Laser",
                "finish": "Geanodiseerd",
                "RAL COLOR": None,
                "Aantal": 1,
            },
        ]
    ).to_csv(bom_path, index=False)

    for pn, content in {"PN1": "one", "PN2": "two"}.items():
        (src / f"{pn}.pdf").write_text(content, encoding="utf-8")

    loaded = load_bom(str(bom_path))

    assert list(loaded["Finish"]) == ["Poedercoating", "Geanodiseerd"]
    assert list(loaded["RAL color"]) == ["RAL 9005", ""]

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        loaded,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        copy_finish_exports=True,
    )

    assert cnt == 2

    finish_dir_1 = dest / (
        "Finish-"
        + _normalize_finish_folder("Poedercoating")
        + "-"
        + _normalize_finish_folder("RAL 9005")
    )
    finish_dir_2 = dest / ("Finish-" + _normalize_finish_folder("Geanodiseerd"))

    assert (finish_dir_1 / "PN1.pdf").is_file()
    assert (finish_dir_2 / "PN2.pdf").is_file()


@pytest.mark.parametrize("zip_parts", [False, True])
def test_finish_exports_written(tmp_path, zip_parts):
    src = tmp_path / "src"
    dest = tmp_path / "dest"
    src.mkdir()
    dest.mkdir()

    (src / "PN1.pdf").write_text("one", encoding="utf-8")
    (src / "PN2.pdf").write_text("two", encoding="utf-8")
    (src / "PN3.pdf").write_text("three", encoding="utf-8")
    (src / "PN4.pdf").write_text("four", encoding="utf-8")

    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Production": "Laser",
                "Finish": "Poedercoating",
                "RAL color": "RAL 9005",
                "Aantal": 1,
            },
            {
                "PartNumber": "PN2",
                "Production": "Laser",
                "Finish": "Anodized/Marine",
                "RAL color": "RAL-9010",
                "Aantal": 1,
            },
            {
                "PartNumber": "PN1",
                "Production": "Laser",
                "Finish": "Poedercoating",
                "RAL color": "RAL 9005",
                "Aantal": 3,
            },
            {
                "PartNumber": "PN3",
                "Production": "Laser",
                "Finish": " ",
                "RAL color": "RAL 9016",
                "Aantal": 2,
            },
            {
                "PartNumber": "PN4",
                "Production": "Laser",
                "Finish": "Geanodiseerd",
                "RAL color": "",
                "Aantal": 1,
            },
        ]
    )

    cnt, _ = copy_per_production_and_orders(
        str(src),
        str(dest),
        df,
        [".pdf"],
        _make_db(),
        {"Laser": "ACME"},
        {},
        {},
        False,
        zip_parts=zip_parts,
        copy_finish_exports=True,
    )

    assert cnt == 4

    finish_dir_1_name = (
        "Finish-"
        + _normalize_finish_folder("Poedercoating")
        + "-"
        + _normalize_finish_folder("RAL 9005")
    )
    finish_dir_2_name = (
        "Finish-"
        + _normalize_finish_folder("Anodized/Marine")
        + "-"
        + _normalize_finish_folder("RAL-9010")
    )
    finish_dir_1 = dest / finish_dir_1_name
    finish_dir_2 = dest / finish_dir_2_name
    finish_dir_3_name = "Finish-" + _normalize_finish_folder("Geanodiseerd")
    finish_dir_3 = dest / finish_dir_3_name

    assert (finish_dir_1 / "PN1.pdf").is_file()
    assert sorted(p.name for p in finish_dir_1.iterdir()) == ["PN1.pdf"]
    assert (finish_dir_2 / "PN2.pdf").is_file()
    assert (finish_dir_3 / "PN4.pdf").is_file()
    finish_dirs = sorted(
        p.name for p in dest.iterdir() if p.is_dir() and p.name.startswith("Finish-")
    )
    assert finish_dirs == sorted(
        [finish_dir_1_name, finish_dir_2_name, finish_dir_3_name]
    )


@pytest.mark.parametrize(
    "flag, settings_value, expected",
    [
        (None, True, True),
        ("--finish-folders", False, True),
        ("--no-finish-folders", True, False),
    ],
)
def test_cli_finish_flag(monkeypatch, tmp_path, flag, settings_value, expected):
    parser = build_parser()
    args_list = [
        "copy-per-prod",
        "--source",
        str(tmp_path / "src"),
        "--dest",
        str(tmp_path / "dst"),
        "--bom",
        str(tmp_path / "bom.xlsx"),
        "--exts",
        "pdf",
    ]
    if flag:
        args_list.append(flag)
    args = parser.parse_args(args_list)

    (tmp_path / "src").mkdir()
    (tmp_path / "dst").mkdir()

    df = pd.DataFrame(
        [
            {
                "PartNumber": "PN1",
                "Description": "",
                "Production": "Laser",
                "Aantal": 1,
            }
        ]
    )
    monkeypatch.setattr(cli, "load_bom", lambda path: df)

    sdb = _make_db()
    monkeypatch.setattr(SuppliersDB, "load", classmethod(lambda cls, path: sdb))
    cdb = ClientsDB([])
    monkeypatch.setattr(ClientsDB, "load", classmethod(lambda cls, path: cdb))
    ddb = DeliveryAddressesDB([])
    monkeypatch.setattr(DeliveryAddressesDB, "load", classmethod(lambda cls, path: ddb))

    monkeypatch.setattr(
        AppSettings,
        "load",
        classmethod(lambda cls, path=...: AppSettings(copy_finish_exports=settings_value)),
    )

    captured = {}

    def fake_copy(*_args, **kwargs):
        captured.update(kwargs)
        return 0, {}

    monkeypatch.setattr(cli, "copy_per_production_and_orders", fake_copy)

    cli_copy_per_prod(args)

    assert captured["copy_finish_exports"] is expected
