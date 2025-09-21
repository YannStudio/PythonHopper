import datetime
import os
from pathlib import Path

import pandas as pd
import pytest

from helpers import create_export_bundle
from models import Supplier
from orders import copy_per_production_and_orders
from suppliers_db import SuppliersDB


def test_create_export_bundle_formats_and_sanitises(tmp_path):
    ts = datetime.datetime(2023, 1, 2, 3, 4, 5)
    bundle = create_export_bundle(
        tmp_path,
        "Client: Demo/Project",
        timestamp=ts,
    )
    expected_name = "20230102-030405_Client_Demo_Project"
    assert bundle["name"] == expected_name
    assert Path(bundle["path"]).is_dir()
    latest = tmp_path / "latest"
    assert latest.is_symlink()
    assert os.readlink(latest) == expected_name


def test_create_export_bundle_suffix_and_limits(tmp_path):
    ts = datetime.datetime(2023, 6, 1, 12, 0, 0)
    first = create_export_bundle(tmp_path, "Project", timestamp=ts)
    assert first["name"].endswith("_Project")
    second = create_export_bundle(tmp_path, "Project", timestamp=ts)
    assert second["name"].endswith("_Project-2")

    # Force the helper to exhaust attempts
    third_root = tmp_path / "other"
    third_root.mkdir()
    conflict_name = ts.strftime("%Y%m%d-%H%M%S")
    (third_root / conflict_name).mkdir()
    with pytest.raises(RuntimeError):
        create_export_bundle(third_root, "", timestamp=ts, max_attempts=1)


def test_create_export_bundle_permission_error(tmp_path, monkeypatch):
    root = tmp_path / "locked"
    root.mkdir()

    original_makedirs = os.makedirs

    def fake_makedirs(path, *args, **kwargs):
        if os.path.abspath(path) != os.path.abspath(root):
            raise PermissionError("no write access")
        return original_makedirs(path, *args, **kwargs)

    monkeypatch.setattr(os, "makedirs", fake_makedirs)

    with pytest.raises(PermissionError):
        create_export_bundle(root, "demo")


def test_copy_per_production_and_orders_uses_bundle(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    db = SuppliersDB()
    db.upsert(Supplier.from_any({"supplier": "ACME"}))

    src = tmp_path / "src"
    src.mkdir()
    (src / "PN1.pdf").write_text("dummy")

    bom_df = pd.DataFrame(
        [{"PartNumber": "PN1", "Description": "", "Production": "Laser", "Aantal": 1}]
    )

    dest_root = tmp_path / "bundles"
    ts = datetime.datetime(2023, 7, 1, 8, 0, 0)
    bundle = create_export_bundle(dest_root, "Run", timestamp=ts)

    cnt, chosen, bundle_info = copy_per_production_and_orders(
        str(src),
        str(bundle["path"]),
        bom_df,
        [".pdf"],
        db,
        {},
        {},
        {},
        False,
        client=None,
        delivery_map={},
        bundle=bundle,
    )

    assert cnt == 1
    assert chosen == {"Laser": "ACME"}
    assert bundle_info["path"] == bundle["path"]
    prod_file = Path(bundle_info["path"]) / "Laser" / "PN1.pdf"
    assert prod_file.exists()
    latest = dest_root / "latest"
    assert latest.is_symlink()
    assert os.readlink(latest) == bundle_info["name"]
