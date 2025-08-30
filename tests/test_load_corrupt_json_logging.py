import logging
import pytest

from suppliers_db import SuppliersDB
from clients_db import ClientsDB
from delivery_addresses_db import DeliveryAddressesDB


@pytest.mark.parametrize("loader", [
    SuppliersDB.load,
    ClientsDB.load,
    DeliveryAddressesDB.load,
])
def test_corrupted_json_logs_error(tmp_path, caplog, loader):
    path = tmp_path / "data.json"
    path.write_text("{bad json")
    with caplog.at_level(logging.ERROR):
        loader(str(path))
    assert any("Error loading" in rec.message for rec in caplog.records)
