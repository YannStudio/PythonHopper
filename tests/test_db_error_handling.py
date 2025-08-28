import json
import pytest

from clients_db import ClientsDB, CLIENTS_DB_FILE
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from delivery_addresses_db import (
    DeliveryAddressesDB,
    DELIVERY_ADDRESSES_DB_FILE,
)


def _write_corrupt(path):
    path.write_text('{ bad json', encoding='utf-8')


@pytest.mark.parametrize(
    "cls, filename",
    [
        (ClientsDB, CLIENTS_DB_FILE),
        (SuppliersDB, SUPPLIERS_DB_FILE),
        (DeliveryAddressesDB, DELIVERY_ADDRESSES_DB_FILE),
    ],
)
def test_load_corrupt_json(tmp_path, monkeypatch, cls, filename):
    monkeypatch.chdir(tmp_path)
    file_path = tmp_path / filename
    _write_corrupt(file_path)
    with pytest.raises(RuntimeError) as exc_info:
        cls.load(filename)
    msg = str(exc_info.value)
    assert filename in msg
    # ensure original JSON error is attached
    assert isinstance(exc_info.value.__cause__, json.JSONDecodeError)
