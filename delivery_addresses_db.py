import json
import logging
import os
from dataclasses import asdict
from typing import List, Optional

from models import DeliveryAddress

DELIVERY_ADDRESSES_DB_FILE = "delivery_addresses_db.json"

logger = logging.getLogger(__name__)


class DeliveryAddressesDB:
    def __init__(self, addresses: Optional[List[DeliveryAddress]] = None):
        self.addresses: List[DeliveryAddress] = addresses or []

    @staticmethod
    def load(path: str = DELIVERY_ADDRESSES_DB_FILE) -> "DeliveryAddressesDB":
        if not os.path.exists(path):
            return DeliveryAddressesDB()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)


    def save(self, path: str = DELIVERY_ADDRESSES_DB_FILE) -> None:
        data = {"addresses": [asdict(a) for a in self.addresses]}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def addresses_sorted(self) -> List[DeliveryAddress]:

    def _idx_by_name(self, name: str) -> int:
        for i, a in enumerate(self.addresses):
            if a.name.strip().lower() == str(name).strip().lower():
                return i
        return -1

    def upsert(self, addr: DeliveryAddress) -> None:
        i = self._idx_by_name(addr.name)
        if i >= 0:

        else:
            self.addresses.append(addr)

    def remove(self, name: str) -> bool:
        i = self._idx_by_name(name)
        if i >= 0:
            self.addresses.pop(i)
            return True
        return False


    def get(self, name: str) -> Optional[DeliveryAddress]:
        i = self._idx_by_name(name)
        return self.addresses[i] if i >= 0 else None
