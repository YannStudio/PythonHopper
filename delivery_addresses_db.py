import os
import json
from dataclasses import asdict
from typing import List, Optional

from models import DeliveryAddress
from clients_db import ClientsDB, CLIENTS_DB_FILE

DELIVERY_DB_FILE = "delivery_addresses_db.json"


class DeliveryAddressesDB:
    def __init__(self, addresses: Optional[List[DeliveryAddress]] = None):
        self.addresses: List[DeliveryAddress] = addresses or []

    @staticmethod
    def _copy_from_clients() -> List[DeliveryAddress]:
        cdb = ClientsDB.load(CLIENTS_DB_FILE)
        addresses: List[DeliveryAddress] = []
        for c in cdb.clients:
            if c.address:
                addresses.append(
                    DeliveryAddress(
                        name=c.name,
                        address=c.address,
                        favorite=c.favorite,
                    )
                )
        return addresses

    @staticmethod
    def load(path: str = DELIVERY_DB_FILE) -> "DeliveryAddressesDB":
        if not os.path.exists(path):
            return DeliveryAddressesDB(DeliveryAddressesDB._copy_from_clients())
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                recs = data
            else:
                recs = data.get("addresses", [])
            addresses = []
            for rec in recs:
                try:
                    if isinstance(rec, dict) and "client" in rec:
                        rec = {k: v for k, v in rec.items() if k != "client"}
                    addresses.append(DeliveryAddress.from_any(rec))
                except Exception:
                    pass
            if not addresses:
                addresses = DeliveryAddressesDB._copy_from_clients()
            return DeliveryAddressesDB(addresses)
        except Exception:
            return DeliveryAddressesDB(DeliveryAddressesDB._copy_from_clients())

    def save(self, path: str = DELIVERY_DB_FILE) -> None:
        data = {"addresses": [asdict(a) for a in self.addresses]}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def addresses_sorted(self) -> List[DeliveryAddress]:
        return sorted(self.addresses, key=lambda a: (not a.favorite, a.name.lower()))

    def find(self, query: str) -> List[DeliveryAddress]:
        q = (query or "").strip().lower()
        if not q:
            return self.addresses_sorted()
        res = []
        for a in self.addresses:
            hay = " ".join([a.name or "", a.address or "", a.remarks or ""]).lower()
            if q in hay:
                res.append(a)
        res.sort(key=lambda a: (not a.favorite, a.name.lower()))
        return res

    def display_name(self, a: DeliveryAddress) -> str:
        return f"{'★ ' if a.favorite else ''}{a.name}"

    def _idx_by_name(self, name: str) -> int:
        for i, a in enumerate(self.addresses):
            if a.name.strip().lower() == str(name).strip().lower():
                return i
        return -1

    def upsert(self, addr: DeliveryAddress, old_name: Optional[str] = None) -> None:
        key = old_name or addr.name
        i = self._idx_by_name(key)
        if i >= 0:
            cur = self.addresses[i]
            for f in asdict(addr):
                # Always overwrite existing values, even with None/"" to clear fields
                setattr(cur, f, getattr(addr, f))
            if old_name and old_name.strip().lower() != addr.name.strip().lower():
                self.addresses.pop(i)
                self.addresses.append(cur)
        else:
            self.addresses.append(addr)

    def remove(self, name: str) -> bool:
        i = self._idx_by_name(name)
        if i >= 0:
            self.addresses.pop(i)
            return True
        return False

    def toggle_fav(self, name: str) -> bool:
        i = self._idx_by_name(name)
        if i >= 0:
            self.addresses[i].favorite = not self.addresses[i].favorite
            return True
        return False

    def get(self, name: str) -> Optional[DeliveryAddress]:
        i = self._idx_by_name(name)
        return self.addresses[i] if i >= 0 else None
