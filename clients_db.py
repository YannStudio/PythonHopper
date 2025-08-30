import os
import json
import logging
from dataclasses import asdict
from typing import List, Optional

from models import Client

CLIENTS_DB_FILE = "clients_db.json"

logger = logging.getLogger(__name__)


class ClientsDB:
    def __init__(self, clients: Optional[List[Client]] = None):
        self.clients: List[Client] = clients or []

    @staticmethod
    def load(path: str = CLIENTS_DB_FILE) -> "ClientsDB":
        if not os.path.exists(path):
            return ClientsDB()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                recs = data
            else:
                recs = data.get("clients", [])
            clients = []
            for rec in recs:
                try:
                    clients.append(Client.from_any(rec))
                except Exception:
                    pass
            return ClientsDB(clients)
        except (OSError, json.JSONDecodeError) as exc:
            logger.error("Error loading clients DB: %s", exc)
            return ClientsDB()

    def save(self, path: str = CLIENTS_DB_FILE) -> None:
        data = {"clients": [asdict(c) for c in self.clients]}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def clients_sorted(self) -> List[Client]:
        return sorted(self.clients, key=lambda c: (not c.favorite, c.name.lower()))

    def find(self, query: str) -> List[Client]:
        q = (query or "").strip().lower()
        if not q:
            return self.clients_sorted()
        res = []
        for c in self.clients:
            hay = " ".join([
                c.name or "",
                c.address or "",
                c.vat or "",
                c.email or "",
            ]).lower()
            if q in hay:
                res.append(c)
        res.sort(key=lambda c: (not c.favorite, c.name.lower()))
        return res

    def display_name(self, c: Client) -> str:
        return f"{'★ ' if c.favorite else ''}{c.name}"

    def _idx_by_name(self, name: str) -> int:
        for i, c in enumerate(self.clients):
            if c.name.strip().lower() == str(name).strip().lower():
                return i
        return -1

    def upsert(self, client: Client) -> None:
        i = self._idx_by_name(client.name)
        if i >= 0:
            cur = self.clients[i]
            for f in asdict(client):
                val = getattr(client, f)
                if val not in (None, ""):
                    setattr(cur, f, val)
        else:
            self.clients.append(client)

    def remove(self, name: str) -> bool:
        i = self._idx_by_name(name)
        if i >= 0:
            self.clients.pop(i)
            return True
        return False

    def toggle_fav(self, name: str) -> bool:
        i = self._idx_by_name(name)
        if i >= 0:
            self.clients[i].favorite = not self.clients[i].favorite
            return True
        return False

    def get(self, name: str) -> Optional[Client]:
        i = self._idx_by_name(name)
        return self.clients[i] if i >= 0 else None
