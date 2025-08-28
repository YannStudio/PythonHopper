import json
import logging
import os
from dataclasses import asdict
from typing import Dict, List, Optional

from models import Supplier

SUPPLIERS_DB_FILE = "suppliers_db.json"

logger = logging.getLogger(__name__)


class SuppliersDB:
    def __init__(self, suppliers: List[Supplier] = None, defaults_by_production: Dict[str, str] = None):
        self.suppliers: List[Supplier] = suppliers or []
        self.defaults_by_production: Dict[str, str] = defaults_by_production or {}

    @staticmethod
    def load(path: str = SUPPLIERS_DB_FILE) -> "SuppliersDB":
        if not os.path.exists(path):
            return SuppliersDB()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):  # backward compat
                sups = [Supplier(supplier=s) for s in data]
                return SuppliersDB(sups, {})
            sups_raw = data.get("suppliers", [])
            sups = []
            for idx, rec in enumerate(sups_raw):
                try:
                    sups.append(Supplier.from_any(rec))
                except Exception as e:
                    print(
                        f"Fout bij leverancier record {idx}: {e}; data={rec}"
                    )
            defaults = data.get("defaults_by_production", {}) or {}
            return SuppliersDB(sups, defaults)
        except Exception as e:
            raise RuntimeError(f"Fout bij laden van {path}") from e

    def save(self, path: str = SUPPLIERS_DB_FILE) -> None:
        data = {
            "suppliers": [asdict(s) for s in self.suppliers],
            "defaults_by_production": self.defaults_by_production,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def suppliers_sorted(self) -> List[Supplier]:
        return sorted(self.suppliers, key=lambda s: (not s.favorite, s.supplier.lower()))

    def find(self, query: str) -> List[Supplier]:
        q = (query or "").strip().lower()
        if not q:
            return self.suppliers_sorted()
        L = []
        for s in self.suppliers:
            hay = " ".join([
                s.supplier or "",
                s.description or "",
                s.supplier_id or "",
                s.adres_1 or "",
                s.adres_2 or "",
                s.postcode or "",
                s.gemeente or "",
                s.land or "",
                s.btw or "",
                s.contact_sales or "",
                s.sales_email or "",
                s.phone or "",
            ]).lower()
            if q in hay:
                L.append(s)
        L.sort(key=lambda s: (not s.favorite, s.supplier.lower()))
        return L

    def display_name(self, s: Supplier) -> str:
        return f"{'★ ' if s.favorite else ''}{s.supplier}"

    def _idx_by_name(self, name: str) -> int:
        for i, s in enumerate(self.suppliers):
            if s.supplier.strip().lower() == str(name).strip().lower():
                return i
        return -1

    def add(self, supplier: Supplier) -> bool:
        if self._idx_by_name(supplier.supplier) >= 0:
            return False
        self.suppliers.append(supplier)
        return True

    def upsert(self, supplier: Supplier) -> None:
        """Update bestaande met dezelfde naam, anders voeg toe."""
        i = self._idx_by_name(supplier.supplier)
        if i >= 0:
            cur = self.suppliers[i]
            for f in asdict(supplier):
                val = getattr(supplier, f)
                if val not in (None, ""):
                    setattr(cur, f, val)
        else:
            self.suppliers.append(supplier)

    def remove(self, name: str) -> bool:
        i = self._idx_by_name(name)
        if i >= 0:
            self.suppliers.pop(i)
            self.defaults_by_production = {k: v for k, v in self.defaults_by_production.items() if v.lower() != name.lower()}
            return True
        return False

    def clear_all(self):
        self.suppliers = []
        self.defaults_by_production = {}

    def toggle_fav(self, name: str) -> bool:
        i = self._idx_by_name(name)
        if i >= 0:
            self.suppliers[i].favorite = not self.suppliers[i].favorite
            return True
        return False

    def set_default(self, production: str, supplier_name: str) -> None:
        if self._idx_by_name(supplier_name) < 0:
            raise ValueError(f"Supplier '{supplier_name}' not found")
        self.defaults_by_production[str(production)] = supplier_name

    def get_default(self, production: str) -> Optional[str]:
        return self.defaults_by_production.get(str(production))
