import os
import json
from dataclasses import asdict
from typing import List, Dict, Optional

from models import Supplier
from app_paths import data_file
from helpers import favorite_prefix

SUPPLIERS_DB_FILE = data_file("suppliers_db.json")
_FAVORITE_PREFIX = favorite_prefix()


class SuppliersDB:
    def __init__(
        self,
        suppliers: List[Supplier] = None,
        defaults_by_production: Dict[str, str] = None,
        defaults_by_finish: Dict[str, str] = None,
    ):
        self.suppliers: List[Supplier] = suppliers or []
        self.defaults_by_production: Dict[str, str] = defaults_by_production or {}
        self.defaults_by_finish: Dict[str, str] = defaults_by_finish or {}

    @staticmethod
    def load(path: str = SUPPLIERS_DB_FILE) -> "SuppliersDB":
        if not os.path.exists(path):
            return SuppliersDB()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):  # backward compat
                sups = [Supplier(supplier=s) for s in data]
                return SuppliersDB(sups, {}, {})
            sups_raw = data.get("suppliers", [])
            sups = []
            for rec in sups_raw:
                try:
                    sups.append(Supplier.from_any(rec))
                except Exception:
                    pass
            defaults = data.get("defaults_by_production", {}) or {}
            finish_defaults = data.get("defaults_by_finish", {}) or {}
            return SuppliersDB(sups, defaults, finish_defaults)
        except Exception:
            return SuppliersDB()

    def save(self, path: str = SUPPLIERS_DB_FILE) -> None:
        data = {
            "suppliers": [asdict(s) for s in self.suppliers],
            "defaults_by_production": self.defaults_by_production,
            "defaults_by_finish": self.defaults_by_finish,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def suppliers_sorted(self) -> List[Supplier]:
        return sorted(self.suppliers, key=lambda s: (not s.favorite, s.supplier.lower()))

    def find(self, query: str, product_type_filter: Optional[str] = None, product_desc_filter: Optional[str] = None) -> List[Supplier]:
        """
        Find suppliers by query and optional product filters.
        Uses AND logic: all filters must match.
        
        Args:
            query: Text to search in supplier name and other fields
            product_type_filter: Optional product type filter (exact match)
            product_desc_filter: Optional product description filter (exact match, requires product_type_filter to be set)
        """
        q = (query or "").strip().lower()
        L = []
        
        # Normalize filter values
        pt_filter = (product_type_filter or "").strip().lower()
        pd_filter = (product_desc_filter or "").strip().lower()
        
        for s in self.suppliers:
            # Apply product type filter if provided (exact match)
            if pt_filter:
                supplier_product_type = (s.product_type or "").strip().lower()
                if supplier_product_type != pt_filter:
                    continue
            
            # Apply product description filter if provided (exact match)
            # AND logic: this is only checked if the product type already matched
            if pd_filter:
                supplier_product_desc = (s.product_description or "").strip().lower()
                if supplier_product_desc != pd_filter:
                    continue
            
            # If no query, include this supplier (all filters passed)
            if not q:
                L.append(s)
            else:
                # Check if supplier name starts with query (prefix match on name)
                supplier_name = (s.supplier or "").lower()
                if supplier_name.startswith(q):
                    L.append(s)
                else:
                    # Also search in other fields as substring
                    hay = " ".join([
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
                        s.product_type or "",
                        s.product_description or "",
                    ]).lower()
                    if q in hay:
                        L.append(s)
        
        L.sort(key=lambda s: (not s.favorite, s.supplier.lower()))
        return L

    def display_name(self, s: Supplier) -> str:
        prefix = _FAVORITE_PREFIX if s.favorite else ""
        return f"{prefix}{s.supplier}"

    def get_unique_product_types(self) -> List[str]:
        """Get list of unique product types from all suppliers, sorted."""
        types = set()
        for s in self.suppliers:
            if s.product_type and s.product_type.strip():
                types.add(s.product_type.strip())
        return sorted(list(types))

    def get_unique_product_descriptions(self) -> List[str]:
        """Get list of unique product descriptions from all suppliers, sorted."""
        descs = set()
        for s in self.suppliers:
            if s.product_description and s.product_description.strip():
                descs.add(s.product_description.strip())
        return sorted(list(descs))

    def get_product_descriptions_for_type(self, product_type: str) -> List[str]:
        """Get unique product descriptions filtered by product type, sorted."""
        if not product_type or not product_type.strip():
            return self.get_unique_product_descriptions()
        
        descs = set()
        for s in self.suppliers:
            if s.product_type and s.product_type.strip().lower() == product_type.strip().lower():
                if s.product_description and s.product_description.strip():
                    descs.add(s.product_description.strip())
        return sorted(list(descs))

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
            self.defaults_by_production = {
                k: v for k, v in self.defaults_by_production.items() if v.lower() != name.lower()
            }
            self.defaults_by_finish = {
                k: v for k, v in self.defaults_by_finish.items() if v.lower() != name.lower()
            }
            return True
        return False

    def clear_all(self):
        self.suppliers = []
        self.defaults_by_production = {}
        self.defaults_by_finish = {}

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

    def set_default_finish(self, finish_key: str, supplier_name: str) -> None:
        if self._idx_by_name(supplier_name) < 0:
            raise ValueError(f"Supplier '{supplier_name}' not found")
        self.defaults_by_finish[str(finish_key)] = supplier_name

    def get_default_finish(self, finish_key: str) -> Optional[str]:
        return self.defaults_by_finish.get(str(finish_key))
