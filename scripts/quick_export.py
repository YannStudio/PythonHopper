"""Quick script to programmatically create an order export and verify outputs.
Run from project root with: python scripts/quick_export.py
"""
import pathlib
import tempfile
import os
import importlib.util
from pathlib import Path

# Load excel_writer from file path (orders is not a proper package in-source)
base = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(base))

# Ensure a package namespace for `orders` so relative imports inside the
# module files (e.g. "from . import core") resolve correctly.
orders_pkg = type(sys)("orders")
orders_pkg.__path__ = [str(base / "orders")]
sys.modules["orders"] = orders_pkg

# Load orders.core first so relative import in excel_writer works
core_path = base / "orders" / "core.py"
if core_path.exists():
    spec_core = importlib.util.spec_from_file_location("orders.core", str(core_path))
    mod_core = importlib.util.module_from_spec(spec_core)
    sys.modules["orders.core"] = mod_core
    spec_core.loader.exec_module(mod_core)  # type: ignore

# Now load excel_writer as part of the orders package
excel_path = base / "orders" / "excel_writer.py"
spec = importlib.util.spec_from_file_location("orders.excel_writer", str(excel_path))
mod = importlib.util.module_from_spec(spec)
sys.modules["orders.excel_writer"] = mod
spec.loader.exec_module(mod)  # type: ignore
write_order_excel = mod.write_order_excel

# Try to load PDF generator similarly
pdf_path = base / "orders" / "pdf_writer.py"
if pdf_path.exists():
    spec2 = importlib.util.spec_from_file_location("orders.pdf_writer", str(pdf_path))
    mod2 = importlib.util.module_from_spec(spec2)
    sys.modules["orders.pdf_writer"] = mod2
    spec2.loader.exec_module(mod2)  # type: ignore
    generate_pdf_order_platypus = getattr(mod2, "generate_pdf_order_platypus", None)
else:
    generate_pdf_order_platypus = None

from models import Supplier
from delivery_addresses_db import DeliveryAddressesDB

TMP = pathlib.Path(tempfile.mkdtemp(prefix="fh_export_"))
print("Using tmp dir:", TMP)

items = [
    {"PartNumber": "P-001", "Description": "Test plaat", "Materiaal": "Staal", "Aantal": 2},
    {"PartNumber": "P-002", "Description": "Test profiel", "Materiaal": "Alu", "Aantal": 5},
]

company_info = {"name": "Testbedrijf", "address": "Straat 1", "vat": "BE0123456789", "email": "t@t.t", "website": "https://example.test", "accent_color": "#000000", "logo_path": None, "logo_crop": None}

supplier = Supplier.from_any({"supplier": "ACME", "adres_1": "Leverstraat 2", "btw": "BE123"})

delivery = None

excel_path = TMP / "test_order.xlsx"
pdf_path = TMP / "test_order.pdf"

# Try write
try:
    write_order_excel(str(excel_path), items, company_info, supplier, delivery)
    print("Excel write attempted:", excel_path)
except Exception as e:
    print("Excel writer raised:", e)

if generate_pdf_order_platypus is not None:
    try:
        generate_pdf_order_platypus(str(pdf_path), company_info, supplier, "Context label", items)
        print("PDF write attempted:", pdf_path)
    except Exception as e:
        print("PDF writer raised:", e)
else:
    print("PDF generator not available in this environment.")

# Check files
for p in (excel_path, pdf_path):
    exists = p.exists()
    size = p.stat().st_size if exists else 0
    print(p.name, "exists:", exists, "size:", size)

print("Done.")
