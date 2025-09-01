import ast
import pathlib
import types
from typing import List, Dict, Optional

import gui
from suppliers_db import SuppliersDB
from delivery_addresses_db import DeliveryAddressesDB
from models import Supplier, DeliveryAddress


class DummyCombo:
    def __init__(self, value=""):
        self.value = value
        self.values = []

    def get(self):
        return self.value

    def __setitem__(self, key, val):
        if key == "values":
            self.values = list(val)

    def set(self, val):
        self.value = val


def _load_supplier_frame():
    source = pathlib.Path("gui.py").read_text()
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    frame_cls = next(
        n for n in start.body if isinstance(n, ast.ClassDef) and n.name == "SupplierSelectionFrame"
    )
    module_ast = ast.Module(body=[frame_cls], type_ignores=[])
    code = compile(module_ast, "<gui_extract>", "exec")
    tk_stub = types.SimpleNamespace(
        Frame=type("Frame", (), {}),
        Toplevel=type("Toplevel", (), {}),
        BooleanVar=lambda value=None: None,
        StringVar=lambda value=None: None,
        Label=type("Label", (), {}),
        Entry=type("Entry", (), {}),
        Checkbutton=type("Checkbutton", (), {}),
        Button=type("Button", (), {}),
        LabelFrame=type("LabelFrame", (), {}),
    )
    ttk_stub = types.SimpleNamespace(Combobox=type("Combobox", (), {}))
    ns = {
        "tk": tk_stub,
        "ttk": ttk_stub,
        "List": List,
        "Dict": Dict,
        "Optional": Optional,
        "Supplier": Supplier,
        "DeliveryAddress": DeliveryAddress,
        "SuppliersDB": SuppliersDB,
        "DeliveryAddressesDB": DeliveryAddressesDB,
        "sort_supplier_options": gui.sort_supplier_options,
        "_norm": gui._norm,
    }
    exec(code, ns)
    return ns["SupplierSelectionFrame"]


SupplierSelectionFrame = _load_supplier_frame()


class DummySel:
    _display_list = SupplierSelectionFrame._display_list
    _refresh_options = SupplierSelectionFrame._refresh_options
    _on_combo_type = SupplierSelectionFrame._on_combo_type
    _resolve_text_to_supplier = SupplierSelectionFrame._resolve_text_to_supplier

    def __init__(self, sdb):
        self.db = sdb
        self.delivery_db = DeliveryAddressesDB([])
        self.rows = [("Prod", DummyCombo())]
        self.delivery_combos = {}
        self._preview_supplier = None

    def _update_preview_for_text(self, text):
        self._preview_supplier = self._resolve_text_to_supplier(text)


def test_unaccented_filter_and_selects_supplier():
    sdb = SuppliersDB([Supplier(supplier="Café"), Supplier(supplier="Other")])
    sel = DummySel(sdb)
    sel._refresh_options(initial=True)
    combo = sel.rows[0][1]
    combo.set("Cafe")
    sel._on_combo_type(types.SimpleNamespace(keysym="Return"), "Prod", combo)
    assert combo.values == ["Café"]
    assert combo.get() == "Café"
    assert sel._preview_supplier and sel._preview_supplier.supplier == "Café"
