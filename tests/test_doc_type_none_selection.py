import ast
import pathlib
import types
from typing import Dict, List, Optional

from suppliers_db import SuppliersDB
from models import Supplier, Client, DeliveryAddress
from delivery_addresses_db import DeliveryAddressesDB
from orders import _prefix_for_doc_type


class DummyCombo:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class DummyVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


def _load_supplier_frame():
    source = pathlib.Path("gui.py").read_text()
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    sup_cls = next(
        n for n in start.body if isinstance(n, ast.ClassDef) and n.name == "SupplierSelectionFrame"
    )
    module_ast = ast.Module(body=[sup_cls], type_ignores=[])
    code = compile(module_ast, "<gui_extract>", "exec")
    tk_stub = types.SimpleNamespace(
        Frame=type("Frame", (), {}),
        Tk=type("Tk", (), {}),
        Toplevel=type("Toplevel", (), {}),
        BooleanVar=lambda value=None: None,
        StringVar=lambda value=None: None,
        Label=type("Label", (), {}),
        Entry=type("Entry", (), {}),
        Checkbutton=type("Checkbutton", (), {}),
        Button=type("Button", (), {}),
        LabelFrame=type("LabelFrame", (), {}),
    )
    ttk_stub = types.SimpleNamespace(
        Combobox=type("Combobox", (), {}),
        Treeview=type("Treeview", (), {}),
        Scrollbar=type("Scrollbar", (), {}),
        Style=type("Style", (), {}),
    )
    ns = {
        "tk": tk_stub,
        "ttk": ttk_stub,
        "List": List,
        "Dict": Dict,
        "Optional": Optional,
        "Supplier": Supplier,
        "Client": Client,
        "DeliveryAddress": DeliveryAddress,
        "SuppliersDB": SuppliersDB,
        "DeliveryAddressesDB": DeliveryAddressesDB,
        "_prefix_for_doc_type": _prefix_for_doc_type,
    }
    exec(code, ns)
    return ns["SupplierSelectionFrame"]


SupplierSelectionFrame = _load_supplier_frame()


def test_supplier_geen_sets_doc_type_to_geen():
    class DummySel:
        _on_combo_change = SupplierSelectionFrame._on_combo_change
        _on_doc_type_change = SupplierSelectionFrame._on_doc_type_change

        def __init__(self):
            self.rows = [("Prod", DummyCombo("(geen)"))]
            self.doc_vars = {"Prod": DummyVar("Bestelbon")}
            self.doc_num_vars = {"Prod": DummyVar("")}
            self._update_preview_from_any_combo = lambda: None

    sel = DummySel()
    sel._on_combo_change()
    assert sel.doc_vars["Prod"].get() == "Geen"

    sel.rows[0][1].set("Other")
    sel._on_combo_change()
    assert sel.doc_vars["Prod"].get() == "Bestelbon"

