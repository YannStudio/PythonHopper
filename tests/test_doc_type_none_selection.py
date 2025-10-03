import ast
import pathlib
import types
from typing import Dict, List, Optional

from suppliers_db import SuppliersDB
from models import Supplier, Client, DeliveryAddress
from delivery_addresses_db import DeliveryAddressesDB
# ``_prefix_for_doc_type`` is imported in the main application from
# ``orders``, but importing that module requires heavy dependencies like
# ``pandas``.  Re-implement the small helper here to keep tests lightweight.
def _prefix_for_doc_type(doc_type: str) -> str:
    t = (doc_type or "").strip().lower()
    if t.startswith("standaard"):
        return "BB-"
    if t.startswith("bestel"):
        return "BB-"
    if t.startswith("offerte"):
        return "OFF-"
    return ""


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


def test_confirm_persists_geen_doc_type():
    class DummySel:
        _on_combo_change = SupplierSelectionFrame._on_combo_change
        _on_doc_type_change = SupplierSelectionFrame._on_doc_type_change
        _confirm = SupplierSelectionFrame._confirm

        def __init__(self):
            self.rows = [("Prod", DummyCombo("(geen)"))]
            self.doc_vars = {"Prod": DummyVar("Bestelbon")}
            self.doc_num_vars = {"Prod": DummyVar("")}
            self.delivery_vars = {"Prod": DummyVar("Geen")}
            self.project_number_var = DummyVar("")
            self.project_name_var = DummyVar("")
            self.remember_var = DummyVar(0)
            self._update_preview_from_any_combo = lambda: None
            self.callback_args = None

        def callback(
            self,
            sel_map,
            doc_map,
            doc_num_map,
            delivery_map,
            project_number,
            project_name,
            remember,
        ):
            self.callback_args = (
                sel_map,
                doc_map,
                doc_num_map,
                delivery_map,
                project_number,
                project_name,
                remember,
            )

    sel = DummySel()
    sel._on_combo_change()
    sel._confirm()
    assert sel.callback_args is not None
    _, doc_map, *_ = sel.callback_args
    assert doc_map["Prod"] == "Geen"

