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


def _normalize_doc_number(value, doc_type):
    prefix = _prefix_for_doc_type(doc_type)
    text = ("" if value is None else str(value)).strip()
    if not text:
        return ""
    if not prefix:
        return text

    prefix_upper = prefix.upper()
    compact = prefix_upper.replace("-", "")
    text_upper = text.upper()

    if text_upper.startswith(prefix_upper):
        remainder = text[len(prefix) :].lstrip(" -_")
        remainder_upper = remainder.upper()
        if remainder_upper.startswith(prefix_upper):
            remainder = remainder[len(prefix) :].lstrip(" -_")
            return prefix + remainder
        if compact and remainder_upper.startswith(compact):
            remainder = remainder[len(compact) :].lstrip(" -_")
            return prefix + remainder
        return prefix + remainder

    if compact and text_upper.startswith(compact):
        remainder = text[len(compact) :].lstrip(" -_")
        return prefix + remainder

    return text


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
        "_normalize_doc_number": _normalize_doc_number,
    }
    exec(code, ns)
    return ns["SupplierSelectionFrame"]


SupplierSelectionFrame = _load_supplier_frame()


def test_supplier_none_sets_doc_type_to_standard():
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
    assert sel.doc_vars["Prod"].get() == "Standaard bon"

    sel.rows[0][1].set("Other")
    sel._on_combo_change()
    assert sel.doc_vars["Prod"].get() == "Bestelbon"


def test_confirm_persists_standard_doc_type():
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
    assert doc_map["Prod"] == "Standaard bon"


def test_supplier_empty_string_sets_doc_type_to_standard():
    class DummySel:
        _on_combo_change = SupplierSelectionFrame._on_combo_change
        _on_doc_type_change = SupplierSelectionFrame._on_doc_type_change

        def __init__(self):
            self.rows = [("Prod", DummyCombo(""))]
            self.doc_vars = {"Prod": DummyVar("Bestelbon")}
            self.doc_num_vars = {"Prod": DummyVar("")}
            self._update_preview_from_any_combo = lambda: None

    sel = DummySel()
    sel._on_combo_change()
    assert sel.doc_vars["Prod"].get() == "Standaard bon"


def test_clear_saved_suppliers_sets_doc_type_to_standard():
    class DummyDB:
        def __init__(self):
            self.defaults_by_production = {"Prod": "Some"}
            self.defaults_by_finish = {"Fin": "Other"}
            self.saved = False

        def save(self):
            self.saved = True

        def suppliers_sorted(self):
            return []

    class DummySel:
        _clear_saved_suppliers = SupplierSelectionFrame._clear_saved_suppliers
        _on_combo_change = SupplierSelectionFrame._on_combo_change
        _on_doc_type_change = SupplierSelectionFrame._on_doc_type_change

        def __init__(self):
            self.db = DummyDB()
            combo = DummyCombo("Leverancier")
            self.rows = [("Prod", combo)]
            self.doc_vars = {"Prod": DummyVar("Bestelbon")}
            self.doc_num_vars = {"Prod": DummyVar("")}
            self.delivery_combos = {"Prod": DummyCombo("Leveradres wordt nog meegedeeld")}
            self.delivery_vars = {"Prod": DummyVar("Leveradres wordt nog meegedeeld")}
            self._update_preview_from_any_combo = lambda: None
            self._doc_type_prefixes = {"BB-"}

    sel = DummySel()
    sel._clear_saved_suppliers()

    assert sel.db.saved is True
    assert sel.rows[0][1].get() == "(geen)"
    assert sel.doc_vars["Prod"].get() == "Standaard bon"
    assert sel.delivery_combos["Prod"].get() == "Geen"

