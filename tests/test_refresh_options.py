import ast
import pathlib
import types
from typing import List, Dict, Optional

from suppliers_db import SuppliersDB
from delivery_addresses_db import DeliveryAddressesDB
from clients_db import ClientsDB
from models import Supplier, DeliveryAddress, Client


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


def _load_gui_classes():
    source = pathlib.Path("gui.py").read_text()
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    sup_cls = next(
        n for n in start.body if isinstance(n, ast.ClassDef) and n.name == "SupplierSelectionFrame"
    )
    app_cls = next(
        n for n in start.body if isinstance(n, ast.ClassDef) and n.name == "App"
    )
    module_ast = ast.Module(body=[sup_cls, app_cls], type_ignores=[])
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
    }
    exec(code, ns)
    return ns["SupplierSelectionFrame"], ns["App"]


SupplierSelectionFrame, App = _load_gui_classes()


def make_dummy_selection():
    sdb = SuppliersDB([Supplier(supplier="Old")])
    ddb = DeliveryAddressesDB([DeliveryAddress(name="Addr1")])

    class DummySel:
        _display_list = SupplierSelectionFrame._display_list
        _refresh_options = SupplierSelectionFrame._refresh_options

        def __init__(self):
            self.db = sdb
            self.delivery_db = ddb
            self.rows = [("Prod", DummyCombo("Old"))]
            self.delivery_combos = {"Prod": DummyCombo("Addr1")}

    return DummySel(), sdb, ddb


def test_supplier_and_delivery_refresh_preserves_selection():
    sel, sdb, ddb = make_dummy_selection()
    sel._refresh_options(initial=True)
    combo = sel.rows[0][1]
    dcombo = sel.delivery_combos["Prod"]

    sdb.upsert(Supplier(supplier="New"))
    combo.set("Old")
    sel._refresh_options()
    assert any("New" in v for v in combo.values)
    assert combo.get() == "Old"

    ddb.upsert(DeliveryAddress(name="Addr2"))
    dcombo.set("Addr1")
    sel._refresh_options()
    assert any("Addr2" in v for v in dcombo.values)
    assert dcombo.get() == "Addr1"


def test_refresh_clients_combo_preserves_selection():
    cdb = ClientsDB([Client(name="Alice"), Client(name="Bob")])

    class DummyApp:
        _refresh_clients_combo = App._refresh_clients_combo

        def __init__(self):
            self.client_db = cdb
            self.client_combo = DummyCombo("Alice")

    app = DummyApp()
    app._refresh_clients_combo()
    cdb.upsert(Client(name="Carol"))
    app._refresh_clients_combo()
    assert any("Carol" in v for v in app.client_combo.values)
    assert app.client_combo.get() == "Alice"
