import ast
import pathlib
import types
import unicodedata
from typing import List, Dict, Optional

from suppliers_db import SuppliersDB
from delivery_addresses_db import DeliveryAddressesDB
from models import Supplier, DeliveryAddress


def _norm(text: str) -> str:
    return (
        unicodedata.normalize("NFKD", text)
        .encode("ASCII", "ignore")
        .decode("ASCII")
        .lower()
    )


def sort_supplier_options(
    options: List[str],
    suppliers: List[Supplier],
    disp_to_name: Dict[str, str],
) -> List[str]:
    fav_map = {_norm(s.supplier): s.favorite for s in suppliers}

    def sort_key(opt: str):
        name = disp_to_name.get(opt, opt)
        n = _norm(name)
        return (not fav_map.get(n, False), n)

    return sorted(options, key=sort_key)


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
        "sort_supplier_options": sort_supplier_options,
        "_norm": _norm,
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

    def _populate_cards(self, options, prod):
        self.last_populate = (options, prod)


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


def test_populate_cards_not_called_for_empty_text():
    sdb = SuppliersDB([Supplier(supplier="Alpha")])
    sel = DummySel(sdb)
    sel._refresh_options(initial=True)
    combo = sel.rows[0][1]
    combo.set("")
    calls = []

    def fake_populate(options, _prod):
        calls.append(options)

    sel._populate_cards = fake_populate
    sel._on_combo_type(types.SimpleNamespace(keysym="a"), "Prod", combo)
    assert calls == [[]]


class DummyWidget:
    def __init__(self, master=None, **kwargs):
        self.master = master
        self.kwargs = kwargs
        self.children = []
        self.bindings = {}
        if master is not None:
            master.children.append(self)

    def winfo_children(self):
        return list(self.children)

    def grid(self, **kwargs):
        self.grid_kwargs = kwargs

    def pack(self, **kwargs):
        self.pack_kwargs = kwargs

    def bind(self, event, handler):
        self.bindings[event] = handler

    def destroy(self):
        pass

    def grid_columnconfigure(self, *args, **kwargs):
        pass

    def grid_rowconfigure(self, *args, **kwargs):
        pass


class DummyFrame(DummyWidget):
    pass


class DummyLabel(DummyWidget):
    pass


def _load_supplier_frame_with_widgets():
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
        Frame=DummyFrame,
        Toplevel=type("Toplevel", (), {}),
        BooleanVar=lambda value=None: None,
        StringVar=lambda value=None: None,
        Label=DummyLabel,
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
        "sort_supplier_options": sort_supplier_options,
        "_norm": _norm,
    }
    exec(code, ns)
    return ns["SupplierSelectionFrame"]


SupplierSelectionFrameWidgets = _load_supplier_frame_with_widgets()


def test_populate_cards_formatting():
    s = Supplier(supplier="Foo", description="Bar")
    sel = types.SimpleNamespace(
        cards_frame=DummyFrame(),
        _resolve_text_to_supplier=lambda _opt: s,
        _on_card_click=lambda *args: None,
    )
    SupplierSelectionFrameWidgets._populate_cards(sel, ["Foo"], "Prod")
    assert sel.cards_frame.children, "Card was not created"
    card = sel.cards_frame.children[0]
    assert card.kwargs.get("highlightbackground") == "#444444"
    assert card.kwargs.get("highlightcolor") == "#444444"
    labels = card.children
    assert labels[0].kwargs.get("text") == "Foo"
    assert labels[0].kwargs.get("font")[2] == "bold"
    assert "(" not in labels[1].kwargs.get("text", "")
