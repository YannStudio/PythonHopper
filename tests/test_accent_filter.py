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
    class Widget:
        def __init__(self, master=None, **kwargs):
            self.master = master
            self.kwargs = kwargs
            self.children = []
            self.bindings = {}
            self.grid_kwargs = {}
            self.pack_kwargs = {}
            if master is not None:
                master.children.append(self)

        def grid(self, **kwargs):
            self.grid_kwargs = kwargs

        def pack(self, **kwargs):
            self.pack_kwargs = kwargs

        def bind(self, event, handler):
            self.bindings[event] = handler

        def destroy(self):
            pass

        def winfo_children(self):
            return list(self.children)

        def grid_columnconfigure(self, *_args, **_kwargs):
            pass

        def grid_rowconfigure(self, *_args, **_kwargs):
            pass

    class Frame(Widget):
        pass

    class Label(Widget):
        pass

    tk_stub = types.SimpleNamespace(
        Frame=Frame,
        Toplevel=Frame,
        BooleanVar=lambda value=None: None,
        StringVar=lambda value=None: None,
        Label=Label,
        Entry=Frame,
        Checkbutton=Frame,
        Button=Frame,
        LabelFrame=Frame,
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
    cls = ns["SupplierSelectionFrame"]
    cls._Frame = Frame
    return cls


SupplierSelectionFrame = _load_supplier_frame()


class DummySel:
    _display_list = SupplierSelectionFrame._display_list
    _refresh_options = SupplierSelectionFrame._refresh_options
    _on_combo_type = SupplierSelectionFrame._on_combo_type
    _resolve_text_to_supplier = SupplierSelectionFrame._resolve_text_to_supplier
    _populate_cards = SupplierSelectionFrame._populate_cards

    def __init__(self, sdb):
        self.db = sdb
        self.delivery_db = DeliveryAddressesDB([])
        self.rows = [("Prod", DummyCombo())]
        self.delivery_combos = {}
        self._preview_supplier = None
        self.cards_frame = SupplierSelectionFrame._Frame(None)

    def _update_preview_for_text(self, text):
        self._preview_supplier = self._resolve_text_to_supplier(text)

    def _on_card_click(self, option, production):
        pass


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
    assert calls == []


def test_populate_cards_card_format():
    sdb = SuppliersDB([Supplier(supplier="Alpha", description="Desc", adres_1="Addr")])
    sel = DummySel(sdb)
    sel._populate_cards(["Alpha"], "Prod")
    cards = sel.cards_frame.children
    assert len(cards) == 1
    card = cards[0]
    assert card.kwargs.get("highlightbackground") == "#444444"
    assert card.grid_kwargs.get("sticky") == "w"
    name_lbl = card.children[0]
    assert "bold" in name_lbl.kwargs.get("font", ("", "", ""))
    desc_lbl = card.children[1]
    assert desc_lbl.kwargs.get("text") == "Desc"
    assert "(" not in desc_lbl.kwargs.get("text")
