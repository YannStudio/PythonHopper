import ast
import pathlib
import types
from typing import Dict, List, Optional

from helpers import _to_str
from models import Client, DeliveryAddress, Supplier
from suppliers_db import SuppliersDB
from delivery_addresses_db import DeliveryAddressesDB


def _load_supplier_frame():
    source = pathlib.Path("gui.py").read_text(encoding="utf-8")
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
        TclError=Exception,
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
        "_to_str": _to_str,
    }
    exec(code, ns)
    return ns["SupplierSelectionFrame"]


SupplierSelectionFrame = _load_supplier_frame()


class DummySel:
    _parse_selection_key = staticmethod(SupplierSelectionFrame._parse_selection_key)
    _is_groupable_kind = staticmethod(SupplierSelectionFrame._is_groupable_kind)
    _resolve_group_root = staticmethod(SupplierSelectionFrame._resolve_group_root)
    _group_code_from_index = staticmethod(SupplierSelectionFrame._group_code_from_index)
    _base_row_label = SupplierSelectionFrame._base_row_label
    _group_root_code_map = SupplierSelectionFrame._group_root_code_map
    _group_root_color_map = SupplierSelectionFrame._group_root_color_map
    _group_followers_by_root = SupplierSelectionFrame._group_followers_by_root
    _group_row_label = SupplierSelectionFrame._group_row_label
    _group_visual_spec = SupplierSelectionFrame._group_visual_spec

    GROUP_ACCENT_COLORS = SupplierSelectionFrame.GROUP_ACCENT_COLORS

    def __init__(self):
        self.rows = [
            ("production::Laser", None),
            ("production::Plooien", None),
            ("production::Lassen", None),
            ("finish::Poeder", None),
        ]
        self.row_meta = {
            "production::Laser": {"base_display": "Laser"},
            "production::Plooien": {"base_display": "Plooien"},
            "production::Lassen": {"base_display": "Lassen"},
            "finish::Poeder": {"base_display": "Poedercoating"},
        }


def test_group_code_from_index_supports_multiple_letters():
    assert SupplierSelectionFrame._group_code_from_index(0) == "A"
    assert SupplierSelectionFrame._group_code_from_index(25) == "Z"
    assert SupplierSelectionFrame._group_code_from_index(26) == "AA"


def test_group_row_label_uses_bon_code_for_master_choice():
    sel = DummySel()

    label = sel._group_row_label("production::Laser", {"production::Plooien": "production::Laser"})

    assert label == "Bon A - Laser"


def test_group_visual_spec_marks_master_and_follower_with_same_group():
    sel = DummySel()
    group_links = {"production::Plooien": "production::Laser"}

    master = sel._group_visual_spec("production::Laser", group_links)
    follower = sel._group_visual_spec("production::Plooien", group_links)
    separate = sel._group_visual_spec("production::Lassen", group_links)

    assert master["grouped"] is True
    assert master["is_root"] is True
    assert "[Bon A]" in master["text"]
    assert master["accent"]

    assert follower["grouped"] is True
    assert follower["is_root"] is False
    assert "[Volgt A]" in follower["text"]
    assert follower["accent"] == master["accent"]

    assert separate["grouped"] is False
    assert separate["text"] == "Lassen"
