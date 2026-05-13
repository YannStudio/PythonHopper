import ast
import pathlib
import types
from typing import List, Dict, Optional

import pandas as pd

from suppliers_db import SuppliersDB
from delivery_addresses_db import DeliveryAddressesDB
from clients_db import ClientsDB
from helpers import _to_str, strip_favorite_marker
from models import Supplier, DeliveryAddress, Client
from app_settings import FileExtensionSetting


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


class DummyVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


def _load_gui_classes():
    source = pathlib.Path("gui.py").read_text(encoding="utf-8")
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
        "_to_str": _to_str,
        "strip_favorite_marker": strip_favorite_marker,
        "FileExtensionSetting": FileExtensionSetting,
        "prepare_custom_bom_for_main": lambda df, _current: df,
    }
    exec(code, ns)
    return ns["SupplierSelectionFrame"], ns["App"]


SupplierSelectionFrame, App = _load_gui_classes()


def make_dummy_selection():
    sdb = SuppliersDB([Supplier(supplier="Old")])
    ddb = DeliveryAddressesDB([DeliveryAddress(name="Addr1")])

    class DummySel:
        CLIENT_DELIVERY_PRESET = SupplierSelectionFrame.CLIENT_DELIVERY_PRESET
        DELIVERY_PRESETS = SupplierSelectionFrame.DELIVERY_PRESETS
        _display_list = SupplierSelectionFrame._display_list
        _resolve_current_client = SupplierSelectionFrame._resolve_current_client
        _default_delivery_value = SupplierSelectionFrame._default_delivery_value
        _delivery_options = SupplierSelectionFrame._delivery_options
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


def test_delivery_refresh_keeps_client_address_option_but_defaults_to_none():
    cdb = ClientsDB([Client(name="Alice", address="Kerkstraat 1, Gent")])

    class DummySel:
        CLIENT_DELIVERY_PRESET = SupplierSelectionFrame.CLIENT_DELIVERY_PRESET
        DELIVERY_PRESETS = SupplierSelectionFrame.DELIVERY_PRESETS
        _display_list = SupplierSelectionFrame._display_list
        _resolve_current_client = SupplierSelectionFrame._resolve_current_client
        _default_delivery_value = SupplierSelectionFrame._default_delivery_value
        _delivery_options = SupplierSelectionFrame._delivery_options
        _refresh_options = SupplierSelectionFrame._refresh_options

        def __init__(self):
            self.db = SuppliersDB([])
            self.delivery_db = DeliveryAddressesDB([])
            self.clients_db = cdb
            self.client_var = DummyVar("Alice")
            self.rows = []
            self.delivery_combos = {"Prod": DummyCombo("")}

    sel = DummySel()
    sel._refresh_options(initial=True)

    dcombo = sel.delivery_combos["Prod"]
    assert SupplierSelectionFrame.CLIENT_DELIVERY_PRESET in dcombo.values
    assert dcombo.get() == "Geen"


def test_app_resolves_client_address_delivery_choice():
    cdb = ClientsDB([Client(name="Alice", address="Kerkstraat 1, Gent")])

    class DummyApp:
        _current_client = App._current_client
        _resolve_delivery_choice = App._resolve_delivery_choice

        def __init__(self):
            self.client_db = cdb
            self.delivery_db = DeliveryAddressesDB([])
            self.client_var = DummyVar("Alice")

    app = DummyApp()
    delivery = app._resolve_delivery_choice(SupplierSelectionFrame.CLIENT_DELIVERY_PRESET)

    assert delivery is not None
    assert delivery.name == "Alice"
    assert delivery.address == "Kerkstraat 1, Gent"


def test_preset_status_message_and_indicator_are_exposed_per_row():
    class DummySel:
        _base_row_label = SupplierSelectionFrame._base_row_label
        _preset_indicator_suffix = SupplierSelectionFrame._preset_indicator_suffix
        _preset_status_message = SupplierSelectionFrame._preset_status_message

        def __init__(self):
            self.row_meta = {
                "production::Laser cutting": {
                    "base_display": "Laser cutting",
                    "display": "Laser cutting",
                    "identifier": "Laser cutting",
                }
            }
            self._preset_state_by_key = {
                "production::Laser cutting": {
                    "applied_rule_names": ["Klant X laser"],
                    "field_names": ["leverancier", "documenttype"],
                }
            }

    sel = DummySel()

    assert sel._preset_indicator_suffix("production::Laser cutting") == " [Preset]"
    message = sel._preset_status_message("production::Laser cutting")
    assert "Laser cutting" in message
    assert "Klant X laser" in message
    assert "leverancier" in message


def test_supplier_selection_frame_uses_scrollable_rows_canvas():
    source = pathlib.Path("gui.py").read_text(encoding="utf-8")
    mod = ast.parse(source)
    start = next(
        node for node in mod.body if isinstance(node, ast.FunctionDef) and node.name == "start_gui"
    )
    sup_cls = next(
        n for n in start.body if isinstance(n, ast.ClassDef) and n.name == "SupplierSelectionFrame"
    )
    init_fn = next(
        n for n in sup_cls.body if isinstance(n, ast.FunctionDef) and n.name == "__init__"
    )

    assert any(
        isinstance(node, ast.Assign)
        and any(
            isinstance(target, ast.Attribute) and target.attr == "selection_rows_canvas"
            for target in node.targets
        )
        and isinstance(node.value, ast.Call)
        and isinstance(node.value.func, ast.Attribute)
        and node.value.func.attr == "Canvas"
        for node in ast.walk(init_fn)
    )


def test_apply_custom_bom_refreshes_order_tab_when_present():
    class DummySelFrame:
        def winfo_exists(self):
            return True

    class DummyNotebook:
        def __init__(self):
            self.selected = None

        def select(self, tab):
            self.selected = tab

    class DummyApp:
        _apply_custom_bom_to_main = App._apply_custom_bom_to_main

        def __init__(self):
            self.custom_bom_tab = None
            self.bom_df = pd.DataFrame({"PartNumber": ["OLD"]})
            self.main_frame = object()
            self.nb = DummyNotebook()
            self.sel_frame = DummySelFrame()
            self.status_var = DummyVar("")
            self.flags_saved = None
            self.tree_refreshed = False
            self.custom_synced = False
            self.refresh_calls = []

        def _store_custom_row_flags(self, df, flags):
            self.flags_saved = (df, list(flags))

        def _refresh_tree(self):
            self.tree_refreshed = True

        def _sync_custom_bom_from_main(self):
            self.custom_synced = True

        def _show_supplier_selection_tab(self, **kwargs):
            self.refresh_calls.append(kwargs)
            return object()

    app = DummyApp()
    custom_df = pd.DataFrame({"PartNumber": ["PN-001"]})

    app._apply_custom_bom_to_main(custom_df)

    assert app.flags_saved is not None
    assert app.tree_refreshed is True
    assert app.custom_synced is True
    assert app.nb.selected is app.main_frame
    assert app.refresh_calls == [
        {"select_tab": False, "prompt_opticutter": False}
    ]
    assert app.status_var.get().endswith("Bestelbonnen bijgewerkt.")


def test_apply_custom_bom_recalculates_file_status_when_source_available():
    class DummyNotebook:
        def __init__(self):
            self.selected = None

        def select(self, tab):
            self.selected = tab

    class DummyApp:
        _apply_custom_bom_to_main = App._apply_custom_bom_to_main

        def __init__(self):
            self.custom_bom_tab = None
            self.bom_df = pd.DataFrame(
                [{"PartNumber": "OLD", "Production": "Laser", "Status": "oud"}]
            )
            self.main_frame = object()
            self.nb = DummyNotebook()
            self.sel_frame = None
            self.status_var = DummyVar("")
            self.source_folder = "bron"
            self.file_status_exts = None
            self.tree_refreshed = False
            self.custom_synced = False

        def _store_custom_row_flags(self, df, flags):
            self.flags_saved = (df, list(flags))

        def _selected_exts(self):
            return [".pdf"]

        def _update_bom_file_status(self, exts):
            self.file_status_exts = list(exts)
            self.bom_df["Bestanden gevonden"] = ["pdf"]
            self.bom_df["Status"] = ["fresh"]
            self.bom_df["Link"] = [""]

        def _refresh_tree(self):
            self.tree_refreshed = True

        def _sync_custom_bom_from_main(self):
            self.custom_synced = True

    app = DummyApp()
    custom_df = pd.DataFrame({"PartNumber": ["PN-001"], "Production": ["Plooien"]})

    app._apply_custom_bom_to_main(custom_df)

    assert app.bom_df.loc[0, "Production"] == "Plooien"
    assert app.bom_df.loc[0, "Status"] == "fresh"
    assert app.file_status_exts == [".pdf"]
    assert app.tree_refreshed is False
    assert "Bestandscontrole bijgewerkt." in app.status_var.get()
