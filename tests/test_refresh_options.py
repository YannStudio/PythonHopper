import ast
import datetime
import os
import pathlib
import sys
import threading
import types
from typing import List, Dict, Mapping, Optional

import pandas as pd

from suppliers_db import SuppliersDB
from delivery_addresses_db import DeliveryAddressesDB
from clients_db import ClientsDB
from helpers import _to_str, strip_favorite_marker
from models import Supplier, DeliveryAddress, Client
from app_settings import FileExtensionSetting
from opticutter import analyse_profiles


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
        "Mapping": Mapping,
        "Optional": Optional,
        "os": os,
        "Supplier": Supplier,
        "Client": Client,
        "DeliveryAddress": DeliveryAddress,
        "SuppliersDB": SuppliersDB,
        "DeliveryAddressesDB": DeliveryAddressesDB,
        "_to_str": _to_str,
        "strip_favorite_marker": strip_favorite_marker,
        "FileExtensionSetting": FileExtensionSetting,
        "prepare_custom_bom_for_main": lambda df, _current: df,
        "datetime": datetime,
        "pd": pd,
        "sys": sys,
        "analyse_profiles": analyse_profiles,
        "threading": threading,
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


def test_price_columns_are_hidden_by_default_and_toggleable():
    class DummySel:
        PRICE_COLUMN_KEYS = SupplierSelectionFrame.PRICE_COLUMN_KEYS
        _refresh_visible_column_keys = SupplierSelectionFrame._refresh_visible_column_keys
        _repack_visible_columns = SupplierSelectionFrame._repack_visible_columns
        set_price_columns_visible = SupplierSelectionFrame.set_price_columns_visible

        def __init__(self):
            self._column_keys = [
                "export_check",
                "label",
                "group_combo",
                "supplier_combo",
                "doc_combo",
                "en1090_widget",
                "doc_entry",
                "unit_price_entry",
                "total_price_entry",
                "vat_combo",
                "line_price_button",
                "remark_entry",
                "delivery_combo",
            ]
            self._en1090_enabled = True
            self._price_columns_visible = False
            self._visible_column_keys = []
            self._row_widget_maps = []
            self._header_aligned = True
            self._header_alignment_pending = True
            self.price_columns_visible_var = DummyVar(False)
            self.header_repack_count = 0
            self.row_repack_count = 0

        def _repack_header_columns(self):
            self.header_repack_count += 1

        def _repack_all_rows(self):
            self.row_repack_count += 1

        def _schedule_header_alignment(self, _row_widgets):
            raise AssertionError("No rows should be scheduled in this test")

    sel = DummySel()
    sel._refresh_visible_column_keys()

    assert "unit_price_entry" not in sel._visible_column_keys
    assert "total_price_entry" not in sel._visible_column_keys
    assert "vat_combo" not in sel._visible_column_keys
    assert "line_price_button" not in sel._visible_column_keys

    sel.set_price_columns_visible(True)

    assert sel.price_columns_visible_var.get() == 1
    assert sel.header_repack_count == 1
    assert sel.row_repack_count == 1
    assert [
        key for key in sel._visible_column_keys if key in sel.PRICE_COLUMN_KEYS
    ] == list(sel.PRICE_COLUMN_KEYS)


def test_price_columns_auto_show_for_existing_price_state():
    price_state = types.SimpleNamespace(
        pricing={"production::Laser": {"unit_price": "12,50"}},
        vat_rates={},
    )
    default_state = types.SimpleNamespace(pricing={}, vat_rates={"row": "21"})
    reduced_vat_state = types.SimpleNamespace(pricing={}, vat_rates={"row": "6"})

    assert SupplierSelectionFrame._state_has_price_values(price_state) is True
    assert SupplierSelectionFrame._state_has_price_values(default_state) is False
    assert SupplierSelectionFrame._state_has_price_values(reduced_vat_state) is True


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


def test_loading_custom_bom_keeps_original_related_bom_source(tmp_path):
    original_bom = tmp_path / "ProjectX-BOM.xlsx"
    custom_bom = tmp_path / "BOM-FileHopper-Temp.xlsx"

    class DummyApp:
        _load_bom_from_path = App._load_bom_from_path
        _apply_loaded_bom = App._apply_loaded_bom

        def __init__(self):
            self.bom_source_path = str(original_bom)
            self.status_var = DummyVar("")
            self.bom_df = None
            self.flags_saved = None
            self.tree_refreshed = False
            self.custom_synced = False

        def _store_custom_row_flags(self, df, flags):
            self.flags_saved = (df, list(flags))

        def _refresh_tree(self):
            self.tree_refreshed = True

        def _sync_custom_bom_from_main(self):
            self.custom_synced = True

    globals_dict = App._load_bom_from_path.__globals__
    previous_loader = globals_dict.get("load_bom")
    globals_dict["load_bom"] = lambda _path: pd.DataFrame({"PartNumber": ["PN-001"]})
    try:
        app = DummyApp()
        app._load_bom_from_path(str(custom_bom), mark_as_custom=True)
    finally:
        if previous_loader is None:
            globals_dict.pop("load_bom", None)
        else:
            globals_dict["load_bom"] = previous_loader

    assert app.bom_source_path == str(original_bom)
    assert app.flags_saved[1] == [True]
    assert app.tree_refreshed is True
    assert app.custom_synced is True


def test_sync_custom_bom_defers_when_tab_is_not_created():
    class DummyApp:
        _sync_custom_bom_from_main = App._sync_custom_bom_from_main

        def __init__(self):
            self.custom_bom_tab = None
            self.bom_df = pd.DataFrame({"PartNumber": ["PN-001"]})
            self._custom_bom_needs_sync = False
            self.created_tab = False

        def _autofill_custom_bom_enabled(self):
            return True

        def _ensure_custom_bom_tab(self):
            self.created_tab = True
            return object()

    app = DummyApp()

    app._sync_custom_bom_from_main()

    assert app.created_tab is False
    assert app._custom_bom_needs_sync is True


def test_ensure_custom_bom_tab_applies_pending_main_sync():
    class DummyBOMCustomTab:
        MAIN_COLUMN_ORDER = ("PartNumber",)

        def __init__(self, *_args, **_kwargs):
            self.loaded = []

        def load_from_main_dataframe(self, df):
            self.loaded.append(df.copy(deep=True))

    class DummyNotebook:
        def __init__(self):
            self.added = None

        def add(self, tab, **_kwargs):
            self.added = tab

    class DummyApp:
        _ensure_custom_bom_tab = App._ensure_custom_bom_tab
        _autofill_custom_bom_enabled = App._autofill_custom_bom_enabled
        _load_current_bom_into_custom_tab = App._load_current_bom_into_custom_tab

        def __init__(self):
            self.nb = DummyNotebook()
            self.custom_bom_tab = None
            self._custom_bom_placeholder = None
            self._custom_bom_needs_sync = True
            self.bom_df = pd.DataFrame({"PartNumber": ["PN-001"]})
            self.autofill_custom_bom_var = DummyVar(1)
            self.settings = types.SimpleNamespace(autofill_custom_bom=True)

        def _on_custom_bom_ready(self, *_args, **_kwargs):
            return None

        def _apply_custom_bom_to_main(self, *_args, **_kwargs):
            return None

    globals_dict = App._ensure_custom_bom_tab.__globals__
    previous_tab = globals_dict.get("BOMCustomTab")
    globals_dict["BOMCustomTab"] = DummyBOMCustomTab
    try:
        app = DummyApp()
        tab = app._ensure_custom_bom_tab()
    finally:
        if previous_tab is None:
            globals_dict.pop("BOMCustomTab", None)
        else:
            globals_dict["BOMCustomTab"] = previous_tab

    assert tab is app.nb.added
    assert app.custom_bom_tab is tab
    assert tab.loaded[0].loc[0, "PartNumber"] == "PN-001"
    assert app._custom_bom_needs_sync is False


def test_opticutter_refresh_is_deferred_until_tab_is_selected():
    class DummyNotebook:
        def __init__(self):
            self.selected = "main"

        def select(self, *_args):
            return self.selected

    class DummyApp:
        _is_opticutter_tab = App._is_opticutter_tab
        _refresh_opticutter_if_needed = App._refresh_opticutter_if_needed
        _request_opticutter_refresh = App._request_opticutter_refresh

        def __init__(self):
            self.nb = DummyNotebook()
            self.opticutter_frame = "opticutter"
            self._opticutter_needs_refresh = False
            self._opticutter_analysis_stale = False
            self.opticutter_last_analysis = "cached-analysis"
            self.refresh_count = 0
            self.background_started = False

        def _refresh_opticutter_table(self, **_kwargs):
            self.refresh_count += 1

        def _start_background_opticutter_analysis_refresh(self):
            self.background_started = True

    app = DummyApp()

    app._request_opticutter_refresh()
    assert app.refresh_count == 0
    assert app._opticutter_needs_refresh is True
    assert app.background_started is True

    app.nb.selected = "opticutter"
    app._refresh_opticutter_if_needed(app.nb.select())

    assert app.refresh_count == 1
    assert app._opticutter_needs_refresh is False


def test_opticutter_analysis_state_keeps_order_scenarios_ready():
    class DummyApp:
        _apply_opticutter_analysis_state = App._apply_opticutter_analysis_state

        def __init__(self):
            self.opticutter_last_analysis = None
            self.opticutter_profile_custom_lengths = {}
            self.opticutter_profile_selection_choice = {}
            self.opticutter_profile_selection_scenarios = {}

    key = ("Koker 50x50", "S235", "Zaag")
    profile = types.SimpleNamespace(
        key=key,
        scenarios={"6000": object(), "12000": object()},
        best_choice="6000",
    )
    analysis = types.SimpleNamespace(profiles=[profile])

    app = DummyApp()
    app._apply_opticutter_analysis_state(analysis)

    assert app.opticutter_last_analysis is analysis
    assert app.opticutter_profile_selection_scenarios[key] == profile.scenarios
    assert app.opticutter_profile_selection_choice[key] == "6000"


def test_leaving_dirty_opticutter_tab_confirms_update():
    class DummyNotebook:
        def __init__(self):
            self.selected = "main"

        def select(self, *_args):
            return self.selected

    class DummyApp:
        _handle_tab_changed = App._handle_tab_changed
        _is_opticutter_tab = App._is_opticutter_tab
        _handle_opticutter_tab_transition = App._handle_opticutter_tab_transition

        def __init__(self):
            self.nb = DummyNotebook()
            self.opticutter_frame = "opticutter"
            self._custom_bom_placeholder = None
            self._last_selected_notebook_tab = "opticutter"
            self._opticutter_dirty = True
            self.confirmed = False

        def _confirm_opticutter_update(self):
            self.confirmed = True
            self._opticutter_dirty = False
            return True

        def _refresh_opticutter_if_needed(self, *_args, **_kwargs):
            return None

    app = DummyApp()
    event = types.SimpleNamespace(widget=app.nb)

    app._handle_tab_changed(event)

    assert app.confirmed is True
    assert app._opticutter_dirty is False
    assert app._last_selected_notebook_tab == "main"


def test_leaving_dirty_opticutter_tab_via_custom_placeholder_confirms_update():
    class DummyNotebook:
        def __init__(self):
            self.selected = "placeholder"
            self.selected_to = None

        def select(self, tab=None):
            if tab is not None:
                self.selected_to = tab
                self.selected = tab
            return self.selected

    class DummyApp:
        _handle_tab_changed = App._handle_tab_changed
        _is_opticutter_tab = App._is_opticutter_tab
        _handle_opticutter_tab_transition = App._handle_opticutter_tab_transition

        def __init__(self):
            self.nb = DummyNotebook()
            self.opticutter_frame = "opticutter"
            self._custom_bom_placeholder = "placeholder"
            self._last_selected_notebook_tab = "opticutter"
            self._opticutter_dirty = True
            self.confirmed = False

        def _ensure_custom_bom_tab(self):
            return "custom-tab"

        def _confirm_opticutter_update(self):
            self.confirmed = True
            self._opticutter_dirty = False
            return True

    app = DummyApp()
    event = types.SimpleNamespace(widget=app.nb)

    app._handle_tab_changed(event)

    assert app.confirmed is True
    assert app.nb.selected_to == "custom-tab"
    assert app._last_selected_notebook_tab == "custom-tab"


def test_confirm_opticutter_update_sets_status_and_toast():
    class DummyLabel:
        def __init__(self):
            self.configured = {}

        def configure(self, **kwargs):
            self.configured.update(kwargs)

    class DummyApp:
        _confirm_opticutter_update = App._confirm_opticutter_update
        _format_opticutter_update_message = App._format_opticutter_update_message
        _set_opticutter_update_status = App._set_opticutter_update_status

        def __init__(self):
            self._opticutter_dirty = True
            self._opticutter_refresh_after_id = None
            self.opticutter_last_analysis = types.SimpleNamespace(profiles=[1, 2, 3])
            self.opticutter_update_status_var = DummyVar("")
            self.opticutter_update_status_label = DummyLabel()
            self.status_var = DummyVar("")
            self.toasts = []

        def _show_transient_toast(self, message):
            self.toasts.append(message)

    app = DummyApp()

    assert app._confirm_opticutter_update() is True
    assert app._opticutter_dirty is False
    assert app.status_var.get() == "Opticutter bijgewerkt voor 3 profielen."
    assert app.toasts == ["Opticutter bijgewerkt voor 3 profielen."]
    assert app.opticutter_update_status_var.get().startswith("Laatst bijgewerkt ")
    assert app.opticutter_update_status_label.configured["fg"] == "#2F855A"


def test_opticutter_dirty_marker_updates_inline_status():
    class DummyLabel:
        def __init__(self):
            self.configured = {}

        def configure(self, **kwargs):
            self.configured.update(kwargs)

    class DummyApp:
        _mark_opticutter_dirty = App._mark_opticutter_dirty
        _set_opticutter_update_status = App._set_opticutter_update_status

        def __init__(self):
            self._opticutter_dirty = False
            self.opticutter_update_status_var = DummyVar("")
            self.opticutter_update_status_label = DummyLabel()

    app = DummyApp()

    app._mark_opticutter_dirty()

    assert app._opticutter_dirty is True
    assert app.opticutter_update_status_var.get() == "Wijzigingen actief"
    assert app.opticutter_update_status_label.configured["fg"] == "#B7791F"
