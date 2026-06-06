import math

import pytest

import manual_order_tab
from manual_order_tab import (
    _ensure_integer_quantity,
    _ManualRowWidgets,
    ManualOrderTab,
    SearchableCombobox,
)


@pytest.mark.parametrize(
    "value,expected",
    [
        ("", ""),
        (None, ""),
        (5, 5),
        (5.0, 5),
        (5.6, 6),
        ("7", 7),
        ("8,0", 8),
        ("9.2", 9),
    ],
)
def test_ensure_integer_quantity_basic(value, expected):
    assert _ensure_integer_quantity(value) == expected


def test_ensure_integer_quantity_invalid_values():
    invalid = math.inf
    assert _ensure_integer_quantity(invalid) is invalid
    text = "abc"
    assert _ensure_integer_quantity(text) == text


@pytest.mark.parametrize(
    "key,expected",
    [
        ("Aantal", True),
        ("aantal", True),
        ("AANTAL", True),
        ("qty", True),
        ("Quantity", True),
        ("Profiel", False),
        ("Gewicht", False),
    ],
)
def test_is_quantity_key_detection(key, expected):
    assert ManualOrderTab._is_quantity_key(key) is expected


def test_ensure_column_metrics_marks_integer_column():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab._entry_char_pixels = 8
    qty_column = {"key": "Aantal", "numeric": True, "width": 10}
    ManualOrderTab._ensure_column_metrics(tab, qty_column)
    assert qty_column.get("integer") is True

    length_column = {"key": "Lengte", "numeric": True, "width": 10}
    ManualOrderTab._ensure_column_metrics(tab, length_column)
    assert "integer" not in length_column


def test_collect_items_multiplies_surface_and_weight_by_quantity():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab.current_columns = [
        {"key": "Aantal", "numeric": True},
        {"key": "Oppervlakte", "numeric": True, "total_surface": True},
        {"key": "Gewicht", "numeric": True, "total_weight": True},
    ]
    tab.rows = [
        type(
            "Row",
            (),
            {
                "vars": {
                    "Aantal": _DummyVar("4"),
                    "Oppervlakte": _DummyVar("0,50"),
                    "Gewicht": _DummyVar("7,80"),
                }
            },
        )()
    ]

    payload = ManualOrderTab._collect_items(tab)

    assert payload["total_surface"] == pytest.approx(2.0)
    assert payload["total_weight"] == pytest.approx(31.2)


def test_build_document_basename_ignores_prefix_only_doc_number():
    assert (
        ManualOrderTab.build_document_basename(
            "BB-",
            "Leien-60x80",
            doc_type="Bestelbon",
        )
        == "Leien-60x80"
    )


def test_attach_entry_overflow_tooltip_retains_tooltip(monkeypatch):
    created = []

    class FakeTooltip:
        def __init__(self, widget, text_provider):
            self.widget = widget
            self.text_provider = text_provider
            created.append(self)

    monkeypatch.setattr(manual_order_tab, "_OverflowTooltip", FakeTooltip)
    tab = ManualOrderTab.__new__(ManualOrderTab)
    entry = object()

    tooltip = ManualOrderTab._attach_entry_overflow_tooltip(
        tab,
        entry,
        lambda: "C:/very/long/source/path/file.xlsx",
    )

    assert tooltip is created[0]
    assert tab._header_overflow_tooltips == [tooltip]
    assert tooltip.widget is entry
    assert tooltip.text_provider() == "C:/very/long/source/path/file.xlsx"


def test_attach_entry_overflow_tooltip_can_store_on_row(monkeypatch):
    class FakeTooltip:
        def __init__(self, widget, text_provider):
            self.widget = widget
            self.text_provider = text_provider

    monkeypatch.setattr(manual_order_tab, "_OverflowTooltip", FakeTooltip)
    tab = ManualOrderTab.__new__(ManualOrderTab)
    row_tooltips = []

    tooltip = ManualOrderTab._attach_entry_overflow_tooltip(
        tab,
        object(),
        lambda: "lange waarde",
        store=row_tooltips,
    )

    assert row_tooltips == [tooltip]
    assert not hasattr(tab, "_header_overflow_tooltips")


class _DummySelectionCombo:
    def __init__(self):
        self.clear_calls = 0
        self.cursor = None
        self.bindings = {}
        self.after_idle_calls = 0

    def selection_clear(self, *args):
        if args:
            raise AssertionError("ttk.Combobox.selection_clear expects no args")
        self.clear_calls += 1

    def icursor(self, index):
        self.cursor = index

    def bind(self, sequence, callback, add=None):
        self.bindings[sequence] = (callback, add)

    def after_idle(self, callback):
        self.after_idle_calls += 1
        callback()


def test_clear_combobox_text_selection_uses_ttk_signature():
    combo = _DummySelectionCombo()

    ManualOrderTab._clear_combobox_text_selection(combo)

    assert combo.clear_calls == 1
    assert combo.cursor == manual_order_tab.tk.END


def test_install_combobox_selection_reset_clears_after_selection():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    combo = _DummySelectionCombo()

    ManualOrderTab._install_combobox_selection_reset(tab, combo)
    combo.bindings["<<ComboboxSelected>>"][0](object())

    assert combo.bindings["<FocusIn>"][1] == "+"
    assert combo.bindings["<ButtonRelease-1>"][1] == "+"
    assert combo.clear_calls == 2
    assert combo.after_idle_calls == 1
    assert combo.cursor == manual_order_tab.tk.END


class _DummyVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


def test_manual_order_row_price_links_unit_and_total_price():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    row = _ManualRowWidgets(
        frame=None,
        vars={
            "Aantal": _DummyVar("4"),
            "Eenheidsprijs": _DummyVar("12,50"),
            "Totaalprijs": _DummyVar(""),
        },
        entries={},
        remove_btn=None,
    )

    ManualOrderTab._on_row_price_change(tab, row, "Eenheidsprijs")

    assert row.vars["Totaalprijs"].get() == "50.00"
    assert row.price_auto_field == "Totaalprijs"

    row.vars["Totaalprijs"].set("80")
    ManualOrderTab._on_row_price_change(tab, row, "Totaalprijs")

    assert row.vars["Eenheidsprijs"].get() == "20.00"
    assert row.price_auto_field == "Eenheidsprijs"


def test_manual_order_row_price_recalculates_when_quantity_changes():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    row = _ManualRowWidgets(
        frame=None,
        vars={
            "Aantal": _DummyVar("2"),
            "Eenheidsprijs": _DummyVar("12.50"),
            "Totaalprijs": _DummyVar(""),
        },
        entries={},
        remove_btn=None,
    )
    ManualOrderTab._on_row_price_change(tab, row, "Eenheidsprijs")

    row.vars["Aantal"].set("4")
    ManualOrderTab._on_row_quantity_change(tab, row)

    assert row.vars["Totaalprijs"].get() == "50.00"


def test_manual_order_vat_summary_rows_use_total_prices():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    layout = [
        {"key": "Description", "label": "Omschrijving"},
        {"key": "Aantal", "label": "Aantal"},
        {"key": "Eenheidsprijs", "label": "Eenheidsprijs"},
        {"key": "Totaalprijs", "label": "Totaalprijs"},
    ]
    items = [
        {"Description": "A", "Aantal": 2, "Totaalprijs": "20.00"},
        {"Description": "B", "Aantal": 1, "Totaalprijs": "5"},
    ]

    result = ManualOrderTab._append_vat_summary_rows(tab, items, layout, "21")

    assert result[-3]["Description"] == "Subtotaal excl. BTW"
    assert result[-3]["Totaalprijs"] == "25.00"
    assert result[-2]["Description"] == "BTW 21%"
    assert result[-2]["Totaalprijs"] == "5.25"
    assert result[-1]["Description"] == "Totaal incl. BTW"
    assert result[-1]["Totaalprijs"] == "30.25"


def test_manual_order_visible_columns_hide_price_fields():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab.price_columns_visible_var = _DummyVar(False)
    columns = [
        {"key": "Description", "label": "Omschrijving"},
        {"key": "Aantal", "label": "Aantal"},
        {"key": "Eenheidsprijs", "label": "Prijs/st. (\u20ac)"},
        {"key": "Totaalprijs", "label": "Totaal (\u20ac)"},
    ]

    visible = ManualOrderTab._visible_columns_for(tab, columns)

    assert [column["key"] for column in visible] == ["Description", "Aantal"]


def test_manual_order_visible_columns_keep_price_fields():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab.price_columns_visible_var = _DummyVar(True)
    columns = [
        {"key": "Description", "label": "Omschrijving"},
        {"key": "Eenheidsprijs", "label": "Prijs/st. (\u20ac)"},
        {"key": "Totaalprijs", "label": "Totaal (\u20ac)"},
    ]

    visible = ManualOrderTab._visible_columns_for(tab, columns)

    assert [column["key"] for column in visible] == [
        "Description",
        "Eenheidsprijs",
        "Totaalprijs",
    ]


def test_manual_order_apply_template_keeps_full_price_columns_when_hidden():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab.current_template_name = ""
    tab._template_rows_cache = {}
    tab._template_layout_cache = {}
    tab._all_columns = []
    tab._all_row_values = []
    tab.price_columns_visible_var = _DummyVar(False)
    columns = [
        {"key": "Description", "label": "Omschrijving"},
        {"key": "Eenheidsprijs", "label": "Prijs/st. (\u20ac)"},
        {"key": "Totaalprijs", "label": "Totaal (\u20ac)"},
    ]
    added_rows = []
    tab._clone_columns = lambda _template: [dict(column) for column in columns]
    tab._ensure_column_metrics = lambda _column: None
    tab._clear_rows = lambda: None
    tab._render_header = lambda: None
    tab.add_row = lambda values=None: added_rows.append(values)
    tab._update_totals = lambda: None

    ManualOrderTab._apply_template(tab, "Standaard", store_previous=False)

    assert [column["key"] for column in tab._all_columns] == [
        "Description",
        "Eenheidsprijs",
        "Totaalprijs",
    ]
    assert [column["key"] for column in tab.current_columns] == ["Description"]
    assert added_rows == [{}]


class _DummyText:
    def __init__(self, value=""):
        self.value = value

    def get(self, *_args):
        return self.value


class _DummySupplierCombo:
    def commit_typed_value(self):
        return True


def test_manual_order_export_omits_empty_price_columns(monkeypatch):
    monkeypatch.setattr(manual_order_tab.messagebox, "showinfo", lambda *a, **k: None)
    monkeypatch.setattr(manual_order_tab.messagebox, "showwarning", lambda *a, **k: None)
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab.supplier_combo = _DummySupplierCombo()
    tab.price_columns_visible_var = _DummyVar(True)
    tab.vat_rate_var = _DummyVar("21")
    tab.current_template_name = "Standaard"
    tab.current_columns = [
        {"key": "Description", "label": "Omschrijving"},
        {"key": "Aantal", "label": "Aantal", "numeric": True},
        {"key": "Eenheidsprijs", "label": "Prijs/st. (\u20ac)", "numeric": True},
        {"key": "Totaalprijs", "label": "Totaal (\u20ac)", "numeric": True},
    ]
    tab.rows = [
        _ManualRowWidgets(
            frame=None,
            vars={
                "Description": _DummyVar("Onderdeel"),
                "Aantal": _DummyVar("4"),
                "Eenheidsprijs": _DummyVar(""),
                "Totaalprijs": _DummyVar(""),
            },
            entries={},
            remove_btn=None,
        )
    ]
    tab.remark_text = _DummyText("")
    tab.doc_type_var = _DummyVar("Bestelbon")
    tab.doc_number_var = _DummyVar("BB001")
    tab.client_var = _DummyVar("Klant")
    tab.supplier_var = _DummyVar("Leverancier")
    tab.delivery_var = _DummyVar("Geen")
    tab.context_label_var = _DummyVar("Document")
    captured = {}
    tab._on_export = lambda payload: captured.setdefault("payload", payload)

    ManualOrderTab._handle_export(tab)

    export_payload = captured["payload"]
    exported_keys = [column["key"] for column in export_payload["column_layout"]]
    assert "Eenheidsprijs" not in exported_keys
    assert "Totaalprijs" not in exported_keys
    assert export_payload["items"] == [{"Description": "Onderdeel", "Aantal": 4}]
    assert export_payload["show_prices"] is False
    assert export_payload["vat_rate"] == ""


class _DummyCombo:
    def __init__(self):
        self.values = []

    def configure(self, **kwargs):
        if "values" in kwargs:
            self.values = list(kwargs["values"])


class _DummySearchableCombo(_DummyCombo):
    def __init__(self):
        super().__init__()
        self.set_choices_calls = []

    def set_choices(self, values):
        self.values = list(values)
        self.set_choices_calls.append(list(values))


class _FakeSuppliersDb:
    def suppliers_sorted(self):
        return ["Leverancier A", "Leverancier B"]

    def display_name(self, supplier):
        return supplier


def test_refresh_data_uses_supplier_choices():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab.clients_db = None
    tab.suppliers_db = _FakeSuppliersDb()
    tab.delivery_db = None
    tab.client_var = _DummyVar("")
    tab.client_combo = _DummyCombo()
    tab.supplier_var = _DummyVar("Leverancier B")
    tab.supplier_combo = _DummyCombo()
    tab.delivery_var = _DummyVar("")
    tab.delivery_combo = _DummyCombo()

    ManualOrderTab.refresh_data(tab)

    assert tab.supplier_combo.values == ["Geen", "Leverancier A", "Leverancier B"]
    assert tab.supplier_var.get() == "Leverancier B"


def _make_searchable_combo(values, current=""):
    combo = SearchableCombobox.__new__(SearchableCombobox)
    combo._all_values = list(values)
    combo._normalized_values = [
        (value, SearchableCombobox._normalize_text(value))
        for value in combo._all_values
    ]
    combo._last_query = ""
    combo._last_valid_value = "Geen"
    combo._readonly_mode = True
    combo.current = current
    combo.configured_values = []
    combo.dropdown_calls = []
    combo.cursor_at_end = False
    combo.focused = False
    combo.selection_is_present = False
    combo.after_callbacks = {}
    combo.cancelled_after_ids = []

    def configure(**kwargs):
        if "values" in kwargs:
            combo.configured_values = list(kwargs["values"])

    def get():
        return combo.current

    def set_value(value):
        combo.current = value

    def delete(_start, _end):
        combo.current = ""

    def insert(_index, value):
        combo.current = value

    def selection_present():
        return combo.selection_is_present

    def selection_clear():
        combo.selection_is_present = False

    def icursor(_index):
        combo.cursor_at_end = True

    def focus_set():
        combo.focused = True

    def after_idle(callback):
        callback()

    def after(_delay, callback):
        after_id = f"after-{len(combo.after_callbacks) + 1}"
        combo.after_callbacks[after_id] = callback
        return after_id

    def after_cancel(after_id):
        combo.cancelled_after_ids.append(after_id)
        combo.after_callbacks.pop(after_id, None)

    class _FakeTk:
        def call(self, *args):
            combo.dropdown_calls.append(args)

    combo.configure = configure
    combo.get = get
    combo.set = set_value
    combo.delete = delete
    combo.insert = insert
    combo.selection_present = selection_present
    combo.selection_clear = selection_clear
    combo.icursor = icursor
    combo.focus_set = focus_set
    combo.after_idle = after_idle
    combo.after = after
    combo.after_cancel = after_cancel
    combo.tk = _FakeTk()
    combo._w = "combo"
    return combo


def test_searchable_combobox_matches_accentless_supplier_text():
    combo = _make_searchable_combo(["Geen", "Café Metaal", "Delta Works"])

    assert SearchableCombobox._filter_values(combo, "cafe") == ["Café Metaal"]


def test_searchable_combobox_commits_typed_supplier_to_best_match():
    combo = _make_searchable_combo(
        ["Geen", "Alpha Lasers", "Metaalwerken NV"],
        current="metaal",
    )

    assert SearchableCombobox.commit_typed_value(combo) is True
    assert combo.current == "Metaalwerken NV"
    assert combo._last_valid_value == "Metaalwerken NV"


def test_searchable_combobox_prefers_supplier_over_empty_choice_for_partial_match():
    combo = _make_searchable_combo(
        ["Geen", "Govaerts Staal"],
        current="g",
    )

    assert SearchableCombobox.commit_typed_value(combo) is True
    assert combo.current == "Govaerts Staal"


def test_searchable_combobox_typing_first_letter_filters_and_opens_dropdown():
    combo = _make_searchable_combo(
        ["Geen", "Antwerp Steel", "Apple Metaal", "Beta Works"],
        current="a",
    )
    event = type("Event", (), {"keysym": "a", "char": "a", "state": 0})()

    SearchableCombobox._on_key_release(combo, event)

    assert combo.configured_values == ["Antwerp Steel", "Apple Metaal"]
    assert ("ttk::combobox::Post", "combo") in combo.dropdown_calls
    assert combo.current == "a"
    assert combo.cursor_at_end is True
    assert combo.focused is True


def test_searchable_combobox_button_press_restores_all_choices_and_posts_dropdown():
    combo = _make_searchable_combo([
        "Geen",
        "Alpha Lasers",
        "Beta Works",
    ], current="a")
    combo.identify = lambda x, y: "arrow"
    combo.event_generate = lambda event: combo.dropdown_calls.append(event)

    SearchableCombobox._on_button_press(combo, type("Event", (), {"x": 0, "y": 0})())

    assert combo.configured_values == ["Geen", "Alpha Lasers", "Beta Works"]
    assert ("ttk::combobox::Post", "combo") in combo.dropdown_calls


def test_searchable_combobox_post_dropdown_fallback_uses_down_event():
    combo = _make_searchable_combo(["Geen", "Alpha Lasers"])

    def fail_call(*args):
        raise Exception("post failed")

    combo.tk.call = fail_call
    combo.event_generate = lambda event: combo.dropdown_calls.append(event)

    SearchableCombobox._post_dropdown(combo)

    assert combo.dropdown_calls == ["<Down>"]


def test_searchable_combobox_keypress_replaces_existing_supplier_choice():
    combo = _make_searchable_combo(
        ["Geen", "Antwerp Steel", "Beta Works"],
        current="Geen",
    )
    event = type("Event", (), {"keysym": "a", "char": "a", "state": 0})()

    SearchableCombobox._on_key_press(combo, event)

    assert combo.current == ""


def test_searchable_combobox_focus_out_waits_for_dropdown_selection():
    combo = _make_searchable_combo(
        ["Geen", "Alpha Lasers", "Metaalwerken NV"],
        current="met",
    )

    SearchableCombobox._on_focus_out(combo, object())

    assert combo.current == "met"
    assert combo.dropdown_calls == []
    assert combo._focus_out_after_id == "after-1"

    combo.current = "Metaalwerken NV"
    SearchableCombobox._on_selection(combo, object())

    assert combo.current == "Metaalwerken NV"
    assert combo._last_valid_value == "Metaalwerken NV"
    assert combo.cancelled_after_ids == ["after-1"]


def test_searchable_combobox_focus_out_commits_after_delay():
    combo = _make_searchable_combo(
        ["Geen", "Alpha Lasers", "Metaalwerken NV"],
        current="metaal",
    )

    SearchableCombobox._on_focus_out(combo, object())
    combo.after_callbacks["after-1"]()

    assert combo.current == "Metaalwerken NV"
    assert combo._last_valid_value == "Metaalwerken NV"
