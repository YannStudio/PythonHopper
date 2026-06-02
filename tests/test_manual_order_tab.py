import math

import pytest

import manual_order_tab
from manual_order_tab import _ensure_integer_quantity, ManualOrderTab, SearchableCombobox


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


class _DummyVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


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


def test_refresh_data_uses_searchable_supplier_choices():
    tab = ManualOrderTab.__new__(ManualOrderTab)
    tab.clients_db = None
    tab.suppliers_db = _FakeSuppliersDb()
    tab.delivery_db = None
    tab.client_var = _DummyVar("")
    tab.client_combo = _DummyCombo()
    tab.supplier_var = _DummyVar("Leverancier B")
    tab.supplier_combo = _DummySearchableCombo()
    tab.delivery_var = _DummyVar("")
    tab.delivery_combo = _DummyCombo()

    ManualOrderTab.refresh_data(tab)

    assert tab.supplier_combo.set_choices_calls == [
        ["Geen", "Leverancier A", "Leverancier B"]
    ]
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
