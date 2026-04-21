import math

import pytest

from manual_order_tab import _ensure_integer_quantity, ManualOrderTab


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
