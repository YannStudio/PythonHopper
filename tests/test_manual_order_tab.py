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
