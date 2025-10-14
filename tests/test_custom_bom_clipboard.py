"""Tests voor clipboardhelper van de Custom BOM-tab."""

import pandas as pd

from bom_custom_tab import _dataframe_slice_to_clipboard


def test_clipboard_slice_basic_matrix() -> None:
    df = pd.DataFrame(
        [
            ["A1", "B1", "C1"],
            ["A2", "B2", pd.NA],
            ["A3", "B3", "C3"],
        ]
    )

    result = _dataframe_slice_to_clipboard(df, [0, 1], [0, 2])

    assert result == "A1\tC1\nA2\t"


def test_clipboard_slice_skips_invalid_indices() -> None:
    df = pd.DataFrame([["R1", "C1"], ["R2", "C2"]])

    result = _dataframe_slice_to_clipboard(df, [0, 5], [0, 1, 4])

    assert result == "R1\tC1"


def test_clipboard_slice_empty_when_nothing_selected() -> None:
    df = pd.DataFrame([["only", "value"]])

    assert _dataframe_slice_to_clipboard(df, [], []) == ""
