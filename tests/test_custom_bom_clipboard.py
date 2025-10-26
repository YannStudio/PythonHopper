"""Tests voor clipboardhelper van de Custom BOM-tab."""

from types import SimpleNamespace

import pandas as pd

from bom_custom_tab import BOMCustomTab, _dataframe_slice_to_clipboard


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


def _build_tab(df: pd.DataFrame):
    tab = object.__new__(BOMCustomTab)
    tab.table_model = SimpleNamespace(df=df.copy(deep=True))
    tab._dataframe = tab.table_model.df.copy(deep=True)
    tab.DEFAULT_EMPTY_ROWS = BOMCustomTab.DEFAULT_EMPTY_ROWS
    tab.HEADERS = BOMCustomTab.HEADERS
    tab.undo_stack = []
    tab.max_undo = 5
    tab._suspend_history = False
    status_updates = []

    def update_status(self, text: str) -> None:
        status_updates.append(text)

    def refresh(self) -> None:
        pass

    def ensure_rows(self, minimum: int) -> None:
        pass

    tab._update_status = update_status.__get__(tab, BOMCustomTab)
    tab._refresh_table = refresh.__get__(tab, BOMCustomTab)
    tab._ensure_minimum_rows = ensure_rows.__get__(tab, BOMCustomTab)
    return tab, status_updates


def test_clear_cells_clears_values_and_records_undo() -> None:
    df = pd.DataFrame({"A": ["one"], "B": ["two"]})
    tab, statuses = _build_tab(df)

    count = tab._clear_cells(
        [0],
        [0, 1],
        undo_action="test",
        empty_status="geen selectie",
        success_status="{count} cellen gewist",
        no_change_status="geen wijziging",
    )

    assert count == 2
    assert tab.table_model.df.iloc[0, 0] == ""
    assert tab.table_model.df.iloc[0, 1] == ""
    assert tab.undo_stack[-1].action == "test"
    assert statuses[-1] == "2 cellen gewist"


def test_clear_cells_handles_empty_selection() -> None:
    df = pd.DataFrame({"A": ["one"]})
    tab, statuses = _build_tab(df)

    count = tab._clear_cells(
        [],
        [0],
        undo_action="test",
        empty_status="geen selectie",
        success_status="{count} cellen gewist",
        no_change_status="geen wijziging",
    )

    assert count == 0
    assert statuses[-1] == "geen selectie"
    assert tab.undo_stack == []


def test_clear_cells_skips_when_values_already_empty() -> None:
    df = pd.DataFrame({"A": [""], "B": [" "]})
    tab, statuses = _build_tab(df)

    count = tab._clear_cells(
        [0],
        [0, 1],
        undo_action="test",
        empty_status="geen selectie",
        success_status="{count} cellen gewist",
        no_change_status="geen wijziging",
    )

    assert count == 1
    assert statuses[-1] == "1 cellen gewist"
    assert len(tab.undo_stack) == 1


def test_delete_rows_removes_rows_and_updates_status() -> None:
    df = pd.DataFrame(
        {header: [f"{header}-1", f"{header}-2", f"{header}-3"] for header in BOMCustomTab.HEADERS}
    )
    tab, statuses = _build_tab(df)

    count = tab._delete_rows([1])

    assert count == 1
    assert len(tab.table_model.df) == 2
    assert tab.table_model.df.iloc[0, 0] == f"{BOMCustomTab.HEADERS[0]}-1"
    assert tab.table_model.df.iloc[1, 0] == f"{BOMCustomTab.HEADERS[0]}-3"
    assert tab.undo_stack[-1].action == "rijen verwijderen"
    assert statuses[-1] == "1 rij verwijderd."


def test_delete_rows_handles_invalid_selection() -> None:
    df = pd.DataFrame({header: [f"{header}-1"] for header in BOMCustomTab.HEADERS})
    tab, statuses = _build_tab(df)

    count = tab._delete_rows([-1, 5])

    assert count == 0
    assert statuses[-1] == "Geen rijen geselecteerd om te verwijderen."
    assert len(tab.table_model.df) == 1
