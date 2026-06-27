from decimal import Decimal
from types import SimpleNamespace

import pandas as pd

import bom_custom_tab
from bom_custom_tab import (
    BOMCustomTab,
    _find_matching_rows,
    _replace_matching_cells,
)


class _DummyTable:
    def __init__(self, rows):
        self.multiplerowlist = list(rows)
        self.currentrow = rows[0] if rows else None

    def _commit_active_edit(self):
        return True


class _DummyVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


def _build_tab(df: pd.DataFrame, *, selected_rows=None):
    tab = object.__new__(BOMCustomTab)
    tab.HEADERS = BOMCustomTab.HEADERS
    tab.DEFAULT_EMPTY_ROWS = BOMCustomTab.DEFAULT_EMPTY_ROWS
    tab.FIND_MATCH_MODE_LABELS = BOMCustomTab.FIND_MATCH_MODE_LABELS
    tab.FIND_ALL_COLUMNS_LABEL = BOMCustomTab.FIND_ALL_COLUMNS_LABEL
    tab.FIND_SCOPE_ALL_LABEL = BOMCustomTab.FIND_SCOPE_ALL_LABEL
    tab.FIND_SCOPE_SELECTION_LABEL = BOMCustomTab.FIND_SCOPE_SELECTION_LABEL
    tab.table_model = SimpleNamespace(df=df.copy(deep=True))
    tab._dataframe = tab.table_model.df.copy(deep=True)
    tab._qty_multiplier_reference = pd.Series(
        [""] * len(tab.table_model.df.index),
        index=tab.table_model.df.index,
        dtype=object,
    )
    tab._current_qty_multiplier = Decimal(1)
    tab.undo_stack = []
    tab.redo_stack = []
    tab.max_undo = 5
    tab.table = _DummyTable(selected_rows or [])
    statuses = []

    def update_status(self, text: str) -> None:
        statuses.append(text)

    def refresh(self) -> None:
        pass

    def ensure_rows(self, minimum: int) -> None:
        pass

    def set_dataframe(
        self,
        df: pd.DataFrame,
        *,
        reset_multiplier_reference: bool = True,
        update_multiplier_entry: bool = True,
    ) -> None:
        self.table_model.df = df.copy(deep=True)
        self._dataframe = self.table_model.df.copy(deep=True)

    def capture_qty_reference(self) -> pd.Series:
        return pd.Series(
            [""] * len(self.table_model.df.index),
            index=self.table_model.df.index,
            dtype=object,
        )

    def sync_qty_reference(self, _cells) -> None:
        pass

    def push_undo(self, *args, **kwargs) -> None:
        pass

    def update_preview(self, *_args) -> None:
        pass

    tab._update_status = update_status.__get__(tab, BOMCustomTab)
    tab._refresh_table = refresh.__get__(tab, BOMCustomTab)
    tab._ensure_minimum_rows = ensure_rows.__get__(tab, BOMCustomTab)
    tab._set_dataframe = set_dataframe.__get__(tab, BOMCustomTab)
    tab._capture_qty_reference_snapshot = capture_qty_reference.__get__(
        tab, BOMCustomTab
    )
    tab._sync_qty_reference_for_cells = sync_qty_reference.__get__(tab, BOMCustomTab)
    tab._push_undo = push_undo.__get__(tab, BOMCustomTab)
    tab._update_find_replace_preview = update_preview.__get__(tab, BOMCustomTab)
    return tab, statuses


def test_find_matching_rows_supports_exact_and_all_columns() -> None:
    df = pd.DataFrame(
        [
            {"Material": "Aluminium", "Description": "Plaat"},
            {"Material": "Staal", "Description": "Aluminium hoek"},
            {"Material": "RVS", "Description": "Buis"},
        ]
    )

    assert _find_matching_rows(df, ["Material"], "aluminium", match_mode="exact") == [0]
    assert _find_matching_rows(df, None, "aluminium", match_mode="contains") == [0, 1]


def test_find_matching_rows_respects_case_sensitive_flag() -> None:
    df = pd.DataFrame(
        [
            {"Material": "Dummy Part"},
            {"Material": "Dummy part"},
        ]
    )

    assert _find_matching_rows(
        df,
        ["Material"],
        "Dummy Part",
        match_mode="contains",
        case_sensitive=True,
    ) == [0]
    assert _find_matching_rows(
        df,
        ["Material"],
        "Dummy Part",
        match_mode="contains",
        case_sensitive=False,
    ) == [0, 1]


def test_find_replace_preview_explains_case_sensitive_near_misses() -> None:
    df = pd.DataFrame({"Material": ["Dummy part"]})
    tab, _statuses = _build_tab(df)
    tab._find_replace_query_var = _DummyVar("Dummy Part")
    tab._find_replace_column_var = _DummyVar("Material")
    tab._find_replace_scope_var = _DummyVar(BOMCustomTab.FIND_SCOPE_ALL_LABEL)
    tab._find_replace_match_mode_var = _DummyVar("Bevat")
    tab._find_replace_case_sensitive_var = _DummyVar(True)

    matches, message = BOMCustomTab._describe_find_matches(tab)

    assert matches == []
    assert message == "0 rijen met exact deze hoofdletters; 1 rij zonder dit vinkje."


def test_replace_reports_case_sensitive_near_miss_when_no_cells_changed() -> None:
    df = pd.DataFrame({"Material": ["Dummy part"]})
    tab, statuses = _build_tab(df)
    tab._find_replace_query_var = _DummyVar("Dummy Part")
    tab._find_replace_replacement_var = _DummyVar("Dummy part")
    tab._find_replace_column_var = _DummyVar("Material")
    tab._find_replace_scope_var = _DummyVar(BOMCustomTab.FIND_SCOPE_ALL_LABEL)
    tab._find_replace_match_mode_var = _DummyVar("Bevat")
    tab._find_replace_case_sensitive_var = _DummyVar(True)

    BOMCustomTab._replace_find_replace_matches(tab)

    assert statuses[-1] == "0 rijen met exact deze hoofdletters; 1 rij zonder dit vinkje."


def test_replace_matching_cells_supports_case_only_replacements() -> None:
    df = pd.DataFrame({"Material": ["Dummy Part"]})

    updated, changed_cells, changed_rows = _replace_matching_cells(
        df,
        ["Material"],
        "Dummy Part",
        "Dummy part",
        match_mode="contains",
        case_sensitive=True,
    )

    assert updated["Material"].tolist() == ["Dummy part"]
    assert changed_cells == [(0, 0)]
    assert changed_rows == [0]


def test_replace_matching_cells_limits_changes_to_selected_rows() -> None:
    df = pd.DataFrame(
        [
            {"Material": "Aluminium", "Description": "Plaat"},
            {"Material": "Aluminium", "Description": "Koker"},
            {"Material": "Staal", "Description": "Buis"},
        ]
    )

    updated, changed_cells, changed_rows = _replace_matching_cells(
        df,
        ["Material"],
        "Aluminium",
        "S235",
        match_mode="exact",
        row_scope=[1, 2],
    )

    assert updated["Material"].tolist() == ["Aluminium", "S235", "Staal"]
    assert changed_cells == [(1, 0)]
    assert changed_rows == [1]


def test_delete_selected_rows_from_toolbar_removes_selected_rows(monkeypatch) -> None:
    df = pd.DataFrame(
        {
            header: [f"{header}-1", f"{header}-2", f"{header}-3"]
            for header in BOMCustomTab.HEADERS
        }
    )
    tab, statuses = _build_tab(df, selected_rows=[1])
    monkeypatch.setattr(bom_custom_tab.messagebox, "askyesno", lambda *a, **k: True)

    BOMCustomTab._delete_selected_rows_from_toolbar(tab)

    assert len(tab.table_model.df.index) == 2
    assert tab.table_model.df.iloc[0, 0] == f"{BOMCustomTab.HEADERS[0]}-1"
    assert tab.table_model.df.iloc[1, 0] == f"{BOMCustomTab.HEADERS[0]}-3"
    assert statuses[-1] == "1 rij verwijderd."
