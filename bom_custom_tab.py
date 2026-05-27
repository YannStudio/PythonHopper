"""Tkinter-tab voor het bewerken en exporteren van een custom BOM.

Installatie en uitvoeren
========================
1. Activeer de gewenste virtualenv.
2. Installeer de GUI-afhankelijkheden met ``pip install -r requirements.txt``.
3. ``pandastable`` is vereist voor het spreadsheet-grid. Installeer handmatig met
   ``pip install pandastable`` wanneer deze nog niet beschikbaar is.
4. Start de applicatie via ``python -m gui`` of ``python main.py``.

Variant B – tijdelijke export
=============================
De knop *"Exporteren"* schrijft altijd naar een tijdelijke CSV in
``%LOCALAPPDATA%/<appnaam>/temp/custom_bom.csv`` (of ``~/.local/share/<appnaam>/temp``
op niet-Windows platformen). Er wordt geen bestandsdialoog geopend; een eventueel
``on_custom_bom_ready``-callback of het Tk-event ``<<CustomBOMReady>>`` krijgt het
pad aangereikt.

Callback- en eventcontract
==========================
* ``on_custom_bom_ready(path: pathlib.Path, row_count: int) -> None`` wordt
  aangeroepen na een geslaagde export.
* Zonder callback wordt ``self.last_temp_csv_path`` gezet en het Tk-event
  ``<<CustomBOMReady>>`` uitgestuurd. Het event bevat ``event.data`` met het pad
  naar de CSV. Een handler kan het pad ophalen via ``event.data`` of via
  ``tab.get_last_export_path()``.

Undo-structuur
==============
Elke actie (plakken, verwijderen, bewerken, leegmaken) bewaart een snapshot van
het volledige sheet vóór de wijziging en aanvullende metadata (bijvoorbeeld de
gewijzigde cellen). De undo-stack bevat maximaal 50 stappen en wordt door
``Ctrl+Z`` in omgekeerde volgorde verwerkt.

CSV-schema
==========
Er worden zestien vaste kolommen geëxporteerd in deze volgorde:
``PartNumber, Description, QTY., Profile, Length profile, Thickness, Production, Material, Finish, RAL color, Weight (kg), Surface Area (m²), Supplier, Supplier code, Manufacturer, Manufacturer code``.
Velden worden met een komma gescheiden en automatisch gequote door ``csv.writer``.

Notebook-integratie
===================
``BOMCustomTab`` is een ``ttk.Frame``. Voeg de tab toe aan een bestaande
``ttk.Notebook`` met::

    from bom_custom_tab import BOMCustomTab

    tab = BOMCustomTab(notebook, app_name="Filehopper")
    notebook.add(tab, text="Custom BOM")

``tab.get_last_export_path()`` geeft het pad van de laatst weggeschreven CSV
terug (of ``None`` wanneer er nog niet geëxporteerd is)."""

from __future__ import annotations

import csv
import os
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

try:
    from pandastable import Table, TableModel
except ModuleNotFoundError as exc:  # pragma: no cover - afhankelijk van installatie
    _PANDASTABLE_IMPORT_ERROR: Optional[BaseException] = exc

    class _TableStub:
        def __init__(self, *args, **kwargs) -> None:
            raise RuntimeError(_PANDASTABLE_ERROR) from _PANDASTABLE_IMPORT_ERROR

    class _TableModelStub:
        def __init__(self, *args, **kwargs) -> None:
            raise RuntimeError(_PANDASTABLE_ERROR) from _PANDASTABLE_IMPORT_ERROR

    Table = _TableStub  # type: ignore[assignment]
    TableModel = _TableModelStub  # type: ignore[assignment]
    _PANDASTABLE_AVAILABLE = False
else:
    _PANDASTABLE_IMPORT_ERROR = None
    _PANDASTABLE_AVAILABLE = True

_PANDASTABLE_ERROR = (
    "De module 'pandastable' is niet geïnstalleerd. "
    "Voer 'pip install pandastable' uit voordat u de Filehopper GUI start."
)

CellCoord = Tuple[int, int]


def _ensure_pandastable_available() -> None:
    if _PANDASTABLE_AVAILABLE:
        return

    try:
        import tkinter as _tk
        from tkinter import messagebox as _messagebox

        _root = _tk.Tk()
        _root.withdraw()
        try:
            _messagebox.showerror("pandastable ontbreekt", _PANDASTABLE_ERROR)
        finally:
            _root.destroy()
    except Exception:
        # Val stilletjes terug op een consolefout wanneer Tkinter niet beschikbaar is
        pass
    raise RuntimeError(_PANDASTABLE_ERROR) from _PANDASTABLE_IMPORT_ERROR


@dataclass
class UndoEntry:
    """Snapshot van een wijziging."""

    action: str
    frame: pd.DataFrame
    cells: Sequence[CellCoord]
    qty_reference: Optional[pd.Series] = None
    qty_multiplier: Decimal = field(default_factory=lambda: Decimal(1))


def _dataframe_slice_to_clipboard(
    frame: pd.DataFrame, rows: Sequence[int], cols: Sequence[int]
) -> str:
    """Formatteer een DataFrame-deel naar tabgescheiden klembordtekst."""

    if not rows or not cols:
        return ""

    safe_rows = [index for index in rows if 0 <= index < len(frame.index)]
    safe_cols = [index for index in cols if 0 <= index < len(frame.columns)]

    if not safe_rows or not safe_cols:
        return ""

    lines: List[str] = []
    for row in safe_rows:
        cells: List[str] = []
        for col in safe_cols:
            value = frame.iat[row, col]
            normalized = "" if pd.isna(value) else str(value)
            cells.append(normalized)
        lines.append("\t".join(cells))
    return "\n".join(lines)


def _coerce_row_indices(
    frame: pd.DataFrame, row_scope: Optional[Sequence[int]] = None
) -> List[int]:
    row_count = len(frame.index)
    if row_scope is None:
        return list(range(row_count))

    rows: List[int] = []
    seen: set[int] = set()
    for raw in row_scope:
        try:
            index = int(raw)
        except (TypeError, ValueError):
            continue
        if 0 <= index < row_count and index not in seen:
            seen.add(index)
            rows.append(index)
    return rows


def _coerce_search_columns(
    frame: pd.DataFrame, columns: Optional[Sequence[str]] = None
) -> List[str]:
    if columns is None:
        return list(frame.columns)

    resolved: List[str] = []
    seen: set[str] = set()
    for column in columns:
        if column in frame.columns and column not in seen:
            seen.add(column)
            resolved.append(column)
    return resolved


def _build_match_pattern(
    query: str,
    *,
    match_mode: str = "contains",
    case_sensitive: bool = False,
) -> Optional[re.Pattern[str]]:
    if not query:
        return None

    escaped = re.escape(query)
    if match_mode == "exact":
        expression = rf"^{escaped}$"
    elif match_mode == "startswith":
        expression = rf"^{escaped}"
    else:
        expression = escaped

    flags = 0 if case_sensitive else re.IGNORECASE
    return re.compile(expression, flags)


def _cell_matches_query(
    value: object,
    query: str,
    *,
    match_mode: str = "contains",
    case_sensitive: bool = False,
) -> bool:
    pattern = _build_match_pattern(
        query,
        match_mode=match_mode,
        case_sensitive=case_sensitive,
    )
    if pattern is None:
        return False
    text = "" if pd.isna(value) else str(value)
    return pattern.search(text) is not None


def _replace_query_in_text(
    text: str,
    query: str,
    replacement: str,
    *,
    match_mode: str = "contains",
    case_sensitive: bool = False,
) -> str:
    pattern = _build_match_pattern(
        query,
        match_mode=match_mode,
        case_sensitive=case_sensitive,
    )
    if pattern is None:
        return text
    if match_mode == "exact":
        return replacement if pattern.fullmatch(text) is not None else text
    if match_mode == "startswith":
        return pattern.sub(replacement, text, count=1)
    return pattern.sub(replacement, text)


def _find_matching_rows(
    frame: pd.DataFrame,
    columns: Optional[Sequence[str]],
    query: str,
    *,
    match_mode: str = "contains",
    case_sensitive: bool = False,
    row_scope: Optional[Sequence[int]] = None,
) -> List[int]:
    if frame.empty or not query:
        return []

    search_columns = _coerce_search_columns(frame, columns)
    if not search_columns:
        return []

    matches: List[int] = []
    for row_index in _coerce_row_indices(frame, row_scope):
        row = frame.iloc[row_index]
        for column in search_columns:
            if _cell_matches_query(
                row.get(column, ""),
                query,
                match_mode=match_mode,
                case_sensitive=case_sensitive,
            ):
                matches.append(row_index)
                break
    return matches


def _replace_matching_cells(
    frame: pd.DataFrame,
    columns: Optional[Sequence[str]],
    query: str,
    replacement: str,
    *,
    match_mode: str = "contains",
    case_sensitive: bool = False,
    row_scope: Optional[Sequence[int]] = None,
) -> Tuple[pd.DataFrame, List[CellCoord], List[int]]:
    if frame.empty or not query:
        return frame.copy(deep=True), [], []

    search_columns = _coerce_search_columns(frame, columns)
    if not search_columns:
        return frame.copy(deep=True), [], []

    updated = frame.copy(deep=True)
    changed_cells: List[CellCoord] = []
    changed_rows: List[int] = []
    seen_rows: set[int] = set()

    for row_index in _coerce_row_indices(updated, row_scope):
        for column in search_columns:
            col_index = updated.columns.get_loc(column)
            current = updated.iat[row_index, col_index]
            current_text = "" if pd.isna(current) else str(current)
            if not _cell_matches_query(
                current_text,
                query,
                match_mode=match_mode,
                case_sensitive=case_sensitive,
            ):
                continue
            replaced = _replace_query_in_text(
                current_text,
                query,
                replacement,
                match_mode=match_mode,
                case_sensitive=case_sensitive,
            )
            if replaced == current_text:
                continue
            updated.iat[row_index, col_index] = replaced
            changed_cells.append((row_index, col_index))
            if row_index not in seen_rows:
                seen_rows.add(row_index)
                changed_rows.append(row_index)

    return updated, changed_cells, changed_rows


class _UndoableTableModel(TableModel):
    """Uitbreiding van :class:`pandastable.TableModel` met undo-notificaties."""

    def __init__(
        self,
        dataframe: pd.DataFrame,
        on_change: Callable[[pd.DataFrame, pd.DataFrame, CellCoord], None],
    ) -> None:
        super().__init__(dataframe=dataframe)
        self._on_change = on_change

    def setValueAt(  # type: ignore[override]
        self,
        value,
        rowIndex,
        columnIndex,
        df: Optional[pd.DataFrame] = None,
    ) -> bool:
        """Sla waarden op en informeer de undo-structuur bij wijzigingen."""

        target_df = self.df if df is None else df

        row = int(rowIndex)
        col = int(columnIndex)
        if row < 0 or col < 0:
            return False
        if row >= len(target_df) or col >= len(target_df.columns):
            return False

        if value is None:
            normalized_value = ""
        elif isinstance(value, float) and pd.isna(value):
            normalized_value = ""
        elif isinstance(value, str):
            normalized_value = value
        else:
            normalized_value = str(value)

        current_value = target_df.iat[row, col]
        normalized_current = "" if pd.isna(current_value) else str(current_value)
        if normalized_current == normalized_value:
            return True

        before_snapshot = self.df.copy(deep=True)

        target_df.iat[row, col] = normalized_value

        if target_df is not self.df:
            row_label = target_df.index[row]
            col_label = target_df.columns[col]
            try:
                self.df.loc[row_label, col_label] = normalized_value
            except Exception:
                # Fallback voor niet-unieke indexen of ontbrekende labels
                loc = self.df.index.get_loc(row_label)
                if isinstance(loc, slice):
                    base_row = loc.start
                else:
                    try:
                        base_row = int(loc)
                    except TypeError:
                        base_row = int(loc[0])
                self.df.iat[base_row, col] = normalized_value

            loc = self.df.index.get_loc(row_label)
            if isinstance(loc, slice):
                effective_row = int(loc.start)
            else:
                try:
                    effective_row = int(loc)
                except TypeError:
                    effective_row = int(loc[0])
        else:
            effective_row = row

        if self._on_change is not None:
            after_snapshot = self.df.copy(deep=True)
            self._on_change(before_snapshot, after_snapshot, (effective_row, col))
        return True


class _UndoAwareTable(Table):
    """Tabel die undo-/paste-acties doorverwijst naar :class:`BOMCustomTab`."""

    def __init__(self, parent: tk.Widget, owner: "BOMCustomTab", **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self._owner = owner
        self._active_edit: Optional[CellCoord] = None
        self._skip_focus_commit = False
        self._drag_selection_active = False

        for sequence in (
            "<Control-c>",
            "<Control-C>",
            "<Control-Insert>",
            "<Command-c>",
            "<Command-C>",
            "<Meta-c>",
            "<Meta-C>",
        ):
            self.bind(sequence, self._on_copy_shortcut, add="+")
        for sequence in (
            "<Control-x>",
            "<Control-X>",
            "<Shift-Delete>",
            "<Command-x>",
            "<Command-X>",
            "<Meta-x>",
            "<Meta-X>",
        ):
            self.bind(sequence, self.cut, add="+")
        for sequence in (
            "<Control-v>",
            "<Control-V>",
            "<Command-v>",
            "<Command-V>",
            "<Meta-v>",
            "<Meta-V>",
            "<<Paste>>",
            "<Shift-Insert>",
        ):
            self.bind(sequence, self.paste, add="+")

        self._install_custom_bindings()

    # ------------------------------------------------------------------
    # Pandastable overrides delegating to BOMCustomTab

    def copy(self, event=None):  # type: ignore[override]
        if not self._commit_active_edit():
            return "break"
        base = super()
        copy_func = getattr(base, "copy", None)
        if callable(copy_func):
            return copy_func(event)
        return "break"

    def cut(self, event=None):  # type: ignore[override]
        if not self._commit_active_edit():
            return "break"

        rows, cols = self._selected_cell_coordinates()
        copied_cells = len(rows) * len(cols) if rows and cols else 0
        if not copied_cells:
            self._owner._update_status("Geen cellen geselecteerd om te knippen.")
            return "break"

        clipboard_text = _dataframe_slice_to_clipboard(
            self._owner.table_model.df, rows, cols
        )
        try:
            self.clipboard_clear()
            self.clipboard_append(clipboard_text)
        except tk.TclError:
            return "break"

        self._owner._clear_cells(
            rows,
            cols,
            undo_action="knippen",
            empty_status="Geen cellen geselecteerd om te knippen.",
            success_status="{count} cellen geknipt.",
            no_change_status="Geen cellen geknipt.",
        )
        return "break"

    def paste(self, event=None):  # type: ignore[override]
        return self._owner._on_paste(event)

    def clearData(self, event=None):  # type: ignore[override]
        return self._owner._clear_selection(event)

    def undo(self, event=None):  # type: ignore[override]
        return self._owner._on_undo(event)

    def drawCellEntry(self, row, col, text=None):  # type: ignore[override]
        result = super().drawCellEntry(row, col, text=text)
        entry = getattr(self, "cellentry", None)
        if entry is not None:
            self._ensure_entry_bindings(entry)
        self._active_edit = (row, col)
        return result

    def handleCellEntry(self, row, col):  # type: ignore[override]
        self._skip_focus_commit = True
        try:
            return super().handleCellEntry(row, col)
        finally:
            self.after_idle(self._reset_focus_commit_guard)
            self._active_edit = None

    # ------------------------------------------------------------------
    # Custom binding helpers

    def _install_custom_bindings(self) -> None:
        self.bind("<Button-1>", self._on_primary_button, add="+")
        self.bind("<B1-Motion>", self._extend_drag_selection, add="+")
        self.bind("<ButtonRelease-1>", self._finalise_drag_selection, add="+")
        self.bind("<Key>", self._handle_table_key, add="+")
        self.bind("<Return>", self._on_return_key, add="+")
        self.bind("<KP_Enter>", self._on_return_key, add="+")
        self.bind("<Tab>", self._on_tab_key, add="+")
        try:
            self.bind("<ISO_Left_Tab>", self._on_shift_tab_key, add="+")
        except tk.TclError:
            # ``ISO_Left_Tab`` is specific to X11/Unix platforms.  The binding
            # does not exist on Windows and older Tk versions, which would
            # otherwise raise a TclError during initialisation.  In that case
            # we silently ignore the sequence; ``<Shift-Tab>`` (bound below)
            # continues to provide the reverse-tab behaviour.
            pass
        self.bind("<Shift-Tab>", self._on_shift_tab_key, add="+")


    def _ensure_entry_bindings(self, entry: tk.Entry) -> None:
        if getattr(entry, "_fh_bindings", False):
            return
        entry.bind("<FocusOut>", self._on_entry_focus_out, add="+")
        entry.bind("<Return>", self._on_entry_return, add="+")
        entry.bind("<KP_Enter>", self._on_entry_return, add="+")
        setattr(entry, "_fh_bindings", True)

    # ------------------------------------------------------------------
    # Mouse helpers

    def _on_primary_button(self, event: tk.Event):
        if self._active_edit is not None:
            committed = self._commit_active_edit()
            if not committed:
                return "break"
        # ``pandastable`` onthoudt welke bron (rijhoofd, kolomhoofd, cel)
        # voor het laatst is aangeklikt.  Wanneer een gebruiker eerst op het
        # rijhoofd klikt en daarna in het raster, bleef de bron op "row"
        # staan waardoor ``Delete`` alsnog de volledige rij leegmaakte.
        # Door hier expliciet terug te schakelen naar de standaardwaarde
        # herkennen we het daarna als een cel-selectie en worden alleen de
        # gekozen cellen geleegd.
        try:
            self.setLeftClickSrc("")
        except AttributeError:
            pass
        self._drag_selection_active = True
        return None

    def _extend_drag_selection(self, event: tk.Event) -> None:
        if not self._drag_selection_active:
            return
        if self.multiplerowlist and self.multiplecollist:
            self.drawMultipleCells()

    def _finalise_drag_selection(self, event: tk.Event) -> None:
        if self._drag_selection_active and self.multiplerowlist and self.multiplecollist:
            self.drawMultipleCells()
        self._drag_selection_active = False

    # ------------------------------------------------------------------
    # Keyboard helpers

    def _modifier_is_pressed(self, event: tk.Event) -> bool:
        state = getattr(event, "state", 0)
        control_mask = 0x0004
        meta_mask = 0x0008
        command_mask = 0x100000
        if state & (control_mask | meta_mask | command_mask):
            return True
        if event.keysym in {
            "Control_L",
            "Control_R",
            "Alt_L",
            "Alt_R",
            "Meta_L",
            "Meta_R",
            "Command",
            "Option_L",
            "Option_R",
        }:
            return True
        return False

    def _handle_table_key(self, event: tk.Event):
        if self._modifier_is_pressed(event):
            return None
        if event.keysym in {"Shift_L", "Shift_R", "Caps_Lock"}:
            return None
        if event.keysym in {"Left", "Right", "Up", "Down"}:
            return None
        if event.keysym in {"Return", "KP_Enter", "Tab"}:
            return None

        if event.char and event.char.isprintable():
            return self._start_direct_edit(event.char)
        if event.keysym in {"BackSpace", "Delete"}:
            return self._start_direct_edit("")
        return None

    def _start_direct_edit(self, initial: str):
        row = self.currentrow if self.currentrow is not None else 0
        col = self.currentcol if self.currentcol is not None else 0
        self.setSelectedRow(row)
        self.setSelectedCol(col)
        self.drawSelectedRow()
        self.drawSelectedRect(row, col)
        self.drawCellEntry(row, col)
        entry = getattr(self, "cellentry", None)
        if entry is not None:
            entry.delete(0, tk.END)
            if initial:
                entry.insert(0, initial)
            entry.icursor(tk.END)
            entry.focus_set()
        return "break"

    def _on_return_key(self, event: tk.Event):
        if self._active_edit is None:
            if self.currentrow is None or self.currentcol is None:
                return None
            self.drawCellEntry(self.currentrow, self.currentcol)
            return "break"
        self._commit_active_edit()
        self.focus_set()
        return "break"

    def _on_tab_key(self, event: tk.Event):
        if not self._commit_active_edit():
            return "break"
        self._move_selection(0, 1)
        return "break"

    def _on_shift_tab_key(self, event: tk.Event):
        if not self._commit_active_edit():
            return "break"
        self._move_selection(0, -1)
        return "break"

    def _on_entry_return(self, event: tk.Event):
        self._commit_active_edit(trigger_widget=event.widget)
        self.focus_set()
        return "break"

    def _on_entry_tab(self, event: tk.Event):
        if self._commit_active_edit(trigger_widget=event.widget):
            self.focus_set()
            self._move_selection(0, 1)
        return "break"

    def _on_entry_shift_tab(self, event: tk.Event):
        if self._commit_active_edit(trigger_widget=event.widget):
            self.focus_set()
            self._move_selection(0, -1)
        return "break"

    def _reset_focus_commit_guard(self) -> None:
        self._skip_focus_commit = False

    def _on_entry_focus_out(self, event: tk.Event) -> None:
        if self._skip_focus_commit:
            return
        self._commit_active_edit(trigger_widget=event.widget)

    def _commit_active_edit(self, trigger_widget: Optional[tk.Widget] = None) -> bool:
        if self._active_edit is None:
            return True

        entry = getattr(self, "cellentry", None)
        if entry is None:
            self._active_edit = None
            return True

        if trigger_widget is not None and trigger_widget is not entry:
            return True

        row, col = self._active_edit
        value = getattr(self, "cellentryvar", tk.StringVar()).get()

        if self.filtered == 1:
            self.delete("entry")
            self._active_edit = None
            return True

        result = self.model.setValueAt(value, row, col, df=None)
        if result is False:
            dtype = self.model.getColumnType(col)
            msg = (
                f"This column is {dtype} and is not compatible with the value {value}."
                " Change data type first."
            )
            messagebox.showwarning(
                "Incompatible type", msg, parent=self.parentframe
            )
            entry.after_idle(entry.focus_set)
            return False

        self.drawText(row, col, value, align=self.align)
        self.delete("entry")
        self._active_edit = None
        return True

    def _move_selection(self, delta_row: int, delta_col: int) -> None:
        total_rows = self.rows or self.model.getRowCount()
        total_cols = self.cols or self.model.getColumnCount()
        if total_rows <= 0 or total_cols <= 0:
            return

        current_row = self.currentrow if self.currentrow is not None else 0
        current_col = self.currentcol if self.currentcol is not None else 0

        new_row = max(0, min(total_rows - 1, current_row + delta_row))
        new_col = current_col + delta_col

        if new_col >= total_cols:
            new_col = 0
            new_row = min(total_rows - 1, new_row + 1)
        elif new_col < 0:
            new_col = total_cols - 1
            new_row = max(0, new_row - 1)

        self.setSelectedRow(new_row)
        self.setSelectedCol(new_col)
        self.drawSelectedRow()
        self.drawSelectedRect(new_row, new_col)
        self.rowheader.drawSelectedRows(new_row)

    def _selected_cell_coordinates(self) -> Tuple[List[int], List[int]]:
        rows, cols = self._owner._collect_selection()
        if not rows:
            current_row = getattr(self, "currentrow", None)
            if current_row is not None:
                rows = [int(current_row)]
        if not cols:
            current_col = getattr(self, "currentcol", None)
            if current_col is not None:
                cols = [int(current_col)]
        rows = sorted(dict.fromkeys(rows))
        cols = sorted(dict.fromkeys(cols))
        row_count = len(self._owner.table_model.df.index)
        col_count = len(self._owner.table_model.df.columns)
        rows = [index for index in rows if 0 <= index < row_count]
        cols = [index for index in cols if 0 <= index < col_count]
        return rows, cols

    def _on_copy_shortcut(self, event=None):
        if not self._commit_active_edit():
            return "break"

        rows, cols = self._selected_cell_coordinates()
        clipboard_text = _dataframe_slice_to_clipboard(
            self._owner.table_model.df, rows, cols
        )

        try:
            self.clipboard_clear()
            self.clipboard_append(clipboard_text)
        except tk.TclError:
            return "break"

        copied_cells = len(rows) * len(cols) if rows and cols else 0
        if copied_cells:
            self._owner._update_status(f"{copied_cells} cellen gekopieerd.")
        else:
            self._owner._update_status("Geen cellen beschikbaar om te kopiëren.")
        return "break"



class BOMCustomTab(ttk.Frame):
    """Tabblad met spreadsheet-functionaliteit voor custom BOM-data."""

    FIND_ALL_COLUMNS_LABEL: str = "Alle kolommen"
    FIND_SCOPE_ALL_LABEL: str = "Hele tabel"
    FIND_SCOPE_SELECTION_LABEL: str = "Geselecteerde rijen"
    FIND_MATCH_MODE_LABELS: Dict[str, str] = {
        "Bevat": "contains",
        "Exact": "exact",
        "Begint met": "startswith",
    }
    HEADERS: Tuple[str, ...] = (
        "PartNumber",
        "Description",
        "QTY.",
        "Profile",
        "Length profile",
        "Thickness",
        "Production",
        "Material",
        "Finish",
        "RAL color",
        "Weight (kg)",
        "Surface Area (m²)",
        "Supplier",
        "Supplier code",
        "Manufacturer",
        "Manufacturer code",
    )
    QTY_COLUMN_INDEX: int = HEADERS.index("QTY.")
    TEMPLATE_DEFAULT_FILENAME: str = "BOM-FileHopper-Temp.xlsx"
    DEFAULT_EMPTY_ROWS: int = 20
    MAIN_TO_CUSTOM_COLUMN_MAP: Dict[str, str] = {
        "PartNumber": "PartNumber",
        "Description": "Description",
        "Profile": "Profile",
        "Length profile": "Length profile",
        "Plate thickness": "Thickness",
        "Production": "Production",
        "Materiaal": "Material",
        "Finish": "Finish",
        "RAL color": "RAL color",
        "Aantal": "QTY.",
        "Oppervlakte": "Surface Area (m²)",
        "Gewicht": "Weight (kg)",
        "Supplier": "Supplier",
        "Supplier code": "Supplier code",
        "Manufacturer": "Manufacturer",
        "Manufacturer code": "Manufacturer code",
    }
    CUSTOM_TO_MAIN_COLUMN_MAP: Dict[str, str] = {
        custom: main for main, custom in MAIN_TO_CUSTOM_COLUMN_MAP.items()
    }
    MAIN_COLUMN_ORDER: Tuple[str, ...] = (
        "PartNumber",
        "Description",
        "Profile",
        "Length profile",
        "Plate thickness",
        "Production",
        "Materiaal",
        "Supplier",
        "Supplier code",
        "Manufacturer",
        "Manufacturer code",
        "Finish",
        "RAL color",
        "Aantal",
        "Oppervlakte",
        "Gewicht",
    )

    def __init__(
        self,
        master: tk.Widget,
        *,
        app_name: str = "Filehopper",
        on_custom_bom_ready: Optional[Callable[[Path, int], None]] = None,
        on_push_to_main: Optional[Callable[[pd.DataFrame], None]] = None,
        event_target: Optional[tk.Misc] = None,
        max_undo: int = 50,
    ) -> None:
        _ensure_pandastable_available()

        super().__init__(master)
        self.configure(padding=(12, 12))
        self.app_name = app_name
        self.on_custom_bom_ready = on_custom_bom_ready
        self.on_push_to_main = on_push_to_main
        self.event_target = event_target
        self.max_undo = max_undo
        self.undo_stack: List[UndoEntry] = []
        self.redo_stack: List[UndoEntry] = []
        self.last_temp_csv_path: Optional[Path] = None
        self._suspend_history = False
        self._qty_multiplier_reference: pd.Series = pd.Series(dtype=object)
        self._current_qty_multiplier: Decimal = Decimal(1)
        self._baseline_dataframe: pd.DataFrame = pd.DataFrame()
        self._baseline_qty_reference: pd.Series = pd.Series(dtype=object)
        self._baseline_multiplier: Decimal = Decimal(1)
        self._find_replace_dialog: Optional[tk.Toplevel] = None
        self._find_replace_query_var: Optional[tk.StringVar] = None
        self._find_replace_replacement_var: Optional[tk.StringVar] = None
        self._find_replace_column_var: Optional[tk.StringVar] = None
        self._find_replace_scope_var: Optional[tk.StringVar] = None
        self._find_replace_match_mode_var: Optional[tk.StringVar] = None
        self._find_replace_case_sensitive_var: Optional[tk.BooleanVar] = None
        self._find_replace_summary_var: Optional[tk.StringVar] = None

        self.status_var = tk.StringVar(value="")
        self.qty_multiplier_var = tk.StringVar(value="")

        self._build_toolbar()
        self._build_sheet()
        self._update_status("Gereed.")

    # ------------------------------------------------------------------
    # UI-opbouw
    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=8, pady=6)

        controls = ttk.Frame(bar)
        controls.pack(fill="x")
        controls.columnconfigure(1, weight=1)

        button_style = dict(
            bg="#FADFA8",
            activebackground="#F4C46C",
            fg="black",
            activeforeground="black",
            highlightthickness=0,
        )

        left_controls = ttk.Frame(controls)
        left_controls.grid(row=0, column=0, sticky="w")

        update_btn = tk.Button(
            left_controls,
            text="Update Main BOM",
            command=self._push_to_main,
            **button_style,
        )
        update_btn.pack(side="left", padx=(0, 4))
        self._update_main_btn = update_btn

        clear_btn = tk.Button(
            left_controls,
            text="Clear BOM",
            command=self._confirm_clear,
            **button_style,
        )
        clear_btn.pack(side="left", padx=(0, 4))

        undo_btn = ttk.Button(
            left_controls, text="Undo", command=self._handle_toolbar_undo
        )
        undo_btn.pack(side="left", padx=(0, 4))

        redo_btn = ttk.Button(
            left_controls, text="Redo", command=self._handle_toolbar_redo
        )
        redo_btn.pack(side="left", padx=(0, 4))

        reset_btn = ttk.Button(
            left_controls,
            text="Reset naar origineel",
            command=self._reset_to_baseline,
        )
        reset_btn.pack(side="left", padx=(0, 4))

        find_btn = ttk.Button(
            left_controls,
            text="Zoek / vervang",
            command=self._open_find_replace_dialog,
        )
        find_btn.pack(side="left", padx=(0, 4))

        delete_selection_btn = ttk.Button(
            left_controls,
            text="Verwijder selectie",
            command=self._delete_selected_rows_from_toolbar,
        )
        delete_selection_btn.pack(side="left", padx=(0, 4))

        status_label = ttk.Label(controls, textvariable=self.status_var, anchor="w")
        status_label.grid(row=0, column=1, sticky="ew", padx=(4, 4))

        multiplier_container = ttk.Frame(controls)
        multiplier_container.grid(row=0, column=2, sticky="e", padx=(4, 0))

        export_container = ttk.Frame(controls)
        export_container.grid(row=0, column=3, sticky="e", padx=(4, 0))

        ttk.Label(multiplier_container, text="Totaal aantal:").pack(side="left", padx=(0, 4))
        multiplier_entry = ttk.Entry(
            multiplier_container,
            textvariable=self.qty_multiplier_var,
            width=6,
            justify="right",
        )
        multiplier_entry.pack(side="left", padx=(0, 4))
        multiplier_entry.bind("<Return>", self._apply_qty_multiplier)

        apply_btn = ttk.Button(
            multiplier_container,
            text="Pas toe",
            command=self._apply_qty_multiplier,
        )
        apply_btn.pack(side="left")

        export_btn = ttk.Button(
            export_container, text="Exporteren", command=self._export_temp
        )
        export_btn.pack(side="left")

        template_btn = ttk.Button(
            export_container,
            text="Download template",
            command=self._download_template,
        )
        template_btn.pack(side="left", padx=(4, 0))

    def _build_sheet(self) -> None:
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        container.rowconfigure(1, weight=1)
        container.columnconfigure(1, weight=1)

        initial_df = self._create_empty_dataframe(self.DEFAULT_EMPTY_ROWS)
        self._dataframe = initial_df
        self.table_model = _UndoableTableModel(
            initial_df.copy(deep=True), self._on_model_change
        )
        self.table = _UndoAwareTable(
            container,
            self,
            model=self.table_model,
            showstatusbar=False,
            enable_menus=False,
            editable=True,
        )
        self.table.show()
        for sequence in (
            "<Control-z>",
            "<Control-Z>",
            "<Command-z>",
            "<Command-Z>",
            "<Meta-z>",
            "<Meta-Z>",
        ):
            self.table.bind(sequence, self._on_undo, add="+")
        for sequence in (
            "<Control-y>",
            "<Control-Y>",
            "<Control-Shift-z>",
            "<Control-Shift-Z>",
            "<Command-y>",
            "<Command-Y>",
            "<Command-Shift-z>",
            "<Command-Shift-Z>",
            "<Meta-y>",
            "<Meta-Y>",
            "<Meta-Shift-z>",
            "<Meta-Shift-Z>",
        ):
            self.table.bind(sequence, self._on_redo, add="+")
        for sequence in ("<Delete>", "<BackSpace>"):
            self.table.bind(sequence, self._clear_selection, add="+")
        for sequence in (
            "<Control-h>",
            "<Control-H>",
            "<Command-h>",
            "<Command-H>",
            "<Meta-h>",
            "<Meta-H>",
        ):
            self.table.bind(sequence, self._open_find_replace_dialog, add="+")

        self._reset_qty_multiplier_reference()
        self._store_baseline_state()


    # ------------------------------------------------------------------
    # Helpers
    def _update_status(self, text: str) -> None:
        self.status_var.set(text)

    def _format_decimal(self, value: Decimal) -> str:
        text = format(value, "f")
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        return text or "0"

    def _set_current_multiplier(self, value: Decimal, *, update_entry: bool = True) -> None:
        self._current_qty_multiplier = value
        if update_entry and hasattr(self, "qty_multiplier_var"):
            self.qty_multiplier_var.set(self._format_decimal(value))

    def _make_qty_series(self, frame: Optional[pd.DataFrame] = None) -> pd.Series:
        target = frame if frame is not None else self.table_model.df
        values: List[str] = []
        for value in target.iloc[:, self.QTY_COLUMN_INDEX]:
            if pd.isna(value):
                values.append("")
            else:
                values.append(str(value).strip())
        return pd.Series(values, index=target.index, dtype=object)

    def _align_qty_reference(self) -> pd.Series:
        df = self.table_model.df
        if not hasattr(self, "_qty_multiplier_reference"):
            self._qty_multiplier_reference = pd.Series(dtype=object)
        reference = self._qty_multiplier_reference
        if not isinstance(reference, pd.Series):
            reference = self._make_qty_series(df)
        else:
            reference = reference.reindex(df.index, fill_value="")
        self._qty_multiplier_reference = reference
        if not hasattr(self, "_current_qty_multiplier"):
            self._current_qty_multiplier = Decimal(1)
        return self._qty_multiplier_reference

    def _reset_qty_multiplier_reference(self, *, update_entry: bool = True) -> None:
        self._qty_multiplier_reference = self._make_qty_series(self.table_model.df)
        self._set_current_multiplier(Decimal(1), update_entry=update_entry)

    def _store_baseline_state(self) -> None:
        """Bewaar de huidige dataset als referentie voor de resetknop."""

        self._baseline_dataframe = self.table_model.df.copy(deep=True)
        self._baseline_qty_reference = self._capture_qty_reference_snapshot()
        self._baseline_multiplier = self._current_qty_multiplier

    def _capture_qty_reference_snapshot(self) -> pd.Series:
        return self._align_qty_reference().copy(deep=True)

    def _update_qty_reference_for_row(self, row: int, value: Any) -> None:
        reference = self._align_qty_reference()
        if row not in reference.index:
            return
        normalized = "" if pd.isna(value) else str(value).strip()
        if not normalized:
            reference.iloc[row] = ""
            return
        raw_value = normalized.replace(",", ".")
        multiplier = self._current_qty_multiplier
        if multiplier == 0:
            reference.iloc[row] = normalized
            return
        try:
            qty_decimal = Decimal(raw_value)
        except InvalidOperation:
            reference.iloc[row] = normalized
            return
        if multiplier == 1:
            base_decimal = qty_decimal
        else:
            base_decimal = qty_decimal / multiplier
        formatted = format(base_decimal, "f")
        if "." in formatted:
            formatted = formatted.rstrip("0").rstrip(".")
        reference.iloc[row] = formatted or "0"

    def _sync_qty_reference_for_cells(self, cells: Sequence[CellCoord]) -> None:
        if not cells:
            return
        reference = self._align_qty_reference()
        for row, col in cells:
            if col != self.QTY_COLUMN_INDEX:
                continue
            if row not in reference.index:
                continue
            value = self.table_model.df.iat[row, col]
            self._update_qty_reference_for_row(row, value)

    def _restore_history_entry(self, entry: UndoEntry) -> None:
        self._set_dataframe(
            entry.frame,
            reset_multiplier_reference=False,
            update_multiplier_entry=False,
        )
        self._dataframe = self.table_model.df.copy(deep=True)
        if entry.qty_reference is not None:
            self._qty_multiplier_reference = entry.qty_reference.copy(deep=True)
        else:
            self._qty_multiplier_reference = self._make_qty_series(self.table_model.df)
        self._align_qty_reference()
        self._set_current_multiplier(entry.qty_multiplier)

    def _create_empty_dataframe(self, rows: int) -> pd.DataFrame:
        data = {header: [""] * rows for header in self.HEADERS}
        return pd.DataFrame(data, columns=self.HEADERS)

    def _set_dataframe(
        self,
        df: pd.DataFrame,
        *,
        reset_multiplier_reference: bool = True,
        update_multiplier_entry: bool = True,
    ) -> None:
        normalized = df.reindex(columns=self.HEADERS, fill_value="")
        self._suspend_history = True
        try:
            self.table_model.df = normalized
            self._refresh_table()
        finally:
            self._suspend_history = False
        self._dataframe = self.table_model.df.copy(deep=True)
        if reset_multiplier_reference:
            self._reset_qty_multiplier_reference(update_entry=update_multiplier_entry)
        else:
            self._align_qty_reference()

    def _refresh_table(self) -> None:
        self.table.updateModel(self.table_model)
        self.table.redraw()

    def _append_empty_rows(self, count: int) -> None:
        if count <= 0:
            return
        extension = self._create_empty_dataframe(count)
        self.table_model.df = pd.concat(
            [self.table_model.df, extension], ignore_index=True
        )
        self._align_qty_reference()

    def _ensure_minimum_rows(self, minimum: int = 1) -> None:
        if minimum <= 0:
            return
        current_rows = len(self.table_model.df)
        if current_rows >= minimum:
            return
        missing = minimum - current_rows
        extension = self._create_empty_dataframe(missing)
        extended = pd.concat([self.table_model.df, extension], ignore_index=True)
        self._set_dataframe(extended, reset_multiplier_reference=False, update_multiplier_entry=False)

    def _snapshot_data(self, frame: Optional[pd.DataFrame] = None) -> List[List[str]]:
        df = frame if frame is not None else self.table_model.df
        normalized = df.reindex(columns=self.HEADERS)
        return normalized.fillna("").astype(str).values.tolist()

    def _commit_pending_table_edit(self) -> bool:
        commit = getattr(getattr(self, "table", None), "_commit_active_edit", None)
        if not callable(commit):
            return True
        return bool(commit())

    def _push_undo(
        self,
        action: str,
        frame: pd.DataFrame,
        cells: Sequence[CellCoord],
        *,
        qty_reference: Optional[pd.Series] = None,
        qty_multiplier: Optional[Decimal] = None,
    ) -> None:
        snapshot = frame.copy(deep=True)
        if qty_reference is not None:
            reference_snapshot = qty_reference.copy(deep=True)
        else:
            reference_snapshot = self._capture_qty_reference_snapshot()
        multiplier_snapshot = (
            qty_multiplier if qty_multiplier is not None else self._current_qty_multiplier
        )
        if not hasattr(self, "undo_stack"):
            self.undo_stack = []
        if not hasattr(self, "redo_stack"):
            self.redo_stack = []
        if not hasattr(self, "max_undo"):
            self.max_undo = 50
        self.undo_stack.append(
            UndoEntry(
                action=action,
                frame=snapshot,
                cells=list(cells),
                qty_reference=reference_snapshot,
                qty_multiplier=multiplier_snapshot,
            )
        )
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _clear_cells(
        self,
        rows: Sequence[int],
        cols: Sequence[int],
        *,
        undo_action: str,
        empty_status: str,
        success_status: str,
        no_change_status: str,
    ) -> int:
        if not rows or not cols:
            self._update_status(empty_status)
            return 0

        before = self.table_model.df.copy(deep=True)
        before_reference = self._capture_qty_reference_snapshot()
        changed: List[CellCoord] = []
        self._suspend_history = True
        try:
            for row in rows:
                for col in cols:
                    current = self.table_model.df.iat[row, col]
                    normalized = "" if pd.isna(current) else str(current)
                    if normalized:
                        self.table_model.df.iat[row, col] = ""
                        changed.append((row, col))
        finally:
            self._suspend_history = False

        after = self.table_model.df.copy(deep=True)
        if changed and not before.equals(after):
            self._push_undo(
                undo_action, before, changed, qty_reference=before_reference
            )
            self._dataframe = after
            self._refresh_table()
            self._sync_qty_reference_for_cells(changed)
            self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
            self._update_status(success_status.format(count=len(changed)))
            return len(changed)

        self._update_status(no_change_status)
        return 0

    def _selected_row_indices(self) -> List[int]:
        if not hasattr(self, "table"):
            return []
        table = self.table
        rows = getattr(table, "multiplerowlist", None) or []
        safe_rows = _coerce_row_indices(self.table_model.df, rows)
        if safe_rows:
            return safe_rows

        current = getattr(table, "currentrow", None)
        if current is None:
            return []
        return _coerce_row_indices(self.table_model.df, [current])

    def _select_rows_in_table(self, rows: Sequence[int]) -> None:
        if not hasattr(self, "table"):
            return
        safe_rows = _coerce_row_indices(self.table_model.df, rows)
        table = self.table
        try:
            table.clearSelected()
        except Exception:
            pass

        table.multiplerowlist = []
        table.multiplecollist = []
        if not safe_rows:
            table.currentrow = None
            try:
                table.redraw()
            except Exception:
                pass
            return

        first = safe_rows[0]
        try:
            table.movetoSelection(row=first, col=0)
        except Exception:
            table.currentrow = first
            table.currentcol = 0

        table.setSelectedRows(safe_rows)
        table.currentrow = first
        table.startrow = first
        table.endrow = safe_rows[-1]

        try:
            table.drawMultipleRows(safe_rows)
        except Exception:
            try:
                table.drawSelectedRow()
            except Exception:
                pass
        try:
            table.rowheader.drawSelectedRows(safe_rows)
        except Exception:
            pass
        try:
            table.setLeftClickSrc("row")
        except Exception:
            setattr(table, "_Table__last_left_click_src", "row")
        try:
            table.focus_set()
        except Exception:
            pass

    def _resolve_find_replace_columns(self) -> List[str]:
        selected = (
            self._find_replace_column_var.get().strip()
            if self._find_replace_column_var is not None
            else ""
        )
        if not selected or selected == self.FIND_ALL_COLUMNS_LABEL:
            return list(self.HEADERS)
        if selected in self.HEADERS:
            return [selected]
        return []

    def _resolve_find_replace_row_scope(self) -> Optional[List[int]]:
        scope = (
            self._find_replace_scope_var.get().strip()
            if self._find_replace_scope_var is not None
            else self.FIND_SCOPE_ALL_LABEL
        )
        if scope == self.FIND_SCOPE_SELECTION_LABEL:
            return self._selected_row_indices()
        return None

    def _resolve_find_replace_match_mode(self) -> str:
        label = (
            self._find_replace_match_mode_var.get().strip()
            if self._find_replace_match_mode_var is not None
            else ""
        )
        return self.FIND_MATCH_MODE_LABELS.get(label, "contains")

    def _describe_find_matches(self) -> Tuple[List[int], str]:
        query = (
            self._find_replace_query_var.get().strip()
            if self._find_replace_query_var is not None
            else ""
        )
        if not query:
            return [], "Vul zoektekst in."

        columns = self._resolve_find_replace_columns()
        if not columns:
            return [], "Geen geldige kolommen geselecteerd."

        row_scope = self._resolve_find_replace_row_scope()
        if (
            row_scope is not None
            and self._find_replace_scope_var is not None
            and self._find_replace_scope_var.get().strip()
            == self.FIND_SCOPE_SELECTION_LABEL
            and not row_scope
        ):
            return [], "Geen geselecteerde rijen als zoekbereik."

        matches = _find_matching_rows(
            self.table_model.df,
            columns,
            query,
            match_mode=self._resolve_find_replace_match_mode(),
            case_sensitive=bool(
                self._find_replace_case_sensitive_var.get()
            )
            if self._find_replace_case_sensitive_var is not None
            else False,
            row_scope=row_scope,
        )
        count = len(matches)
        label = "rij" if count == 1 else "rijen"
        return matches, f"{count} {label} voldoen aan de zoekcriteria."

    def _update_find_replace_preview(self, *_args) -> None:
        if self._find_replace_summary_var is None:
            return
        _matches, message = self._describe_find_matches()
        self._find_replace_summary_var.set(message)

    def _close_find_replace_dialog(self) -> None:
        dialog = self._find_replace_dialog
        self._find_replace_dialog = None
        self._find_replace_query_var = None
        self._find_replace_replacement_var = None
        self._find_replace_column_var = None
        self._find_replace_scope_var = None
        self._find_replace_match_mode_var = None
        self._find_replace_case_sensitive_var = None
        self._find_replace_summary_var = None
        if dialog is None:
            return
        try:
            dialog.destroy()
        except tk.TclError:
            pass

    def _open_find_replace_dialog(self, event=None):
        existing = self._find_replace_dialog
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    query_entry = getattr(existing, "_fh_query_entry", None)
                    if query_entry is not None:
                        query_entry.focus_set()
                        query_entry.selection_range(0, tk.END)
                    self._update_find_replace_preview()
                    return "break" if event is not None else None
            except tk.TclError:
                self._find_replace_dialog = None

        dialog = tk.Toplevel(self)
        dialog.title("Zoek / vervang")
        dialog.transient(self.winfo_toplevel())
        dialog.resizable(False, False)
        dialog.protocol("WM_DELETE_WINDOW", self._close_find_replace_dialog)
        dialog.bind(
            "<Destroy>",
            lambda e, win=dialog: self._close_find_replace_dialog()
            if e.widget is win
            else None,
            add="+",
        )

        self._find_replace_dialog = dialog
        self._find_replace_query_var = tk.StringVar()
        self._find_replace_replacement_var = tk.StringVar()
        self._find_replace_column_var = tk.StringVar(
            value="Material" if "Material" in self.HEADERS else self.HEADERS[0]
        )
        self._find_replace_scope_var = tk.StringVar(value=self.FIND_SCOPE_ALL_LABEL)
        self._find_replace_match_mode_var = tk.StringVar(value="Bevat")
        self._find_replace_case_sensitive_var = tk.BooleanVar(value=False)
        self._find_replace_summary_var = tk.StringVar(value="Vul zoektekst in.")

        body = ttk.Frame(dialog, padding=12)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        ttk.Label(body, text="Kolom:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        column_combo = ttk.Combobox(
            body,
            textvariable=self._find_replace_column_var,
            values=[self.FIND_ALL_COLUMNS_LABEL, *self.HEADERS],
            state="readonly",
            width=24,
        )
        column_combo.grid(row=0, column=1, sticky="ew", pady=(0, 6))

        ttk.Label(body, text="Zoek naar:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        query_entry = ttk.Entry(body, textvariable=self._find_replace_query_var, width=32)
        query_entry.grid(row=1, column=1, sticky="ew", pady=6)
        dialog._fh_query_entry = query_entry

        ttk.Label(body, text="Vervang door:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        replacement_entry = ttk.Entry(
            body,
            textvariable=self._find_replace_replacement_var,
            width=32,
        )
        replacement_entry.grid(row=2, column=1, sticky="ew", pady=6)

        ttk.Label(body, text="Zoektype:").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=6)
        mode_combo = ttk.Combobox(
            body,
            textvariable=self._find_replace_match_mode_var,
            values=list(self.FIND_MATCH_MODE_LABELS.keys()),
            state="readonly",
            width=24,
        )
        mode_combo.grid(row=3, column=1, sticky="ew", pady=6)

        ttk.Label(body, text="Bereik:").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=6)
        scope_combo = ttk.Combobox(
            body,
            textvariable=self._find_replace_scope_var,
            values=[self.FIND_SCOPE_ALL_LABEL, self.FIND_SCOPE_SELECTION_LABEL],
            state="readonly",
            width=24,
        )
        scope_combo.grid(row=4, column=1, sticky="ew", pady=6)

        ttk.Checkbutton(
            body,
            text="Hoofdlettergevoelig",
            variable=self._find_replace_case_sensitive_var,
        ).grid(row=5, column=1, sticky="w", pady=(4, 6))

        ttk.Label(
            body,
            textvariable=self._find_replace_summary_var,
            anchor="w",
            foreground="#4A4A4A",
        ).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(4, 10))

        actions = ttk.Frame(body)
        actions.grid(row=7, column=0, columnspan=2, sticky="e")

        ttk.Button(
            actions,
            text="Selecteer treffers",
            command=self._select_find_replace_matches,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            actions,
            text="Vervang treffers",
            command=self._replace_find_replace_matches,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            actions,
            text="Verwijder selectie",
            command=self._delete_selected_rows_from_toolbar,
        ).pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="Sluiten", command=self._close_find_replace_dialog).pack(
            side="left"
        )

        for variable in (
            self._find_replace_query_var,
            self._find_replace_column_var,
            self._find_replace_scope_var,
            self._find_replace_match_mode_var,
            self._find_replace_case_sensitive_var,
        ):
            variable.trace_add("write", self._update_find_replace_preview)

        query_entry.bind("<Return>", self._select_find_replace_matches, add="+")
        replacement_entry.bind("<Return>", self._replace_find_replace_matches, add="+")
        dialog.bind("<Escape>", lambda _e: self._close_find_replace_dialog(), add="+")

        self._update_find_replace_preview()
        query_entry.focus_set()
        return "break" if event is not None else None

    def _select_find_replace_matches(self, event=None):
        if not self.table._commit_active_edit():
            return "break" if event is not None else None

        matches, message = self._describe_find_matches()
        if not matches:
            self._select_rows_in_table([])
            self._update_status(message)
            self._update_find_replace_preview()
            return "break" if event is not None else None

        self._select_rows_in_table(matches)
        count = len(matches)
        label = "rij" if count == 1 else "rijen"
        self._update_status(f"{count} {label} geselecteerd.")
        self._update_find_replace_preview()
        return "break" if event is not None else None

    def _replace_find_replace_matches(self, event=None):
        if not self.table._commit_active_edit():
            return "break" if event is not None else None

        query = (
            self._find_replace_query_var.get().strip()
            if self._find_replace_query_var is not None
            else ""
        )
        if not query:
            self._update_status("Vul eerst zoektekst in.")
            self._update_find_replace_preview()
            return "break" if event is not None else None

        columns = self._resolve_find_replace_columns()
        row_scope = self._resolve_find_replace_row_scope()
        if (
            row_scope is not None
            and self._find_replace_scope_var is not None
            and self._find_replace_scope_var.get().strip()
            == self.FIND_SCOPE_SELECTION_LABEL
            and not row_scope
        ):
            self._update_status("Geen geselecteerde rijen als zoekbereik.")
            self._update_find_replace_preview()
            return "break" if event is not None else None

        replacement = (
            self._find_replace_replacement_var.get()
            if self._find_replace_replacement_var is not None
            else ""
        )
        before = self.table_model.df.copy(deep=True)
        before_reference = self._capture_qty_reference_snapshot()
        updated, changed_cells, changed_rows = _replace_matching_cells(
            before,
            columns,
            query,
            replacement,
            match_mode=self._resolve_find_replace_match_mode(),
            case_sensitive=bool(
                self._find_replace_case_sensitive_var.get()
            )
            if self._find_replace_case_sensitive_var is not None
            else False,
            row_scope=row_scope,
        )
        if not changed_cells:
            self._update_status("Geen waarden vervangen.")
            self._update_find_replace_preview()
            return "break" if event is not None else None

        self._push_undo(
            "zoek/vervang",
            before,
            changed_cells,
            qty_reference=before_reference,
        )
        self._set_dataframe(
            updated,
            reset_multiplier_reference=False,
            update_multiplier_entry=False,
        )
        self._dataframe = self.table_model.df.copy(deep=True)
        self._sync_qty_reference_for_cells(changed_cells)
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._select_rows_in_table(changed_rows)

        cell_count = len(changed_cells)
        row_count = len(changed_rows)
        cell_label = "cel" if cell_count == 1 else "cellen"
        row_label = "rij" if row_count == 1 else "rijen"
        self._update_status(
            f"{cell_count} {cell_label} vervangen in {row_count} {row_label}."
        )
        self._update_find_replace_preview()
        return "break" if event is not None else None

    def _apply_qty_multiplier(self, event=None):
        if not self.table._commit_active_edit():
            return "break" if event is not None else None

        raw_value = self.qty_multiplier_var.get().strip()
        if not raw_value:
            self._update_status("Voer een totaal aantal in om te vermenigvuldigen.")
            return "break" if event is not None else None

        normalized_value = raw_value.replace(",", ".")
        try:
            factor = Decimal(normalized_value)
        except InvalidOperation:
            self._update_status(f"Ongeldig totaal aantal: {raw_value}.")
            return "break" if event is not None else None

        if factor <= 0:
            self._update_status("Het totaal aantal moet groter dan nul zijn.")
            return "break" if event is not None else None

        before = self.table_model.df.copy(deep=True)
        reference = self._align_qty_reference()
        reference_values = list(reference.tolist())
        changed: List[CellCoord] = []
        errors = 0

        self._suspend_history = True
        try:
            for row_index, base_value in enumerate(reference_values):
                if pd.isna(base_value):
                    normalized_base = ""
                else:
                    normalized_base = str(base_value).strip()
                if not normalized_base:
                    continue

                try:
                    base_decimal = Decimal(normalized_base.replace(",", "."))
                except InvalidOperation:
                    errors += 1
                    continue

                new_value = base_decimal * factor
                formatted = format(new_value, "f")
                if "." in formatted:
                    formatted = formatted.rstrip("0").rstrip(".")
                if not formatted:
                    formatted = "0"

                current = self.table_model.df.iat[row_index, self.QTY_COLUMN_INDEX]
                normalized_current = "" if pd.isna(current) else str(current).strip()
                if normalized_current != formatted:
                    self.table_model.df.iat[row_index, self.QTY_COLUMN_INDEX] = formatted
                    changed.append((row_index, self.QTY_COLUMN_INDEX))
        finally:
            self._suspend_history = False

        after = self.table_model.df.copy(deep=True)
        if changed and not before.equals(after):
            self._push_undo("QTY vermenigvuldigen", before, changed)
            self._dataframe = after
            self._refresh_table()
            self._set_current_multiplier(factor)
            message = f"QTY. vermenigvuldigd met {raw_value}."
            if errors:
                message += f" {errors} rijen met ongeldige QTY.-waarden overgeslagen."
            self._update_status(message)
        else:
            self._set_current_multiplier(factor)
            if errors and not changed:
                self._update_status(
                    "Alle QTY.-waarden waren ongeldig en zijn niet aangepast."
                )
            else:
                self._update_status("Geen QTY.-waarden gewijzigd.")

        return "break" if event is not None else None

    def _collect_selection(self) -> Tuple[List[int], List[int]]:
        table = self.table
        rows = list(dict.fromkeys(table.multiplerowlist)) if table.multiplerowlist else []
        cols = list(dict.fromkeys(table.multiplecollist)) if table.multiplecollist else []
        last_src = getattr(table, "_Table__last_left_click_src", "")
        row_count = len(self.table_model.df)

        if last_src == "row":
            if not rows:
                current = table.currentrow
                if current is not None:
                    rows = [int(current)]
            if not cols:
                cols = list(range(len(self.HEADERS)))
        elif last_src == "column":
            if not cols:
                current = table.currentcol
                if current is not None:
                    cols = [int(current)]
            if not rows:
                rows = list(range(row_count))
        else:
            if not rows:
                current = table.currentrow
                if current is not None:
                    rows = [int(current)]
            if not cols:
                current = table.currentcol
                if current is not None:
                    cols = [int(current)]

        rows = [r for r in rows if 0 <= r < row_count]
        cols = [c for c in cols if 0 <= c < len(self.HEADERS)]
        return rows, cols

    def _on_model_change(
        self, before: pd.DataFrame, after: pd.DataFrame, cell: CellCoord
    ) -> None:
        if self._suspend_history:
            return
        if before.equals(after):
            return
        before_reference = self._capture_qty_reference_snapshot()
        self._push_undo("bewerking", before, [cell], qty_reference=before_reference)
        self._dataframe = after
        target_minimum = max(self.DEFAULT_EMPTY_ROWS, cell[0] + 2)
        self._ensure_minimum_rows(target_minimum)
        row, col = cell
        if col == self.QTY_COLUMN_INDEX:
            value = after.iat[row, col]
            self._update_qty_reference_for_row(row, value)
        self._update_status(f"Cel ({row + 1}, {col + 1}) bijgewerkt.")

    def _clear_selection(self, event=None):
        if not self.table._commit_active_edit():
            return "break"
        rows, cols = self._collect_selection()
        last_src = getattr(self.table, "_Table__last_left_click_src", "")
        if last_src == "row":
            self._delete_rows(rows)
            return "break"
        self._clear_cells(
            rows,
            cols,
            undo_action="leegmaken",
            empty_status="Geen cellen geselecteerd om te legen.",
            success_status="{count} cellen geleegd.",
            no_change_status="Geen wijzigingen bij legen.",
        )
        return "break"

    def _delete_rows(self, rows: Sequence[int]) -> int:
        total_rows = len(self.table_model.df)
        unique_rows = sorted({row for row in rows if 0 <= row < total_rows})
        if not unique_rows:
            self._update_status("Geen rijen geselecteerd om te verwijderen.")
            return 0

        before = self.table_model.df.copy(deep=True)
        before_reference = self._capture_qty_reference_snapshot()
        remaining = before.drop(index=unique_rows).reset_index(drop=True)
        if remaining.equals(before):
            self._update_status("Geen rijen geselecteerd om te verwijderen.")
            return 0

        after_reference = before_reference.drop(index=unique_rows, errors="ignore")
        after_reference = after_reference.reset_index(drop=True)

        self._push_undo(
            "rijen verwijderen", before, [], qty_reference=before_reference
        )
        self._qty_multiplier_reference = after_reference
        self._set_dataframe(
            remaining, reset_multiplier_reference=False, update_multiplier_entry=False
        )
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)

        count = len(unique_rows)
        label = "rij" if count == 1 else "rijen"
        self._update_status(f"{count} {label} verwijderd.")
        return count

    def _delete_selected_rows_from_toolbar(self, event=None):
        if not self.table._commit_active_edit():
            return "break" if event is not None else None

        rows = self._selected_row_indices()
        if not rows:
            self._update_status("Selecteer eerst rijen om te verwijderen.")
            return "break" if event is not None else None

        count = len(rows)
        label = "rij" if count == 1 else "rijen"
        if not messagebox.askyesno(
            "Bevestigen",
            f"{count} {label} verwijderen?",
            parent=self,
        ):
            self._update_status("Verwijderen geannuleerd.")
            return "break" if event is not None else None

        self._delete_rows(rows)
        self._update_find_replace_preview()
        return "break" if event is not None else None

    def _parse_clipboard_text(self, text: str) -> List[List[str]]:
        import io

        delimiter = "\t"
        reader = csv.reader(io.StringIO(text), delimiter=delimiter)
        rows: List[List[str]] = []
        for row in reader:
            if not row:
                rows.append([""])
                continue
            rows.append([cell.strip() for cell in row])
        return rows

    def _on_paste(self, event=None, *, clipboard_text: Optional[str] = None):
        if not self.table._commit_active_edit():
            return "break"
        if clipboard_text is None:
            try:
                raw = self.table.clipboard_get()
            except tk.TclError:
                self._update_status("Klembordinhoud kon niet gelezen worden.")
                return "break"
        else:
            raw = clipboard_text

        parsed = self._parse_clipboard_text(raw)
        while parsed and all(cell.strip() == "" for cell in parsed[-1]):
            parsed.pop()
        if not parsed:
            self._update_status("Geen gegevens gevonden om te plakken.")
            return "break"
        if not any(cell.strip() for row in parsed for cell in row):
            self._update_status("Klembord bevat alleen lege waarden.")
            return "break"

        rows, cols = self._collect_selection()

        def _coerce_index(value: Optional[Any]) -> Optional[int]:
            if value is None:
                return None
            try:
                return int(value)
            except (TypeError, ValueError):
                return None

        active_row = _coerce_index(getattr(self.table, "currentrow", None))
        active_col = _coerce_index(getattr(self.table, "currentcol", None))

        start_row = active_row if active_row is not None else (min(rows) if rows else 0)
        start_col = active_col if active_col is not None else (min(cols) if cols else 0)
        start_row = max(start_row, 0)
        start_col = max(start_col, 0)

        required_rows = max(self.DEFAULT_EMPTY_ROWS, start_row + len(parsed))
        self._ensure_minimum_rows(required_rows)

        before = self.table_model.df.copy(deep=True)
        before_reference = self._capture_qty_reference_snapshot()
        changed: List[CellCoord] = []
        self._suspend_history = True
        try:
            for r_offset, row_values in enumerate(parsed):
                target_row = start_row + r_offset
                if target_row >= len(self.table_model.df):
                    self._append_empty_rows(target_row + 1 - len(self.table_model.df))
                for c_offset, raw_value in enumerate(row_values):
                    target_col = start_col + c_offset
                    if target_col >= len(self.HEADERS):
                        continue
                    normalized = raw_value.strip()
                    current = self.table_model.df.iat[target_row, target_col]
                    normalized_current = "" if pd.isna(current) else str(current)
                    if normalized_current != normalized:
                        self.table_model.df.iat[target_row, target_col] = normalized
                        changed.append((target_row, target_col))
        finally:
            self._suspend_history = False

        after = self.table_model.df.copy(deep=True)
        if changed and not before.equals(after):
            self._push_undo("plakken", before, changed, qty_reference=before_reference)
            self._dataframe = after
            self._refresh_table()
            self._sync_qty_reference_for_cells(changed)
            max_row = max(row for row, _ in changed)
            self._ensure_minimum_rows(max(self.DEFAULT_EMPTY_ROWS, max_row + 2))
            self._update_status(f"{len(changed)} cellen geplakt.")
        else:
            self._update_status("Geen nieuwe waarden geplakt.")
        return "break"

    # ------------------------------------------------------------------
    # Template
    def _download_template(self) -> None:
        """Vraag een doelpad en schrijf een leeg Excel-sjabloon."""

        path_str = filedialog.asksaveasfilename(
            parent=self,
            title="BOM-template opslaan",
            defaultextension=".xlsx",
            filetypes=(("Excel-werkboek", "*.xlsx"), ("Alle bestanden", "*.*")),
            initialfile=self.default_template_filename(),
        )
        if not path_str:
            self._update_status("Download geannuleerd.")
            return

        target_path = Path(path_str)
        try:
            self.write_template_workbook(target_path)
        except Exception as exc:
            messagebox.showerror("Opslaan mislukt", str(exc), parent=self)
            self._update_status("Fout bij opslaan van template.")
            return

        messagebox.showinfo(
            "Template opgeslagen",
            f"Leeg BOM-sjabloon opgeslagen als:\n{target_path}",
            parent=self,
        )
        self._update_status(f"Template opgeslagen naar {target_path}.")


    def _confirm_clear(self) -> None:
        before_df = self.table_model.df.copy(deep=True)
        before_reference = self._capture_qty_reference_snapshot()
        trimmed = before_df.replace("", pd.NA).dropna(how="all")
        if trimmed.empty:
            self._update_status("Sheet was al leeg.")
            return
        if not messagebox.askyesno("Bevestigen", "Alle custom BOM-data verwijderen?", parent=self):
            return
        self._push_undo("clear", before_df, [], qty_reference=before_reference)
        cleared = self._create_empty_dataframe(self.DEFAULT_EMPTY_ROWS)
        self._set_dataframe(cleared)
        self._update_status("Custom BOM geleegd.")

    def _reset_to_baseline(self) -> None:
        """Herstel de huidige tabel naar de laatst opgeslagen beginsituatie."""

        baseline_df = getattr(self, "_baseline_dataframe", None)
        baseline_reference = getattr(self, "_baseline_qty_reference", None)
        baseline_multiplier = getattr(self, "_baseline_multiplier", Decimal(1))

        if not isinstance(baseline_df, pd.DataFrame):
            self._update_status("Geen oorspronkelijke waarden beschikbaar om te herstellen.")
            return

        current_df = self.table_model.df.copy(deep=True)
        current_reference = self._capture_qty_reference_snapshot()

        same_df = current_df.equals(baseline_df)
        same_multiplier = self._current_qty_multiplier == baseline_multiplier
        if isinstance(baseline_reference, pd.Series):
            baseline_series = baseline_reference
            same_reference = current_reference.equals(baseline_series)
        else:
            baseline_series = None
            same_reference = False

        if same_df and same_multiplier and same_reference:
            self._update_status("Custom BOM komt al overeen met oorspronkelijke waarden.")
            return

        self._push_undo(
            "reset",
            current_df,
            [],
            qty_reference=current_reference,
            qty_multiplier=self._current_qty_multiplier,
        )

        entry = UndoEntry(
            action="reset",
            frame=baseline_df.copy(deep=True),
            cells=[],
            qty_reference=baseline_series.copy(deep=True) if baseline_series is not None else None,
            qty_multiplier=baseline_multiplier,
        )
        self._restore_history_entry(entry)
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._update_status("Custom BOM teruggezet naar oorspronkelijke waarden.")

    # ------------------------------------------------------------------
    # Undo
    def _handle_toolbar_undo(self) -> None:
        self._on_undo()

    def _handle_toolbar_redo(self) -> None:
        self._on_redo()

    def _on_undo(self, event=None):
        if not hasattr(self, "undo_stack"):
            self.undo_stack = []
        if not hasattr(self, "redo_stack"):
            self.redo_stack = []
        if not self.undo_stack:
            self._update_status("Niets om ongedaan te maken.")
            return "break"
        current_snapshot = self.table_model.df.copy(deep=True)
        current_reference = self._capture_qty_reference_snapshot()
        current_multiplier = self._current_qty_multiplier
        entry = self.undo_stack.pop()
        self.redo_stack.append(
            UndoEntry(
                action=entry.action,
                frame=current_snapshot,
                cells=[],
                qty_reference=current_reference,
                qty_multiplier=current_multiplier,
            )
        )
        if len(self.redo_stack) > self.max_undo:
            self.redo_stack.pop(0)
        self._restore_history_entry(entry)
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._update_status(f"{entry.action.capitalize()} ongedaan gemaakt.")
        return "break"

    def _on_redo(self, event=None):
        if not hasattr(self, "redo_stack"):
            self.redo_stack = []
        if not hasattr(self, "undo_stack"):
            self.undo_stack = []
        if not self.redo_stack:
            self._update_status("Niets om te herhalen.")
            return "break"

        current_snapshot = self.table_model.df.copy(deep=True)
        current_reference = self._capture_qty_reference_snapshot()
        current_multiplier = self._current_qty_multiplier
        entry = self.redo_stack.pop()
        self.undo_stack.append(
            UndoEntry(
                action=entry.action,
                frame=current_snapshot,
                cells=[],
                qty_reference=current_reference,
                qty_multiplier=current_multiplier,
            )
        )
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)
        self._restore_history_entry(entry)
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._update_status(f"{entry.action.capitalize()} opnieuw toegepast.")
        return "break"

    # ------------------------------------------------------------------
    # Export
    def _push_to_main(self) -> None:
        if not self._commit_pending_table_edit():
            self._update_status("Werk de actieve cel af voordat je de Main-BOM bijwerkt.")
            return

        if self.on_push_to_main is None:
            messagebox.showinfo(
                "Niet beschikbaar",
                (
                    "Deze knop is alleen actief wanneer de Custom BOM-tab "
                    "aan de hoofdinterface is gekoppeld."
                ),
                parent=self,
            )
            return

        main_df = self.build_main_dataframe()
        if main_df.empty:
            messagebox.showwarning(
                "Geen gegevens",
                "Er zijn geen rijen met gegevens om naar de Main-tab te sturen.",
                parent=self,
            )
            self._update_status("Geen gegevens om naar Main te sturen.")
            return

        try:
            self.on_push_to_main(main_df)
        except Exception as exc:
            messagebox.showerror("Bijwerken mislukt", str(exc), parent=self)
            self._update_status("Fout bij bijwerken van de hoofd-BOM.")
        else:
            self._update_status(f"Main-BOM geüpdatet met {len(main_df)} rijen.")

    def _export_temp(self) -> None:
        if not self._commit_pending_table_edit():
            self._update_status("Werk de actieve cel af voordat je exporteert.")
            return

        data = self._snapshot_data()
        cleaned = [row[: len(self.HEADERS)] for row in data]
        non_empty_rows = [row for row in cleaned if any(cell.strip() for cell in row)]
        row_count = len(non_empty_rows)
        if row_count == 0:
            self._update_status("Geen gegevens om te exporteren.")
            return

        path = self._resolve_default_export_path()
        self._write_csv(path, non_empty_rows)

        if self.on_custom_bom_ready is not None:
            self.on_custom_bom_ready(path, row_count)
            self._update_status(f"Custom BOM geëxporteerd naar {path}.")
            return

        self.last_temp_csv_path = path
        target = self.event_target or self.winfo_toplevel()
        if target is not None:
            target.event_generate("<<CustomBOMReady>>", when="tail", data=str(path))
        self._update_status(f"Tijdelijke CSV klaar ({row_count} rijen).")

    def _resolve_default_export_path(self) -> Path:
        base = os.environ.get("LOCALAPPDATA")
        if base:
            root = Path(base)
        else:
            root = Path.home() / ".local" / "share"
        folder = root / self.app_name / "temp"
        folder.mkdir(parents=True, exist_ok=True)
        return folder / "custom_bom.csv"

    def _write_csv(self, path: Path, rows: Sequence[Sequence[str]]) -> None:
        with path.open("w", encoding="utf-8", newline="") as fh:
            writer = csv.writer(fh, delimiter=",", quoting=csv.QUOTE_MINIMAL)
            writer.writerow(self.HEADERS)
            for row in rows:
                cells = [cell.strip() for cell in row]
                while len(cells) < len(self.HEADERS):
                    cells.append("")
                writer.writerow(cells[: len(self.HEADERS)])

    @classmethod
    def write_template_workbook(cls, target: Path) -> None:
        """Schrijf een leeg Excel-sjabloon met de standaardkolommen."""

        normalized = Path(target)
        if not normalized.parent.exists():
            normalized.parent.mkdir(parents=True, exist_ok=True)
        df = pd.DataFrame(columns=cls.HEADERS)
        df.to_excel(normalized, index=False)

    @classmethod
    def default_template_filename(cls) -> str:
        """Geef de standaard bestandsnaam voor het templatesjabloon terug."""

        return cls.TEMPLATE_DEFAULT_FILENAME

    # ------------------------------------------------------------------
    # API
    def get_last_export_path(self) -> Optional[Path]:
        """Geef het pad van de laatst geëxporteerde CSV terug."""

        return self.last_temp_csv_path

    def clear_history(self) -> None:
        """Verwijder alle undo-stappen."""

        if hasattr(self, "undo_stack"):
            self.undo_stack.clear()
        if hasattr(self, "redo_stack"):
            self.redo_stack.clear()

    def load_from_main_dataframe(self, df: pd.DataFrame) -> None:
        """Vul de Custom BOM met gegevens uit de hoofd-BOM."""

        if df is None:
            return
        if not isinstance(df, pd.DataFrame):
            raise TypeError("df must be a pandas.DataFrame")

        row_count = len(df.index)
        target_rows = max(self.DEFAULT_EMPTY_ROWS, row_count + 1)
        fresh = self._create_empty_dataframe(target_rows)
        if row_count:
            normalized = df.fillna("")
            for main_col, custom_col in self.MAIN_TO_CUSTOM_COLUMN_MAP.items():
                if main_col not in normalized.columns:
                    continue
                values = [str(value).strip() for value in normalized[main_col]]
                col_index = fresh.columns.get_loc(custom_col)
                fresh.iloc[: len(values), col_index] = values

        self._set_dataframe(fresh)
        self._store_baseline_state()
        self.clear_history()
        self._update_status("Custom BOM gevuld vanuit hoofd-BOM.")

    def build_main_dataframe(self) -> pd.DataFrame:
        """Zet de huidige Custom BOM om naar het hoofd-BOM formaat."""

        column_order = list(self.MAIN_COLUMN_ORDER)
        empty = pd.DataFrame(columns=column_order)

        snapshot = pd.DataFrame(self._snapshot_data(), columns=self.HEADERS)
        trimmed = snapshot.replace("", pd.NA).dropna(how="all")
        if trimmed.empty:
            return empty

        trimmed = trimmed.fillna("").map(lambda value: str(value).strip())
        result = pd.DataFrame(index=trimmed.index)
        for custom_col, main_col in self.CUSTOM_TO_MAIN_COLUMN_MAP.items():
            if custom_col in trimmed.columns:
                series = trimmed[custom_col]
            else:
                series = pd.Series([""] * len(trimmed), index=trimmed.index)
            result[main_col] = series.astype(str).map(lambda value: value.strip())

        if "PartNumber" not in result.columns:
            return empty

        result = result[result["PartNumber"].str.strip() != ""]
        if result.empty:
            return empty

        if "Aantal" in result.columns:
            qty = pd.to_numeric(result["Aantal"], errors="coerce").fillna(1).astype(int)
            result["Aantal"] = qty.clip(lower=1, upper=999)

        result = result.reindex(columns=column_order, fill_value="")
        result.reset_index(drop=True, inplace=True)
        return result
