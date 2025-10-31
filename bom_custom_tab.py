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
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

try:
    from pandastable import Table, TableModel
except ModuleNotFoundError as exc:  # pragma: no cover - afhankelijk van installatie
    class _TableStub:
        def __init__(self, *args, **kwargs) -> None:
            raise RuntimeError(_PANDASTABLE_ERROR) from exc

    class _TableModelStub:
        def __init__(self, *args, **kwargs) -> None:
            raise RuntimeError(_PANDASTABLE_ERROR) from exc

    Table = _TableStub  # type: ignore[assignment]
    TableModel = _TableModelStub  # type: ignore[assignment]
    _PANDASTABLE_IMPORT_ERROR: Optional[BaseException] = exc
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
    qty_multiplier_factor: Decimal = Decimal(1)
    qty_multiplier_value: str = ""


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
        self.last_temp_csv_path: Optional[Path] = None
        self._suspend_history = False

        self.status_var = tk.StringVar(value="")
        self.qty_multiplier_var = tk.StringVar(value="")
        self._qty_multiplier_last_factor: Decimal = Decimal(1)
        self._reset_qty_multiplier_state(clear_entry=True)

        self._build_toolbar()
        self._build_sheet()
        self._update_status("Gereed.")

    # ------------------------------------------------------------------
    # UI-opbouw
    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=8, pady=6)

        clear_btn = ttk.Button(bar, text="Clear Custom BOM", command=self._confirm_clear)
        clear_btn.pack(side="left", padx=(0, 6))

        export_btn = ttk.Button(bar, text="Exporteren", command=self._export_temp)
        export_btn.pack(side="left", padx=(0, 6))

        update_btn = ttk.Button(bar, text="Update Main BOM", command=self._push_to_main)
        update_btn.pack(side="left", padx=(0, 6))
        self._update_main_btn = update_btn

        template_btn = ttk.Button(bar, text="Download template", command=self._download_template)
        template_btn.pack(side="left", padx=(0, 6))

        ttk.Label(bar, textvariable=self.status_var, anchor="w").pack(
            side="left", fill="x", expand=True
        )

        multiplier_container = ttk.Frame(bar)
        multiplier_container.pack(side="right")

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
        for sequence in ("<Delete>", "<BackSpace>"):
            self.table.bind(sequence, self._clear_selection, add="+")


    # ------------------------------------------------------------------
    # Helpers
    def _update_status(self, text: str) -> None:
        self.status_var.set(text)

    def _reset_qty_multiplier_state(self, *, clear_entry: bool = False) -> None:
        self._qty_multiplier_last_factor = Decimal(1)
        if clear_entry:
            self.qty_multiplier_var.set("")

    def _create_empty_dataframe(self, rows: int) -> pd.DataFrame:
        data = {header: [""] * rows for header in self.HEADERS}
        return pd.DataFrame(data, columns=self.HEADERS)

    def _set_dataframe(self, df: pd.DataFrame) -> None:
        normalized = df.reindex(columns=self.HEADERS, fill_value="")
        self._suspend_history = True
        try:
            self.table_model.df = normalized
            self._refresh_table()
        finally:
            self._suspend_history = False
        self._dataframe = self.table_model.df.copy(deep=True)

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

    def _ensure_minimum_rows(self, minimum: int = 1) -> None:
        if minimum <= 0:
            return
        current_rows = len(self.table_model.df)
        if current_rows >= minimum:
            return
        missing = minimum - current_rows
        extension = self._create_empty_dataframe(missing)
        extended = pd.concat([self.table_model.df, extension], ignore_index=True)
        self._set_dataframe(extended)

    def _snapshot_data(self, frame: Optional[pd.DataFrame] = None) -> List[List[str]]:
        df = frame if frame is not None else self.table_model.df
        normalized = df.reindex(columns=self.HEADERS)
        return normalized.fillna("").astype(str).values.tolist()

    def _push_undo(self, action: str, frame: pd.DataFrame, cells: Sequence[CellCoord]) -> None:
        snapshot = frame.copy(deep=True)
        self.undo_stack.append(
            UndoEntry(
                action=action,
                frame=snapshot,
                cells=list(cells),
                qty_multiplier_factor=self._qty_multiplier_last_factor,
                qty_multiplier_value=self.qty_multiplier_var.get(),
            )
        )
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)

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
            self._push_undo(undo_action, before, changed)
            self._dataframe = after
            self._refresh_table()
            self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
            self._update_status(success_status.format(count=len(changed)))
            return len(changed)

        self._update_status(no_change_status)
        return 0

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
        changed: List[CellCoord] = []
        errors = 0
        previous_factor = self._qty_multiplier_last_factor
        if previous_factor <= 0:
            previous_factor = Decimal(1)

        self._suspend_history = True
        try:
            for row in range(len(self.table_model.df)):
                current = self.table_model.df.iat[row, self.QTY_COLUMN_INDEX]
                normalized_current = "" if pd.isna(current) else str(current).strip()
                if not normalized_current:
                    continue

                try:
                    qty_value = Decimal(normalized_current.replace(",", "."))
                except InvalidOperation:
                    errors += 1
                    continue

                base_value = qty_value
                if previous_factor != Decimal(1):
                    try:
                        base_value = qty_value / previous_factor
                    except InvalidOperation:
                        base_value = qty_value

                new_value = base_value * factor
                formatted = format(new_value, "f").rstrip("0").rstrip(".")
                if not formatted:
                    formatted = "0"

                if formatted != normalized_current:
                    self.table_model.df.iat[row, self.QTY_COLUMN_INDEX] = formatted
                    changed.append((row, self.QTY_COLUMN_INDEX))
        finally:
            self._suspend_history = False

        after = self.table_model.df.copy(deep=True)
        if changed and not before.equals(after):
            self._push_undo("QTY vermenigvuldigen", before, changed)
            self._dataframe = after
            self._refresh_table()
            message = f"QTY. vermenigvuldigd met {raw_value}."
            if errors:
                message += f" {errors} rijen met ongeldige QTY.-waarden overgeslagen."
            self._update_status(message)
        else:
            if errors and not changed:
                self._update_status(
                    "Alle QTY.-waarden waren ongeldig en zijn niet aangepast."
                )
            else:
                self._update_status("Geen QTY.-waarden gewijzigd.")

        self._qty_multiplier_last_factor = factor

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
        self._push_undo("bewerking", before, [cell])
        self._dataframe = after
        target_minimum = max(self.DEFAULT_EMPTY_ROWS, cell[0] + 2)
        self._ensure_minimum_rows(target_minimum)
        row, col = cell
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
        remaining = before.drop(index=unique_rows).reset_index(drop=True)
        if remaining.equals(before):
            self._update_status("Geen rijen geselecteerd om te verwijderen.")
            return 0

        self._push_undo("rijen verwijderen", before, [])
        self._set_dataframe(remaining)
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)

        count = len(unique_rows)
        label = "rij" if count == 1 else "rijen"
        self._update_status(f"{count} {label} verwijderd.")
        return count

    def _parse_clipboard_text(self, text: str) -> List[List[str]]:
        import io

        delimiter = "\t"
        if "\t" not in text:
            if ";" in text:
                delimiter = ";"
            elif "," in text:
                delimiter = ","
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
            self._push_undo("plakken", before, changed)
            self._dataframe = after
            self._refresh_table()
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
        trimmed = before_df.replace("", pd.NA).dropna(how="all")
        if trimmed.empty:
            self._update_status("Sheet was al leeg.")
            return
        if not messagebox.askyesno("Bevestigen", "Alle custom BOM-data verwijderen?", parent=self):
            return
        self._push_undo("clear", before_df, [])
        cleared = self._create_empty_dataframe(self.DEFAULT_EMPTY_ROWS)
        self._set_dataframe(cleared)
        self._reset_qty_multiplier_state(clear_entry=True)
        self._update_status("Custom BOM geleegd.")

    # ------------------------------------------------------------------
    # Undo
    def _on_undo(self, event=None):
        if not self.undo_stack:
            self._update_status("Niets om ongedaan te maken.")
            return "break"
        entry = self.undo_stack.pop()
        self._set_dataframe(entry.frame)
        self._qty_multiplier_last_factor = entry.qty_multiplier_factor
        self.qty_multiplier_var.set(entry.qty_multiplier_value)
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._update_status(f"{entry.action.capitalize()} ongedaan gemaakt.")
        return "break"

    # ------------------------------------------------------------------
    # Export
    def _push_to_main(self) -> None:
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

        self.undo_stack.clear()

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
        self._reset_qty_multiplier_state(clear_entry=True)
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

        trimmed = trimmed.fillna("").applymap(lambda value: str(value).strip())
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
