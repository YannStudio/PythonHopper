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
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple

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

        # ``pandastable.Table`` implementeert ``__getattr__`` en lijkt daarmee
        # sommige "_on_*"-attributen te verbergen wanneer ze op de instantie
        # worden opgevraagd tijdens ``__init__``. Gebruik daarom een lambda die
        # het klasse-attribuut aanspreekt zodat de handler beschikbaar blijft. 
        self.bind(
            "<KeyPress>",
            lambda event, table=self: type(table)._on_table_keypress(table, event),
            add="+",
        )
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

    def _ensure_entry_bindings(self, entry: tk.Widget) -> None:
        if getattr(entry, "_undoaware_bindings", False):  # pragma: no cover - Tk internals
            return

        entry.bind("<FocusOut>", self._on_entry_focus_out, add="+")
        entry.bind("<Return>", self._on_entry_return)
        entry.bind("<KP_Enter>", self._on_entry_return)
        entry.bind("<Tab>", self._on_entry_tab)
        entry.bind("<ISO_Left_Tab>", self._on_entry_shift_tab)
        entry.bind("<Shift-Tab>", self._on_entry_shift_tab)
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
            entry.bind(sequence, self._on_entry_clipboard_paste, add="+")
        setattr(entry, "_undoaware_bindings", True)

    def handle_left_click(self, event):  # type: ignore[override]
        if self._active_edit is not None:
            committed = self._commit_active_edit()
            if not committed:
                return

        return super().handle_left_click(event)

    def handleCellEntry(self, row, col):  # type: ignore[override]
        self._skip_focus_commit = True
        try:
            return super().handleCellEntry(row, col)
        finally:
            self.after_idle(self._reset_focus_commit_guard)
            self._active_edit = None

    def _reset_focus_commit_guard(self) -> None:
        self._skip_focus_commit = False

    def _on_entry_focus_out(self, event: tk.Event) -> None:
        if self._skip_focus_commit:
            return
        self._commit_active_edit(trigger_widget=event.widget)

    def _on_entry_clipboard_paste(self, event: tk.Event) -> Optional[str]:
        try:
            text = event.widget.clipboard_get()
        except tk.TclError:
            try:
                text = self.clipboard_get()
            except tk.TclError:
                return None

        parsed = self._owner._parse_clipboard_text(text)
        while parsed and all(cell.strip() == "" for cell in parsed[-1]):
            parsed.pop()

        if not parsed:
            return None

        if len(parsed) == 1 and len(parsed[0]) == 1:
            return None

        if not self._commit_active_edit(trigger_widget=event.widget):
            return "break"

        try:
            self.focus_set()
        except Exception:  # pragma: no cover - focus issues only in GUI
            pass
        return self._owner._on_paste(None, clipboard_text=text)

    def cut(self, event=None):  # type: ignore[override]
        self._on_copy_shortcut(event)
        self._owner._clear_selection(event)
        return "break"

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

    def _on_table_keypress(self, event: tk.Event) -> Optional[str]:
        if event.keysym in {"Return", "KP_Enter"}:
            if self._active_edit is None:
                if self._begin_edit(select_all=True):
                    return "break"
                return None
            if not self._commit_active_edit():
                return "break"
            self._move_vertical(1)
            return "break"

        if not event.char:
            return None
        if not self._should_start_direct_edit(event):
            return None
        if self._start_edit_with_char(event.char):
            return "break"
        return None

    def _should_start_direct_edit(self, event: tk.Event) -> bool:
        if not event.char or not event.char.isprintable():
            return False
        state = event.state or 0
        # Control (0x4) and Command/Meta modifiers (0x10 and 0x20000) should not trigger typing
        modifier_mask = 0x4 | 0x10 | 0x20000
        if state & modifier_mask:
            return False
        return event.keysym not in {"BackSpace", "Delete"}

    def _start_edit_with_char(self, char: str) -> bool:
        if not self._commit_active_edit():
            return True
        entry = self._begin_edit(initial_text=char)
        return entry is not None

    def _begin_edit(
        self,
        *,
        initial_text: Optional[str] = None,
        select_all: bool = False,
    ) -> Optional[tk.Widget]:
        row = self.currentrow
        col = self.currentcol
        if row is None or col is None:
            return None
        self.drawCellEntry(int(row), int(col))
        entry = getattr(self, "cellentry", None)
        if entry is None:
            return None
        var = getattr(self, "cellentryvar", None)
        if initial_text is not None and var is not None:
            var.set(initial_text)
            try:
                entry.icursor("end")
            except Exception:
                pass
        elif select_all:
            try:
                entry.selection_range(0, "end")
            except Exception:
                pass
        try:
            entry.focus_set()
        except Exception:
            pass
        return entry

    def _on_entry_return(self, event: tk.Event) -> str:
        if not self._commit_active_edit(trigger_widget=event.widget):
            return "break"
        self._move_vertical(1)
        return "break"

    def _on_entry_tab(self, event: tk.Event) -> str:
        if not self._commit_active_edit(trigger_widget=event.widget):
            return "break"
        self._move_horizontal(1)
        return "break"

    def _on_entry_shift_tab(self, event: tk.Event) -> str:
        if not self._commit_active_edit(trigger_widget=event.widget):
            return "break"
        self._move_horizontal(-1)
        return "break"

    def _move_horizontal(self, delta: int) -> None:
        if self.rows <= 0 or self.cols <= 0:
            return
        row = int(self.currentrow or 0)
        col = int(self.currentcol or 0) + delta
        if col >= self.cols:
            col = 0
            if row < self.rows - 1:
                row += 1
        elif col < 0:
            col = self.cols - 1
            if row > 0:
                row -= 1
        self._select_cell(row, col)

    def _move_vertical(self, delta: int) -> None:
        if self.rows <= 0:
            return
        row = int(self.currentrow or 0) + delta
        row = min(max(row, 0), self.rows - 1)
        col = int(self.currentcol or 0)
        self._select_cell(row, col)

    def _select_cell(self, row: int, col: int) -> None:
        row = min(max(int(row), 0), max(self.rows - 1, 0))
        col = min(max(int(col), 0), max(self.cols - 1, 0))
        self.setSelectedRow(row)
        self.setSelectedCol(col)
        self.drawSelectedRect(row, col)
        self.drawSelectedRow()
        try:
            self.rowheader.drawSelectedRows(row)
        except Exception:  # pragma: no cover - Tk internals
            pass
        try:
            self.colheader.delete("rect")
        except Exception:  # pragma: no cover - Tk internals
            pass
        try:
            self.focus_set()
        except Exception:  # pragma: no cover - focus issues only in GUI
            pass

    def _on_copy_shortcut(self, event=None) -> str:
        if not self._commit_active_edit():
            return "break"
        rows = list(dict.fromkeys(self.multiplerowlist)) if self.multiplerowlist else []
        cols = list(dict.fromkeys(self.multiplecollist)) if self.multiplecollist else []
        if not rows:
            if self.currentrow is not None:
                rows = [int(self.currentrow)]
        if not cols:
            if self.currentcol is not None:
                cols = [int(self.currentcol)]
        if rows and cols:
            super().copy(rows, cols)
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
    TEMPLATE_DEFAULT_FILENAME: str = "BOM-FileHopper-Temp.xlsx"
    DEFAULT_EMPTY_ROWS: int = 20

    def __init__(
        self,
        master: tk.Widget,
        *,
        app_name: str = "Filehopper",
        on_custom_bom_ready: Optional[Callable[[Path, int], None]] = None,
        event_target: Optional[tk.Misc] = None,
        max_undo: int = 50,
    ) -> None:
        _ensure_pandastable_available()

        super().__init__(master)
        self.configure(padding=(12, 12))
        self.app_name = app_name
        self.on_custom_bom_ready = on_custom_bom_ready
        self.event_target = event_target
        self.max_undo = max_undo
        self.undo_stack: List[UndoEntry] = []
        self.last_temp_csv_path: Optional[Path] = None
        self._suspend_history = False

        self.status_var = tk.StringVar(value="")

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

        template_btn = ttk.Button(bar, text="Download template", command=self._download_template)
        template_btn.pack(side="left", padx=(0, 6))

        ttk.Label(bar, textvariable=self.status_var, anchor="w").pack(
            side="left", fill="x", expand=True
        )

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
        self.undo_stack.append(UndoEntry(action=action, frame=snapshot, cells=list(cells)))
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)

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
        if not rows or not cols:
            self._update_status("Geen cellen geselecteerd om te legen.")
            return "break"

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
            self._push_undo("leegmaken", before, changed)
            self._dataframe = after
            self._refresh_table()
            self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
            self._update_status(f"{len(changed)} cellen geleegd.")
        else:
            self._update_status("Geen wijzigingen bij legen.")
        return "break"

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
        start_row = min(rows) if rows else int(self.table.currentrow or 0)
        start_col = min(cols) if cols else int(self.table.currentcol or 0)
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
        self._update_status("Custom BOM geleegd.")

    # ------------------------------------------------------------------
    # Undo
    def _on_undo(self, event=None):
        if not self.undo_stack:
            self._update_status("Niets om ongedaan te maken.")
            return "break"
        entry = self.undo_stack.pop()
        self._set_dataframe(entry.frame)
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._update_status(f"{entry.action.capitalize()} ongedaan gemaakt.")
        return "break"

    # ------------------------------------------------------------------
    # Export
    def _export_temp(self) -> None:
        data = self._snapshot_data()
        cleaned = [row[: len(self.HEADERS)] for row in data]
        non_empty_rows = [row for row in cleaned if any(cell.strip() for cell in row)]
        row_count = len(non_empty_rows)
        if row_count == 0:
            self._update_status("Geen gegevens om te exporteren.")
            return

        path = self._resolve_default_export_path()
        self._write_csv(path, cleaned)

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
