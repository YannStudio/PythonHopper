"""Tkinter-tab voor het bewerken en exporteren van een custom BOM.

Installatie en uitvoeren
========================
1. Activeer de gewenste virtualenv.
2. Installeer de GUI-afhankelijkheden met ``pip install -r requirements.txt``.
3. ``tksheet`` is vereist voor het spreadsheet-grid. Installeer handmatig met
   ``pip install tksheet`` wanneer deze nog niet beschikbaar is.
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
het volledige sheet vóór en na de wijziging en aanvullende metadata
(bijvoorbeeld gewijzigde cellen). De undo-stack bevat maximaal 50 stappen en
wordt door ``Ctrl+Z`` in omgekeerde volgorde verwerkt.

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
terug (of ``None`` wanneer er nog niet geëxporteerd is).
"""

from __future__ import annotations

import csv
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import pandas as pd

try:
    import tksheet
    _TKSHEET_IMPORT_ERROR: Optional[BaseException] = None
except ModuleNotFoundError as exc:  # pragma: no cover - afhankelijk van installatie
    tksheet = None  # type: ignore[assignment]
    _TKSHEET_IMPORT_ERROR = exc

_TKSHEET_ERROR = (
    "De module 'tksheet' is niet geïnstalleerd. "
    "Voer 'pip install tksheet' uit voordat u de Filehopper GUI start."
)


def _ensure_tksheet_available() -> None:
    if tksheet is not None:
        return

    try:
        import tkinter as _tk
        from tkinter import messagebox as _messagebox

        _root = _tk.Tk()
        _root.withdraw()
        try:
            _messagebox.showerror("tksheet ontbreekt", _TKSHEET_ERROR)
        finally:
            _root.destroy()
    except Exception:
        # Val stilletjes terug op een consolefout wanneer Tkinter niet beschikbaar is
        pass
    raise RuntimeError(_TKSHEET_ERROR) from _TKSHEET_IMPORT_ERROR


CellCoord = Tuple[int, int]


@dataclass
class UndoEntry:
    """Snapshot van een wijziging."""

    action: str
    before: List[List[str]]
    after: List[List[str]]
    cells: Sequence[CellCoord]


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
    COLUMN_PADDING: int = 24
    TRAILING_GUTTER: int = 12

    def __init__(
        self,
        master: tk.Widget,
        *,
        app_name: str = "Filehopper",
        on_custom_bom_ready: Optional[Callable[[Path, int], None]] = None,
        event_target: Optional[tk.Misc] = None,
        max_undo: int = 50,
    ) -> None:
        _ensure_tksheet_available()

        super().__init__(master)
        self.configure(padding=(12, 12))
        self.app_name = app_name
        self.on_custom_bom_ready = on_custom_bom_ready
        self.event_target = event_target
        self.max_undo = max_undo
        self.undo_stack: List[UndoEntry] = []
        self._edit_snapshot: Optional[List[List[str]]] = None
        self._edit_cell: Optional[CellCoord] = None
        self.last_temp_csv_path: Optional[Path] = None
        self._sheet_container: Optional[tk.Widget] = None
        self._in_container_resize = False
        self._single_click_after_id: Optional[str] = None
        self._single_click_delay_ms = self._determine_single_click_delay()

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

        ttk.Label(bar, textvariable=self.status_var, anchor="w").pack(side="left", fill="x", expand=True)

    def _build_sheet(self) -> None:
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        self.sheet = tksheet.Sheet(
            container,
            headers=self.HEADERS,
            show_index=False,
        )
        self.sheet.enable_bindings(
            (
                "single_select",
                "drag_select",
                "select_all",
                "column_select",
                "column_width_resize",
                "double_click_column_resize",
                "double_click_select",
                "copy",
                "edit_cell",
                "tab_key",
                "arrowkeys",
                "enter_key",
            )
        )
        self.sheet.set_sheet_data([])
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self.sheet.grid(row=0, column=0, sticky="nsew")
        self._sheet_container = container
        container.bind("<Configure>", self._on_container_resize)

        self.sheet.bind("<Control-v>", self._on_paste)
        self.sheet.bind("<Control-V>", self._on_paste)
        self.sheet.bind("<Command-v>", self._on_paste)
        self.sheet.bind("<Command-V>", self._on_paste)
        self.sheet.bind("<Control-c>", self._on_copy)
        self.sheet.bind("<Control-C>", self._on_copy)
        self.sheet.bind("<Command-c>", self._on_copy)
        self.sheet.bind("<Command-C>", self._on_copy)
        self.sheet.bind("<Control-x>", self._on_cut)
        self.sheet.bind("<Control-X>", self._on_cut)
        self.sheet.bind("<Command-x>", self._on_cut)
        self.sheet.bind("<Command-X>", self._on_cut)
        self.sheet.bind("<Delete>", self._on_delete)
        self.sheet.bind("<Control-z>", self._on_undo)
        self.sheet.bind("<Control-Z>", self._on_undo)
        self.sheet.MT.bind("<ButtonPress-1>", self._on_single_click_press, add="+")
        self.sheet.MT.bind("<ButtonRelease-1>", self._on_single_click_release, add="+")
        self.sheet.MT.bind("<Double-Button-1>", self._on_single_click_cancel, add="+")

        self.sheet.extra_bindings("begin_edit_cell", self._on_begin_edit_cell)
        self.sheet.extra_bindings("end_edit_cell", self._on_end_edit_cell)

        self._auto_resize_columns(range(len(self.HEADERS)))
        self._apply_row_striping()

    # ------------------------------------------------------------------
    # Helpers
    def _update_status(self, text: str) -> None:
        self.status_var.set(text)

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

    def _snapshot_data(self) -> List[List[str]]:
        data = self.sheet.get_sheet_data()
        return [list(map(self._coerce_to_str, row[: len(self.HEADERS)])) for row in data]

    @staticmethod
    def _coerce_to_str(value: object) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        return str(value)

    def _apply_row_striping(self) -> None:
        total_rows = self.sheet.get_total_rows()
        self.sheet.dehighlight_rows(rows="all")
        if total_rows <= 0:
            self.sheet.refresh()
            return

        even_rows = list(range(0, total_rows, 2))
        odd_rows = list(range(1, total_rows, 2))

        if even_rows:
            self.sheet.highlight_rows(
                rows=even_rows,
                bg="#ffffff",
                fg=False,
                highlight_index=False,
                redraw=False,
                overwrite=True,
            )
        if odd_rows:
            self.sheet.highlight_rows(
                rows=odd_rows,
                bg="#f7f7f7",
                fg=False,
                highlight_index=False,
                redraw=False,
                overwrite=True,
            )
        self.sheet.refresh()

    def _calculate_column_min_widths(self, columns: Iterable[int]) -> Dict[int, int]:
        if not hasattr(self, "sheet"):
            return {}
        valid_columns = sorted({col for col in columns if 0 <= col < len(self.HEADERS)})
        if not valid_columns:
            return {}

        try:
            header_font = tkfont.Font(self.sheet, font=self.sheet.header_font)
        except (tk.TclError, AttributeError):
            header_font = tkfont.nametofont("TkDefaultFont")
        try:
            table_font = tkfont.Font(self.sheet, font=self.sheet.table_font)
        except (tk.TclError, AttributeError):
            table_font = tkfont.nametofont("TkDefaultFont")

        min_width = 60
        total_rows = self.sheet.get_total_rows()

        widths: Dict[int, int] = {}

        for col in valid_columns:
            max_width = header_font.measure(self.HEADERS[col])
            for row in range(total_rows):
                cell_value = self._coerce_to_str(self.sheet.get_cell_data(row, col))
                if not cell_value:
                    continue
                cell_width = table_font.measure(cell_value)
                if cell_width > max_width:
                    max_width = cell_width
            target_width = max(min_width, max_width + self.COLUMN_PADDING)
            widths[col] = target_width

        return widths

    def _auto_resize_columns(self, columns: Iterable[int]) -> None:
        widths = self._calculate_column_min_widths(columns)
        if not widths:
            return

        for col in sorted(widths):
            self.sheet.column_width(column=col, width=widths[col], redraw=False)
        self.sheet.refresh()
        self._on_container_resize()

    def _on_container_resize(self, event=None) -> None:
        if self._in_container_resize:
            return
        container = self._sheet_container
        if container is None:
            return
        if not hasattr(self, "sheet"):
            return
        try:
            container_width = 0
            if event is not None and getattr(event, "width", None):
                container_width = int(event.width)
            if container_width <= 0:
                container_width = int(container.winfo_width())
        except (tk.TclError, ValueError, TypeError):
            return
        if container_width <= 1:
            return

        scrollbar_width = 0
        yscroll = getattr(self.sheet, "yscroll", None)
        if yscroll is not None:
            showing = False
            try:
                showing = bool(yscroll.winfo_ismapped())
            except tk.TclError:
                showing = False

            if not showing:
                yscroll_showing = getattr(self.sheet, "yscroll_showing", None)
                try:
                    if callable(yscroll_showing):
                        showing = bool(yscroll_showing())
                    elif yscroll_showing is not None:
                        showing = bool(yscroll_showing)
                except Exception:
                    showing = False

            if showing:
                try:
                    scrollbar_width = int(yscroll.winfo_width())
                    if scrollbar_width <= 0:
                        scrollbar_width = int(yscroll.winfo_reqwidth())
                except (tk.TclError, ValueError, TypeError):
                    try:
                        scrollbar_width = int(yscroll.winfo_reqwidth())
                    except (tk.TclError, ValueError, TypeError):
                        scrollbar_width = 0

        available_width = container_width - scrollbar_width - self.TRAILING_GUTTER
        effective_available_width = max(available_width, 0)

        column_indices = range(len(self.HEADERS))
        min_widths_map = self._calculate_column_min_widths(column_indices)
        if not min_widths_map:
            return

        widths = [min_widths_map.get(idx, 0) for idx in column_indices]
        if not widths:
            return

        total_min_width = sum(widths)

        if effective_available_width > total_min_width and widths:
            extra_width = effective_available_width - total_min_width
            weights = [max(min_widths_map.get(idx, 1), 1) for idx in column_indices]
            total_weight = sum(weights)
            if total_weight <= 0:
                total_weight = len(weights)
                weights = [1] * len(weights)

            base_additions = []
            remainders = []
            for idx, (width, weight) in enumerate(zip(widths, weights)):
                numerator = extra_width * weight
                addition = numerator // total_weight
                base_additions.append(addition)
                remainders.append((numerator % total_weight, idx))

            remainder = extra_width - sum(base_additions)
            if remainder > 0:
                remainders.sort(reverse=True)
                for _, idx in remainders[:remainder]:
                    base_additions[idx] += 1

            widths = [width + addition for width, addition in zip(widths, base_additions)]

        widths = [max(width, min_widths_map.get(idx, 0)) for idx, width in enumerate(widths)]

        self._in_container_resize = True
        try:
            for col, width in enumerate(widths):
                self.sheet.column_width(column=col, width=width, redraw=False)
            self._apply_row_striping()
        finally:
            self._in_container_resize = False

    def _restore_data(self, data: List[List[str]]) -> None:
        trimmed = [row[: len(self.HEADERS)] for row in data]
        self.sheet.set_sheet_data(trimmed)
        if trimmed:
            self._auto_resize_columns(range(len(self.HEADERS)))
            self._apply_row_striping()
            return
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._auto_resize_columns(range(len(self.HEADERS)))
        self._apply_row_striping()

    def _push_undo(self, action: str, before: List[List[str]], after: List[List[str]], cells: Sequence[CellCoord]) -> None:
        if before == after:
            return
        entry = UndoEntry(action=action, before=before, after=after, cells=list(cells))
        self.undo_stack.append(entry)
        if len(self.undo_stack) > self.max_undo:
            self.undo_stack.pop(0)

    def _flash_cells(self, cells: Iterable[CellCoord]) -> None:
        cell_list = list({cell for cell in cells if cell is not None})
        if not cell_list:
            return
        self.sheet.highlight_cells(cells=cell_list, bg="#fff2b6", fg=False)
        self.after(150, lambda: self.sheet.dehighlight_cells(cells=cell_list))

    def _get_selection_start(self) -> CellCoord:
        sel = self.sheet.get_currently_selected()
        if hasattr(sel, "row") and hasattr(sel, "column"):
            return int(sel.row), int(sel.column)
        if isinstance(sel, Sequence):
            ints = [int(v) for v in sel if isinstance(v, int)]
            if len(ints) >= 2:
                return ints[-2], ints[-1]
        cells = self.sheet.get_selected_cells()
        if cells:
            return cells[0]
        return 0, 0

    def _get_selection_bounds(self) -> Optional[Tuple[int, int, int, int]]:
        boxes = self.sheet.get_all_selection_boxes()
        if boxes:
            start_row, start_col, end_row, end_col = boxes[0]
        else:
            row, col = self._get_selection_start()
            start_row = row
            start_col = col
            end_row = row
            end_col = col

        if start_row is None or start_col is None:
            return None
        if start_row < 0 or start_col < 0:
            return None

        total_rows = self.sheet.get_total_rows()
        max_row_index = max(0, total_rows - 1)
        max_col_index = max(0, len(self.HEADERS) - 1)

        start_row = min(max(start_row, 0), max_row_index)
        end_row = min(max(end_row, 0), max_row_index)
        start_col = min(max(start_col, 0), max_col_index)
        end_col = min(max(end_col, 0), max_col_index)

        if start_row > end_row:
            start_row, end_row = end_row, start_row
        if start_col > end_col:
            start_col, end_col = end_col, start_col

        return start_row, start_col, end_row, end_col

    def _selection_to_matrix(self, bounds: Tuple[int, int, int, int]) -> List[List[str]]:
        start_row, start_col, end_row, end_col = bounds
        matrix: List[List[str]] = []
        for row in range(start_row, end_row + 1):
            row_values: List[str] = []
            for col in range(start_col, end_col + 1):
                value = self._coerce_to_str(self.sheet.get_cell_data(row, col))
                row_values.append(value)
            matrix.append(row_values)
        return matrix

    def _ensure_row_capacity(self, required_rows: int) -> None:
        current = self.sheet.get_total_rows()
        if required_rows > current:
            self.sheet.insert_rows(rows=required_rows - current, idx=current)
        self._apply_row_striping()

    def _ensure_minimum_rows(self, minimum: int = 1) -> None:
        """Zorg dat er minimaal ``minimum`` lege rijen beschikbaar zijn."""

        if minimum <= 0:
            return
        current = self.sheet.get_total_rows()
        if current >= minimum:
            return
        self.sheet.insert_rows(rows=minimum - current, idx=current)

    def _event_to_cell(self, event) -> CellCoord:
        if isinstance(event, dict):
            return int(event.get("row", 0)), int(event.get("column", 0))
        if hasattr(event, "row") and hasattr(event, "column"):
            return int(event.row), int(event.column)
        if hasattr(event, "r") and hasattr(event, "c"):
            return int(event.r), int(event.c)
        if isinstance(event, Sequence):
            ints = [int(v) for v in event if isinstance(v, int)]
            if len(ints) >= 2:
                return ints[-2], ints[-1]
        return 0, 0

    # ------------------------------------------------------------------
    # Acties
    def _on_paste(self, event=None):
        try:
            raw = self.clipboard_get()
        except tk.TclError:
            self._update_status("Geen data op het klembord.")
            return "break"

        lines = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        rows: List[List[str]] = []
        for line in lines:
            if line == "":
                continue
            values = [cell.strip() for cell in line.split("\t")]
            if not any(values):
                continue
            rows.append(values)
        if not rows:
            self._update_status("Geen gegevens om te plakken.")
            return "break"

        start_row, start_col = self._get_selection_start()
        before = self._snapshot_data()

        max_cols = len(self.HEADERS)
        required_rows = start_row + len(rows)
        self._ensure_row_capacity(required_rows)

        changed_cells = []
        for r_offset, values in enumerate(rows):
            target_row = start_row + r_offset
            for c_offset, value in enumerate(values):
                target_col = start_col + c_offset
                if target_col >= max_cols:
                    break
                new_val = value.strip()
                old_val = self._coerce_to_str(self.sheet.get_cell_data(target_row, target_col))
                if old_val == new_val:
                    continue
                self.sheet.set_cell_data(target_row, target_col, new_val, redraw=False)
                changed_cells.append((target_row, target_col))
        changed_columns = {col for _, col in changed_cells}
        if changed_columns:
            self._auto_resize_columns(changed_columns)
        self._apply_row_striping()
        after = self._snapshot_data()
        self._push_undo("paste", before, after, changed_cells)

        if changed_cells:
            self._flash_cells(changed_cells)
            rows_touched = len({cell[0] for cell in changed_cells})
            self._update_status(f"{rows_touched} rijen geplakt.")
        else:
            self._update_status("Geen wijzigingen tijdens plakken.")
        return "break"

    def _on_copy(self, event=None):
        bounds = self._get_selection_bounds()
        if not bounds:
            return "break"

        data = self._selection_to_matrix(bounds)
        text = "\n".join("\t".join(row) for row in data)
        self.clipboard_clear()
        self.clipboard_append(text)

        start_row, start_col, end_row, end_col = bounds
        cell_count = (end_row - start_row + 1) * (end_col - start_col + 1)
        suffix = "cel" if cell_count == 1 else "cellen"
        self._update_status(f"{cell_count} {suffix} gekopieerd naar klembord.")
        return "break"

    def _on_cut(self, event=None):
        bounds = self._get_selection_bounds()
        if not bounds:
            return "break"

        before = self._snapshot_data()
        data = self._selection_to_matrix(bounds)
        text = "\n".join("\t".join(row) for row in data)
        self.clipboard_clear()
        self.clipboard_append(text)

        start_row, start_col, end_row, end_col = bounds
        changed: List[CellCoord] = []
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                old_val = self._coerce_to_str(self.sheet.get_cell_data(row, col))
                if old_val == "":
                    continue
                self.sheet.set_cell_data(row, col, "", redraw=False)
                changed.append((row, col))

        if changed:
            self.sheet.refresh()
            self._auto_resize_columns({col for _, col in changed})
        self._apply_row_striping()

        after = self._snapshot_data()
        self._push_undo("cut", before, after, changed)

        if changed:
            self._flash_cells(changed)
            cell_count = len(changed)
            suffix = "cel" if cell_count == 1 else "cellen"
            self._update_status(f"{cell_count} {suffix} geknipt.")
        else:
            self._update_status("Geen waarden om te knippen.")
        return "break"

    def _on_delete(self, event=None):
        cells = list(dict.fromkeys(self.sheet.get_selected_cells()))
        if not cells:
            cells = [self._get_selection_start()]
        before = self._snapshot_data()
        changed = []
        for row, col in cells:
            old_val = self._coerce_to_str(self.sheet.get_cell_data(row, col))
            if old_val == "":
                continue
            self.sheet.set_cell_data(row, col, "", redraw=False)
            changed.append((row, col))
        if changed:
            changed_columns = {col for _, col in changed}
            self._auto_resize_columns(changed_columns)
        self._apply_row_striping()
        after = self._snapshot_data()
        self._push_undo("delete", before, after, changed)
        if changed:
            self._flash_cells(changed)
            self._update_status(f"{len(changed)} cellen gewist.")
        else:
            self._update_status("Geen cellen om te wissen.")
        return "break"

    def _on_single_click_release(self, event) -> None:
        if event is not None:
            if event.state & 0x0001 or event.state & 0x0004 or event.state & 0x0008:
                return
        self._cancel_pending_single_click()
        delay = max(1, int(self._single_click_delay_ms))
        self._single_click_after_id = self.after(delay, self._start_single_click_edit)

    def _start_single_click_edit(self) -> None:
        self._single_click_after_id = None
        text_editor = getattr(self.sheet.MT, "text_editor", None)
        if text_editor is not None and getattr(text_editor, "open", False):
            return

        bounds = self._get_selection_bounds()
        if not bounds:
            return
        start_row, start_col, end_row, end_col = bounds
        if start_row != end_row or start_col != end_col:
            return

        try:
            self.sheet.focus_set()
            self.sheet.MT.focus_set()
        except Exception:
            pass
        try:
            self.sheet.MT.open_cell(ignore_existing_editor=True)
        except Exception:
            pass

    def _on_single_click_press(self, event) -> None:
        self._cancel_pending_single_click()

    def _on_single_click_cancel(self, event=None) -> None:
        self._cancel_pending_single_click()

    def _cancel_pending_single_click(self) -> None:
        if self._single_click_after_id is None:
            return
        try:
            self.after_cancel(self._single_click_after_id)
        except Exception:
            pass
        self._single_click_after_id = None

    def _determine_single_click_delay(self) -> int:
        try:
            value = int(self.tk.call("tk", "getdoubleclicktime"))
            if value <= 0:
                raise ValueError("invalid double-click time")
            return value
        except Exception:
            return 300

    def _confirm_clear(self) -> None:
        data_before = self._snapshot_data()
        if not any(any(cell for cell in row) for row in data_before):
            self._update_status("Sheet was al leeg.")
            return
        if not messagebox.askyesno("Bevestigen", "Alle custom BOM-data verwijderen?", parent=self):
            return
        self.sheet.set_sheet_data([])
        self._ensure_minimum_rows(self.DEFAULT_EMPTY_ROWS)
        self._auto_resize_columns(range(len(self.HEADERS)))
        self._apply_row_striping()
        self._push_undo("clear", data_before, self._snapshot_data(), [])
        coords = [
            (r, c)
            for r, row in enumerate(data_before)
            for c, value in enumerate(row)
            if value
        ]
        self._flash_cells(coords)
        self._update_status("Custom BOM geleegd.")

    def _on_begin_edit_cell(self, event) -> None:
        self._edit_snapshot = self._snapshot_data()
        self._edit_cell = self._event_to_cell(event)

    def _on_end_edit_cell(self, event) -> None:
        if self._edit_snapshot is None or self._edit_cell is None:
            return
        row, col = self._event_to_cell(event)
        after = self._snapshot_data()
        before_val = ""
        try:
            before_val = self._edit_snapshot[row][col]
        except IndexError:
            before_val = ""
        current_val = self._coerce_to_str(self.sheet.get_cell_data(row, col))
        if before_val != current_val:
            self._push_undo("edit", self._edit_snapshot, after, [(row, col)])
            self._update_status(f"Cel ({row + 1}, {col + 1}) bijgewerkt.")
            self._auto_resize_columns([col])
        total_rows = self.sheet.get_total_rows()
        if row >= total_rows - 1:
            self._ensure_minimum_rows(row + 2)
            self._apply_row_striping()
        self._edit_snapshot = None
        self._edit_cell = None

    def _on_undo(self, event=None):
        if not self.undo_stack:
            self._update_status("Niets om ongedaan te maken.")
            return "break"
        entry = self.undo_stack.pop()
        self._restore_data(entry.before)
        self._flash_cells(entry.cells)
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

