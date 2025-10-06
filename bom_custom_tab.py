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
    _PANDASTABLE_IMPORT_ERROR: Optional[BaseException] = None
except ModuleNotFoundError as exc:  # pragma: no cover - afhankelijk van installatie
    Table = None  # type: ignore[assignment]
    TableModel = object  # type: ignore[assignment]
    _PANDASTABLE_IMPORT_ERROR = exc

_PANDASTABLE_ERROR = (
    "De module 'pandastable' is niet geïnstalleerd. "
    "Voer 'pip install pandastable' uit voordat u de Filehopper GUI start."
)

CellCoord = Tuple[int, int]


def _ensure_pandastable_available() -> None:
    if Table is not None:
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

    def setValueAt(self, value, rowIndex, columnIndex) -> bool:  # type: ignore[override]
        before = self.df.copy(deep=True)
        str_value = "" if value is None else str(value)
        changed = super().setValueAt(str_value, rowIndex, columnIndex)
        if changed and self._on_change is not None:
            after = self.df.copy(deep=True)
            self._on_change(before, after, (int(rowIndex), int(columnIndex)))
        return changed


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
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        initial_df = self._create_empty_dataframe(self.DEFAULT_EMPTY_ROWS)
        self._dataframe = initial_df
        self.table_model = _UndoableTableModel(initial_df.copy(deep=True), self._on_model_change)
        self.table = Table(
            container,


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
            self.table.updateModel(self.table_model)
            self.table.redraw()
        finally:
            self._suspend_history = False
        self._dataframe = self.table_model.df.copy(deep=True)

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
