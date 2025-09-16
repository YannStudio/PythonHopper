# excel_like_table.py
"""
Installatie & gebruik
---------------------
pip install PyQt6
python excel_like_table.py
"""
from __future__ import annotations

import sys
from typing import List

import pandas as pd
from PyQt6 import QtCore, QtGui, QtWidgets


COLUMNS = [
    "PartNumber",
    "Description",
    "QTY.",
    "Profile Length",
    "Profile",
    "Thickness",
    "Production",
    "Material",
]


def parse_tsv(text: str) -> List[List[str]]:
    """
    Parseer TSV-tekst (Excel-plakformaat) naar een lijst van rijen.

    Parameters
    ----------
    text : str
        Klembordtekst in tab-gescheiden formaat.

    Returns
    -------
    list[list[str]]
        Elke sublijst bevat de kolomwaarden van één rij.
    """
    rows: List[List[str]] = []
    for line in text.splitlines():
        cells = [c.strip().rstrip("\r") for c in line.split("\t")]
        # negeer eventuele volledig lege laatste rij
        if len(cells) == 1 and cells[0] == "":
            continue
        rows.append(cells)
    return rows


class HighlightTableWidget(QtWidgets.QTableWidget):
    """QTableWidget met tijdelijke klik-highlight."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._pressed_item: QtWidgets.QTableWidgetItem | None = None
        self._orig_brush: QtGui.QBrush | None = None

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        item = self.itemAt(event.pos())
        if item:
            self._pressed_item = item
            self._orig_brush = item.background()
            item.setBackground(QtGui.QColor("#ffe680"))  # gele flash
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent) -> None:
        if self._pressed_item:
            item = self._pressed_item
            orig = self._orig_brush
            QtCore.QTimer.singleShot(150, lambda: item.setBackground(orig))
            self._pressed_item = None
        super().mouseReleaseEvent(event)


class MainWindow(QtWidgets.QMainWindow):
    """Hoofdvenster met Excel-achtige tabel."""

    def __init__(self) -> None:
        super().__init__()
        self._adding_row = False

        # Clear-knop + tabel in verticale layout
        clear_btn = QtWidgets.QPushButton("Clear table")
        clear_btn.clicked.connect(self.clear_table)

        upload_btn = QtWidgets.QPushButton("Upload to main app")
        upload_btn.clicked.connect(self.upload)

        self.table = HighlightTableWidget(10, len(COLUMNS), self)
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection
        )
        self.table.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectItems
        )
        self.table.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.SelectedClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.table.itemChanged.connect(self._add_row_if_needed)
        self.table.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.ResizeMode.Interactive
        )
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                self.table.setItem(r, c, QtWidgets.QTableWidgetItem(""))

        button_row = QtWidgets.QHBoxLayout()
        button_row.addWidget(clear_btn)
        button_row.addWidget(upload_btn)

        layout = QtWidgets.QVBoxLayout()
        layout.addLayout(button_row)
        layout.addWidget(self.table)

        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Sneltoetsen
        paste_shortcut = QtGui.QShortcut(QtGui.QKeySequence("Ctrl+V"), self.table)
        paste_shortcut.activated.connect(self.paste)

        del_shortcut = QtGui.QShortcut(QtGui.QKeySequence("Delete"), self.table)
        del_shortcut.activated.connect(self.delete_cells)

        self.statusBar().showMessage("Klaar")

    # ── Clipboard functionaliteit ──────────────────────────────────────────────
    def paste(self) -> None:
        indexes = self.table.selectedIndexes()
        if indexes:
            start_row = min(i.row() for i in indexes)
            start_col = min(i.column() for i in indexes)
        else:
            start_row = max(self.table.currentRow(), 0)
            start_col = max(self.table.currentColumn(), 0)
        self.paste_from_clipboard(start_row, start_col)

    def paste_from_clipboard(self, start_row: int, start_col: int) -> None:
        text = QtWidgets.QApplication.clipboard().text(
            mode=QtGui.QClipboard.Mode.Clipboard
        )
        data = parse_tsv(text)
        if not data:
            return

        max_cols = self.table.columnCount()
        cells_pasted = 0
        for r_off, row in enumerate(data):
            dest_row = start_row + r_off
            if dest_row >= self.table.rowCount():
                self.table.insertRow(dest_row)
                for c in range(max_cols):
                    self.table.setItem(dest_row, c, QtWidgets.QTableWidgetItem(""))
            for c_off, cell in enumerate(row):
                dest_col = start_col + c_off
                if dest_col >= max_cols:
                    break
                self.table.setItem(dest_row, dest_col, QtWidgets.QTableWidgetItem(cell))
                cells_pasted += 1

        rows = len(data)
        cols = min(max(len(r) for r in data), max_cols - start_col)
        self.statusBar().showMessage(
            f"{cells_pasted} cellen geplakt ({rows}×{cols}) vanaf R{start_row+1}C{start_col+1}",
            3000,
        )
        self._add_row_if_needed()

    # ── Bewerken / wissen ─────────────────────────────────────────────────────
    def delete_cells(self) -> None:
        for idx in self.table.selectedIndexes():
            item = self.table.item(idx.row(), idx.column())
            if item:
                item.setText("")
        self.statusBar().showMessage("Geselecteerde cellen gewist", 2000)

    def clear_table(self) -> None:
        ans = QtWidgets.QMessageBox.question(
            self,
            "Bevestigen",
            "Weet je zeker dat je de tabel wil leegmaken?",
            QtWidgets.QMessageBox.StandardButton.Yes
            | QtWidgets.QMessageBox.StandardButton.No,
        )
        if ans == QtWidgets.QMessageBox.StandardButton.Yes:
            self.table.setRowCount(0)
            for _ in range(10):
                row = self.table.rowCount()
                self.table.insertRow(row)
                for c in range(self.table.columnCount()):
                    self.table.setItem(row, c, QtWidgets.QTableWidgetItem(""))
            self.statusBar().showMessage("Tabel leeggemaakt", 2000)

    def upload(self) -> None:
        rows: list[list[str]] = []
        for r in range(self.table.rowCount()):
            row_values: list[str] = []
            has_data = False
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                value = item.text() if item else ""
                row_values.append(value)
                if value.strip():
                    has_data = True
            if has_data:
                rows.append(row_values)

        df = pd.DataFrame(rows, columns=COLUMNS)

        if len(sys.argv) < 2:
            QtWidgets.QMessageBox.critical(
                self,
                "Upload mislukt",
                "Geen uploadpad opgegeven (sys.argv[1]).",
            )
            return

        csv_path = sys.argv[1]
        df.to_csv(csv_path, index=False)
        print(csv_path, flush=True)
        self.close()
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.quit()

    # ── Hulpmethoden ─────────────────────────────────────────────────────────
    def _add_row_if_needed(self, _item: QtWidgets.QTableWidgetItem | None = None) -> None:
        """Voeg automatische rij toe wanneer laatste rij data bevat."""
        if self._adding_row:
            return
        last = self.table.rowCount() - 1
        has_data = any(
            (self.table.item(last, c) and self.table.item(last, c).text().strip())
            for c in range(self.table.columnCount())
        )
        if has_data:
            self._adding_row = True
            self.table.insertRow(self.table.rowCount())
            for c in range(self.table.columnCount()):
                self.table.setItem(
                    self.table.rowCount() - 1, c, QtWidgets.QTableWidgetItem("")
                )
            self._adding_row = False


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.setWindowTitle("Custom BOM")
    win.resize(900, 400)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
