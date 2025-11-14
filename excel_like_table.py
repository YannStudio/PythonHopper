"""
Excel-achtige BOM tabel widget (aangepast voor header-alignment fix).

Belangrijk:
- Er is één lege kolomkop links toegevoegd om te voorkomen dat headerlabels verschoven
  t.o.v. de data-velden staan (lege kolom boven de regelverwijderknop).
- Plakken / delete / clear-gedrag blijft werken; plakken wordt niet in de lege kolom gezet.
"""
from __future__ import annotations
from typing import List
import sys

from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTableView,
    QHeaderView,
    QMessageBox,
    QStatusBar,
    QStyledItemDelegate,
    QMainWindow,
)
from PyQt6.QtGui import (
    QStandardItemModel,
    QStandardItem,
    QKeySequence,
    QShortcut,
    QColor,
    QClipboard,
)
from PyQt6.QtCore import (
    Qt,
    QTimer,
    QModelIndex,
    QEvent,
    QItemSelectionModel,
    QItemSelection,
)

# Exact kolomnamen in de gevraagde volgorde (data-kolommen)
DATA_COLUMNS = [
    "PartNumber",
    "Description",
    "QTY.",
    "Profile Length",
    "Profile",
    "Thickness",
    "Production",
    "Material",
]
DATA_COL_COUNT = len(DATA_COLUMNS)

# Aantal lege kolommen voor knoppen/rijnummer etc. (hier 1 lege kolom RECHTS)
TRAILING_EMPTY_COLS = 1

# Totale kolommen in het model
TOTAL_COLS = DATA_COL_COUNT + TRAILING_EMPTY_COLS

START_ROWS = 10


def parse_tsv(text: str) -> List[List[str]]:
    """
    Parse clipboard TSV text naar rij-lijsten.
    Splits op newline (splitlines) en tabs; trims trailing whitespace.
    """
    if not text:
        return []
    raw_rows = text.splitlines()
    rows: List[List[str]] = []
    for raw in raw_rows:
        cells = raw.split("\t")
        cells = [c.rstrip("\r").strip() for c in cells]
        rows.append(cells)
    return rows


class ClickHighlightDelegate(QStyledItemDelegate):
    """Tekent tijdelijk highlight voor de index die tijdens mouse-down actief is."""

    def __init__(self, parent: QTableView):
        super().__init__(parent)
        self._parent = parent
        self.highlight_index: QModelIndex | None = None
        self.highlight_color = QColor(200, 230, 255)

    def paint(self, painter, option, index):
        if self.highlight_index is not None and index == self.highlight_index:
            painter.save()
            painter.fillRect(option.rect, self.highlight_color)
            painter.restore()
        super().paint(painter, option, index)


class BomTableWidget(QWidget):
    """Excel-achtige tabel widget voor custom BOMs."""

    def __init__(self, parent=None):
        super().__init__(parent)
        # Model met extra lege kolom RECHTS (voor rij-verwijderknop e.d.)
        self.model = QStandardItemModel(START_ROWS, TOTAL_COLS, self)
        # bouw header labels: data-kolomnamen + lege kolomkop(s)
        header_labels = DATA_COLUMNS + [""] * TRAILING_EMPTY_COLS
        self.model.setHorizontalHeaderLabels(header_labels)

        # Vul initiële lege items
        for r in range(self.model.rowCount()):
            for c in range(self.model.columnCount()):
                it = QStandardItem("")
                # maak lege trailing-kolom niet-bewerkbaar (typisch voor knoppen)
                if c >= DATA_COL_COUNT:
                    it.setEditable(False)
                else:
                    it.setEditable(True)
                self.model.setItem(r, c, it)

        self.view = QTableView(self)
        self.view.setModel(self.model)
        self.view.setSelectionMode(QTableView.SelectionMode.ExtendedSelection)
        self.view.setSelectionBehavior(QTableView.SelectionBehavior.SelectItems)
        self.view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.view.verticalHeader().setVisible(True)
        self.view.setEditTriggers(QTableView.EditTrigger.DoubleClicked | QTableView.EditTrigger.EditKeyPressed)
        self.view.setSortingEnabled(False)  # header niet sorteerbaar

        # Delegate voor click-highlight
        self.delegate = ClickHighlightDelegate(self.view)
        self.view.setItemDelegate(self.delegate)

        # Shortcuts
        self.paste_shortcut = QShortcut(QKeySequence("Ctrl+V"), self)
        self.paste_shortcut.activated.connect(self.on_paste_shortcut)
        self.delete_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self)
        self.delete_shortcut.activated.connect(self.on_delete_shortcut)
        self.f2_shortcut = QShortcut(QKeySequence("F2"), self)
        self.f2_shortcut.activated.connect(self.on_f2_shortcut)

        # Buttons en status
        self.clear_button = QPushButton("Clear table", self)
        self.clear_button.clicked.connect(self.on_clear_table)

        top_layout = QHBoxLayout()
        top_layout.addWidget(self.clear_button)
        top_layout.addStretch()

        self.status = QStatusBar(self)

        layout = QVBoxLayout(self)
        layout.addLayout(top_layout)
        layout.addWidget(self.view)
        layout.addWidget(self.status)
        self.setLayout(layout)

        # Mouse highlight events
        self.view.viewport().installEventFilter(self)

        # Data-change hooking om automatisch rijen toe te voegen wanneer in laatste rij geschreven wordt
        self.model.dataChanged.connect(self.on_data_changed)

    # ---- Paste handling ----
    def on_paste_shortcut(self):
        clip = QApplication.clipboard()
        text = clip.text(mode=QClipboard.Clipboard)
        rows = parse_tsv(text)
        if not rows:
            self.status.showMessage("Leeg klembord of geen tekst gevonden", 2000)
            return
        sel = self.view.selectionModel().selectedIndexes()
        if sel:
            start_row = min(idx.row() for idx in sel)
            start_col = min(idx.column() for idx in sel)
        else:
            cur = self.view.currentIndex()
            start_row = cur.row() if cur.isValid() else 0
            start_col = cur.column() if cur.isValid() else 0
        self.paste_from_clipboard(start_row, start_col, rows)

    def paste_from_clipboard(self, start_row: int, start_col: int, rows: List[List[str]]):
        """
        Plakt rows (lijst van cel-rijen) vanaf start_row/start_col.
        Voegt rijen toe indien nodig. Negeert kolommen buiten DATA_COL_COUNT.
        Zorgt dat plakken niet in de lege trailing-kolom plaatsvindt.
        """
        # Forceer start_col op data-kolom indien user per ongeluk trailing-kolom selecteerde
        if start_col >= DATA_COL_COUNT:
            start_col = DATA_COL_COUNT - 1

        num_rows = len(rows)
        num_cols = max((len(r) for r in rows), default=0)

        required_rows = start_row + num_rows
        if required_rows > self.model.rowCount():
            add_count = required_rows - self.model.rowCount()
            self.model.insertRows(self.model.rowCount(), add_count)
            for r in range(self.model.rowCount() - add_count, self.model.rowCount()):
                for c in range(self.model.columnCount()):
                    if self.model.item(r, c) is None:
                        it = QStandardItem("")
                        if c >= DATA_COL_COUNT:
                            it.setEditable(False)
                        self.model.setItem(r, c, it)

        pasted_cells = 0
        for r_idx, r_vals in enumerate(rows):
            target_row = start_row + r_idx
            for c_idx, val in enumerate(r_vals):
                target_col = start_col + c_idx
                # truncate to data columns only (niet in trailing empty cols)
                if target_col >= DATA_COL_COUNT:
                    break
                item = self.model.item(target_row, target_col)
                if item is None:
                    item = QStandardItem("")
                    self.model.setItem(target_row, target_col, item)
                item.setText(val)
                pasted_cells += 1

        # Selecteer het geplakte gebied (top-left anchor)
        top_left = self.model.index(start_row, start_col)
        bottom_right = self.model.index(start_row + num_rows - 1, min(DATA_COL_COUNT - 1, start_col + num_cols - 1))
        sel = QItemSelection(top_left, bottom_right)
        self.view.selectionModel().select(sel, QItemSelectionModel.SelectionFlag.ClearAndSelect)
        self.view.setCurrentIndex(top_left)

        display_row = start_row + 2  # header is row 1
        display_col = start_col + 1  # 1-based
        self.status.showMessage(f"{pasted_cells} cellen geplakt ({num_rows}×{min(num_cols, DATA_COL_COUNT - start_col)}) vanaf R{display_row}C{display_col}", 5000)

    # ---- Delete / Clear ----
    def on_delete_shortcut(self):
        self.clear_selected_cells()

    def clear_selected_cells(self):
        sel = self.view.selectionModel().selectedIndexes()
        if not sel:
            return
        for idx in sel:
            if not idx.isValid():
                continue
            # Bescherm trailing-kolommen
            if idx.column() >= DATA_COL_COUNT:
                continue
            item = self.model.item(idx.row(), idx.column())
            if item is not None:
                item.setText("")

    def on_clear_table(self):
        reply = QMessageBox.question(self, "Bevestiging", "Weet je zeker dat je de tabel wil leegmaken?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            for r in range(self.model.rowCount()):
                for c in range(self.model.columnCount()):
                    it = self.model.item(r, c)
                    if it is not None and c < DATA_COL_COUNT:
                        it.setText("")
            self.status.showMessage("Tabel geleegd", 3000)

    # ---- Key navigation (Enter/Tab) ----
    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            cur = self.view.currentIndex()
            if cur.isValid():
                next_row = cur.row() + 1
                next_col = cur.column()
                if next_col >= DATA_COL_COUNT:
                    next_col = DATA_COL_COUNT - 1
                if next_row >= self.model.rowCount():
                    self.model.insertRow(self.model.rowCount())
                    for c in range(self.model.columnCount()):
                        if self.model.item(self.model.rowCount() - 1, c) is None:
                            it = QStandardItem("")
                            if c >= DATA_COL_COUNT:
                                it.setEditable(False)
                            self.model.setItem(self.model.rowCount() - 1, c, it)
                next_index = self.model.index(next_row, next_col)
                self.view.setCurrentIndex(next_index)
                self.view.edit(next_index)
                return
        if event.key() == Qt.Key.Key_Tab:
            cur = self.view.currentIndex()
            if cur.isValid():
                next_col = cur.column() + 1
                next_row = cur.row()
                if next_col >= DATA_COL_COUNT:
                    next_col = 0
                    next_row = cur.row() + 1
                if next_row >= self.model.rowCount():
                    self.model.insertRow(self.model.rowCount())
                    for c in range(self.model.columnCount()):
                        if self.model.item(self.model.rowCount() - 1, c) is None:
                            it = QStandardItem("")
                            if c >= DATA_COL_COUNT:
                                it.setEditable(False)
                            self.model.setItem(self.model.rowCount() - 1, c, it)
                next_index = self.model.index(next_row, next_col)
                self.view.setCurrentIndex(next_index)
                self.view.edit(next_index)
                return
        super().keyPressEvent(event)

    # ---- Mouse highlight (temporary op mousedown) ----
    def eventFilter(self, watched, event):
        if watched is self.view.viewport():
            if event.type() == QEvent.Type.MouseButtonPress:
                pos = event.pos()
                idx = self.view.indexAt(pos)
                if idx.isValid():
                    self.delegate.highlight_index = idx
                    self.view.viewport().update(self.view.visualRect(idx))
                return False
            if event.type() == QEvent.Type.MouseButtonRelease:
                if self.delegate.highlight_index is not None:
                    QTimer.singleShot(150, self.clear_temp_highlight)
                return False
        return super().eventFilter(watched, event)

    def clear_temp_highlight(self):
        if self.delegate.highlight_index is not None:
            rect = self.view.visualRect(self.delegate.highlight_index)
            self.delegate.highlight_index = None
            self.view.viewport().update(rect)

    def on_f2_shortcut(self):
        cur = self.view.currentIndex()
        if cur.isValid():
            self.view.edit(cur)

    def on_data_changed(self, topLeft: QModelIndex, bottomRight: QModelIndex, roles=None):
        if bottomRight.row() >= self.model.rowCount() - 1:
            # zorg voor extra lege rij
            self.model.insertRow(self.model.rowCount())
            for c in range(self.model.columnCount()):
                if self.model.item(self.model.rowCount() - 1, c) is None:
                    it = QStandardItem("")
                    if c >= DATA_COL_COUNT:
                        it.setEditable(False)
                    self.model.setItem(self.model.rowCount() - 1, c, it)

    def activate(self):
        self.view.setFocus()


# Standalone entrypoint
def main():
    app = QApplication(sys.argv)
    win = QMainWindow()
    win.setWindowTitle("Custom BOM - Test")
    w = BomTableWidget()
    win.setCentralWidget(w)
    win.resize(1000, 600)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
