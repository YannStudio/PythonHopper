"""Example Tkinter application showcasing custom Pandastable bindings.

The demo focuses on providing Excel-like editing behaviour while keeping
Pandastable's own functionality intact.  Run the module directly to see the
behaviour.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox
from typing import Iterable

import pandas as pd
from pandastable import Table


class DirectEditTable(Table):
    """Table subclass with direct typing, Excel-style shortcuts and selection."""

    def __init__(self, parent: tk.Widget, dataframe: pd.DataFrame, **kwargs) -> None:
        super().__init__(parent, dataframe=dataframe, editable=True, **kwargs)
        self._drag_selection_active = False
        self._install_custom_bindings()

    # ------------------------------------------------------------------
    # Binding helpers
    def _install_custom_bindings(self) -> None:
        """Wire additional keyboard and mouse bindings without overriding defaults."""

        # Mouse drag tracking for rectangular selections.  ``add='+'`` keeps the
        # default Pandastable behaviour wired up by ``Table.doBindings``.
        self.bind("<Button-1>", self._remember_drag_origin, add="+")
        self.bind("<B1-Motion>", self._extend_drag_selection, add="+")
        self.bind("<ButtonRelease-1>", self._finalise_drag_selection, add="+")

        # Direct typing: start an edit whenever the user types a printable key
        # while a cell is selected.  Ctrl/Command modifiers are ignored so that
        # application shortcuts keep functioning.
        self.bind("<Key>", self._maybe_start_direct_edit, add="+")

        # Navigation/edit control keys.
        self.bind("<Return>", self._handle_return_key, add="+")
        self.bind("<KP_Enter>", self._handle_return_key, add="+")
        self.bind("<Tab>", self._handle_tab_key, add="+")
        self.bind("<ISO_Left_Tab>", self._handle_shift_tab_key, add="+")
        self.bind("<Shift-Tab>", self._handle_shift_tab_key, add="+")

        # Excel-style clipboard shortcuts on all major platforms.
        self._bind_shortcut_sequences("c", self.copy)
        self._bind_shortcut_sequences("x", self.cut)
        self._bind_shortcut_sequences("v", self.paste)

    def _bind_shortcut_sequences(self, key: str, handler) -> None:  # type: ignore[override]
        sequences = self._shortcut_sequences(key)
        for sequence in sequences:
            self.bind(sequence, handler, add="+")

    @staticmethod
    def _shortcut_sequences(key: str) -> Iterable[str]:
        base = key.lower()
        upper = key.upper()
        modifiers = ("Control", "Command", "Meta")
        for mod in modifiers:
            yield f"<{mod}-{base}>"
            yield f"<{mod}-{upper}>"

    # ------------------------------------------------------------------
    # Mouse drag handling
    def _remember_drag_origin(self, event) -> None:
        self._drag_selection_active = True

    def _extend_drag_selection(self, event) -> None:
        if not self._drag_selection_active:
            return
        # ``Table.handle_mouse_drag`` already does the heavy lifting, but we make
        # sure the multi-cell rectangle is redrawn as the mouse moves.
        if self.multiplerowlist and self.multiplecollist:
            self.drawMultipleCells()

    def _finalise_drag_selection(self, event) -> None:
        if self._drag_selection_active and self.multiplerowlist and self.multiplecollist:
            # Keep the current rectangle when the mouse is released.
            self.drawMultipleCells()
        self._drag_selection_active = False

    # ------------------------------------------------------------------
    # Keyboard helpers
    def _modifier_is_pressed(self, event) -> bool:
        state = getattr(event, "state", 0)
        control_mask = 0x0004  # Control on Windows/Linux
        meta_mask = 0x0008     # Meta/Alt mask (covers Command on some Tk builds)
        command_mask = 0x100000
        if state & (control_mask | meta_mask | command_mask):
            return True
        if event.keysym in {"Control_L", "Control_R", "Alt_L", "Alt_R", "Meta_L", "Meta_R", "Command"}:
            return True
        return False

    def _maybe_start_direct_edit(self, event) -> str | None:
        if self._modifier_is_pressed(event):
            return None
        if event.keysym in {"Shift_L", "Shift_R", "Caps_Lock"}:
            return None
        if event.keysym in {"Left", "Right", "Up", "Down"}:
            return None
        if event.keysym in {"Return", "KP_Enter", "Tab"}:
            return None
        if event.char and event.char.isprintable():
            return self._start_edit_with_char(event.char)
        if event.keysym in {"BackSpace", "Delete"}:
            return self._start_edit_with_char("")
        return None

    def _start_edit_with_char(self, initial_text: str) -> str:
        self.drawCellEntry(self.currentrow, self.currentcol)
        entry = getattr(self, "cellentry", None)
        if entry is not None:
            entry.delete(0, tk.END)
            if initial_text:
                entry.insert(0, initial_text)
            entry.icursor(tk.END)
        return "break"

    def _handle_return_key(self, event) -> str:
        if hasattr(self, "cellentry"):
            self._commit_editor()
        else:
            self.drawCellEntry(self.currentrow, self.currentcol)
        return "break"

    def _handle_tab_key(self, event) -> str:
        self._commit_editor()
        self._move_selection(0, 1)
        return "break"

    def _handle_shift_tab_key(self, event) -> str:
        self._commit_editor()
        self._move_selection(0, -1)
        return "break"

    def cut(self, event=None):  # type: ignore[override]
        """Cut selection using the native Pandastable implementation when available."""

        base_cut = getattr(super(), "cut", None)
        if callable(base_cut):
            return base_cut(event)
        # Fallback: copy and clear the current selection.
        self.copy(event)
        self.clearData()
        return "break"

    def _commit_editor(self) -> None:
        if not hasattr(self, "cellentry"):
            return
        value = self.cellentryvar.get()
        row = self.currentrow
        col = self.currentcol
        if self.filtered == 1:
            self.delete("entry")
            return
        result = self.model.setValueAt(value, row, col, df=None)
        if result is False:
            dtype = self.model.getColumnType(col)
            messagebox.showwarning(
                "Incompatible type",
                f"This column is {dtype} and cannot accept the value {value}.",
                parent=self.parentframe,
            )
            return
        self.drawText(row, col, value, align=self.align)
        self.delete("entry")

    def _move_selection(self, delta_row: int, delta_col: int) -> None:
        if self.rows == 0 or self.cols == 0:
            return
        row = max(0, min(self.rows - 1, self.currentrow + delta_row))
        col = self.currentcol + delta_col
        if col >= self.cols:
            col = 0
            row = min(self.rows - 1, row + 1)
        elif col < 0:
            col = self.cols - 1
            row = max(0, row - 1)
        self.setSelectedRow(row)
        self.setSelectedCol(col)
        self.drawSelectedRow()
        self.drawSelectedRect(row, col)
        self.rowheader.drawSelectedRows(row)


def main() -> None:
    root = tk.Tk()
    root.title("Pandastable direct edit demo")
    root.geometry("800x300")

    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)

    dataframe = pd.DataFrame(
        {
            "Artikel": ["Staal", "Aluminium", "Koper", "Messing"],
            "Aantal": [10, 5, 3, 12],
            "Prijs": [4.50, 6.80, 7.25, 8.10],
        }
    )

    table = DirectEditTable(frame, dataframe=dataframe, showtoolbar=True, showstatusbar=True)
    table.show()

    root.mainloop()


if __name__ == "__main__":
    main()
