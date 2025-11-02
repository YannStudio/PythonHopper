import os
import datetime
import math
import re
import shutil
import subprocess
import sys
import threading
import unicodedata
from collections import defaultdict, OrderedDict
from dataclasses import dataclass
from copy import deepcopy
from pathlib import Path
from typing import Dict, Iterable, List, Optional, TYPE_CHECKING

import pandas as pd

from app_settings import AppSettings, FileExtensionSetting, FILE_EXTENSION_PRESETS
from helpers import _to_str, _build_file_index, create_export_bundle, ExportBundleResult
from models import Supplier, Client, DeliveryAddress
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from clients_db import ClientsDB, CLIENTS_DB_FILE
from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE
from bom import read_csv_flex, load_bom
from bom_custom_tab import BOMCustomTab
from bom_sync import prepare_custom_bom_for_main
from orders import (
    copy_per_production_and_orders,
    DEFAULT_FOOTER_NOTE,
    combine_pdfs_per_production,
    combine_pdfs_from_source,
    find_related_bom_exports,
    make_bom_export_filename,
    _prefix_for_doc_type,
    _normalize_doc_number,
    _export_bom_workbook,
    describe_finish_combo,
    make_finish_selection_key,
    make_production_selection_key,
    make_opticutter_selection_key,
    make_opticutter_default_key,
    compute_opticutter_order_details,
    parse_selection_key,
    _WINDOWS_MAX_PATH,
)
from opticutter import (
    DEFAULT_KERF_MM,
    LONG_STOCK_LENGTH_MM,
    STOCK_LENGTH_MM,
    StockScenarioResult,
    analyse_profiles,
    parse_length_to_mm,
    prepare_opticutter_export,
)

if TYPE_CHECKING:
    from orders import OpticutterOrderComputation


CLIENT_LOGO_DIR = Path("client_logos")
# A softer brand accent for manufacturing-focused actions.
MANUFACT_BRAND_COLOR = "#F9C74F"


def _norm(text: str) -> str:
    return (
        unicodedata.normalize("NFKD", text)
        .encode("ASCII", "ignore")
        .decode("ASCII")
        .lower()
    )


def sort_supplier_options(
    options: List[str],
    suppliers: List[Supplier],
    disp_to_name: Dict[str, str],
) -> List[str]:
    """Return options sorted with favorites first and then alphabetically.

    Parameters
    ----------
    options: list of display strings
    suppliers: list of Supplier objects from the DB
    disp_to_name: mapping from display string to supplier name
    """

    fav_map = {_norm(s.supplier): s.favorite for s in suppliers}

    def sort_key(opt: str):
        name = disp_to_name.get(opt, opt)
        n = _norm(name)
        return (not fav_map.get(n, False), n)

    return sorted(options, key=sort_key)

def start_gui():
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import ttk, filedialog, messagebox, simpledialog
    try:
        from PIL import Image, ImageTk  # type: ignore
        try:
            RESAMPLE = Image.Resampling.LANCZOS  # type: ignore[attr-defined]
        except AttributeError:  # pragma: no cover - Pillow < 9
            RESAMPLE = Image.LANCZOS
    except Exception:  # pragma: no cover - Pillow might be unavailable in minimal setups
        Image = None  # type: ignore
        ImageTk = None  # type: ignore
        RESAMPLE = None

    TREE_ODD_BG = "#FFFFFF"
    TREE_EVEN_BG = "#F5F5F5"
    STOCK_LENGTH_MM = 6000
    LONG_STOCK_LENGTH_MM = 12000
    DEFAULT_KERF_MM = 5.0

    def _entry_overflows(entry: "tk.Entry", text: str) -> bool:
        """Return True if the Entry content is wider than the widget."""

        if not text:
            return False
        entry.update_idletasks()
        width = entry.winfo_width()
        if width <= 1:
            width = entry.winfo_reqwidth()
        try:
            font = tkfont.nametofont(entry.cget("font"))
        except tk.TclError:
            font = tkfont.nametofont("TkDefaultFont")
        padding = 0
        try:
            padding += float(entry.cget("highlightthickness")) * 2
        except tk.TclError:
            pass
        try:
            padding += float(entry.cget("bd")) * 2
        except tk.TclError:
            pass
        usable_width = max(1, width - int(padding) - 4)
        return font.measure(text) > usable_width

    def _autosize_tree_columns(
        tree: "ttk.Treeview", padding: int = 16
    ) -> None:
        """Resize Treeview columns to fit their contents with padding."""

        if tree is None:
            return

        try:
            font = tkfont.nametofont(tree.cget("font"))
        except tk.TclError:
            font = tkfont.nametofont("TkDefaultFont")

        for column in tree["columns"]:
            heading = tree.heading(column).get("text", "")
            max_width = font.measure(heading)
            for item in tree.get_children(""):
                value = tree.set(item, column)
                if value:
                    max_width = max(max_width, font.measure(str(value)))
            tree.column(column, width=max_width + padding)

    def _scroll_entry_to_end(entry: "tk.Entry", variable: Optional["tk.StringVar"] = None) -> None:
        """Ensure the end of the entry text remains visible."""

        def adjust(*_ignored):
            try:
                entry.icursor("end")
                entry.xview_moveto(1.0)
            except tk.TclError:
                pass

        entry.bind("<FocusIn>", adjust, add="+")
        entry.bind("<Configure>", adjust, add="+")
        entry.after_idle(adjust)
        if variable is not None:
            trace_id = variable.trace_add("write", lambda *_: entry.after_idle(adjust))
            setattr(entry, "_auto_scroll_trace", trace_id)

    _parse_length_to_mm = parse_length_to_mm

    def _bold_digits(text: str) -> str:
        bold_map = {
            "0": "𝟬",
            "1": "𝟭",
            "2": "𝟮",
            "3": "𝟯",
            "4": "𝟰",
            "5": "𝟱",
            "6": "𝟲",
            "7": "𝟳",
            "8": "𝟴",
            "9": "𝟵",
        }
        return "".join(bold_map.get(ch, ch) for ch in str(text))

    @dataclass
    class StockScenarioResult:
        bars: int
        waste_mm: float
        waste_pct: float
        dropped_pieces: int
        cuts: int

    # Stock scenario calculations are provided by opticutter.analyse_profiles

    class _OverflowTooltip:
        """Show a tooltip with full text when an Entry's content overflows."""

        def __init__(self, widget: "tk.Entry", text_provider):
            self.widget = widget
            self._text_provider = text_provider
            self._tipwindow: Optional["tk.Toplevel"] = None
            self._after_id: Optional[str] = None
            widget.bind("<Enter>", self._schedule_show, add="+")
            widget.bind("<Leave>", self._hide, add="+")
            widget.bind("<Destroy>", self._hide, add="+")

        def _schedule_show(self, _event=None):
            self._cancel_scheduled()
            if not self.widget.winfo_viewable():
                return
            self._after_id = self.widget.after(200, self._maybe_show)

        def _maybe_show(self):
            self._after_id = None
            if not self.widget.winfo_exists():
                return
            text = self._text_provider()
            if not text:
                return
            if not _entry_overflows(self.widget, text):
                return
            if self._tipwindow is not None:
                return
            tip = tk.Toplevel(self.widget)
            tip.wm_overrideredirect(True)
            try:
                tip.wm_attributes("-topmost", True)
            except tk.TclError:
                pass
            label = tk.Label(
                tip,
                text=text,
                background="#ffffe0",
                foreground="#444444",
                relief="solid",
                borderwidth=1,
                justify="left",
                padx=4,
                pady=2,
            )
            label.pack()
            x = self.widget.winfo_rootx()
            y = self.widget.winfo_rooty() + self.widget.winfo_height()
            tip.wm_geometry(f"+{x}+{y}")
            self._tipwindow = tip

        def _cancel_scheduled(self):
            if self._after_id is not None:
                try:
                    self.widget.after_cancel(self._after_id)
                except tk.TclError:
                    pass
                self._after_id = None

        def _hide(self, _event=None):
            self._cancel_scheduled()
            if self._tipwindow is not None:
                try:
                    self._tipwindow.destroy()
                except tk.TclError:
                    pass
                self._tipwindow = None

    class _TreeTooltipManager:
        """Attach tooltips to individual Treeview cells."""

        def __init__(self, tree: "ttk.Treeview"):
            self.tree = tree
            self._messages: Dict[tuple[str, str], str] = {}
            self._tipwindow: Optional["tk.Toplevel"] = None
            self._current: Optional[tuple[str, str]] = None
            tree.bind("<Motion>", self._on_motion, add="+")
            tree.bind("<Leave>", self._hide, add="+")
            tree.bind("<Destroy>", self._hide, add="+")

        def clear(self) -> None:
            self._messages.clear()
            self._hide()

        def set(self, item: str, column: str, message: str) -> None:
            key = (item, column)
            if message:
                self._messages[key] = message
            else:
                self._messages.pop(key, None)

        def _on_motion(self, event):
            try:
                item = self.tree.identify_row(event.y)
                column = self.tree.identify_column(event.x)
            except tk.TclError:
                item = ""
                column = ""
            key = (item or "", column or "")
            message = self._messages.get(key)
            if not item or not column or not message:
                self._hide()
                return
            if self._current == key and self._tipwindow is not None:
                return
            x_root = event.x_root if hasattr(event, "x_root") else self.tree.winfo_pointerx()
            y_root = event.y_root if hasattr(event, "y_root") else self.tree.winfo_pointery()
            self._show_tip(x_root + 12, y_root + 12, message)
            self._current = key

        def _show_tip(self, x: int, y: int, message: str) -> None:
            self._hide()
            if not message:
                return
            tip = tk.Toplevel(self.tree)
            tip.wm_overrideredirect(True)
            try:
                tip.wm_attributes("-topmost", True)
            except tk.TclError:
                pass
            label = tk.Label(
                tip,
                text=message,
                background="#ffffe0",
                foreground="#444444",
                relief="solid",
                borderwidth=1,
                justify="left",
                padx=4,
                pady=2,
                anchor="w",
            )
            label.pack()
            tip.wm_geometry(f"+{x}+{y}")
            self._tipwindow = tip

        def _hide(self, _event=None):
            self._current = None
            if self._tipwindow is not None:
                try:
                    self._tipwindow.destroy()
                except tk.TclError:
                    pass
                self._tipwindow = None

    def _place_window_near_parent(win: "tk.Toplevel", parent: "tk.Misc") -> None:
        """Place a popup window on the same screen as its parent."""

        def _apply_geometry() -> None:
            try:
                parent.update_idletasks()
                win.update_idletasks()

                parent_x = parent.winfo_rootx()
                parent_y = parent.winfo_rooty()
                parent_w = parent.winfo_width()
                parent_h = parent.winfo_height()
                if parent_w <= 1 or parent_h <= 1:
                    parent_w = parent.winfo_reqwidth()
                    parent_h = parent.winfo_reqheight()

                win_w = win.winfo_width()
                win_h = win.winfo_height()
                if win_w <= 1 or win_h <= 1:
                    win_w = win.winfo_reqwidth()
                    win_h = win.winfo_reqheight()

                if parent_w > 1 and parent_h > 1:
                    x = parent_x + max(0, (parent_w - win_w) // 2)
                    y = parent_y + max(0, (parent_h - win_h) // 3)
                else:
                    screen_w = win.winfo_screenwidth()
                    screen_h = win.winfo_screenheight()
                    x = max(0, (screen_w - win_w) // 2)
                    y = max(0, (screen_h - win_h) // 3)

                screen_w = win.winfo_screenwidth()
                screen_h = win.winfo_screenheight()
                x = max(0, min(screen_w - win_w, x))
                y = max(0, min(screen_h - win_h, y))

                win.wm_geometry(f"+{int(x)}+{int(y)}")
            except tk.TclError:
                pass

        try:
            win.after_idle(_apply_geometry)
        except tk.TclError:
            _apply_geometry()

    class ClientsManagerFrame(tk.Frame):
        def __init__(self, master, db: ClientsDB, on_change=None):
            super().__init__(master)
            self.configure(padx=12, pady=12)
            self.db = db
            self.on_change = on_change

            cols = ("Naam", "Adres", "BTW", "E-mail")
            self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
            for c in cols:
                self.tree.heading(c, text=c)
                self.tree.column(c, width=160, anchor="w")
            self.tree.pack(fill="both", expand=True, padx=8, pady=8)
            self.tree.bind("<Double-1>", lambda _e: self.edit_sel())

            btns = tk.Frame(self)
            btns.pack(fill="x")
            tk.Button(btns, text="Toevoegen", command=self.add_client).pack(side="left", padx=4)
            tk.Button(btns, text="Bewerken", command=self.edit_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Verwijderen", command=self.remove_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Favoriet ★", command=self.toggle_fav_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Importeer CSV", command=self.import_csv).pack(side="left", padx=4)
            self.refresh()

        def refresh(self):
            for it in self.tree.get_children():
                self.tree.delete(it)
            for idx, c in enumerate(self.db.clients_sorted()):
                name = self.db.display_name(c)
                vals = (name, c.address or "", c.vat or "", c.email or "")
                tag = "odd" if idx % 2 == 0 else "even"
                self.tree.insert("", "end", values=vals, tags=(tag,))
            self.tree.tag_configure("odd", background=TREE_ODD_BG)
            self.tree.tag_configure("even", background=TREE_EVEN_BG)

        def _sel_name(self):
            sel = self.tree.selection()
            if not sel:
                return None
            vals = self.tree.item(sel[0], "values")
            return vals[0].replace("★ ", "", 1)

        def _open_edit_dialog(self, client: Optional[Client] = None):
            win = tk.Toplevel(self)
            win.title("Opdrachtgever")
            win.columnconfigure(1, weight=1)
            fields = [
                ("Naam", "name"),
                ("Adres", "address"),
                ("BTW", "vat"),
                ("E-mail", "email"),
            ]
            entries: Dict[str, tk.Entry] = {}
            for i, (lbl, key) in enumerate(fields):
                tk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=4, pady=2)
                ent = tk.Entry(win, width=40)
                ent.grid(row=i, column=1, padx=4, pady=2, sticky="ew")
                if client:
                    ent.insert(0, _to_str(getattr(client, key)))
                entries[key] = ent
            fav_var = tk.BooleanVar(value=client.favorite if client else False)
            tk.Checkbutton(win, text="Favoriet", variable=fav_var).grid(
                row=len(fields), column=1, sticky="w", padx=4, pady=2
            )

            logo_path_var = tk.StringVar(
                value=(client.logo_path if client and client.logo_path else "")
            )
            logo_crop_state = (
                dict(client.logo_crop) if client and client.logo_crop else None
            )

            logo_frame = tk.LabelFrame(win, text="Logo")
            logo_frame.grid(
                row=len(fields) + 1,
                column=0,
                columnspan=2,
                sticky="ew",
                padx=4,
                pady=(6, 2),
            )
            logo_frame.columnconfigure(0, weight=1)

            preview_label = tk.Label(
                logo_frame,
                text="Geen logo",
                relief="sunken",
                width=32,
                height=8,
                anchor="center",
                justify="center",
            )
            preview_label.grid(row=0, column=0, rowspan=4, sticky="nsew", padx=4, pady=4)

            def resolve_logo_path(path_str: str) -> Optional[Path]:
                if not path_str:
                    return None
                p = Path(path_str)
                if not p.is_absolute():
                    p = Path.cwd() / p
                return p

            def update_preview() -> None:
                path_str = logo_path_var.get().strip()
                if not path_str or Image is None:
                    preview_label.configure(text="Geen logo", image="")
                    preview_label.image = None  # type: ignore[attr-defined]
                    return
                abs_path = resolve_logo_path(path_str)
                if not abs_path or not abs_path.exists():
                    preview_label.configure(text="Logo niet gevonden", image="")
                    preview_label.image = None  # type: ignore[attr-defined]
                    return
                try:
                    with Image.open(abs_path) as src:  # type: ignore[union-attr]
                        img = src.convert("RGBA")
                except Exception:
                    preview_label.configure(text="Kan logo niet laden", image="")
                    preview_label.image = None  # type: ignore[attr-defined]
                    return
                crop = logo_crop_state
                if crop and all(k in crop for k in ("left", "top", "right", "bottom")):
                    left = max(0, min(img.width, int(crop.get("left", 0))))
                    top = max(0, min(img.height, int(crop.get("top", 0))))
                    right = max(left + 1, min(img.width, int(crop.get("right", img.width))))
                    bottom = max(top + 1, min(img.height, int(crop.get("bottom", img.height))))
                    img = img.crop((left, top, right, bottom))
                thumb = img.copy()
                if RESAMPLE is not None:
                    thumb.thumbnail((220, 120), RESAMPLE)
                else:  # pragma: no cover - fallback without Pillow resampling enum
                    thumb.thumbnail((220, 120))
                photo = ImageTk.PhotoImage(thumb)  # type: ignore[union-attr]
                preview_label.configure(image=photo, text="")
                preview_label.image = photo  # type: ignore[attr-defined]

            def upload_logo() -> None:
                path = filedialog.askopenfilename(
                    filetypes=[
                        ("Afbeeldingen", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
                        ("Alle bestanden", "*.*"),
                    ]
                )
                if not path:
                    return
                dest_dir = CLIENT_LOGO_DIR
                dest_dir.mkdir(exist_ok=True)
                ext = Path(path).suffix or ".png"
                base = entries["name"].get().strip() or Path(path).stem
                safe = re.sub(r"[^a-z0-9]+", "_", base.lower()).strip("_") or "logo"
                timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                dest = dest_dir / f"{safe}_{timestamp}{ext}"
                try:
                    shutil.copy2(path, dest)
                except Exception as exc:
                    messagebox.showerror(
                        "Fout", f"Kan logo niet kopiëren: {exc}", parent=win
                    )
                    return
                nonlocal logo_crop_state
                logo_crop_state = None
                logo_path_var.set(dest.as_posix())
                update_preview()

            def clear_logo() -> None:
                nonlocal logo_crop_state
                logo_path_var.set("")
                logo_crop_state = None
                update_preview()

            def crop_logo() -> None:
                if Image is None:
                    messagebox.showwarning(
                        "Niet beschikbaar",
                        "Pillow is vereist om te kunnen bijsnijden.",
                        parent=win,
                    )
                    return
                path_str = logo_path_var.get().strip()
                if not path_str:
                    messagebox.showinfo(
                        "Geen logo",
                        "Upload eerst een logo voordat je gaat bijsnijden.",
                        parent=win,
                    )
                    return
                abs_path = resolve_logo_path(path_str)
                if not abs_path or not abs_path.exists():
                    messagebox.showerror(
                        "Onbekend pad",
                        "Het logobestand kan niet gevonden worden.",
                        parent=win,
                    )
                    return
                try:
                    with Image.open(abs_path) as src_img:  # type: ignore[union-attr]
                        base_img = src_img.convert("RGBA")
                except Exception as exc:
                    messagebox.showerror(
                        "Fout", f"Kan logo niet openen: {exc}", parent=win
                    )
                    return

                crop_win = tk.Toplevel(win)
                crop_win.title("Bijsnijden logo")
                crop_win.transient(win)
                crop_win.resizable(False, False)

                max_w, max_h = 600, 400
                if base_img.width == 0 or base_img.height == 0:
                    messagebox.showerror(
                        "Fout", "Afbeelding heeft ongeldige afmetingen.", parent=win
                    )
                    crop_win.destroy()
                    return
                scale = min(max_w / base_img.width, max_h / base_img.height, 1.0)
                disp_w = max(1, int(round(base_img.width * scale)))
                disp_h = max(1, int(round(base_img.height * scale)))
                if scale != 1.0:
                    disp_img = base_img.resize(
                        (disp_w, disp_h), RESAMPLE or Image.BICUBIC  # type: ignore[union-attr]
                    )
                else:
                    disp_img = base_img.copy()
                photo = ImageTk.PhotoImage(disp_img)  # type: ignore[union-attr]
                canvas = tk.Canvas(
                    crop_win, width=disp_w, height=disp_h, highlightthickness=0
                )
                canvas.pack(padx=8, pady=8)
                canvas.create_image(0, 0, anchor="nw", image=photo)
                canvas.image = photo  # type: ignore[attr-defined]
                canvas.configure(cursor="crosshair")

                ratio = base_img.width / base_img.height if base_img.height else 1.0
                current_box = [0.0, 0.0, float(disp_w), float(disp_h)]
                if logo_crop_state:
                    left = max(0, min(base_img.width, int(logo_crop_state.get("left", 0))))
                    top = max(0, min(base_img.height, int(logo_crop_state.get("top", 0))))
                    right = max(
                        left + 1,
                        min(base_img.width, int(logo_crop_state.get("right", base_img.width))),
                    )
                    bottom = max(
                        top + 1,
                        min(
                            base_img.height,
                            int(logo_crop_state.get("bottom", base_img.height)),
                        ),
                    )
                    current_box = [
                        left / base_img.width * disp_w,
                        top / base_img.height * disp_h,
                        right / base_img.width * disp_w,
                        bottom / base_img.height * disp_h,
                    ]

                rect_id = None
                start_point = [0.0, 0.0]

                def draw_rect() -> None:
                    nonlocal rect_id
                    if rect_id is not None:
                        canvas.delete(rect_id)
                    rect_id = canvas.create_rectangle(
                        current_box[0],
                        current_box[1],
                        current_box[2],
                        current_box[3],
                        outline="#ff007f",
                        width=2,
                    )

                def clamp(x: float, y: float) -> tuple[float, float]:
                    return (
                        max(0.0, min(float(disp_w), x)),
                        max(0.0, min(float(disp_h), y)),
                    )

                def update_box(x0: float, y0: float, x1: float, y1: float) -> None:
                    x1, y1 = clamp(x1, y1)
                    dx = x1 - x0
                    dy = y1 - y0
                    if abs(dx) < 1 and abs(dy) < 1:
                        return
                    target_ratio = ratio if ratio > 0 else 1.0
                    abs_dx = abs(dx)
                    abs_dy = abs(dy)
                    if abs_dx == 0 and abs_dy == 0:
                        return
                    if abs_dx / target_ratio >= abs_dy:
                        width = dx
                        height = (abs(dx) / target_ratio) * (1 if dy >= 0 else -1)
                    else:
                        height = dy
                        width = (abs(dy) * target_ratio) * (1 if dx >= 0 else -1)
                    x_min = x0 if width >= 0 else x0 + width
                    x_max = x_min + abs(width)
                    y_min = y0 if height >= 0 else y0 + height
                    y_max = y_min + abs(height)
                    if x_min < 0:
                        shift = -x_min
                        x_min = 0
                        x_max += shift
                    if x_max > disp_w:
                        shift = x_max - disp_w
                        x_max = disp_w
                        x_min -= shift
                    if y_min < 0:
                        shift = -y_min
                        y_min = 0
                        y_max += shift
                    if y_max > disp_h:
                        shift = y_max - disp_h
                        y_max = disp_h
                        y_min -= shift
                    x_min = max(0.0, min(float(disp_w), x_min))
                    x_max = max(0.0, min(float(disp_w), x_max))
                    y_min = max(0.0, min(float(disp_h), y_min))
                    y_max = max(0.0, min(float(disp_h), y_max))
                    if x_max - x_min < 1 or y_max - y_min < 1:
                        return
                    current_box[0] = x_min
                    current_box[1] = y_min
                    current_box[2] = x_max
                    current_box[3] = y_max
                    draw_rect()

                def on_press(evt):
                    start_point[0], start_point[1] = clamp(evt.x, evt.y)

                def on_drag(evt):
                    update_box(start_point[0], start_point[1], evt.x, evt.y)

                canvas.bind("<Button-1>", on_press)
                canvas.bind("<B1-Motion>", on_drag)
                canvas.bind("<ButtonRelease-1>", on_drag)

                draw_rect()

                tk.Label(
                    crop_win,
                    text="Klik en sleep om het logo bij te snijden. Volledige selectie = geen crop.",
                ).pack(padx=8, pady=(0, 6))

                btns = tk.Frame(crop_win)
                btns.pack(pady=6)

                def reset_full() -> None:
                    current_box[0] = 0.0
                    current_box[1] = 0.0
                    current_box[2] = float(disp_w)
                    current_box[3] = float(disp_h)
                    draw_rect()

                def apply_crop() -> None:
                    nonlocal logo_crop_state
                    x_scale = base_img.width / disp_w
                    y_scale = base_img.height / disp_h
                    left = int(round(current_box[0] * x_scale))
                    top = int(round(current_box[1] * y_scale))
                    right = int(round(current_box[2] * x_scale))
                    bottom = int(round(current_box[3] * y_scale))
                    left = max(0, min(base_img.width, left))
                    top = max(0, min(base_img.height, top))
                    right = max(left + 1, min(base_img.width, right))
                    bottom = max(top + 1, min(base_img.height, bottom))
                    if (
                        left <= 0
                        and top <= 0
                        and right >= base_img.width
                        and bottom >= base_img.height
                    ):
                        logo_crop_state = None
                    else:
                        logo_crop_state = {
                            "left": left,
                            "top": top,
                            "right": right,
                            "bottom": bottom,
                        }
                    crop_win.destroy()
                    update_preview()

                tk.Button(btns, text="Volledige afbeelding", command=reset_full).pack(
                    side="left", padx=4
                )
                tk.Button(btns, text="Opslaan", command=apply_crop).pack(
                    side="left", padx=4
                )
                tk.Button(btns, text="Annuleer", command=crop_win.destroy).pack(
                    side="left", padx=4
                )

                _place_window_near_parent(crop_win, win)
                crop_win.grab_set()
                crop_win.focus_set()

            tk.Button(logo_frame, text="Upload", command=upload_logo).grid(
                row=0, column=1, sticky="ew", padx=4, pady=2
            )
            tk.Button(logo_frame, text="Bijsnijden", command=crop_logo).grid(
                row=1, column=1, sticky="ew", padx=4, pady=2
            )
            tk.Button(logo_frame, text="Verwijder", command=clear_logo).grid(
                row=2, column=1, sticky="ew", padx=4, pady=2
            )

            update_preview()

            def _save():
                rec = {k: e.get().strip() for k, e in entries.items()}
                rec["favorite"] = fav_var.get()
                rec["logo_path"] = logo_path_var.get().strip()
                rec["logo_crop"] = logo_crop_state
                if not rec["name"]:
                    messagebox.showwarning("Let op", "Naam is verplicht.", parent=win)
                    return
                c = Client.from_any(rec)
                self.db.upsert(c)
                self.db.save(CLIENTS_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()
                win.destroy()

            btnf = tk.Frame(win)
            btnf.grid(row=len(fields) + 2, column=0, columnspan=2, pady=6)
            tk.Button(btnf, text="Opslaan", command=_save).pack(side="left", padx=4)
            tk.Button(btnf, text="Annuleer", command=win.destroy).pack(side="left", padx=4)
            win.transient(self)
            _place_window_near_parent(win, self)
            win.grab_set()
            entries["name"].focus_set()

        def add_client(self):
            self._open_edit_dialog(None)

        def edit_sel(self):
            n = self._sel_name()
            if not n:
                return
            c = self.db.get(n)
            if c:
                self._open_edit_dialog(c)

        def remove_sel(self):
            n = self._sel_name()
            if not n:
                return
            if messagebox.askyesno("Bevestigen", f"Verwijder '{n}'?", parent=self):
                if self.db.remove(n):
                    self.db.save(CLIENTS_DB_FILE)
                    self.refresh()
                    if self.on_change:
                        self.on_change()

        def toggle_fav_sel(self):
            n = self._sel_name()
            if not n:
                return
            if self.db.toggle_fav(n):
                self.db.save(CLIENTS_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()

        def import_csv(self):
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("Excel","*.xlsx;*.xls")])
            if not path:
                return
            try:
                if path.lower().endswith((".xls", ".xlsx")):
                    df = pd.read_excel(path)
                else:
                    try:
                        df = pd.read_csv(path, encoding="latin1", sep=";")
                    except Exception:
                        df = read_csv_flex(path)
            except Exception as e:
                messagebox.showerror("Fout", str(e))
                return
            changed = 0
            for _, row in df.iterrows():
                try:
                    rec = {k: row[k] for k in df.columns if k in row}
                    c = Client.from_any(rec)
                    self.db.upsert(c)
                    changed += 1
                except Exception:
                    pass
            self.db.save(CLIENTS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()
            messagebox.showinfo("Import", f"Verwerkt (upsert): {changed}")

    class DeliveryAddressesManagerFrame(tk.Frame):
        def __init__(self, master, db: DeliveryAddressesDB, on_change=None):
            super().__init__(master)
            self.configure(padx=12, pady=12)
            self.db = db
            self.on_change = on_change

            cols = ("Naam", "Adres", "Opmerkingen")
            self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
            for c in cols:
                self.tree.heading(c, text=c)
                self.tree.column(c, width=160, anchor="w")
            self.tree.pack(fill="both", expand=True, padx=8, pady=8)
            self.tree.bind("<Double-1>", lambda _e: self.edit_sel())

            btns = tk.Frame(self)
            btns.pack(fill="x")
            tk.Button(btns, text="Toevoegen", command=self.add_address).pack(side="left", padx=4)
            tk.Button(btns, text="Bewerken", command=self.edit_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Verwijderen", command=self.remove_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Favoriet ★", command=self.toggle_fav_sel).pack(side="left", padx=4)
            self.refresh()

        def refresh(self):
            for it in self.tree.get_children():
                self.tree.delete(it)
            for idx, a in enumerate(self.db.addresses_sorted()):
                name = self.db.display_name(a)
                vals = (name, a.address or "", a.remarks or "")
                tag = "odd" if idx % 2 == 0 else "even"
                self.tree.insert("", "end", values=vals, tags=(tag,))
            self.tree.tag_configure("odd", background=TREE_ODD_BG)
            self.tree.tag_configure("even", background=TREE_EVEN_BG)

        def _sel_name(self):
            sel = self.tree.selection()
            if not sel:
                return None
            vals = self.tree.item(sel[0], "values")
            return vals[0].replace("★ ", "", 1)

        def _open_edit_dialog(self, addr: Optional[DeliveryAddress] = None):
            win = tk.Toplevel(self)
            win.title("Leveradres")
            fields = [
                ("Naam", "name"),
                ("Adres", "address"),
                ("Opmerkingen", "remarks"),
            ]
            entries = {}
            for i, (lbl, key) in enumerate(fields):
                tk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=4, pady=2)
                ent = tk.Entry(win, width=40)
                ent.grid(row=i, column=1, padx=4, pady=2)
                if addr:
                    ent.insert(0, _to_str(getattr(addr, key)))
                entries[key] = ent
            fav_var = tk.BooleanVar(value=addr.favorite if addr else False)
            tk.Checkbutton(win, text="Favoriet", variable=fav_var).grid(row=len(fields), column=1, sticky="w", padx=4, pady=2)

            def _save():
                # Convert blank strings to None so cleared fields overwrite old data
                rec = {k: (e.get().strip() or None) for k, e in entries.items()}
                rec["favorite"] = fav_var.get()
                if not rec["name"]:
                    messagebox.showwarning("Let op", "Naam is verplicht.", parent=win)
                    return
                a = DeliveryAddress.from_any(rec)
                self.db.upsert(a)
                self.db.save(DELIVERY_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()
                win.destroy()

            btnf = tk.Frame(win)
            btnf.grid(row=len(fields)+1, column=0, columnspan=2, pady=6)
            tk.Button(btnf, text="Opslaan", command=_save).pack(side="left", padx=4)
            tk.Button(btnf, text="Annuleer", command=win.destroy).pack(side="left", padx=4)
            win.transient(self)
            _place_window_near_parent(win, self)
            win.grab_set()
            entries["name"].focus_set()

        def add_address(self):
            self._open_edit_dialog(None)

        def edit_sel(self):
            n = self._sel_name()
            if not n:
                return
            a = self.db.get(n)
            if a:
                self._open_edit_dialog(a)

        def remove_sel(self):
            n = self._sel_name()
            if not n:
                return
            if messagebox.askyesno("Bevestigen", f"Verwijder '{n}'?", parent=self):
                if self.db.remove(n):
                    self.db.save(DELIVERY_DB_FILE)
                    self.refresh()
                    if self.on_change:
                        self.on_change()

        def toggle_fav_sel(self):
            n = self._sel_name()
            if not n:
                return
            if self.db.toggle_fav(n):
                self.db.save(DELIVERY_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()


    class SuppliersManagerFrame(tk.Frame):
        def __init__(self, master, db: SuppliersDB, on_change=None):
            super().__init__(master)
            self.configure(padx=12, pady=12)
            self.db = db
            self.on_change = on_change
            search = tk.Frame(self)
            search.pack(fill="x", padx=8, pady=(8, 0))
            tk.Label(search, text="Zoek:").pack(side="left")
            self.search_var = tk.StringVar()
            entry = tk.Entry(search, textvariable=self.search_var)
            entry.pack(side="left", fill="x", expand=True)
            self.search_var.trace_add("write", lambda *_: self.refresh())
            cols = ("Supplier", "BTW", "E-mail", "Tel", "Adres 1", "Adres 2")
            self.tree = ttk.Treeview(self, columns=cols, show="headings")
            for c in cols:
                self.tree.heading(c, text=c)
                self.tree.column(c, anchor="w")
            self.tree.pack(fill="both", expand=True, padx=8, pady=8)
            btns = tk.Frame(self)
            btns.pack(fill="x")
            tk.Button(btns, text="Toevoegen", command=self.add_supplier).pack(side="left", padx=4)
            tk.Button(btns, text="Bewerken", command=self.edit_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Verwijderen", command=self.remove_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Update uit CSV (merge)", command=self.merge_csv).pack(side="left", padx=4)
            tk.Button(btns, text="Favoriet ★", command=self.toggle_fav_sel).pack(side="left", padx=4)
            self.refresh()

        def suspend_search_filter(self) -> str:
            """Temporarily clear the search box and return the previous query."""

            current = self.search_var.get()
            if current:
                self.search_var.set("")
            return current

        def restore_search_filter(self, value: str) -> None:
            """Restore a previously cleared search query, if any."""

            if value:
                self.search_var.set(value)

        def refresh(self):
            for r in self.tree.get_children():
                self.tree.delete(r)
            q = self.search_var.get()
            sups = self.db.find(q)
            for i, s in enumerate(sups):
                vals = (
                    ("★ " if s.favorite else "") + (s.supplier or ""),
                    s.btw or "",
                    s.sales_email or "",
                    s.phone or "",
                    s.adres_1 or "",
                    s.adres_2 or "",
                )
                tag = "odd" if i % 2 else "even"
                self.tree.insert("", "end", iid=s.supplier, values=vals, tags=(tag,))
            self.tree.tag_configure("odd", background=TREE_ODD_BG)
            self.tree.tag_configure("even", background=TREE_EVEN_BG)

        def _sel_name(self):
            sel = self.tree.selection()
            if not sel:
                return None
            name = self.tree.item(sel[0], "values")[0]
            return name.replace("★ ", "", 1)

        def _sel_supplier(self) -> Optional[Supplier]:
            n = self._sel_name()
            if not n:
                return None
            for s in self.db.suppliers:
                if s.supplier == n:
                    return s
            return None

        def add_supplier(self):
            name = simpledialog.askstring("Nieuwe leverancier", "Naam:", parent=self)
            if not name:
                return
            s = Supplier.from_any({"supplier": name})
            self.db.upsert(s)
            self.db.save(SUPPLIERS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()

        def remove_sel(self):
            n = self._sel_name()
            if not n:
                return
            if messagebox.askyesno("Bevestigen", f"Verwijder '{n}'?", parent=self):
                if self.db.remove(n):
                    self.db.save(SUPPLIERS_DB_FILE)
                    self.refresh()
                    if self.on_change:
                        self.on_change()

        def toggle_fav_sel(self):
            n = self._sel_name()
            if not n:
                return
            if self.db.toggle_fav(n):
                self.db.save(SUPPLIERS_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()

        def merge_csv(self):
            path = filedialog.askopenfilename(
                parent=self,
                title="CSV bestand",
                filetypes=[("CSV", "*.csv"), ("Alle bestanden", "*.*")],
            )
            if not path:
                return
            try:
                df = read_csv_flex(path)
                for rec in df.to_dict(orient="records"):
                    try:
                        sup = Supplier.from_any(rec)
                        self.db.upsert(sup)
                    except Exception:
                        pass
                self.db.save(SUPPLIERS_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()
            except Exception as e:
                messagebox.showerror("Fout", str(e), parent=self)

        class _EditDialog(tk.Toplevel):
            def __init__(self, master, supplier: Supplier):
                super().__init__(master)
                self.title("Leverancier bewerken")
                self.result = None
                fields = [
                    ("supplier", "Naam"),
                    ("description", "Beschrijving"),
                    ("supplier_id", "ID"),
                    ("adres_1", "Adres 1"),
                    ("adres_2", "Adres 2"),
                    ("postcode", "Postcode"),
                    ("gemeente", "Gemeente"),
                    ("land", "Land"),
                    ("btw", "BTW"),
                    ("contact_sales", "Contact"),
                    ("sales_email", "E-mail"),
                    ("phone", "Tel"),
                ]
                self.vars = {}
                for i, (f, lbl) in enumerate(fields):
                    tk.Label(self, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=4, pady=2)
                    var = tk.StringVar(value=getattr(supplier, f) or "")
                    tk.Entry(self, textvariable=var, width=40).grid(row=i, column=1, padx=4, pady=2)
                    self.vars[f] = var
                btn = tk.Frame(self)
                btn.grid(row=len(fields), column=0, columnspan=2, pady=4)
                tk.Button(btn, text="Opslaan", command=self._ok).pack(side="left", padx=4)
                tk.Button(btn, text="Annuleer", command=self.destroy).pack(side="left", padx=4)
                self.transient(master)
                _place_window_near_parent(self, master)
                self.grab_set()

            def _ok(self):
                data = {f: v.get().strip() or None for f, v in self.vars.items()}
                try:
                    self.result = Supplier.from_any(data)
                except Exception as e:
                    messagebox.showerror("Fout", str(e), parent=self)
                    return
                self.destroy()

        def edit_sel(self):
            s = self._sel_supplier()
            if not s:
                return
            dlg = self._EditDialog(self, s)
            self.wait_window(dlg)
            if dlg.result:
                self.db.upsert(dlg.result)
                self.db.save(SUPPLIERS_DB_FILE)
                self.refresh()
                if self.on_change:
                    self.on_change()

    @dataclass
    class SupplierSelectionState:
        selections: Dict[str, str]
        doc_types: Dict[str, str]
        doc_numbers: Dict[str, str]
        remarks: Dict[str, str]
        deliveries: Dict[str, str]
        remember: bool = True

    class SupplierSelectionFrame(tk.Frame):
        """Per productie: type-to-filter of dropdown; rechts detailkaart (klik = selecteer).
           Knoppen altijd zichtbaar onderaan.
        """

        LABEL_COLUMN_WIDTH = 30

        @staticmethod
        def _install_supplier_focus_behavior(combo: "ttk.Combobox") -> None:
            """Selecteer automatisch alle tekst bij eerste focus of placeholder."""

            def _handle_focus(event):
                widget = event.widget
                try:
                    current = widget.get()
                except tk.TclError:
                    current = ""

                first_focus = not getattr(widget, "_supplier_focus_seen", False)
                placeholder = current.strip().lower() in {"(geen)", "geen"}

                if first_focus or placeholder:
                    def _select_all():
                        try:
                            widget.selection_range(0, "end")
                        except tk.TclError:
                            return

                    widget.after_idle(_select_all)

                widget._supplier_focus_seen = True

            combo.bind("<FocusIn>", _handle_focus, add="+")

        @staticmethod
        def _set_combo_value(combo: "ttk.Combobox", value: str) -> None:
            """Update combobox text and reset focus selection tracking."""

            combo.set(value)
            setattr(combo, "_supplier_focus_seen", False)

        def set_opticutter_notice(self, message: str) -> None:
            text = (message or "").strip()
            if text:
                self._opticutter_notice_var.set(text)
                self._opticutter_notice_label.pack(fill="x", pady=(0, 6))
            else:
                self._opticutter_notice_var.set("")
                self._opticutter_notice_label.pack_forget()

        def __init__(
            self,
            master,
            productions: List[str],
            finishes: List[Dict[str, str]],
            db: SuppliersDB,
            delivery_db: DeliveryAddressesDB,
            callback,
            project_number_var: tk.StringVar,
            project_name_var: tk.StringVar,
            opticutter_details: Dict[str, "OpticutterOrderComputation"] | None = None,
            initial_state: Optional["SupplierSelectionState"] = None,
        ):
            super().__init__(master)
            self.configure(padx=12, pady=12)
            self.db = db
            self.delivery_db = delivery_db
            self.callback = callback
            self.project_number_var = project_number_var
            self.project_name_var = project_name_var
            self.opticutter_details = opticutter_details or {}
            self._preview_supplier: Optional[Supplier] = None
            self._active_key: Optional[str] = None  # laatst gefocuste rij
            self.sel_vars: Dict[str, tk.StringVar] = {}
            self.doc_vars: Dict[str, tk.StringVar] = {}
            self.doc_num_vars: Dict[str, tk.StringVar] = {}
            self.remark_vars: Dict[str, tk.StringVar] = {}
            self.delivery_vars: Dict[str, tk.StringVar] = {}
            self.delivery_combos: Dict[str, ttk.Combobox] = {}
            self.row_meta: Dict[str, Dict[str, str]] = {}
            self.finish_entries = finishes

            # Grid layout: content (row=0, weight=1), buttons (row=1)
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)

            content = tk.Frame(self)
            content.grid(row=0, column=0, sticky="nsew", padx=10, pady=6)
            content.grid_columnconfigure(0, weight=1)
            content.grid_rowconfigure(0, weight=0)
            content.grid_rowconfigure(1, weight=1)

            # Left: per productie comboboxen
            left = tk.Frame(content)
            left.grid(row=0, column=0, sticky="nw")

            self._opticutter_notice_var = tk.StringVar(value="")
            self._opticutter_notice_label = tk.Label(
                left,
                textvariable=self._opticutter_notice_var,
                anchor="w",
                justify="left",
                wraplength=520,
                foreground="#B15C00",
            )
            self._opticutter_notice_label.pack(fill="x", pady=(0, 6))
            self._opticutter_notice_label.pack_forget()

            # Project info entries above production rows
            proj_container = tk.Frame(left)
            proj_container.pack(fill="x", pady=(0, 6))
            proj_container.grid_columnconfigure(0, weight=0)
            proj_container.grid_columnconfigure(1, weight=1)

            proj_frame = tk.LabelFrame(
                proj_container,
                text="Projectgegevens",
                labelanchor="n",
                padx=12,
                pady=10,
            )
            proj_frame.grid(row=0, column=0, sticky="nw")

            clear_btn_container = tk.Frame(proj_container)
            clear_btn_container.grid(row=0, column=1, sticky="new", padx=(12, 0))
            clear_btn_container.grid_columnconfigure(0, weight=1)

            tk.Button(
                clear_btn_container,
                text="Clear list",
                command=self._clear_saved_suppliers,
            ).grid(row=0, column=0, sticky="e", pady=(2, 0))

            readonly_bg = "#f0f0f0"

            pn_row = tk.Frame(proj_frame)
            pn_row.pack(fill="x", pady=3)
            project_number_label = tk.Label(
                pn_row,
                text="Projectnr.",
                width=18,
                anchor="w",
            )
            project_number_label.pack(side="left")
            field_border = "#d8d8d8"
            field_kwargs = dict(
                width=50,
                anchor="w",
                background=readonly_bg,
                relief="flat",
                borderwidth=0,
                padx=6,
                pady=2,
                highlightthickness=1,
                highlightbackground=field_border,
                highlightcolor=field_border,
            )
            project_number_value = tk.Label(
                pn_row,
                textvariable=self.project_number_var,
                **field_kwargs,
            )
            project_number_value.pack(side="left", padx=(6, 0))
            self._project_number_label = project_number_label
            self._project_number_value = project_number_value

            name_row = tk.Frame(proj_frame)
            name_row.pack(fill="x", pady=3)
            project_name_label = tk.Label(
                name_row,
                text="Projectnaam",
                width=18,
                anchor="w",
            )
            project_name_label.pack(side="left")
            project_name_value = tk.Label(
                name_row,
                textvariable=self.project_name_var,
                **field_kwargs,
            )
            project_name_value.pack(side="left", padx=(6, 0))
            self._project_name_label = project_name_label
            self._project_name_value = project_name_value

            proj_frame.update_idletasks()
            required_height = proj_frame.winfo_reqheight()
            pad_spec = project_number_value.pack_info().get("padx", 0)
            if isinstance(pad_spec, str):
                pad_parts = [int(p) for p in pad_spec.split()]
            elif isinstance(pad_spec, (tuple, list)):
                pad_parts = [int(p) for p in pad_spec]
            elif pad_spec:
                pad_parts = [int(pad_spec)]
            else:
                pad_parts = []
            desired_padding = (
                pad_parts[0] * 2 if len(pad_parts) == 1 else sum(pad_parts)
            )
            width_candidates = [
                project_number_label.winfo_reqwidth()
                + project_number_value.winfo_reqwidth(),
                project_name_label.winfo_reqwidth()
                + project_name_value.winfo_reqwidth(),
            ]
            pad_conf = proj_frame.cget("padx")
            if isinstance(pad_conf, str):
                pad_values = [int(p) for p in pad_conf.split() if p]
            elif isinstance(pad_conf, (tuple, list)):
                pad_values = [int(p) for p in pad_conf]
            elif pad_conf:
                pad_values = [int(pad_conf)]
            else:
                pad_values = []
            if len(pad_values) == 1:
                total_inner_pad = pad_values[0] * 2
            else:
                total_inner_pad = sum(pad_values)
            target_width = max(width_candidates) + desired_padding + total_inner_pad
            proj_frame.grid_propagate(False)
            proj_frame.configure(width=target_width, height=required_height)

            ttk.Separator(left, orient="horizontal").pack(fill="x", pady=(0, 6))

            delivery_opts = [
                "Geen",
                "Bestelling wordt opgehaald",
                "Leveradres wordt nog meegedeeld",
            ] + [
                self.delivery_db.display_name(a)
                for a in self.delivery_db.addresses_sorted()
            ]

            doc_type_opts = [
                "Geen",
                "Bestelbon",
                "Standaard bon",
                "Offerteaanvraag",
            ]
            self._doc_type_prefixes = {
                _prefix_for_doc_type(t) for t in doc_type_opts
            }

            header_row = tk.Frame(left)
            header_row.pack(fill="x", pady=(8, 3))
            header_label_kwargs = dict(
                anchor=tk.W,
                justify=tk.LEFT,
                background=left.cget("bg"),
            )
            header_font = ("TkDefaultFont", 10, "bold")
            header_columns = [
                ("Producttype", self.LABEL_COLUMN_WIDTH, header_font),
                ("Leverancier", 50, None),
                ("Documenttype", 18, None),
                ("Nr.", 12, None),
                ("Opmerking", 24, None),
                ("Leveradres", 50, None),
            ]

            self._header_column_frames: List[tk.Frame] = []
            self._header_labels: List[tk.Label] = []
            self._header_aligned = False
            self._header_alignment_pending = False

            for text, width, font in header_columns:
                label_kwargs = dict(header_label_kwargs)
                if font is not None:
                    label_kwargs["font"] = font
                column_frame = tk.Frame(header_row, background=left.cget("bg"))
                column_frame.pack(side="left", padx=(0, 6))
                label = tk.Label(
                    column_frame,
                    text=text,
                    width=width,
                    **label_kwargs,
                )
                label.pack(fill="x", anchor="w")
                self._header_column_frames.append(column_frame)
                self._header_labels.append(label)

            self.finish_label_by_key: Dict[str, str] = {
                entry.get("key", ""): _to_str(entry.get("label")) or _to_str(entry.get("key"))
                for entry in finishes
            }

            self.rows = []
            self.combo_by_key: Dict[str, ttk.Combobox] = {}

            def add_row(display_text: str, sel_key: str, metadata: Dict[str, str]):
                row = tk.Frame(left)
                row.pack(fill="x", pady=3)
                row_label = tk.Label(
                    row,
                    text=display_text,
                    width=self.LABEL_COLUMN_WIDTH,
                    anchor="w",
                )
                row_label.pack(side="left", padx=(0, 6))
                var = tk.StringVar()
                self.sel_vars[sel_key] = var
                combo = ttk.Combobox(row, textvariable=var, state="normal", width=50)
                combo.pack(side="left", padx=(0, 6))
                combo.bind("<<ComboboxSelected>>", self._on_combo_change)
                combo.bind(
                    "<FocusIn>", lambda _e, key=sel_key: self._on_focus_key(key)
                )
                self._install_supplier_focus_behavior(combo)
                combo.bind(
                    "<KeyRelease>",
                    lambda ev, key=sel_key, c=combo: self._on_combo_type(ev, key, c),
                )

                doc_var = tk.StringVar(value="Bestelbon")
                self.doc_vars[sel_key] = doc_var
                doc_combo = ttk.Combobox(
                    row,
                    textvariable=doc_var,
                    values=doc_type_opts,
                    state="readonly",
                    width=18,
                )
                doc_combo.pack(side="left", padx=(0, 6))
                doc_combo.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, key=sel_key: self._on_doc_type_change(key),
                )

                doc_num_var = tk.StringVar()
                self.doc_num_vars[sel_key] = doc_num_var
                doc_entry = tk.Entry(row, textvariable=doc_num_var, width=12)
                doc_entry.pack(side="left", padx=(0, 6))

                remark_var = tk.StringVar()
                self.remark_vars[sel_key] = remark_var
                remark_entry = tk.Entry(row, textvariable=remark_var, width=24)
                remark_entry.pack(side="left", padx=(0, 6))
                _scroll_entry_to_end(remark_entry, remark_var)
                _OverflowTooltip(remark_entry, lambda v=remark_var: v.get().strip())

                dvar = tk.StringVar(value="Geen")
                self.delivery_vars[sel_key] = dvar
                dcombo = ttk.Combobox(
                    row,
                    textvariable=dvar,
                    values=delivery_opts,
                    state="readonly",
                    width=50,
                )
                dcombo.pack(side="left", padx=(0, 6))
                self.delivery_combos[sel_key] = dcombo

                self.rows.append((sel_key, combo))
                self.combo_by_key[sel_key] = combo
                self.row_meta[sel_key] = metadata

                self._schedule_header_alignment(
                    {
                        "label": row_label,
                        "supplier_combo": combo,
                        "doc_combo": doc_combo,
                        "doc_entry": doc_entry,
                        "remark_entry": remark_entry,
                        "delivery_combo": dcombo,
                    }
                )

            for prod in productions:
                key = make_production_selection_key(prod)
                add_row(
                    prod,
                    key,
                    {"kind": "production", "identifier": prod, "display": prod},
                )

                comp = (self.opticutter_details or {}).get(prod)
                if comp and getattr(comp, "selection_count", 0):
                    summary_bits: List[str] = []
                    if getattr(comp, "total_bars", 0):
                        summary_bits.append(f"{comp.total_bars} staven")
                    if comp.total_weight_kg is not None:
                        summary_bits.append(f"{comp.total_weight_kg:.1f} kg")
                    summary_text = ", ".join(summary_bits)
                    label_text = "  ↳ Brutemateriaal"
                    if summary_text:
                        label_text = f"{label_text} ({summary_text})"
                    add_row(
                        label_text,
                        make_opticutter_selection_key(prod),
                        {
                            "kind": "opticutter",
                            "identifier": prod,
                            "display": f"{prod} – Brutemateriaal",
                            "summary": summary_text,
                        },
                    )

            if finishes:
                ttk.Separator(left, orient="horizontal").pack(fill="x", pady=(12, 6))
                finishes_header = tk.Frame(left)
                finishes_header.pack(fill="x")
                tk.Label(
                    finishes_header,
                    text="Afwerkingen",
                    width=self.LABEL_COLUMN_WIDTH,
                    anchor="w",
                    background=left.cget("bg"),
                    font=("TkDefaultFont", 10, "bold"),
                ).pack(side="left", padx=(0, 6))
                for entry in finishes:
                    finish_key = entry.get("key", "")
                    if not finish_key:
                        continue
                    sel_key = make_finish_selection_key(finish_key)
                    label_text = _to_str(entry.get("label")) or finish_key
                    add_row(
                        label_text,
                        sel_key,
                        {
                            "kind": "finish",
                            "identifier": finish_key,
                            "display": label_text,
                        },
                    )

            # Container voor kaarten
            preview_frame = tk.LabelFrame(
                content,
                text="Leverancier details\n(klik om te selecteren)",
                labelanchor="n",
            )
            preview_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
            preview_frame.grid_rowconfigure(0, weight=1)
            preview_frame.grid_columnconfigure(0, weight=1)

            self.cards_frame = tk.Frame(preview_frame)
            self.cards_frame.grid(row=0, column=0, sticky="nsew", pady=(8, 0))

            # Mapping voor combobox per selectie
            self.combo_by_key = getattr(self, "combo_by_key", {})

            # Buttons bar (altijd zichtbaar)
            btns = tk.Frame(self)
            btns.grid(row=1, column=0, sticky="ew", padx=10, pady=(6, 10))
            btns.grid_columnconfigure(0, weight=1)
            self.remember_var = tk.BooleanVar(value=True)
            tk.Checkbutton(
                btns,
                text="Onthoud keuze per selectie",
                variable=self.remember_var,
            ).grid(row=0, column=0, sticky="w")
            self.cancel_button = tk.Button(btns, text="Annuleer", command=self._cancel)
            self.cancel_button.grid(row=0, column=1, sticky="e", padx=(4, 0))
            self.confirm_button = tk.Button(btns, text="Bevestig", command=self._confirm)
            self.confirm_button.grid(row=0, column=2, sticky="e")
            self.status_var = tk.StringVar(value="")
            self.status_label = tk.Label(
                btns,
                textvariable=self.status_var,
                anchor="w",
                justify="left",
            )
            self.status_label.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(6, 0))

            # Init
            self._refresh_options(initial=True)
            if initial_state is not None:
                try:
                    self.apply_state(initial_state)
                except Exception:
                    pass
            self._update_preview_from_any_combo()

        def serialize_state(self) -> "SupplierSelectionState":
            selections: Dict[str, str] = {}
            for sel_key, combo in self.rows:
                try:
                    selections[sel_key] = combo.get()
                except tk.TclError:
                    continue

            doc_types = {key: var.get() for key, var in self.doc_vars.items()}
            doc_numbers = {
                key: var.get() for key, var in self.doc_num_vars.items()
            }
            remarks = {key: var.get() for key, var in self.remark_vars.items()}
            deliveries = {
                key: var.get() for key, var in self.delivery_vars.items()
            }

            remember = bool(self.remember_var.get()) if hasattr(self, "remember_var") else True

            return SupplierSelectionState(
                selections=selections,
                doc_types=doc_types,
                doc_numbers=doc_numbers,
                remarks=remarks,
                deliveries=deliveries,
                remember=remember,
            )

        def apply_state(self, state: "SupplierSelectionState") -> None:
            set_combo_value = getattr(
                type(self), "_set_combo_value", SupplierSelectionFrame._set_combo_value
            )
            for sel_key, combo in self.rows:
                if sel_key in state.selections:
                    set_combo_value(combo, state.selections[sel_key])

            for sel_key, value in state.doc_types.items():
                if sel_key in self.doc_vars:
                    self.doc_vars[sel_key].set(value)

            for sel_key, value in state.doc_numbers.items():
                if sel_key in self.doc_num_vars:
                    doc_var = self.doc_vars.get(sel_key)
                    doc_type_text = doc_var.get() if doc_var else "Bestelbon"
                    normalized = _normalize_doc_number(value, doc_type_text)
                    self.doc_num_vars[sel_key].set(normalized)

            for sel_key, value in state.remarks.items():
                if sel_key in self.remark_vars:
                    self.remark_vars[sel_key].set(value)

            for sel_key, value in state.deliveries.items():
                dcombo = self.delivery_combos.get(sel_key)
                if dcombo is not None:
                    try:
                        dcombo.set(value)
                    except tk.TclError:
                        pass

            if hasattr(self, "remember_var"):
                try:
                    self.remember_var.set(1 if state.remember else 0)
                except tk.TclError:
                    pass

        def _schedule_header_alignment(self, row_widgets: Dict[str, "tk.Misc"]) -> None:
            if not getattr(self, "_header_column_frames", None):
                return
            if self._header_aligned or self._header_alignment_pending:
                return

            def _do_align() -> None:
                self._align_header_columns(row_widgets)

            self._header_alignment_pending = True
            self.after_idle(_do_align)

        def _align_header_columns(self, row_widgets: Dict[str, "tk.Misc"]) -> None:
            try:
                column_widgets = [
                    row_widgets["label"],
                    row_widgets["supplier_combo"],
                    row_widgets["doc_combo"],
                    row_widgets["doc_entry"],
                    row_widgets["remark_entry"],
                    row_widgets["delivery_combo"],
                ]
            except KeyError:
                self._header_alignment_pending = False
                return

            self.update_idletasks()

            for frame, label, widget in zip(
                self._header_column_frames,
                self._header_labels,
                column_widgets,
            ):
                frame.pack_propagate(False)
                width = widget.winfo_width() or widget.winfo_reqwidth()
                height = label.winfo_reqheight() or widget.winfo_reqheight()
                frame.configure(width=width, height=height)
                label.configure(anchor="w", width=0)
                label.pack_configure(fill="x")

            self._header_aligned = True
            self._header_alignment_pending = False

        def _clear_saved_suppliers(self) -> None:
            self.db.defaults_by_production.clear()
            self.db.defaults_by_finish.clear()
            self.db.save()

            set_combo_value = getattr(
                type(self), "_set_combo_value", SupplierSelectionFrame._set_combo_value
            )
            for _sel_key, combo in self.rows:
                set_combo_value(combo, "(geen)")

            for sel_key in self.doc_vars:
                self.doc_vars[sel_key].set("Standaard bon")

            for dcombo in self.delivery_combos.values():
                dcombo.set("Geen")

            for rvar in getattr(self, "remark_vars", {}).values():
                rvar.set("")

            self._on_combo_change()

        def _on_focus_key(self, sel_key: str):
            self._active_key = sel_key

        def _display_list(self) -> List[str]:
            sups = self.db.suppliers_sorted()
            opts = [self.db.display_name(s) for s in sups]
            opts.insert(0, "(geen)")
            return opts

        @staticmethod
        def _parse_selection_key(key: str) -> tuple[str, str]:
            """Safely resolve a selection key even when helper imports are missing."""

            try:
                return parse_selection_key(key)
            except Exception:
                pass

            if "::" in key:
                prefix, identifier = key.split("::", 1)
                if prefix in ("production", "finish", "opticutter"):
                    return prefix, identifier

            return "production", key

        def _refresh_options(self, initial=False):
            self._base_options = self._display_list()
            self._disp_to_name = {}
            src = self.db.suppliers_sorted()
            for s in src:
                self._disp_to_name[self.db.display_name(s)] = s.supplier

            set_combo_value = getattr(
                type(self), "_set_combo_value", SupplierSelectionFrame._set_combo_value
            )
            for sel_key, combo in self.rows:
                typed = combo.get()
                combo["values"] = self._base_options
                parser = getattr(
                    self,
                    "_parse_selection_key",
                    SupplierSelectionFrame._parse_selection_key,
                )
                kind, identifier = parser(sel_key)
                if kind == "production":
                    lower_name = identifier.strip().lower()
                    if lower_name in ("dummy part", "nan", "spare part"):
                        set_combo_value(combo, self._base_options[0])
                        continue
                    name = self.db.get_default(identifier)
                elif kind == "opticutter":
                    name = self.db.get_default(make_opticutter_default_key(identifier))
                else:
                    name = self.db.get_default_finish(identifier)
                if typed:
                    set_combo_value(combo, typed)
                    continue
                if not name and initial:
                    favs = [x for x in src if x.favorite]
                    name = (
                        favs[0].supplier
                        if favs
                        else (src[0].supplier if src else "")
                    )
                disp = None
                for k, v in self._disp_to_name.items():
                    if v and name and v.lower() == name.lower():
                        disp = k
                        break
                if disp:
                    set_combo_value(combo, disp)
                elif self._base_options:
                    set_combo_value(
                        combo,
                        self._base_options[1]
                        if len(self._base_options) > 1
                        else self._base_options[0],
                    )

            delivery_opts = [
                "Geen",
                "Bestelling wordt opgehaald",
                "Leveradres wordt nog meegedeeld",
            ] + [
                self.delivery_db.display_name(a)
                for a in self.delivery_db.addresses_sorted()
            ]
            for sel_key, dcombo in self.delivery_combos.items():
                cur = dcombo.get()
                dcombo["values"] = delivery_opts
                if cur:
                    dcombo.set(cur)

        def _on_combo_change(self, _evt=None):
            for sel_key, combo in self.rows:
                doc_var = self.doc_vars.get(sel_key)
                if not doc_var:
                    continue
                raw_val = combo.get().strip()
                norm_val = raw_val.lower()
                if not raw_val or norm_val in ("(geen)", "geen"):
                    doc_var.set("Standaard bon")
                else:
                    doc_var.set("Bestelbon")
                self._on_doc_type_change(sel_key)
            self._update_preview_from_any_combo()

        def _on_doc_type_change(self, sel_key: str):
            doc_var = self.doc_vars.get(sel_key)
            doc_num_var = self.doc_num_vars.get(sel_key)
            if not doc_var or not doc_num_var:
                return
            cur = doc_num_var.get()
            prefix = _prefix_for_doc_type(doc_var.get())
            prefixes = getattr(self, "_doc_type_prefixes", {prefix})
            if not cur or cur in prefixes:
                doc_num_var.set(prefix)

        def _on_combo_type(self, evt, sel_key: str, combo):
            self._active_key = sel_key
            text = _norm(combo.get().strip())
            if not hasattr(self, "_base_options"):
                return
            if evt.keysym in ("Up", "Down", "Escape"):
                return
            if not text:
                combo["values"] = self._base_options
                for ch in self.cards_frame.winfo_children():
                    ch.destroy()
                self._update_preview_for_text("")
                return
            filtered = [
                opt for opt in self._base_options if _norm(opt).startswith(text)
            ]
            filtered = sort_supplier_options(
                filtered, self.db.suppliers, getattr(self, "_disp_to_name", {})
            )
            combo["values"] = filtered
            self._populate_cards(filtered, sel_key)
            if evt.keysym == "Return" and len(filtered) == 1:
                set_combo_value = getattr(
                    type(self),
                    "_set_combo_value",
                    SupplierSelectionFrame._set_combo_value,
                )
                set_combo_value(combo, filtered[0])
                self._update_preview_for_text(filtered[0])
            else:
                self._update_preview_for_text(combo.get())

        def _resolve_text_to_supplier(self, text: str) -> Optional[Supplier]:
            if not text:
                return None
            norm_text = _norm(text)
            if hasattr(self, "_disp_to_name"):
                for disp, name in self._disp_to_name.items():
                    if _norm(disp) == norm_text:
                        for s in self.db.suppliers:
                            if _norm(s.supplier) == _norm(name):
                                return s
            for s in self.db.suppliers:
                if _norm(s.supplier) == norm_text:
                    return s
            cand = [
                s for s in self.db.suppliers if _norm(s.supplier).startswith(norm_text)
            ]
            if cand:
                return sorted(cand, key=lambda x: (not x.favorite, _norm(x.supplier)))[0]
            cand = [
                s for s in self.db.suppliers if norm_text in _norm(s.supplier)
            ]
            if cand:
                return sorted(cand, key=lambda x: (not x.favorite, _norm(x.supplier)))[0]
            return None

        def _update_preview_for_text(self, text: str):
            s = self._resolve_text_to_supplier(text)
            self._preview_supplier = s

        def _update_preview_from_any_combo(self):
            for sel_key, combo in self.rows:
                t = combo.get()
                if t:
                    self._active_key = sel_key
                    self._update_preview_for_text(t)
                    self._populate_cards([t], sel_key)
                    return
            self._preview_supplier = None
            self._populate_cards([], self._active_key if self._active_key else None)

        def _on_card_click(self, option: str, sel_key: str):
            combo = self.combo_by_key.get(sel_key)
            if combo:
                set_combo_value = getattr(
                    type(self), "_set_combo_value", SupplierSelectionFrame._set_combo_value
                )
                set_combo_value(combo, option)
            self._active_key = sel_key
            self._update_preview_for_text(option)
            self._populate_cards([option], sel_key)

        def _populate_cards(self, options, sel_key):
            for ch in self.cards_frame.winfo_children():
                ch.destroy()
            if not options:
                return
            cols = 3
            for i in range(cols):
                self.cards_frame.grid_columnconfigure(i, weight=0)
            for idx, opt in enumerate(options):
                s = self._resolve_text_to_supplier(opt)
                if not s:
                    continue
                r, c = divmod(idx, cols)
                self.cards_frame.grid_rowconfigure(r, weight=0)
                border = "#444444"
                card = tk.Frame(
                    self.cards_frame,
                    highlightbackground=border,
                    highlightcolor=border,
                    highlightthickness=2,
                    cursor="hand2",
                )
                card.grid(row=r, column=c, padx=4, pady=4, sticky="w")
                widgets = []
                name_lbl = tk.Label(
                    card,
                    text=s.supplier,
                    justify="left",
                    anchor="w",
                    font=("TkDefaultFont", 10, "bold"),
                )
                name_lbl.pack(anchor="w", padx=4, pady=(4, 0))
                widgets.append(name_lbl)
                if s.description:
                    desc_lbl = tk.Label(
                        card, text=s.description, justify="left", anchor="w"
                    )
                    desc_lbl.pack(anchor="w", padx=4)
                    widgets.append(desc_lbl)
                if s.adres_1 or s.adres_2:
                    addr_line = (
                        f"{s.adres_1}, {s.adres_2}"
                        if (s.adres_1 and s.adres_2)
                        else (s.adres_1 or s.adres_2)
                    )
                    addr_lbl = tk.Label(card, text=addr_line, justify="left", anchor="w")
                    addr_lbl.pack(anchor="w", padx=4, pady=(0, 4))
                    widgets.append(addr_lbl)
                handler = lambda _e, o=opt, key=sel_key: self._on_card_click(o, key)
                card.bind("<Button-1>", handler)
                for w in widgets:
                    w.bind("<Button-1>", handler)

        def set_busy(self, busy: bool, message: Optional[str] = None) -> None:
            confirm = getattr(self, "confirm_button", None)
            cancel = getattr(self, "cancel_button", None)

            if confirm is not None:
                try:
                    confirm.configure(state="disabled" if busy else "normal")
                except tk.TclError:
                    pass

            if cancel is not None:
                try:
                    cancel.configure(state="normal")
                except tk.TclError:
                    pass

            if message is not None:
                self.status_var.set(message)

        def update_status(self, message: str) -> None:
            self.status_var.set(message)

        def _cancel(self):
            parent_app = getattr(self.master, "master", None)
            if parent_app is not None:
                try:
                    parent_app._last_supplier_selection_state = self.serialize_state()
                except Exception:
                    pass
            if self.master:
                try:
                    self.master.forget(self)
                except Exception:
                    pass
                if hasattr(self.master, "select") and hasattr(self.master.master, "main_frame"):
                    self.master.select(self.master.master.main_frame)
                if hasattr(self.master.master, "sel_frame"):
                    self.master.master.sel_frame = None
            self.destroy()

        def _confirm(self):
            """Collect selected suppliers per production and return via callback."""
            import inspect

            sel_map: Dict[str, str] = {}
            doc_map: Dict[str, str] = {}
            for sel_key, combo in self.rows:
                typed = combo.get().strip()
                if not typed or typed.lower() in ("(geen)", "geen"):
                    sel_map[sel_key] = ""
                else:
                    s = self._resolve_text_to_supplier(typed)
                    if s:
                        sel_map[sel_key] = s.supplier
                doc_var = self.doc_vars.get(sel_key)
                doc_map[sel_key] = doc_var.get() if doc_var else "Bestelbon"

            doc_num_map: Dict[str, str] = {}
            delivery_map: Dict[str, str] = {}
            remarks_map: Dict[str, str] = {}
            remark_vars = getattr(self, "remark_vars", {})
            for sel_key, _combo in self.rows:
                raw_doc_num = self.doc_num_vars[sel_key].get().strip()
                doc_type_text = doc_map.get(sel_key, "Bestelbon")
                normalized_doc_num = _normalize_doc_number(
                    raw_doc_num, doc_type_text
                )
                if normalized_doc_num != raw_doc_num:
                    self.doc_num_vars[sel_key].set(normalized_doc_num)
                doc_num_map[sel_key] = normalized_doc_num
                delivery_map[sel_key] = self.delivery_vars.get(
                    sel_key, tk.StringVar(value="Geen")
                ).get()
                remark_var = remark_vars.get(sel_key)
                remarks_map[sel_key] = remark_var.get().strip() if remark_var else ""

            project_number = self.project_number_var.get().strip()
            project_name = self.project_name_var.get().strip()

            remember_flag = bool(self.remember_var.get())
            callback = self.callback
            use_new_signature = False
            sig_params = None
            try:
                sig = inspect.signature(callback)
            except (ValueError, TypeError):
                sig = None
            if sig is not None:
                params = list(sig.parameters.values())
                if params and params[0].name == "self" and params[0].kind in (
                    inspect.Parameter.POSITIONAL_ONLY,
                    inspect.Parameter.POSITIONAL_OR_KEYWORD,
                ):
                    params = params[1:]
                sig_params = params
                if any(
                    p.kind in (inspect.Parameter.VAR_POSITIONAL, inspect.Parameter.VAR_KEYWORD)
                    for p in params
                ):
                    use_new_signature = True
                elif len(params) >= 8:
                    use_new_signature = True

            if use_new_signature:
                try:
                    callback(
                        sel_map,
                        doc_map,
                        doc_num_map,
                        delivery_map,
                        remarks_map,
                        project_number,
                        project_name,
                        remember_flag,
                    )
                    return
                except TypeError as exc:
                    msg = str(exc)
                    if not (
                        "positional" in msg
                        or "keyword" in msg
                        or (sig_params is not None and len(sig_params) >= 8)
                    ):
                        raise
                    use_new_signature = False

            if not use_new_signature:
                callback(
                    sel_map,
                    doc_map,
                    doc_num_map,
                    delivery_map,
                    project_number,
                    project_name,
                    remember_flag,
                )

    class SettingsFrame(tk.Frame):
        def __init__(self, master, app: "App"):
            super().__init__(master)
            self.app = app
            self.extensions: List[FileExtensionSetting] = deepcopy(
                self.app.settings.file_extensions
            )

            self.configure(padx=12, pady=12)
            self.columnconfigure(0, weight=1)
            self.rowconfigure(3, weight=1)

            export_options = tk.LabelFrame(
                self, text="Exportopties", labelanchor="n"
            )
            export_options.grid(row=0, column=0, sticky="ew")
            export_options.columnconfigure(0, weight=1)

            def _add_option(
                parent: tk.Widget,
                text: str,
                description: str,
                variable: "tk.IntVar",
            ) -> None:
                row = parent.grid_size()[1]
                container = tk.Frame(parent)
                container.grid(row=row, column=0, sticky="ew", padx=12, pady=(6, 2))
                container.columnconfigure(0, weight=1)
                tk.Checkbutton(
                    container,
                    text=text,
                    variable=variable,
                    anchor="w",
                    justify="left",
                ).grid(row=0, column=0, sticky="w")
                tk.Label(
                    container,
                    text=description,
                    justify="left",
                    anchor="w",
                    wraplength=520,
                    foreground="#555555",
                ).grid(row=1, column=0, sticky="ew", padx=(28, 0))

            _add_option(
                export_options,
                "Exporteer bewerkte BOM naar exportmap",
                (
                    "Bewaar automatisch een Excel-bestand van de huidige BOM in de "
                    "hoofdfolder van elke export. Alle wijzigingen die je in Filehopper "
                    "hebt aangebracht, zoals verwijderde rijen, worden meegeschreven."
                ),
                self.app.export_bom_var,
            )

            _add_option(
                export_options,
                "Exporteer gerelateerde exportbestanden naar exportmap",
                (
                    "Zoek in de geselecteerde extensies naar bestanden waarvan de naam "
                    "overeenkomt met de BOM en kopieer ze naast het BOM-bestand. "
                    "Handig voor extra documenten zoals PDF's of STEP-bestanden."
                ),
                self.app.export_related_files_var,
            )

            _add_option(
                export_options,
                "Maak snelkoppeling naar nieuwste exportmap",
                (
                    "Na het exporteren wordt er een snelkoppeling met de naam 'latest'"
                    " geplaatst in de exportmap. Deze verwijst altijd naar de"
                    " meest recente export zodat je die snel kunt openen."
                ),
                self.app.bundle_latest_var,
            )
            _add_option(
                export_options,
                "Testrun: toon alleen doelmap (niets wordt gekopieerd)",
                (
                    "Voer een proefrun uit zonder bestanden te kopiëren. Je ziet"
                    " welke doelmap gebruikt zou worden, maar er worden geen"
                    " bestanden aangemaakt of overschreven."
                ),
                self.app.bundle_dry_run_var,
            )
            _add_option(
                export_options,
                "Vul Custom BOM automatisch na het laden van de hoofd-BOM",
                (
                    "Wanneer je een BOM opent, wordt dezelfde inhoud ook in de"
                    " Custom BOM-tab geplaatst zodat je daar meteen kunt"
                    " aanpassen en bijwerken."
                ),
                self.app.autofill_custom_bom_var,
            )

            template_frame = tk.LabelFrame(
                self,
                text="BOM-template",
            )
            template_frame.grid(
                row=1,
                column=0,
                sticky="ew",
                padx=0,
                pady=(12, 0),
            )
            template_frame.columnconfigure(0, weight=1)

            tk.Label(
                template_frame,
                text=(
                    "Download een leeg Excel-sjabloon met alle kolommen van de BOM."
                    " Handig om gegevens vooraf in te vullen of te delen met collega's."
                ),
                justify="left",
                anchor="w",
                wraplength=480,
            ).grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))

            tk.Button(
                template_frame,
                text="Download BOM template",
                command=self._download_bom_template,
            ).grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))

            footer_frame = tk.LabelFrame(
                self,
                text="Bestelbon/offerte onderschrift",
                labelanchor="n",
            )
            footer_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
            footer_frame.columnconfigure(0, weight=1)
            footer_frame.rowconfigure(1, weight=1)

            tk.Label(
                footer_frame,
                text=(
                    "Pas hier het onderschrift aan dat onderaan de bestelbon of"
                    " offerteaanvraag wordt geplaatst."
                ),
                justify="left",
                anchor="w",
                wraplength=520,
            ).grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 4))

            footer_text_container = tk.Frame(footer_frame)
            footer_text_container.grid(
                row=1,
                column=0,
                sticky="nsew",
                padx=12,
                pady=(0, 4),
            )
            footer_text_container.columnconfigure(0, weight=1)
            footer_text_container.rowconfigure(0, weight=1)

            self.footer_note_text = tk.Text(
                footer_text_container,
                height=5,
                wrap="word",
            )
            self.footer_note_text.grid(row=0, column=0, sticky="nsew")

            footer_scrollbar = tk.Scrollbar(
                footer_text_container,
                orient="vertical",
                command=self.footer_note_text.yview,
            )
            footer_scrollbar.grid(row=0, column=1, sticky="ns")
            self.footer_note_text.configure(yscrollcommand=footer_scrollbar.set)
            self._reload_footer_note()

            footer_btns = tk.Frame(footer_frame)
            footer_btns.grid(row=2, column=0, sticky="e", padx=12, pady=(0, 8))
            tk.Button(footer_btns, text="Opslaan", command=self._save_footer_note).pack(
                side="left", padx=4
            )
            tk.Button(
                footer_btns,
                text="Reset naar standaard",
                command=self._reset_footer_note,
            ).pack(side="left", padx=4)

            extensions_frame = tk.LabelFrame(
                self, text="Bestandstypen", labelanchor="n"
            )
            extensions_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
            extensions_frame.columnconfigure(0, weight=1)
            extensions_frame.rowconfigure(1, weight=1)

            tk.Label(
                extensions_frame,
                text=(
                    "Beheer hier welke bestandstypen beschikbaar zijn op het hoofdscherm.\n"
                    "Voeg extensies toe of verwijder ze naar wens."
                ),
                justify="left",
                anchor="w",
                wraplength=520,
            ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(8, 4))

            list_container = tk.Frame(extensions_frame)
            list_container.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=12)
            list_container.columnconfigure(0, weight=1)
            list_container.rowconfigure(0, weight=1)

            list_frame = tk.Frame(list_container)
            list_frame.grid(row=0, column=0, sticky="nsew")
            list_frame.columnconfigure(0, weight=1)

            self.listbox = tk.Listbox(list_frame, activestyle="none")
            self.listbox.grid(row=0, column=0, sticky="nsew")
            scrollbar = tk.Scrollbar(list_frame, command=self.listbox.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")
            self.listbox.configure(yscrollcommand=scrollbar.set)
            self.listbox.bind("<Double-Button-1>", lambda _e: self._edit_selected())

            move_btns = tk.Frame(list_container)
            move_btns.grid(row=0, column=1, sticky="ns", padx=(8, 0))
            move_btns.grid_rowconfigure(0, weight=1)
            move_btns.grid_rowconfigure(3, weight=1)
            move_btns.grid_columnconfigure(0, weight=1)
            tk.Button(
                move_btns,
                text="▲",
                width=3,
                command=lambda: self._move_selected(-1),
            ).grid(row=1, column=0, pady=2, sticky="nsew")
            tk.Button(
                move_btns,
                text="▼",
                width=3,
                command=lambda: self._move_selected(1),
            ).grid(row=2, column=0, pady=2, sticky="nsew")

            btns = tk.Frame(extensions_frame)
            btns.grid(row=2, column=0, columnspan=2, sticky="ew", padx=12, pady=(8, 12))
            tk.Button(btns, text="Toevoegen", command=self._add_extension).pack(
                side="left", padx=4
            )
            tk.Button(btns, text="Bewerken", command=self._edit_selected).pack(
                side="left", padx=4
            )
            tk.Button(btns, text="Verwijderen", command=self._remove_selected).pack(
                side="left", padx=4
            )

            self._refresh_list()

        def _refresh_list(self) -> None:
            self.listbox.delete(0, tk.END)
            if not self.extensions:
                self.listbox.insert(0, "Geen bestandstypen gedefinieerd.")
                self.listbox.itemconfig(0, foreground="#777777")
                self._update_listbox_height(1)
                self._update_listbox_width()
                return
            for ext in self.extensions:
                status = "✓" if ext.enabled else "✗"
                patterns = ", ".join(ext.patterns)
                self.listbox.insert(tk.END, f"{status} {ext.label} — {patterns}")
            self._update_listbox_height(len(self.extensions))
            self._update_listbox_width()

        def _reload_footer_note(self) -> None:
            text = self.app.footer_note_var.get()
            self.footer_note_text.delete("1.0", "end")
            if text:
                self.footer_note_text.insert("1.0", text)

        def _current_footer_text(self) -> str:
            raw = self.footer_note_text.get("1.0", "end-1c")
            return raw.replace("\r\n", "\n")

        def _save_footer_note(self) -> None:
            note = self._current_footer_text().strip()
            self.app.update_footer_note(note)
            self._reload_footer_note()

        def _reset_footer_note(self) -> None:
            self.app.update_footer_note(DEFAULT_FOOTER_NOTE)
            self._reload_footer_note()

        def _download_bom_template(self) -> None:
            path_str = filedialog.asksaveasfilename(
                parent=self,
                title="BOM-template opslaan",
                defaultextension=".xlsx",
                filetypes=(("Excel-werkboek", "*.xlsx"), ("Alle bestanden", "*.*")),
                initialfile=BOMCustomTab.default_template_filename(),
            )
            if not path_str:
                return

            target_path = Path(path_str)
            try:
                BOMCustomTab.write_template_workbook(target_path)
            except Exception as exc:
                messagebox.showerror("Opslaan mislukt", str(exc), parent=self)
                return

            messagebox.showinfo(
                "Template opgeslagen",
                (
                    "Het lege BOM-sjabloon is opgeslagen. Vul het formulier in en"
                    " importeer de gegevens later in de Custom BOM-tab.\n\n"
                    f"Locatie: {target_path}"
                ),
                parent=self,
            )

        def _update_listbox_height(self, item_count: int) -> None:
            visible = max(1, item_count)
            height = min(max(visible, 3), 10)
            self.listbox.configure(height=height)

        def _update_listbox_width(self) -> None:
            items = self.listbox.get(0, tk.END)
            if not items:
                self.listbox.configure(width=28)
                return
            max_len = max(len(item) for item in items)
            width = max(28, min(64, max_len + 4))
            self.listbox.configure(width=width)

        def _selected_index(self) -> Optional[int]:
            if not self.extensions:
                return None
            sel = self.listbox.curselection()
            if not sel:
                return None
            idx = int(sel[0])
            if idx >= len(self.extensions):
                return None
            return idx

        def _selected_extension(self) -> Optional[FileExtensionSetting]:
            idx = self._selected_index()
            if idx is None:
                return None
            return self.extensions[idx]

        def _ensure_unique_key(self, key: str, exclude_index: Optional[int] = None) -> str:
            existing = {
                ext.key
                for idx, ext in enumerate(self.extensions)
                if exclude_index is None or idx != exclude_index
            }
            if key not in existing:
                return key
            base = key
            suffix = 2
            while True:
                candidate = f"{base}_{suffix}"
                if candidate not in existing:
                    return candidate
                suffix += 1

        def _persist(self) -> None:
            self.app.apply_file_extensions(deepcopy(self.extensions))
            self.extensions = deepcopy(self.app.settings.file_extensions)
            self._refresh_list()

        def _add_extension(self) -> None:
            self._open_extension_dialog("Bestandstype toevoegen", None)

        def _edit_selected(self) -> None:
            ext = self._selected_extension()
            if ext is None:
                return
            self._open_extension_dialog("Bestandstype bewerken", ext)

        def _remove_selected(self) -> None:
            idx = self._selected_index()
            if idx is None:
                return
            ext = self.extensions[idx]
            if not messagebox.askyesno(
                "Bevestigen",
                f"Verwijder '{ext.label}' van de lijst?",
                parent=self,
            ):
                return
            del self.extensions[idx]
            self._persist()

        def _move_selected(self, offset: int) -> None:
            idx = self._selected_index()
            if idx is None:
                return
            new_idx = idx + offset
            if new_idx < 0 or new_idx >= len(self.extensions):
                return
            self.extensions[idx], self.extensions[new_idx] = (
                self.extensions[new_idx],
                self.extensions[idx],
            )
            self._persist()
            if 0 <= new_idx < len(self.extensions):
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(new_idx)
                self.listbox.activate(new_idx)
                self.listbox.see(new_idx)

        def _open_extension_dialog(
            self, title: str, existing: Optional[FileExtensionSetting]
        ) -> None:
            win = tk.Toplevel(self)
            win.title(title)
            win.transient(self)
            _place_window_near_parent(win, self)
            win.grab_set()

            def _normalize_extensions(values) -> List[str]:
                cleaned: List[str] = []
                seen = set()
                for raw in values:
                    if not isinstance(raw, str):
                        continue
                    ext = raw.strip().lower()
                    if not ext:
                        continue
                    ext = ext.lstrip(".")
                    if not ext or ext in seen:
                        continue
                    cleaned.append(ext)
                    seen.add(ext)
                return cleaned

            tk.Label(win, text="Naam:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
            name_var = tk.StringVar(value=existing.label if existing else "")
            tk.Entry(win, textvariable=name_var, width=40).grid(
                row=0, column=1, padx=4, pady=4
            )

            tk.Label(win, text="Extensies (komma of spatie gescheiden):").grid(
                row=1, column=0, sticky="e", padx=4, pady=4
            )
            patterns_text = ", ".join(existing.patterns) if existing else ""
            patterns_var = tk.StringVar(value=patterns_text)
            tk.Entry(win, textvariable=patterns_var, width=28).grid(
                row=1, column=1, padx=(4, 12), pady=(4, 8)
            )

            tk.Label(win, text="Preset:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
            no_preset_label = "(Geen preset)"
            preset_choices = [no_preset_label, *FILE_EXTENSION_PRESETS.keys()]
            preset_var = tk.StringVar(value=no_preset_label)
            preset_combo = ttk.Combobox(
                win,
                textvariable=preset_var,
                values=preset_choices,
                state="readonly",
                width=32,
            )
            preset_combo.grid(row=2, column=1, sticky="we", padx=4, pady=4)
            preset_info_var = tk.StringVar(value="Selecteer een preset")
            tk.Label(win, textvariable=preset_info_var, anchor="w").grid(
                row=2, column=2, sticky="w", padx=(4, 0), pady=4
            )

            enabled_var = tk.BooleanVar(value=existing.enabled if existing else True)
            tk.Checkbutton(
                win,
                text="Standaard aangevinkt",
                variable=enabled_var,
            ).grid(row=3, column=1, sticky="w", padx=4, pady=4)

            def _update_preset_info(name: str) -> None:
                if name in FILE_EXTENSION_PRESETS:
                    count = len(_normalize_extensions(FILE_EXTENSION_PRESETS[name]))
                    suffix = "s" if count != 1 else ""
                    preset_info_var.set(f"Preset bevat {count} extensie{suffix}")
                else:
                    preset_info_var.set("Selecteer een preset")

            def _on_preset_selected(_event=None) -> None:
                name = preset_var.get()
                if name in FILE_EXTENSION_PRESETS:
                    normalized = _normalize_extensions(FILE_EXTENSION_PRESETS[name])
                    if normalized:
                        patterns_var.set(", ".join(f".{ext}" for ext in normalized))
                        if existing is None or not name_var.get().strip():
                            name_var.set(name)
                _update_preset_info(name)

            preset_combo.bind("<<ComboboxSelected>>", _on_preset_selected)

            def _save() -> None:
                try:
                    new_ext = FileExtensionSetting.from_user_input(
                        name_var.get(),
                        patterns_var.get(),
                        enabled_var.get(),
                        key=existing.key if existing else None,
                    )
                except ValueError as exc:
                    messagebox.showerror("Fout", str(exc), parent=win)
                    return
                if existing is None:
                    new_ext.key = self._ensure_unique_key(new_ext.key)
                    self.extensions.append(new_ext)
                else:
                    idx = self.extensions.index(existing)
                    new_ext.key = self._ensure_unique_key(new_ext.key, exclude_index=idx)
                    self.extensions[idx] = new_ext
                self._persist()
                win.destroy()

            if existing:
                existing_norm = set(_normalize_extensions(existing.patterns))
                for preset_name, preset_exts in FILE_EXTENSION_PRESETS.items():
                    if existing_norm == set(_normalize_extensions(preset_exts)):
                        preset_var.set(preset_name)
                        break

            preset_combo.set(preset_var.get())
            _update_preset_info(preset_var.get())

            btns = tk.Frame(win)
            btns.grid(row=4, column=0, columnspan=3, pady=(8, 4))
            tk.Button(btns, text="Opslaan", command=_save).pack(side="left", padx=4)
            tk.Button(btns, text="Annuleer", command=win.destroy).pack(
                side="left", padx=4
            )

            win.columnconfigure(1, weight=1)
            win.columnconfigure(2, weight=1)
            name_var.set(name_var.get())
            win.resizable(False, False)
            win.wait_visibility()
            win.focus_set()
            win.wait_window()

    class App(tk.Tk):
        _CUSTOM_ROW_ATTR = "_custom_row_flags"
        def __init__(self):
            super().__init__()
            import sys
            style = ttk.Style(self)
            if sys.platform == "darwin":
                style.theme_use("aqua")
            else:
                style.theme_use("clam")

            def _configure_tab_like_button_style():
                tab_layout = deepcopy(style.layout("TNotebook.Tab"))

                def _remove_focus(layout_items):
                    cleaned = []
                    for child_element, child_options in layout_items:
                        if child_element == "Notebook.focus":
                            if child_options:
                                children = child_options.get("children")
                                if children:
                                    cleaned.extend(_remove_focus(children))
                            continue
                        if child_options:
                            new_child_options = deepcopy(child_options)
                            if "children" in new_child_options:
                                new_child_options["children"] = _remove_focus(
                                    new_child_options["children"]
                                )
                            cleaned.append((child_element, new_child_options))
                        else:
                            cleaned.append((child_element, child_options))
                    return cleaned

                if tab_layout:
                    cleaned_layout = _remove_focus(tab_layout)
                    style.layout("TNotebook.Tab", cleaned_layout)
                    style.layout("Tab.TButton", cleaned_layout)

                tab_config = {}
                for opt in ("padding", "background", "foreground", "font", "borderwidth", "relief"):
                    val = style.lookup("TNotebook.Tab", opt)
                    if val not in (None, ""):
                        tab_config[opt] = val
                if tab_config:
                    style.configure("Tab.TButton", **tab_config)

                for opt in ("background", "foreground", "bordercolor", "focuscolor", "lightcolor", "darkcolor"):
                    states = style.map("TNotebook.Tab", opt)
                    if states:
                        style.map("Tab.TButton", **{opt: states})

                padding = style.lookup("TNotebook.Tab", "padding")
                if padding in (None, ""):
                    style.configure("Tab.TButton", padding=(12, 4))

            _configure_tab_like_button_style()
            self.title("Filehopper")
            self.minsize(1024, 720)

            self.db = SuppliersDB.load(SUPPLIERS_DB_FILE)
            self.client_db = ClientsDB.load(CLIENTS_DB_FILE)
            self.delivery_db = DeliveryAddressesDB.load(DELIVERY_DB_FILE)

            self.settings = AppSettings.load()
            self._suspend_save = False

            self.source_folder_var = tk.StringVar(
                master=self, value=self.settings.source_folder
            )
            self.dest_folder_var = tk.StringVar(
                master=self, value=self.settings.dest_folder
            )
            self.project_number_var = tk.StringVar(
                master=self, value=self.settings.project_number
            )
            self.project_name_var = tk.StringVar(
                master=self, value=self.settings.project_name
            )
            self.extension_vars: Dict[str, tk.IntVar] = {}
            self._sync_extension_vars_from_settings()
            self.zip_var = tk.IntVar(
                master=self, value=1 if self.settings.zip_per_production else 0
            )
            self.combine_pdf_per_production_var = tk.IntVar(
                master=self, value=1 if self.settings.combine_pdf_per_production else 0
            )
            self.finish_export_var = tk.IntVar(
                master=self, value=1 if self.settings.copy_finish_exports else 0
            )
            self.zip_finish_var = tk.IntVar(
                master=self, value=1 if self.settings.zip_finish_exports else 0
            )
            self.export_bom_var = tk.IntVar(
                master=self, value=1 if self.settings.export_processed_bom else 0
            )
            self.export_related_files_var = tk.IntVar(
                master=self,
                value=1 if self.settings.export_related_bom_files else 0,
            )
            self.zip_per_finish_var = tk.IntVar(
                master=self,
                value=
                1
                if self.settings.zip_per_production
                and self.settings.zip_finish_exports
                else 0,
            )
            self.export_date_prefix_var = tk.IntVar(
                master=self, value=1 if self.settings.export_date_prefix else 0
            )
            self.export_date_suffix_var = tk.IntVar(
                master=self, value=1 if self.settings.export_date_suffix else 0
            )
            self.export_name_custom_prefix_text = tk.StringVar(
                master=self, value=self.settings.custom_prefix_text
            )
            self.export_name_custom_prefix_enabled_var = tk.IntVar(
                master=self, value=1 if self.settings.custom_prefix_enabled else 0
            )
            self.export_name_custom_suffix_text = tk.StringVar(
                master=self, value=self.settings.custom_suffix_text
            )
            self.export_name_custom_suffix_enabled_var = tk.IntVar(
                master=self, value=1 if self.settings.custom_suffix_enabled else 0
            )
            self.bundle_latest_var = tk.IntVar(
                master=self, value=1 if self.settings.bundle_latest else 0
            )
            self.bundle_dry_run_var = tk.IntVar(
                master=self, value=1 if self.settings.bundle_dry_run else 0
            )
            self.autofill_custom_bom_var = tk.IntVar(
                master=self, value=1 if self.settings.autofill_custom_bom else 0
            )
            self.footer_note_var = tk.StringVar(
                master=self, value=self.settings.footer_note or ""
            )

            self.source_folder = self.source_folder_var.get().strip()
            self.dest_folder = self.dest_folder_var.get().strip()
            self.last_bundle_result: Optional[ExportBundleResult] = None
            self.bom_df: Optional["pd.DataFrame"] = None
            self.bom_source_path: Optional[str] = None
            self.sel_frame: Optional["SupplierSelectionFrame"] = None
            self._last_supplier_selection_state: Optional[SupplierSelectionState] = None

            for var in (
                self.source_folder_var,
                self.dest_folder_var,
                self.project_number_var,
                self.project_name_var,
                self.export_name_custom_prefix_text,
                self.export_name_custom_suffix_text,
            ):
                var.trace_add("write", self._save_settings)
            for var in (
                self.zip_var,
                self.combine_pdf_per_production_var,
                self.finish_export_var,
                self.zip_finish_var,
                self.export_bom_var,
                self.export_related_files_var,
                self.export_date_prefix_var,
                self.export_date_suffix_var,
                self.export_name_custom_prefix_enabled_var,
                self.export_name_custom_suffix_enabled_var,
                self.bundle_latest_var,
                self.bundle_dry_run_var,
                self.autofill_custom_bom_var,
            ):
                var.trace_add("write", self._save_settings)

            self.zip_var.trace_add("write", self._update_zip_per_finish_var)
            self.zip_finish_var.trace_add("write", self._update_zip_per_finish_var)
            self._update_zip_per_finish_var()

            tabs_wrapper = tk.Frame(self)
            tabs_wrapper.pack(fill="both", expand=True, padx=8, pady=(12, 0))

            tabs_background = (
                style.lookup("TNotebook", "background")
                or style.lookup("TFrame", "background")
                or self.cget("background")
            )

            tabs_container = tk.Frame(tabs_wrapper, background=tabs_background)
            tabs_container.pack(fill="both", expand=True)

            self.nb = ttk.Notebook(tabs_container)
            self.nb.pack(fill="both", expand=True)
            self.custom_bom_tab = BOMCustomTab(
                self.nb,
                app_name="Filehopper",
                on_custom_bom_ready=self._on_custom_bom_ready,
                on_push_to_main=self._apply_custom_bom_to_main,
                event_target=self,
            )
            main = tk.Frame(self.nb)
            main.configure(padx=12, pady=12)
            self.nb.add(main, text="Main")
            self.nb.add(self.custom_bom_tab, text="Custom BOM")
            self.opticutter_frame = tk.Frame(self.nb)
            self.opticutter_frame.configure(padx=12, pady=12)
            self.nb.add(self.opticutter_frame, text="Opticutter")

            opticutter_header = tk.Frame(self.opticutter_frame)
            opticutter_header.pack(fill="x", pady=(0, 8))

            header_text = tk.Label(
                opticutter_header,
                text=(
                    "Gebruik deze zaagoptimalisatie om lineaire materialen zoals "
                    "balken, buizen of profielen zo efficiënt mogelijk te verdelen. "
                    "Geef de gewenste stukken door en stel indien nodig de "
                    "zaagbreedte in; het algoritme berekent automatisch het meest "
                    "gunstige zaagplan."
                ),
                justify="left",
                anchor="w",
                wraplength=520,
                font=tkfont.nametofont("TkDefaultFont"),
            )
            header_text.pack(side="left", fill="both", expand=True, padx=(0, 12))

            controls_frame = tk.Frame(opticutter_header)
            controls_frame.pack(side="right", anchor="ne")

            self.opticutter_kerf_var = tk.StringVar(
                master=self.opticutter_frame,
                value=f"{DEFAULT_KERF_MM:g}",
            )
            kerf_frame = tk.Frame(controls_frame)
            kerf_frame.pack(anchor="e")
            tk.Label(kerf_frame, text="Zaagbreedte (mm):").pack(side="left", padx=(0, 6))
            kerf_entry = ttk.Entry(
                kerf_frame,
                textvariable=self.opticutter_kerf_var,
                width=8,
                justify="right",
            )
            kerf_entry.pack(side="right")

            self.opticutter_custom_stock_var = tk.StringVar(master=self.opticutter_frame)
            custom_frame = tk.Frame(controls_frame)
            custom_frame.pack(anchor="e", pady=(6, 0))
            tk.Label(custom_frame, text="Custom stock lengte:").pack(
                side="left", padx=(0, 6)
            )
            custom_entry = ttk.Entry(
                custom_frame,
                textvariable=self.opticutter_custom_stock_var,
                width=12,
                justify="right",
            )
            custom_entry.pack(side="right")

            self._opticutter_refresh_after_id: Optional[str] = None
            self.opticutter_kerf_var.trace_add(
                "write", self._on_opticutter_kerf_change
            )
            self.opticutter_custom_stock_var.trace_add(
                "write", self._on_opticutter_custom_stock_change
            )

            self.opticutter_info_var = tk.StringVar(
                master=self.opticutter_frame,
                value="Laad een BOM om profielen te bekijken.",
            )
            tk.Label(
                self.opticutter_frame,
                textvariable=self.opticutter_info_var,
                anchor="w",
                justify="left",
                font=tkfont.nametofont("TkDefaultFont"),
            ).pack(fill="x", pady=(0, 12))

            opticutter_table_container = tk.Frame(self.opticutter_frame)
            opticutter_table_container.pack(fill="both", expand=True, pady=(0, 8))

            opticutter_left_frame = tk.Frame(opticutter_table_container)
            opticutter_left_frame.pack(
                side="left", fill="both", expand=True, padx=(0, 12)
            )

            opticutter_columns = (
                "PartNumber",
                "Profile",
                "Material",
                "Production",
                "Profile length",
                "QTY.",
            )
            self.opticutter_tree = ttk.Treeview(
                opticutter_left_frame,
                columns=opticutter_columns,
                show="headings",
                selectmode="browse",
            )
            for col in opticutter_columns:
                anchor = "center" if col == "QTY." else "w"
                self.opticutter_tree.heading(col, text=col, anchor=anchor)
                minwidth = 40
                if col in {"Material", "Production"}:
                    minwidth = 120
                elif col == "Profile length":
                    minwidth = 110
                self.opticutter_tree.column(
                    col,
                    anchor=anchor,
                    stretch=False,
                    minwidth=minwidth,
                )

            opticutter_scroll = ttk.Scrollbar(
                opticutter_left_frame,
                orient="vertical",
                command=self.opticutter_tree.yview,
            )
            self.opticutter_tree.configure(yscrollcommand=opticutter_scroll.set)
            self.opticutter_tree.pack(
                side="left", fill="both", expand=True, anchor="w", padx=(0, 4)
            )
            opticutter_scroll.pack(side="left", fill="y")

            opticutter_summary_frame = tk.Frame(opticutter_table_container)
            opticutter_summary_frame.pack(side="left", fill="y")

            summary_common_columns = ("Profile", "Material", "Production")
            summary_metric_columns = ("Bars", "Waste", "Cuts")
            summary_headings = {
                "Profile": "Profiel",
                "Material": "Materiaal",
                "Production": "Productie",
                "Bars": "Staven",
                "Waste": "Afval",
                "Cuts": "Zaagsneden",
            }

            self.opticutter_profile_summary_base_tree: Optional["ttk.Treeview"] = None
            self.opticutter_profile_summary_trees: Dict[str, "ttk.Treeview"] = {}
            self.opticutter_summary_tooltips: Dict[str, _TreeTooltipManager] = {}
            self.opticutter_summary_column_map: Dict[str, Dict[str, str]] = {}
            self.opticutter_summary_frames: Dict[str, "tk.LabelFrame"] = {}

            summary_container = tk.Frame(opticutter_summary_frame)
            summary_container.pack(fill="both", expand=True)

            base_frame = tk.LabelFrame(
                summary_container, text="Profiel, materiaal en productie"
            )
            base_frame.pack(side="left", fill="both", padx=(0, 8))

            base_tree = ttk.Treeview(
                base_frame,
                columns=summary_common_columns,
                show="headings",
                selectmode="none",
                height=8,
            )
            for col in summary_common_columns:
                anchor = "w"
                minwidth = 170 if col == "Profile" else 140
                base_tree.heading(col, text=summary_headings[col], anchor=anchor)
                base_tree.column(
                    col,
                    anchor=anchor,
                    stretch=False,
                    minwidth=minwidth,
                )

            base_scrollbar = ttk.Scrollbar(
                base_frame, orient="vertical", command=base_tree.yview
            )
            base_tree.configure(yscrollcommand=base_scrollbar.set)
            base_tree.pack(side="left", fill="both", expand=True, padx=(0, 4))
            base_scrollbar.pack(side="left", fill="y")

            self.opticutter_profile_summary_base_tree = base_tree

            scenarios_frame = tk.Frame(summary_container)
            scenarios_frame.pack(side="left", fill="both", expand=True)

            summary_sections = [
                ("6m", "6000 mm parameters"),
                ("12m", "12000 mm parameters"),
                ("custom", "Aangepaste lengte"),
            ]

            for index, (section_key, section_title) in enumerate(summary_sections):
                section_frame = tk.LabelFrame(scenarios_frame, text=section_title)
                padx = (0, 8) if index < len(summary_sections) - 1 else (0, 0)
                section_frame.pack(side="left", fill="both", expand=True, padx=padx)

                tree = ttk.Treeview(
                    section_frame,
                    columns=summary_metric_columns,
                    show="headings",
                    selectmode="none",
                    height=8,
                )
                for col in summary_metric_columns:
                    anchor = "center"
                    minwidth = 100 if col == "Waste" else 90
                    tree.heading(col, text=summary_headings[col], anchor=anchor)
                    tree.column(
                        col,
                        anchor=anchor,
                        stretch=False,
                        minwidth=minwidth,
                    )

                scrollbar = ttk.Scrollbar(
                    section_frame, orient="vertical", command=tree.yview
                )
                tree.configure(yscrollcommand=scrollbar.set)
                tree.pack(side="left", fill="both", expand=True, padx=(0, 4))
                scrollbar.pack(side="left", fill="y")

                self.opticutter_profile_summary_trees[section_key] = tree
                self.opticutter_summary_tooltips[section_key] = _TreeTooltipManager(tree)
                self.opticutter_summary_column_map[section_key] = {
                    name: f"#{idx + 1}" for idx, name in enumerate(summary_metric_columns)
                }
                self.opticutter_summary_frames[section_key] = section_frame

            selection_section = tk.LabelFrame(
                self.opticutter_frame, text="Lengte selectie per profiel"
            )
            selection_section.pack(fill="both", expand=True, pady=(8, 0))

            tk.Label(
                selection_section,
                text=(
                    "Kies per profiel welke staaflengte je wilt gebruiken."
                    " De tabel hierboven toont alle scenario's."
                ),
                anchor="w",
                justify="left",
                wraplength=620,
            ).pack(fill="x", padx=8, pady=(8, 6))

            selection_header = tk.Frame(selection_section)
            selection_header.pack(fill="x", padx=8, pady=(0, 2))
            header_font = ("TkDefaultFont", 10, "bold")
            tk.Label(
                selection_header,
                text="Profiel",
                anchor="w",
                width=30,
                font=header_font,
            ).pack(side="left", padx=(0, 6))
            tk.Label(
                selection_header,
                text="Materiaal",
                anchor="w",
                width=20,
                font=header_font,
            ).pack(side="left", padx=(0, 6))
            tk.Label(
                selection_header,
                text="Productie",
                anchor="w",
                width=20,
                font=header_font,
            ).pack(side="left", padx=(0, 6))
            tk.Label(
                selection_header,
                text="Gewenste lengte",
                anchor="w",
                font=header_font,
            ).pack(side="left", fill="x", expand=True)

            selection_body = tk.Frame(selection_section)
            selection_body.pack(fill="both", expand=True, pady=(0, 8))
            self.opticutter_selection_canvas = tk.Canvas(
                selection_body, highlightthickness=0, borderwidth=0, height=200
            )
            self.opticutter_selection_canvas.pack(
                # Keep the column content aligned with the header underline.
                side="left", fill="both", expand=True, padx=0
            )
            self.opticutter_selection_scroll = ttk.Scrollbar(
                selection_body,
                orient="vertical",
                command=self.opticutter_selection_canvas.yview,
            )
            self.opticutter_selection_scroll.pack(side="left", fill="y", padx=(0, 8))
            self.opticutter_selection_canvas.configure(
                yscrollcommand=self.opticutter_selection_scroll.set
            )
            self.opticutter_selection_inner = tk.Frame(
                self.opticutter_selection_canvas
            )
            self.opticutter_selection_window = self.opticutter_selection_canvas.create_window(
                (0, 0), window=self.opticutter_selection_inner, anchor="nw"
            )
            self.opticutter_selection_inner.bind(
                "<Configure>",
                lambda _e: self.opticutter_selection_canvas.configure(
                    scrollregion=self.opticutter_selection_canvas.bbox("all")
                ),
            )
            self.opticutter_selection_canvas.bind(
                "<Configure>",
                lambda e: self.opticutter_selection_canvas.itemconfigure(
                    self.opticutter_selection_window, width=e.width
                ),
            )

            self.opticutter_selection_empty_label = tk.Label(
                self.opticutter_selection_inner,
                text="Laad een BOM om profielselecties te maken.",
                anchor="w",
                justify="left",
            )
            self.opticutter_selection_empty_label.pack(fill="x", padx=8, pady=8)

            self.opticutter_profile_selection_rows: Dict[
                tuple[str, str, str], tk.Frame
            ] = {}
            self.opticutter_profile_selection_labels: Dict[
                tuple[str, str, str], tuple[tk.Label, tk.Label, tk.Label]
            ] = {}
            self.opticutter_profile_selection_vars: Dict[
                tuple[str, str, str], tk.StringVar
            ] = {}
            self.opticutter_profile_selection_combos: Dict[
                tuple[str, str, str], ttk.Combobox
            ] = {}
            self.opticutter_profile_selection_display_map: Dict[
                tuple[str, str, str], OrderedDict[str, str]
            ] = {}
            self.opticutter_profile_selection_value_by_display: Dict[
                tuple[str, str, str], Dict[str, str]
            ] = {}
            self.opticutter_profile_selection_choice: Dict[
                tuple[str, str, str], str
            ] = {}
            self.opticutter_profile_selection_scenarios: Dict[
                tuple[str, str, str], Dict[str, "StockScenarioResult"]
            ] = {}
            self.opticutter_profile_custom_lengths: Dict[
                tuple[str, str, str], int
            ] = {}
            self._opticutter_selection_update_in_progress = False
            self.main_frame = main
            self.clients_frame = ClientsManagerFrame(
                self.nb, self.client_db, on_change=self._on_db_change
            )
            self.clients_frame.configure(padx=12, pady=12)
            self.nb.add(self.clients_frame, text="Klant beheer")
            self.delivery_frame = DeliveryAddressesManagerFrame(
                self.nb, self.delivery_db, on_change=self._on_db_change
            )
            self.delivery_frame.configure(padx=12, pady=12)
            self.nb.add(self.delivery_frame, text="Leveradres beheer")
            self.suppliers_frame = SuppliersManagerFrame(
                self.nb, self.db, on_change=self._on_db_change
            )
            self.suppliers_frame.configure(padx=12, pady=12)
            self.nb.add(self.suppliers_frame, text="Leverancier beheer")

            self.settings_frame = SettingsFrame(self.nb, self)
            self.settings_frame.configure(padx=12, pady=12)
            self.nb.add(self.settings_frame, text="⚙ Settings")

            # Top folders
            top = tk.Frame(main); top.pack(fill="x", padx=8, pady=6)
            FOLDER_ICON = "\U0001F4C1"
            USER_ICON = "\U0001F464"
            label_font = tkfont.nametofont("TkDefaultFont")

            tk.Label(top, text=f"{FOLDER_ICON} Bronmap:", font=label_font).grid(
                row=0, column=0, sticky="w"
            )
            self.src_entry = tk.Entry(top, width=60, textvariable=self.source_folder_var)
            self.src_entry.grid(row=0, column=1, padx=4)
            _scroll_entry_to_end(self.src_entry, self.source_folder_var)
            _OverflowTooltip(self.src_entry, lambda: self.source_folder_var.get().strip())
            tk.Button(top, text="Bladeren", command=self._pick_src).grid(row=0, column=2, padx=4)
            tk.Label(top, text="Projectnr.:").grid(row=0, column=3, sticky="w", padx=(16, 0))
            tk.Entry(top, textvariable=self.project_number_var, width=60).grid(
                row=0, column=4, padx=4, sticky="w"
            )

            tk.Label(top, text=f"{FOLDER_ICON} Bestemmingsmap:", font=label_font).grid(
                row=1, column=0, sticky="w"
            )
            self.dst_entry = tk.Entry(top, width=60, textvariable=self.dest_folder_var)
            self.dst_entry.grid(row=1, column=1, padx=4)
            _scroll_entry_to_end(self.dst_entry, self.dest_folder_var)
            _OverflowTooltip(self.dst_entry, lambda: self.dest_folder_var.get().strip())
            tk.Button(top, text="Bladeren", command=self._pick_dst).grid(row=1, column=2, padx=4)
            tk.Label(top, text="Projectnaam:").grid(row=1, column=3, sticky="w", padx=(16, 0))
            tk.Entry(top, textvariable=self.project_name_var, width=60).grid(
                row=1, column=4, padx=4, sticky="w"
            )

            top.grid_columnconfigure(5, weight=1)
            tk.Button(
                top,
                text="Leegmaken",
                command=self._clear_main_inputs,
            ).grid(row=0, column=5, rowspan=2, sticky="ne", padx=(16, 0))

            tk.Label(top, text=f"{USER_ICON} Opdrachtgever:", font=label_font).grid(
                row=2, column=0, sticky="w", pady=(8, 0)
            )
            self.client_var = tk.StringVar()
            self.client_combo = ttk.Combobox(
                top, textvariable=self.client_var, state="readonly", width=40
            )
            self.client_combo.grid(row=2, column=1, padx=4, pady=(8, 0))
            tk.Button(top, text="Beheer", command=lambda: self.nb.select(self.clients_frame)).grid(
                row=2, column=2, padx=4, pady=(8, 0)
            )
            self._refresh_clients_combo()



            # Filters
            filters_row = tk.Frame(main)
            filters_row.pack(fill="x", padx=8, pady=6)
            filters_row.grid_columnconfigure(0, weight=1)
            filters_row.grid_columnconfigure(1, weight=1)
            filters_row.grid_columnconfigure(2, weight=1)

            filt = tk.LabelFrame(
                filters_row,
                text="Selecteer bestandstypen om te kopiëren",
                labelanchor="n",
            )
            filt.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
            filt.grid_columnconfigure(0, weight=1)

            options_frame_parent = tk.LabelFrame(
                filters_row, text="Geavanceerde opties", labelanchor="n"
            )
            options_frame_parent.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
            options_frame_parent.grid_columnconfigure(0, weight=1)

            export_name_frame = tk.LabelFrame(
                filters_row,
                text="Benaming exportbestand",
                labelanchor="n",
            )
            export_name_frame.grid(row=0, column=2, sticky="nsew")
            export_name_frame.grid_columnconfigure(0, weight=1)

            self.ext_frame = tk.Frame(filt)
            self.ext_frame.grid(row=0, column=0, sticky="nw", padx=8, pady=4)
            options_frame = tk.Frame(options_frame_parent)
            options_frame.grid(row=0, column=0, sticky="nw", padx=8, pady=4)
            export_name_inner = tk.Frame(export_name_frame)
            export_name_inner.grid(row=0, column=0, sticky="nw", padx=8, pady=4)

            self._rebuild_extension_checkbuttons()
            tk.Checkbutton(
                options_frame,
                text="Zip per productie/finish",
                variable=self.zip_per_finish_var,
                anchor="w",
                command=self._toggle_zip_per_finish,
            ).pack(anchor="w", pady=2)
            tk.Checkbutton(
                options_frame,
                text="Finish export",
                variable=self.finish_export_var,
                anchor="w",
            ).pack(anchor="w", pady=2)
            tk.Checkbutton(
                options_frame,
                text="Combineer pdf per productie (uit = één PDF)",
                variable=self.combine_pdf_per_production_var,
                anchor="w",
            ).pack(anchor="w", pady=2)
            tk.Checkbutton(
                export_name_inner,
                text="Datumprefix (YYYYMMDD-)",
                variable=self.export_date_prefix_var,
                anchor="w",
            ).pack(anchor="w", pady=2)
            tk.Checkbutton(
                export_name_inner,
                text="Datumsuffix (-YYYYMMDD)",
                variable=self.export_date_suffix_var,
                anchor="w",
            ).pack(anchor="w", pady=2)
            prefix_row = tk.Frame(export_name_inner)
            prefix_row.pack(anchor="w", fill="x", pady=(8, 2))
            tk.Checkbutton(
                prefix_row,
                text="Aangepaste prefix",
                variable=self.export_name_custom_prefix_enabled_var,
            ).pack(side="left", padx=(0, 4))
            tk.Entry(
                prefix_row,
                textvariable=self.export_name_custom_prefix_text,
            ).pack(side="left", fill="x", expand=True)
            suffix_row = tk.Frame(export_name_inner)
            suffix_row.pack(anchor="w", fill="x", pady=2)
            tk.Checkbutton(
                suffix_row,
                text="Aangepaste suffix",
                variable=self.export_name_custom_suffix_enabled_var,
            ).pack(side="left", padx=(0, 4))
            tk.Entry(
                suffix_row,
                textvariable=self.export_name_custom_suffix_text,
            ).pack(side="left", fill="x", expand=True)
            # Legacy options moved to settings tab

            # BOM controls
            bf = tk.Frame(main); bf.pack(fill="x", padx=8, pady=6)
            tk.Button(bf, text="Laad BOM (CSV/Excel)", command=self._load_bom).pack(side="left", padx=6)
            tk.Button(
                bf,
                text="Custom BOM",
                command=lambda: self.nb.select(self.custom_bom_tab),
            ).pack(side="left", padx=6)
            tk.Button(bf, text="Controleer Bestanden", command=self._check_files).pack(side="left", padx=6)
            tk.Button(bf, text="Clear BOM", command=self._clear_bom).pack(side="left", padx=6)
            tk.Button(
                bf,
                text="Delete",
                command=self._delete_selected_bom_rows,
            ).pack(side="left", padx=6)


            # Tree
            style.configure("Treeview", rowheight=24)
            treef = tk.Frame(main)
            treef.pack(fill="both", expand=True, padx=8, pady=6)
            self.tree = ttk.Treeview(
                treef,
                columns=(
                    "PartNumber",
                    "Description",
                    "Production",
                    "Bestanden gevonden",
                    "Status",
                ),
                show="headings",
                selectmode="extended",
            )
            for col in ("PartNumber","Description","Production","Bestanden gevonden","Status"):
                w = 140
                if col=="Description": w=320
                if col=="Bestanden gevonden": w=180
                if col=="Status": w=120
                anchor = "center" if col=="Status" else "w"
                self.tree.heading(col, text=col, anchor=anchor)
                self.tree.column(col, width=w, anchor=anchor)
            tree_scroll = ttk.Scrollbar(treef, orient="vertical", command=self.tree.yview)
            self.tree.configure(yscrollcommand=tree_scroll.set)
            self.tree.pack(side="left", fill="both", expand=True)
            tree_scroll.pack(side="left", fill="y")
            self.tree.bind("<Button-1>", self._on_tree_click)
            self.tree.bind("<Delete>", self._delete_selected_bom_rows)

            self.tree.bind("<Down>", lambda event: self._move_tree_focus(1))
            self.tree.bind("<Up>", lambda event: self._move_tree_focus(-1))
            self.tree.bind("<Control-Tab>", self._select_next_with_ctrl_tab)
            self.tree.bind("<Control-Shift-Tab>", self._select_prev_with_ctrl_tab)
            try:
                # Some Tk builds (e.g. Linux) use ISO_Left_Tab instead of Shift-Tab.
                self.tree.bind("<Control-ISO_Left_Tab>", self._select_prev_with_ctrl_tab)
            except tk.TclError:
                # Skip the binding on platforms where the keysym is unknown (e.g. Windows).
                pass

            self.item_links: Dict[str, str] = {}

            # Actions
            act = tk.Frame(main); act.pack(fill="x", padx=8, pady=8)
            button_style = dict(
                bg=MANUFACT_BRAND_COLOR,
                activebackground="#F7B538",
                fg="black",
                activeforeground="black",
            )
            tk.Button(
                act, text="Kopieer zonder submappen", command=self._copy_flat, **button_style
            ).pack(side="left", padx=6)
            self.copy_per_prod_button = tk.Button(
                act,
                text="Kopieer per productie + bestelbonnen",
                command=self._copy_per_prod,
                **button_style,
            )
            self.copy_per_prod_button.pack(side="left", padx=6)
            tk.Button(
                act, text="Combine pdf", command=self._combine_pdf, **button_style
            ).pack(side="left", padx=6)

            # Status
            self.status_var = tk.StringVar(value="Klaar")
            tk.Label(main, textvariable=self.status_var, anchor="w").pack(fill="x", padx=8, pady=(0,8))
            self._save_settings()

        def _on_db_change(self):
            self._refresh_clients_combo()
            sel = getattr(self, "sel_frame", None)
            if sel is not None:
                try:
                    if sel.winfo_exists():
                        sel._refresh_options()
                    else:
                        self.sel_frame = None
                except Exception:
                    self.sel_frame = None

        def _refresh_clients_combo(self):
            cur = self.client_combo.get()
            opts = [self.client_db.display_name(c) for c in self.client_db.clients_sorted()]
            self.client_combo["values"] = opts
            if cur in opts:
                self.client_combo.set(cur)
            elif opts:
                self.client_combo.set(opts[0])

        def _clear_main_inputs(self) -> None:
            prev_suspend = getattr(self, "_suspend_save", False)
            self._suspend_save = True
            try:
                self.source_folder_var.set("")
                self.dest_folder_var.set("")
                self.project_number_var.set("")
                self.project_name_var.set("")
            finally:
                self._suspend_save = prev_suspend

            if not prev_suspend:
                self._save_settings()

            try:
                self.src_entry.focus_set()
            except Exception:
                pass

        def _toggle_zip_per_finish(self):
            enabled = bool(self.zip_per_finish_var.get())
            desired = 1 if enabled else 0
            if self.zip_var.get() != desired:
                self.zip_var.set(desired)
            if self.zip_finish_var.get() != desired:
                self.zip_finish_var.set(desired)

        def _update_zip_per_finish_var(self, *_args):
            desired = 1 if (self.zip_var.get() and self.zip_finish_var.get()) else 0
            if self.zip_per_finish_var.get() != desired:
                self.zip_per_finish_var.set(desired)

        def _save_settings(self, *_args):
            if getattr(self, "_suspend_save", False):
                return
            self.source_folder = self.source_folder_var.get().strip()
            self.dest_folder = self.dest_folder_var.get().strip()
            self.settings.source_folder = self.source_folder
            self.settings.dest_folder = self.dest_folder
            self.settings.project_number = self.project_number_var.get().strip()
            self.settings.project_name = self.project_name_var.get().strip()
            self.settings.zip_per_production = bool(self.zip_var.get())
            self.settings.combine_pdf_per_production = bool(
                self.combine_pdf_per_production_var.get()
            )
            self.settings.copy_finish_exports = bool(self.finish_export_var.get())
            self.settings.zip_finish_exports = bool(self.zip_finish_var.get())
            self.settings.export_processed_bom = bool(self.export_bom_var.get())
            self.settings.export_related_bom_files = bool(
                self.export_related_files_var.get()
            )
            self.settings.export_date_prefix = bool(self.export_date_prefix_var.get())
            self.settings.export_date_suffix = bool(self.export_date_suffix_var.get())
            self.settings.custom_prefix_enabled = bool(
                self.export_name_custom_prefix_enabled_var.get()
            )
            self.settings.custom_prefix_text = self.export_name_custom_prefix_text.get().strip()
            self.settings.custom_suffix_enabled = bool(
                self.export_name_custom_suffix_enabled_var.get()
            )
            self.settings.custom_suffix_text = self.export_name_custom_suffix_text.get().strip()
            self.settings.bundle_latest = bool(self.bundle_latest_var.get())
            self.settings.bundle_dry_run = bool(self.bundle_dry_run_var.get())
            self.settings.autofill_custom_bom = bool(
                self.autofill_custom_bom_var.get()
            )
            self.settings.footer_note = self.footer_note_var.get().replace("\r\n", "\n")
            for ext in self.settings.file_extensions:
                var = self.extension_vars.get(ext.key)
                if var is not None:
                    ext.enabled = bool(var.get())
            try:
                self.settings.save()
            except Exception as exc:
                print(f"Kon instellingen niet opslaan: {exc}", file=sys.stderr)

        def _sync_extension_vars_from_settings(self) -> None:
            prev = getattr(self, "_suspend_save", False)
            self._suspend_save = True
            new_vars: Dict[str, tk.IntVar] = {}
            try:
                for ext in self.settings.file_extensions:
                    var = self.extension_vars.get(ext.key)
                    if var is None:
                        var = tk.IntVar(master=self, value=1 if ext.enabled else 0)
                        var.trace_add("write", self._save_settings)
                    else:
                        desired = 1 if ext.enabled else 0
                        if var.get() != desired:
                            var.set(desired)
                    new_vars[ext.key] = var
            finally:
                self._suspend_save = prev
            self.extension_vars = new_vars

        def _rebuild_extension_checkbuttons(self) -> None:
            if not hasattr(self, "ext_frame"):
                return
            for child in self.ext_frame.winfo_children():
                child.destroy()
            if not self.settings.file_extensions:
                tk.Label(
                    self.ext_frame,
                    text="Geen bestandstypen beschikbaar. Voeg ze toe via instellingen.",
                    anchor="w",
                    justify="left",
                ).pack(anchor="w", pady=2)
                return
            for ext in self.settings.file_extensions:
                var = self.extension_vars.get(ext.key)
                if var is None:
                    var = tk.IntVar(master=self, value=1 if ext.enabled else 0)
                    var.trace_add("write", self._save_settings)
                    self.extension_vars[ext.key] = var
                tk.Checkbutton(
                    self.ext_frame, text=ext.label, variable=var, anchor="w"
                ).pack(anchor="w", pady=2)

        def apply_file_extensions(self, extensions: List[FileExtensionSetting]) -> None:
            normalized: List[FileExtensionSetting] = []
            seen_keys = set()
            for ext in extensions:
                if isinstance(ext, FileExtensionSetting):
                    ext_obj = FileExtensionSetting(
                        key=ext.key,
                        label=ext.label,
                        patterns=list(ext.patterns),
                        enabled=bool(ext.enabled),
                    )
                else:
                    try:
                        ext_obj = FileExtensionSetting.from_any(ext)
                    except ValueError:
                        continue
                base_key = ext_obj.key or "ext"
                key = base_key
                suffix = 2
                while key in seen_keys:
                    key = f"{base_key}_{suffix}"
                    suffix += 1
                if key != ext_obj.key:
                    ext_obj = FileExtensionSetting(
                        key=key,
                        label=ext_obj.label,
                        patterns=list(ext_obj.patterns),
                        enabled=ext_obj.enabled,
                    )
                normalized.append(ext_obj)
                seen_keys.add(key)

            self.settings.file_extensions = normalized
            self._sync_extension_vars_from_settings()
            self._rebuild_extension_checkbuttons()
            self._save_settings()

        def update_footer_note(self, text: str) -> None:
            normalized = (text or "").replace("\r\n", "\n")
            prev = getattr(self, "_suspend_save", False)
            self._suspend_save = True
            try:
                self.footer_note_var.set(normalized)
            finally:
                self._suspend_save = prev
            self._save_settings()

        def _pick_src(self):
            from tkinter import filedialog
            p = filedialog.askdirectory()
            if p:
                self.source_folder_var.set(p)
                self._save_settings()

        def _pick_dst(self):
            from tkinter import filedialog
            p = filedialog.askdirectory()
            if p:
                self.dest_folder_var.set(p)
                self._save_settings()

        def _selected_exts(self) -> Optional[List[str]]:
            selected: List[str] = []
            for ext in self.settings.file_extensions:
                var = self.extension_vars.get(ext.key)
                if var is None:
                    continue
                if var.get():
                    selected.extend(ext.patterns)
            return selected or None

        def _store_custom_row_flags(
            self, df: "pd.DataFrame", flags: List[bool]
        ) -> None:
            normalized = list(flags)
            if len(normalized) != len(df.index):
                normalized = [False] * len(df.index)
            df.attrs[self._CUSTOM_ROW_ATTR] = normalized

        def _get_custom_row_flags(self, df: "pd.DataFrame") -> List[bool]:
            flags = list(df.attrs.get(self._CUSTOM_ROW_ATTR, []))
            if len(flags) != len(df.index):
                flags = [False] * len(df.index)
            return flags

        def _ensure_bom_loaded(self) -> bool:
            from tkinter import messagebox

            bom_df = self.bom_df
            if bom_df is None or bom_df.empty:
                messagebox.showwarning("Let op", "Laad eerst een BOM.")
                return False
            return True

        def _autofill_custom_bom_enabled(self) -> bool:
            """Return whether automatic syncing to the Custom BOM is enabled."""

            var = getattr(self, "autofill_custom_bom_var", None)
            if var is not None:
                try:
                    return bool(var.get())
                except tk.TclError:
                    pass
            return bool(getattr(self.settings, "autofill_custom_bom", True))

        def _sync_custom_bom_from_main(self) -> None:
            """Update the Custom BOM tab so it mirrors the main BOM."""

            if not getattr(self, "custom_bom_tab", None):
                return
            if not self._autofill_custom_bom_enabled():
                return

            df = self.bom_df
            if df is None:
                empty = pd.DataFrame(columns=self.custom_bom_tab.MAIN_COLUMN_ORDER)
            else:
                empty = df

            try:
                self.custom_bom_tab.load_from_main_dataframe(empty)
            except Exception as exc:
                print(
                    f"Kon custom BOM niet vullen vanuit hoofd-BOM: {exc}",
                    file=sys.stderr,
                )

        def _load_bom_from_path(self, path: str, *, mark_as_custom: bool = False) -> None:
            df = load_bom(path)
            if "Bestanden gevonden" not in df.columns:
                df["Bestanden gevonden"] = ""
            if "Status" not in df.columns:
                df["Status"] = ""
            if "Link" not in df.columns:
                df["Link"] = ""
            self._store_custom_row_flags(df, [mark_as_custom] * len(df.index))
            self.bom_df = df
            self.bom_source_path = os.path.abspath(path)
            self._refresh_tree()
            self.status_var.set(f"BOM geladen: {len(df)} rijen")
            self._sync_custom_bom_from_main()

        def _load_bom(self):
            from tkinter import filedialog, messagebox

            start_dir = self.source_folder if self.source_folder else os.getcwd()
            path = filedialog.askopenfilename(
                filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")],
                initialdir=start_dir,
            )
            if not path:
                return
            try:
                self._load_bom_from_path(path)
            except Exception as e:
                messagebox.showerror("Fout", str(e))

        def _on_custom_bom_ready(self, path: "Path", _row_count: int) -> None:
            from tkinter import messagebox

            try:
                self._load_bom_from_path(str(path), mark_as_custom=True)
            except Exception as exc:
                messagebox.showerror("Fout", str(exc))
            else:
                self.nb.select(self.main_frame)
                row_count = _row_count
                if row_count:
                    self.status_var.set(
                        f"Aangepaste BOM geladen: {row_count} rijen (Main)"
                    )
                else:
                    self.status_var.set(
                        "Aangepaste BOM geladen. Terug naar Main-tabblad."
                    )

        def _apply_custom_bom_to_main(self, custom_df: "pd.DataFrame") -> None:
            from tkinter import messagebox


            if custom_df is None or custom_df.empty:
                messagebox.showwarning(
                    "Geen gegevens",
                    "Er zijn geen rijen met gegevens om naar de Main-tab te sturen.",
                    parent=self.custom_bom_tab,
                )
                return

            try:
                normalized = prepare_custom_bom_for_main(custom_df, self.bom_df)
            except ValueError as exc:
                messagebox.showerror("Fout", str(exc), parent=self.custom_bom_tab)
                return

            self._store_custom_row_flags(normalized, [True] * len(normalized.index))
            self.bom_df = normalized
            self._refresh_tree()
            self._sync_custom_bom_from_main()
            self.nb.select(self.main_frame)
            self.status_var.set(
                f"Custom BOM wijzigingen toegepast ({len(normalized)} rijen)."
            )

        def _schedule_opticutter_refresh(self) -> None:
            after_id = getattr(self, "_opticutter_refresh_after_id", None)
            if after_id is not None:
                try:
                    self.after_cancel(after_id)
                except tk.TclError:
                    pass
            self._opticutter_refresh_after_id = self.after(200, self._refresh_opticutter_table)

        def _on_opticutter_kerf_change(self, *_args) -> None:
            self._schedule_opticutter_refresh()

        def _on_opticutter_custom_stock_change(self, *_args) -> None:
            self._schedule_opticutter_refresh()

        def _get_opticutter_kerf_mm(self) -> float:
            var = getattr(self, "opticutter_kerf_var", None)
            if var is None:
                return DEFAULT_KERF_MM
            try:
                raw_value = str(var.get()).strip().replace(",", ".")
            except tk.TclError:
                return DEFAULT_KERF_MM
            if not raw_value:
                return DEFAULT_KERF_MM
            try:
                value = float(raw_value)
            except ValueError:
                return DEFAULT_KERF_MM
            return max(0.0, value)

        def _get_opticutter_custom_stock_mm(self) -> Optional[int]:
            var = getattr(self, "opticutter_custom_stock_var", None)
            if var is None:
                return None
            try:
                raw_value = str(var.get()).strip()
            except tk.TclError:
                return None
            if not raw_value:
                return None
            parsed = _parse_length_to_mm(raw_value)
            if parsed is None or parsed <= 0:
                return None
            return int(parsed)

        def _prompt_opticutter_manual_length(
            self, key: tuple[str, str, str]
        ) -> Optional[int]:
            existing = self.opticutter_profile_custom_lengths.get(key)
            initial = f"{existing}" if existing is not None else ""
            while True:
                response = simpledialog.askstring(
                    "Aangepaste lengte",
                    "Voer de gewenste staaflengte in (bijv. 6400 mm):",
                    parent=self,
                    initialvalue=initial,
                )
                if response is None:
                    return None
                parsed = _parse_length_to_mm(response)
                if parsed is None:
                    messagebox.showerror(
                        "Ongeldige lengte",
                        "Voer een geldige lengte in millimeter, centimeter of meter in.",
                        parent=self,
                    )
                    continue
                return parsed

        def _on_opticutter_profile_selection_change(
            self, key: tuple[str, str, str]
        ) -> None:
            if getattr(self, "_opticutter_selection_update_in_progress", False):
                return
            var = self.opticutter_profile_selection_vars.get(key)
            if var is None:
                return
            value_by_display = self.opticutter_profile_selection_value_by_display.get(
                key
            )
            display_by_value = self.opticutter_profile_selection_display_map.get(key)
            if not value_by_display:
                return
            display_value = var.get()
            canonical = value_by_display.get(display_value)
            if canonical is None:
                return
            if canonical == "manual_prompt":
                length_mm = self._prompt_opticutter_manual_length(key)
                if length_mm is None:
                    previous_choice = self.opticutter_profile_selection_choice.get(key)
                    fallback_display = (
                        display_by_value.get(previous_choice)
                        if display_by_value is not None
                        else None
                    )
                    if fallback_display is None and display_by_value:
                        first_value, fallback_display = next(
                            iter(display_by_value.items())
                        )
                        self.opticutter_profile_selection_choice[key] = first_value
                    self._opticutter_selection_update_in_progress = True
                    try:
                        var.set(fallback_display or "")
                    finally:
                        self._opticutter_selection_update_in_progress = False
                    return

                manual_value = f"manual:{length_mm}"
                self.opticutter_profile_custom_lengths[key] = length_mm
                self.opticutter_profile_selection_choice[key] = manual_value
                self._refresh_opticutter_table()
                return
            self.opticutter_profile_selection_choice[key] = canonical

        def _update_opticutter_selection_rows(
            self,
            entries: List[
                tuple[
                    tuple[str, str, str],
                    str,
                    str,
                    str,
                    List[tuple[str, str]],
                    str,
                ]
            ],
        ) -> None:
            container = getattr(self, "opticutter_selection_inner", None)
            if container is None:
                return

            empty_label = getattr(self, "opticutter_selection_empty_label", None)
            rows = self.opticutter_profile_selection_rows
            labels = self.opticutter_profile_selection_labels
            vars_map = self.opticutter_profile_selection_vars
            combos = self.opticutter_profile_selection_combos
            display_map = self.opticutter_profile_selection_display_map
            value_by_display_map = self.opticutter_profile_selection_value_by_display
            custom_lengths = self.opticutter_profile_custom_lengths

            new_keys = {entry[0] for entry in entries}
            for key in list(rows.keys()):
                if key in new_keys:
                    continue
                row = rows.pop(key)
                row.destroy()
                labels.pop(key, None)
                vars_map.pop(key, None)
                combos.pop(key, None)
                display_map.pop(key, None)
                value_by_display_map.pop(key, None)
                self.opticutter_profile_selection_choice.pop(key, None)
                custom_lengths.pop(key, None)

            if not entries:
                if empty_label is not None and not empty_label.winfo_ismapped():
                    empty_label.pack(fill="x", padx=8, pady=8)
                custom_lengths.clear()
                return

            if empty_label is not None:
                empty_label.pack_forget()

            for key, profile_name, material_name, production_name, options, selection in entries:
                row = rows.get(key)
                if row is None:
                    row = tk.Frame(container)
                    row.pack(fill="x", padx=8, pady=2)
                    profile_label = tk.Label(
                        row, text=profile_name or "—", anchor="w", width=30
                    )
                    profile_label.pack(side="left", padx=(0, 6))
                    material_label = tk.Label(
                        row, text=material_name or "—", anchor="w", width=20
                    )
                    material_label.pack(side="left", padx=(0, 6))
                    production_label = tk.Label(
                        row, text=production_name or "—", anchor="w", width=20
                    )
                    production_label.pack(side="left", padx=(0, 6))
                    var = tk.StringVar()
                    combo = ttk.Combobox(
                        row,
                        textvariable=var,
                        state="readonly",
                        width=1,
                    )
                    combo.pack(side="left", fill="x", expand=True)
                    combo.bind(
                        "<<ComboboxSelected>>",
                        lambda _e, item_key=key: self._on_opticutter_profile_selection_change(
                            item_key
                        ),
                    )
                    rows[key] = row
                    labels[key] = (profile_label, material_label, production_label)
                    vars_map[key] = var
                    combos[key] = combo
                else:
                    row.pack_forget()
                    row.pack(fill="x", padx=8, pady=2)
                    profile_label, material_label, production_label = labels[key]
                    profile_label.configure(text=profile_name or "—")
                    material_label.configure(text=material_name or "—")
                    production_label.configure(text=production_name or "—")
                    combo = combos[key]
                    var = vars_map[key]

                display_by_value = OrderedDict(options)
                display_values = list(display_by_value.values())
                value_by_display = {display: value for value, display in display_by_value.items()}
                display_map[key] = display_by_value
                value_by_display_map[key] = value_by_display
                combo_width = max((len(value) for value in display_values), default=1)
                combo.configure(values=display_values, width=combo_width)

                chosen = selection
                if chosen not in display_by_value:
                    chosen = next(iter(display_by_value.keys()), "")
                selected_display = display_by_value.get(chosen, "")
                self._opticutter_selection_update_in_progress = True
                var.set(selected_display)
                self._opticutter_selection_update_in_progress = False
                self.opticutter_profile_selection_choice[key] = chosen

        def _refresh_opticutter_table(self) -> None:
            after_id = getattr(self, "_opticutter_refresh_after_id", None)
            if after_id is not None:
                try:
                    self.after_cancel(after_id)
                except tk.TclError:
                    pass
                self._opticutter_refresh_after_id = None
            tree = getattr(self, "opticutter_tree", None)
            summary_trees = getattr(self, "opticutter_profile_summary_trees", {})
            base_summary_tree = getattr(
                self, "opticutter_profile_summary_base_tree", None
            )
            if tree is None and base_summary_tree is None and not summary_trees:
                return

            if tree is not None:
                for item in tree.get_children():
                    tree.delete(item)
                _autosize_tree_columns(tree)

            if base_summary_tree is not None:
                for item in base_summary_tree.get_children():
                    base_summary_tree.delete(item)

            for summary_tree in summary_trees.values():
                for item in summary_tree.get_children():
                    summary_tree.delete(item)

            tooltip_managers = getattr(self, "opticutter_summary_tooltips", {})
            for manager in tooltip_managers.values():
                manager.clear()
            column_maps = getattr(self, "opticutter_summary_column_map", {})

            info_var = getattr(self, "opticutter_info_var", None)
            default_message = "Laad een BOM om profielen te bekijken."
            if info_var is not None:
                info_var.set(default_message)

            df = self.bom_df
            if df is None:
                self.opticutter_profile_selection_scenarios = {}
                self._update_opticutter_selection_rows([])
                return
            if df.empty:
                if info_var is not None:
                    info_var.set("BOM is leeg. Geen profielen om te tonen.")
                self.opticutter_profile_selection_scenarios = {}
                self._update_opticutter_selection_rows([])
                return

            kerf_mm = self._get_opticutter_kerf_mm()
            custom_stock_mm = self._get_opticutter_custom_stock_mm()

            manual_lengths = dict(self.opticutter_profile_custom_lengths)
            analysis = analyse_profiles(
                df,
                kerf_mm=kerf_mm,
                custom_stock_mm=custom_stock_mm,
                manual_lengths=manual_lengths,
            )
            self.opticutter_last_analysis = analysis

            valid_keys = {profile.key for profile in analysis.profiles}
            for stored_key in list(self.opticutter_profile_custom_lengths.keys()):
                if stored_key not in valid_keys:
                    self.opticutter_profile_custom_lengths.pop(stored_key, None)

            if not analysis.profiles:
                if info_var is not None:
                    message = analysis.error or "Geen profielen gevonden in de BOM."
                    info_var.set(message)
                self.opticutter_profile_selection_scenarios = {}
                self._update_opticutter_selection_rows([])
                return

            total_qty = analysis.total_quantity

            if tree is not None:
                for row in analysis.aggregated_rows:
                    qty = int(row.get("Aantal") or 0)
                    tree.insert(
                        "",
                        "end",
                        values=(
                            row.get("PartNumber", ""),
                            row.get("Profile", ""),
                            row.get("Material", ""),
                            row.get("Production", ""),
                            row.get("Length profile", ""),
                            qty,
                        ),
                    )
                _autosize_tree_columns(tree)

            summary_frames = getattr(self, "opticutter_summary_frames", {})
            custom_frame = summary_frames.get("custom")
            if custom_frame is not None:
                if custom_stock_mm is not None:
                    custom_frame.configure(
                        text=f"Aangepaste lengte ({custom_stock_mm} mm)"
                    )
                else:
                    custom_frame.configure(text="Aangepaste lengte")

            def _format_bars(result: Optional[StockScenarioResult]) -> str:
                if result is None:
                    return "—"
                if result.dropped_pieces:
                    return "❌"
                return _bold_digits(str(result.bars))

            def _format_waste(result: Optional[StockScenarioResult]) -> str:
                if result is None or result.dropped_pieces or result.bars == 0:
                    return "—"
                return f"{result.waste_pct:.1f}%"

            def _format_cuts(result: Optional[StockScenarioResult]) -> str:
                if result is None or result.dropped_pieces or result.bars == 0:
                    return "—"
                return _bold_digits(str(result.cuts))

            def _describe_option(
                length_label: str, result: Optional[StockScenarioResult]
            ) -> str:
                if result is None:
                    return f"{length_label} – geen berekening beschikbaar"
                if result.dropped_pieces or result.bars <= 0:
                    return f"{length_label} – niet mogelijk (stukken te lang)"
                details = [
                    f"{result.bars} staven",
                    f"{result.waste_pct:.1f}% afval",
                    f"{result.cuts} zaagsneden",
                ]
                return f"{length_label} – {', '.join(details)}"

            def _join_blockers(blocker_values: Iterable[str], stock_length: int) -> str:
                values = list(blocker_values)
                if not values:
                    return "Past niet binnen de staaflengte; sommige stukken zijn te lang."
                lines = [
                    "Past niet binnen de staaflengte:",
                    *(f"- {text}" for text in sorted(values)),
                    f"Max. lengte: {stock_length} mm",
                ]
                return "\n".join(lines)

            selection_entries: List[
                tuple[
                    tuple[str, str, str],
                    str,
                    str,
                    str,
                    List[tuple[str, str]],
                    str,
                ]
            ] = []
            selection_scenarios: Dict[
                tuple[str, str, str], Dict[str, StockScenarioResult]
            ] = {}

            if base_summary_tree is not None:
                for profile in analysis.profiles:
                    base_summary_tree.insert(
                        "",
                        "end",
                        values=(
                            profile.profile,
                            profile.material,
                            profile.production,
                        ),
                    )
                _autosize_tree_columns(base_summary_tree)

            for profile in analysis.profiles:
                scenario_map = profile.scenarios
                selection_scenarios[profile.key] = scenario_map

                option_items: List[tuple[str, str]] = [
                    ("input", "Input lengte – per stuk zagen"),
                    (
                        "6000",
                        _describe_option("6000 mm", scenario_map.get("6000")),
                    ),
                    (
                        "12000",
                        _describe_option("12000 mm", scenario_map.get("12000")),
                    ),
                ]

                if "custom" in scenario_map:
                    label = (
                        f"{custom_stock_mm} mm"
                        if custom_stock_mm is not None
                        else "Custom lengte"
                    )
                    option_items.append(
                        ("custom", _describe_option(label, scenario_map.get("custom")))
                    )

                if profile.manual_choice_key and profile.manual_choice_key in scenario_map:
                    manual_label = (
                        f"Aangepaste lengte ({profile.manual_length_mm} mm)"
                        if profile.manual_length_mm is not None
                        else "Aangepaste lengte"
                    )
                    option_items.append(
                        (
                            profile.manual_choice_key,
                            _describe_option(
                                manual_label, scenario_map.get(profile.manual_choice_key)
                            ),
                        )
                    )

                option_items.append(("manual_prompt", "Aangepaste lengte…"))

                previous_choice = self.opticutter_profile_selection_choice.get(
                    profile.key
                )
                available_values = {value for value, _ in option_items}
                selected_value = (
                    previous_choice
                    if previous_choice in available_values
                    else profile.best_choice
                )

                selection_entries.append(
                    (
                        profile.key,
                        profile.profile,
                        profile.material,
                        profile.production,
                        option_items,
                        selected_value,
                    )
                )

                blockers = profile.blockers

                tree_6m = summary_trees.get("6m")
                if tree_6m is not None:
                    result_6m = scenario_map.get("6000")
                    item_id_6m = tree_6m.insert(
                        "",
                        "end",
                        values=(
                            _format_bars(result_6m),
                            _format_waste(result_6m),
                            _format_cuts(result_6m),
                        ),
                    )
                    tooltip_6m = tooltip_managers.get("6m")
                    columns_6m = column_maps.get("6m", {})
                    if tooltip_6m is not None and result_6m is not None:
                        if result_6m.dropped_pieces:
                            column_id = columns_6m.get("Bars")
                            if column_id:
                                tooltip_6m.set(
                                    item_id_6m,
                                    column_id,
                                    _join_blockers(blockers.get("6m", set()), STOCK_LENGTH_MM),
                                )
                            column_id = columns_6m.get("Waste")
                            if column_id:
                                tooltip_6m.set(
                                    item_id_6m,
                                    column_id,
                                    "Afval niet beschikbaar door te lange stukken.",
                                )
                            column_id = columns_6m.get("Cuts")
                            if column_id:
                                tooltip_6m.set(
                                    item_id_6m,
                                    column_id,
                                    "Zaagplan niet beschikbaar door te lange stukken.",
                                )
                        else:
                            column_id = columns_6m.get("Waste")
                            if column_id:
                                tooltip_6m.set(
                                    item_id_6m,
                                    column_id,
                                    f"Totale restlengte: {result_6m.waste_mm:.0f} mm",
                                )
                            column_id = columns_6m.get("Bars")
                            if column_id:
                                tooltip_6m.set(
                                    item_id_6m,
                                    column_id,
                                    f"{result_6m.bars} staaf/staven nodig",
                                )
                            column_id = columns_6m.get("Cuts")
                            if column_id:
                                tooltip_6m.set(
                                    item_id_6m,
                                    column_id,
                                    f"Geschat aantal zaagsneden: {result_6m.cuts}",
                                )

                tree_12m = summary_trees.get("12m")
                if tree_12m is not None:
                    result_12m = scenario_map.get("12000")
                    item_id_12m = tree_12m.insert(
                        "",
                        "end",
                        values=(
                            _format_bars(result_12m),
                            _format_waste(result_12m),
                            _format_cuts(result_12m),
                        ),
                    )
                    tooltip_12m = tooltip_managers.get("12m")
                    columns_12m = column_maps.get("12m", {})
                    if tooltip_12m is not None and result_12m is not None:
                        if result_12m.dropped_pieces:
                            column_id = columns_12m.get("Bars")
                            if column_id:
                                tooltip_12m.set(
                                    item_id_12m,
                                    column_id,
                                    _join_blockers(blockers.get("12m", set()), LONG_STOCK_LENGTH_MM),
                                )
                            column_id = columns_12m.get("Waste")
                            if column_id:
                                tooltip_12m.set(
                                    item_id_12m,
                                    column_id,
                                    "Afval niet beschikbaar door te lange stukken.",
                                )
                            column_id = columns_12m.get("Cuts")
                            if column_id:
                                tooltip_12m.set(
                                    item_id_12m,
                                    column_id,
                                    "Zaagplan niet beschikbaar door te lange stukken.",
                                )
                        else:
                            column_id = columns_12m.get("Waste")
                            if column_id:
                                tooltip_12m.set(
                                    item_id_12m,
                                    column_id,
                                    f"Totale restlengte: {result_12m.waste_mm:.0f} mm",
                                )
                            column_id = columns_12m.get("Bars")
                            if column_id:
                                tooltip_12m.set(
                                    item_id_12m,
                                    column_id,
                                    f"{result_12m.bars} staaf/staven nodig",
                                )
                            column_id = columns_12m.get("Cuts")
                            if column_id:
                                tooltip_12m.set(
                                    item_id_12m,
                                    column_id,
                                    f"Geschat aantal zaagsneden: {result_12m.cuts}",
                                )

                tree_custom = summary_trees.get("custom")
                if tree_custom is not None:
                    result_custom = scenario_map.get("custom")
                    values_custom = (
                        _format_bars(result_custom)
                        if "custom" in scenario_map
                        else "—",
                        _format_waste(result_custom)
                        if "custom" in scenario_map
                        else "—",
                        _format_cuts(result_custom)
                        if "custom" in scenario_map
                        else "—",
                    )
                    item_id_custom = tree_custom.insert("", "end", values=values_custom)
                    tooltip_custom = tooltip_managers.get("custom")
                    columns_custom = column_maps.get("custom", {})
                    if tooltip_custom is not None:
                        if "custom" not in scenario_map:
                            message = "Stel een aangepaste staaflengte in om scenario's te berekenen."
                            for column_key in ("Bars", "Waste", "Cuts"):
                                column_id = columns_custom.get(column_key)
                                if column_id:
                                    tooltip_custom.set(item_id_custom, column_id, message)
                        elif result_custom is not None and result_custom.dropped_pieces:
                            column_id = columns_custom.get("Bars")
                            if column_id:
                                tooltip_custom.set(
                                    item_id_custom,
                                    column_id,
                                    _join_blockers(blockers.get("custom", set()),
                                                  custom_stock_mm if custom_stock_mm is not None else 0),
                                )
                            column_id = columns_custom.get("Waste")
                            if column_id:
                                tooltip_custom.set(
                                    item_id_custom,
                                    column_id,
                                    "Afval niet beschikbaar door te lange stukken.",
                                )
                            column_id = columns_custom.get("Cuts")
                            if column_id:
                                tooltip_custom.set(
                                    item_id_custom,
                                    column_id,
                                    "Zaagplan niet beschikbaar door te lange stukken.",
                                )
                        elif result_custom is not None:
                            column_id = columns_custom.get("Waste")
                            if column_id:
                                tooltip_custom.set(
                                    item_id_custom,
                                    column_id,
                                    f"Totale restlengte: {result_custom.waste_mm:.0f} mm",
                                )
                            column_id = columns_custom.get("Bars")
                            if column_id:
                                tooltip_custom.set(
                                    item_id_custom,
                                    column_id,
                                    f"{result_custom.bars} staaf/staven nodig",
                                )
                            column_id = columns_custom.get("Cuts")
                            if column_id:
                                tooltip_custom.set(
                                    item_id_custom,
                                    column_id,
                                    f"Geschat aantal zaagsneden: {result_custom.cuts}",
                                )

            for summary_tree in summary_trees.values():
                _autosize_tree_columns(summary_tree)

            self.opticutter_profile_selection_scenarios = selection_scenarios
            self._update_opticutter_selection_rows(selection_entries)

            if info_var is not None:
                profile_count = len(analysis.aggregated_rows)
                profile_label = "profiel" if profile_count == 1 else "profielen"
                base_message = (
                    f"{profile_count} {profile_label}, totaal aantal: {total_qty}"
                )

                profile_types = len(analysis.profiles)
                base_message = (
                    f"{base_message} | {profile_types} profieltypen in overzicht"
                )

                base_message = f"{base_message} | Zaagbreedte: {kerf_mm:g} mm"
                if custom_stock_mm is not None:
                    base_message = (
                        f"{base_message} | Custom staaflengte: {custom_stock_mm} mm"
                    )

                warnings: List[str] = []
                if analysis.oversized_profiles_6m:
                    warnings.append("Let op: sommige profielen zijn langer dan 6m.")
                if analysis.oversized_profiles_12m:
                    warnings.append("Let op: sommige profielen zijn langer dan 12m.")
                if analysis.unparsed_lengths:
                    warnings.append(
                        "Sommige profiel lengtes konden niet worden gelezen."
                    )

                if warnings:
                    base_message = base_message + "\n" + " ".join(warnings)

                info_var.set(base_message)

        def _refresh_tree(self):
            self.item_links.clear()
            for it in self.tree.get_children():
                self.tree.delete(it)
            self._refresh_opticutter_table()
            df = self.bom_df
            if df is None:
                self.status_var.set("Geen BOM geladen.")
                return
            if df.empty:
                self.status_var.set("BOM is leeg.")
                return
            for _, row in df.iterrows():
                vals = (
                    row.get("PartNumber", ""),
                    row.get("Description", ""),
                    row.get("Production", ""),
                    row.get("Bestanden gevonden", ""),
                    row.get("Status", ""),
                )
                item = self.tree.insert("", "end", values=vals)
                link = row.get("Link")
                if link:
                    self.item_links[item] = link

        def _delete_selected_bom_rows(self, event=None):
            df = self.bom_df
            if df is None or df.empty:
                return "break" if event is not None else None

            if event is not None:
                try:
                    widget_with_focus = self.focus_get()
                except tk.TclError:
                    widget_with_focus = None
                if widget_with_focus is not self.tree:
                    return None

            selection = self.tree.selection()
            if not selection:
                return "break" if event is not None else None

            row_count = len(df.index)
            item_pairs: List[tuple[int, str]] = []
            for item in selection:
                try:
                    idx = self.tree.index(item)
                except tk.TclError:
                    continue
                item_pairs.append((idx, item))
            if not item_pairs:
                return "break" if event is not None else None

            custom_flags = self._get_custom_row_flags(df)
            has_custom_rows = any(custom_flags)
            sorted_pairs = sorted(item_pairs, key=lambda pair: pair[0])

            removable_pairs = [
                (idx, item)
                for idx, item in sorted_pairs
                if 0 <= idx < row_count
                and idx < len(custom_flags)
                and custom_flags[idx]
            ]

            if not removable_pairs:
                # Allow deleting regular BOM rows as a fallback when there are no
                # custom rows flagged (e.g. when working on a freshly loaded BOM).
                if has_custom_rows:
                    self.status_var.set(
                        "Geen Custom BOM-rijen geselecteerd om te verwijderen."
                    )
                    return "break" if event is not None else None

                removable_pairs = [
                    (idx, item)
                    for idx, item in sorted_pairs
                    if 0 <= idx < row_count
                ]
                if not removable_pairs:
                    return "break" if event is not None else None

            drop_labels = [df.index[idx] for idx, _ in removable_pairs]
            if not drop_labels:
                return "break" if event is not None else None

            removed_positions = {idx for idx, _ in removable_pairs}
            remaining_flags = [
                flag
                for pos, flag in enumerate(custom_flags)
                if pos not in removed_positions
            ]
            updated_df = df.drop(drop_labels).reset_index(drop=True)
            self._store_custom_row_flags(updated_df, remaining_flags)
            self.bom_df = updated_df
            self._sync_custom_bom_from_main()

            target_index = removable_pairs[0][0]
            removed = 0
            for _, item in removable_pairs:
                if item in self.item_links:
                    self.item_links.pop(item, None)
                try:
                    self.tree.delete(item)
                except tk.TclError:
                    continue
                removed += 1

            if removed:
                skipped = len(item_pairs) - removed
                if has_custom_rows:
                    if removed == 1:
                        msg = "1 Custom BOM-rij verwijderd."
                    else:
                        msg = f"{removed} Custom BOM-rijen verwijderd."
                else:
                    if removed == 1:
                        msg = "1 rij verwijderd."
                    else:
                        msg = f"{removed} rijen verwijderd."
                if skipped > 0:
                    suffix = (
                        "1 rij overgeslagen" if skipped == 1 else f"{skipped} rijen overgeslagen"
                    )
                    msg = f"{msg} ({suffix})"
                self.status_var.set(msg)

            remaining_items = list(self.tree.get_children())
            current_selection = self.tree.selection()
            if current_selection:
                try:
                    self.tree.see(current_selection[0])
                except tk.TclError:
                    pass
            elif remaining_items:
                target_index = min(target_index, len(remaining_items) - 1)
                next_item = remaining_items[target_index]
                try:
                    self.tree.selection_set(next_item)
                    self.tree.focus(next_item)
                    self.tree.see(next_item)
                except tk.TclError:
                    pass
            else:
                try:
                    if current_selection:
                        self.tree.selection_remove(*current_selection)
                    self.tree.focus("")
                except tk.TclError:
                    pass

            self._refresh_opticutter_table()


            return "break" if event is not None else None

        def _move_tree_focus(self, direction: int) -> str:
            items = list(self.tree.get_children())
            if not items:
                return "break"

            focus = self.tree.focus()
            if focus in items:
                idx = items.index(focus)
            else:
                idx = -1 if direction >= 0 else len(items)

            idx = max(0, min(len(items) - 1, idx + direction))
            target = items[idx]
            self.tree.selection_set(target)
            self.tree.focus(target)
            self.tree.see(target)
            return "break"

        def _extend_tree_selection(self, direction: int) -> str:
            items = list(self.tree.get_children())
            if not items:
                return "break"

            focus = self.tree.focus()
            if focus not in items:
                focus = items[0] if direction >= 0 else items[-1]
                self.tree.focus(focus)

            self.tree.selection_add(focus)

            idx = items.index(focus)
            idx = max(0, min(len(items) - 1, idx + direction))
            target = items[idx]
            self.tree.selection_add(target)
            self.tree.focus(target)
            self.tree.see(target)
            return "break"

        def _select_next_with_ctrl_tab(self, _event) -> str:
            return self._extend_tree_selection(1)

        def _select_prev_with_ctrl_tab(self, _event) -> str:
            return self._extend_tree_selection(-1)

        def _clear_bom(self):
            from tkinter import messagebox

            if self.bom_df is None:
                messagebox.showwarning("Let op", "Laad eerst een BOM.")
                return
            for col in ("Bestanden gevonden", "Status", "Link"):
                if col in self.bom_df.columns:
                    self.bom_df[col] = ""
            self.bom_df = None
            self.bom_source_path = None
            self._refresh_tree()
            self._sync_custom_bom_from_main()
            self.status_var.set("BOM gewist.")

        def _on_tree_click(self, event):
            item = self.tree.identify_row(event.y)
            col = self.tree.identify_column(event.x)
            if col != "#5" or not item:
                return
            if self.tree.set(item, "Status") != "❌":
                return
            path = self.item_links.get(item)
            if not path or not os.path.exists(path):
                return
            try:
                if sys.platform.startswith("win"):
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", path], check=False)
                else:
                    subprocess.run(["xdg-open", path], check=False)
            except Exception:
                pass

        def _check_files(self):
            from tkinter import messagebox
            if not self._ensure_bom_loaded():
                return
            if not self.source_folder:
                messagebox.showwarning("Let op", "Selecteer een bronmap."); return
            exts = self._selected_exts()
            if not exts:
                messagebox.showwarning("Let op", "Selecteer minstens één bestandstype."); return
            self.status_var.set("Bezig met controleren...")
            self.update_idletasks()
            idx = _build_file_index(self.source_folder, exts)
            sw_idx = _build_file_index(self.source_folder, [".sldprt", ".slddrw"])
            found, status, links = [], [], []
            groups = []
            exts_set = set(e.lower() for e in exts)
            if ".step" in exts_set or ".stp" in exts_set:
                groups.append({".step", ".stp"})
                exts_set -= {".step", ".stp"}
            for e in exts_set:
                groups.append({e})
            for _, row in self.bom_df.iterrows():
                pn = row["PartNumber"]
                hits = idx.get(pn, [])
                hit_exts = {os.path.splitext(h)[1].lower() for h in hits}
                all_present = all(any(ext in hit_exts for ext in g) for g in groups)
                found.append(", ".join(sorted(e.lstrip('.') for e in hit_exts)))
                status.append("✅" if all_present else "❌")
                link = ""
                if not all_present:
                    missing = []
                    for g in groups:
                        if not any(ext in hit_exts for ext in g):
                            missing.extend(g)
                    sw_hits = sw_idx.get(pn, [])
                    drw = next((p for p in sw_hits if p.lower().endswith(".slddrw")), None)
                    prt = next((p for p in sw_hits if p.lower().endswith(".sldprt")), None)
                    if ".pdf" in missing and drw:
                        link = drw
                    elif prt:
                        link = prt
                    elif drw:
                        link = drw
                links.append(link)
            self.bom_df["Bestanden gevonden"] = found
            self.bom_df["Status"] = status
            self.bom_df["Link"] = links
            self._refresh_tree()
            self.status_var.set("Controle klaar.")

        def _copy_flat(self):
            from tkinter import messagebox
            if not self._ensure_bom_loaded():
                return
            exts = self._selected_exts()
            if not exts or not self.source_folder or not self.dest_folder:
                messagebox.showwarning("Let op", "Selecteer bron, bestemming en extensies."); return
            custom_prefix_text = self.export_name_custom_prefix_text.get().strip()
            custom_prefix_enabled = bool(
                self.export_name_custom_prefix_enabled_var.get()
            )
            custom_suffix_text = self.export_name_custom_suffix_text.get().strip()
            custom_suffix_enabled = bool(
                self.export_name_custom_suffix_enabled_var.get()
            )

            tree_items = list(self.tree.get_children()) if hasattr(self, "tree") else []
            part_numbers_for_export: List[str] = []
            seen_part_numbers: set[str] = set()

            if tree_items:
                for item in tree_items:
                    pn = _to_str(self.tree.set(item, "PartNumber")).strip()
                    if pn and pn not in seen_part_numbers:
                        seen_part_numbers.add(pn)
                        part_numbers_for_export.append(pn)
            else:
                df_snapshot = self.bom_df
                if df_snapshot is not None:
                    for _, row in df_snapshot.iterrows():
                        pn = _to_str(row.get("PartNumber")).strip()
                        if pn and pn not in seen_part_numbers:
                            seen_part_numbers.add(pn)
                            part_numbers_for_export.append(pn)

            def work(
                token_prefix_text=custom_prefix_text,
                token_suffix_text=custom_suffix_text,
                token_prefix_enabled=custom_prefix_enabled,
                token_suffix_enabled=custom_suffix_enabled,
                export_part_numbers=tuple(part_numbers_for_export),
                bom_df_snapshot=self.bom_df,
                bom_source=self.bom_source_path,
                export_bom_enabled=bool(self.export_bom_var.get()),
                export_related_enabled=bool(self.export_related_files_var.get()),
            ):
                self.status_var.set("Bundelmap voorbereiden...")
                try:
                    bundle = create_export_bundle(
                        self.dest_folder,
                        self.project_number_var.get().strip() or None,
                        self.project_name_var.get().strip() or None,
                        latest_symlink="latest" if self.bundle_latest_var.get() else False,
                        dry_run=bool(self.bundle_dry_run_var.get()),
                    )
                except Exception as exc:
                    def on_error():
                        messagebox.showerror(
                            "Fout",
                            f"Kon bundelmap niet maken:\n{exc}",
                            parent=self,
                        )
                        self.status_var.set("Bundelmap maken mislukt.")

                    self.after(0, on_error)
                    return

                self.last_bundle_result = bundle
                bundle_dest = bundle.bundle_dir

                if bundle.warnings:
                    warnings = list(bundle.warnings)

                    def show_warnings():
                        messagebox.showwarning("Let op", "\n".join(warnings), parent=self)

                    self.after(0, show_warnings)

                if bundle.dry_run:
                    def on_dry():
                        lines = ["Testrun - doelmap:", bundle_dest]
                        if bundle.latest_symlink:
                            lines.append(f"Snelkoppeling: {bundle.latest_symlink}")
                        messagebox.showinfo("Testrun", "\n".join(lines), parent=self)
                        self.status_var.set(f"Testrun - doelmap: {bundle_dest}")

                    self.after(0, on_dry)
                    return

                self.status_var.set("Kopiëren...")
                idx = _build_file_index(self.source_folder, exts)
                date_prefix = bool(self.export_date_prefix_var.get())
                date_suffix = bool(self.export_date_suffix_var.get())
                prefix_text_clean = (token_prefix_text or "").strip()
                suffix_text_clean = (token_suffix_text or "").strip()
                prefix_active = bool(token_prefix_enabled) and bool(prefix_text_clean)
                suffix_active = bool(token_suffix_enabled) and bool(suffix_text_clean)
                today_date = datetime.date.today()
                date_token = (
                    today_date.strftime("%Y%m%d") if date_prefix or date_suffix else ""
                )
                today_iso = today_date.strftime("%Y-%m-%d")

                def _export_name(fname: str) -> str:
                    if not (
                        date_prefix
                        or date_suffix
                        or prefix_active
                        or suffix_active
                    ):
                        return fname
                    stem, ext = os.path.splitext(fname)
                    prefix_parts = []
                    if date_prefix and date_token:
                        prefix_parts.append(date_token)
                    if prefix_active:
                        prefix_parts.append(prefix_text_clean)
                    suffix_parts = []
                    if date_suffix and date_token:
                        suffix_parts.append(date_token)
                    if suffix_active:
                        suffix_parts.append(suffix_text_clean)
                    parts = prefix_parts + [stem] + suffix_parts
                    new_stem = "-".join([p for p in parts if p])
                    return f"{new_stem}{ext}"
                copied_paths: set[str] = set()
                cnt = 0
                for pn in export_part_numbers:
                    for p in idx.get(pn, []):
                        if p in copied_paths:
                            continue
                        copied_paths.add(p)
                        name = _export_name(os.path.basename(p))
                        dst = os.path.join(bundle_dest, name)
                        shutil.copy2(p, dst)
                        cnt += 1

                bom_written = False
                related_copied = 0

                if export_bom_enabled:
                    if bom_df_snapshot is None:
                        def on_error():
                            messagebox.showerror(
                                "Fout",
                                "Geen BOM beschikbaar om te exporteren.",
                                parent=self,
                            )
                            self.status_var.set("BOM-export mislukt.")

                        self.after(0, on_error)
                        return
                    try:
                        bom_filename = make_bom_export_filename(
                            bom_source,
                            today_iso,
                            _export_name,
                        )
                        _export_bom_workbook(bom_df_snapshot, bundle_dest, bom_filename)
                        bom_written = True
                    except Exception as exc:
                        def on_error():
                            messagebox.showerror(
                                "Fout",
                                f"Kon BOM-export niet opslaan:\n{exc}",
                                parent=self,
                            )
                            self.status_var.set("BOM-export mislukt.")

                        self.after(0, on_error)
                        return

                if bom_written and export_related_enabled and bom_source:
                    try:
                        for src_file in find_related_bom_exports(bom_source, idx):
                            if src_file in copied_paths:
                                continue
                            copied_paths.add(src_file)
                            transformed = _export_name(os.path.basename(src_file))
                            dst = os.path.join(bundle_dest, transformed)
                            shutil.copy2(src_file, dst)
                            related_copied += 1
                    except Exception as exc:
                        def on_error():
                            messagebox.showerror(
                                "Fout",
                                f"Kon gerelateerde exportbestanden kopiëren:\n{exc}",
                                parent=self,
                            )
                            self.status_var.set("Kopiëren mislukt.")

                        self.after(0, on_error)
                        return

                def on_done():
                    status_text = f"Klaar. Gekopieerd: {cnt} → {bundle_dest}"
                    if bom_written:
                        status_text += " (BOM opgeslagen)"
                    if related_copied:
                        status_text += f" (+{related_copied} gerelateerd)"
                    self.status_var.set(status_text)
                    info_lines = ["Bestanden gekopieerd naar:", bundle_dest]
                    if bundle.latest_symlink:
                        info_lines.append(f"Symlink: {bundle.latest_symlink}")
                    details = []
                    if bom_written:
                        details.append("BOM geëxporteerd")
                    if related_copied:
                        details.append(f"Gerelateerde bestanden: {related_copied}")
                    if details:
                        info_lines.append("")
                        info_lines.append(", ".join(details))
                    messagebox.showinfo("Klaar", "\n".join(info_lines), parent=self)
                    try:
                        if sys.platform.startswith("win"):
                            os.startfile(bundle_dest)
                        elif sys.platform == "darwin":
                            subprocess.run(["open", bundle_dest], check=False)
                        else:
                            subprocess.run(["xdg-open", bundle_dest], check=False)
                    except Exception as exc:
                        messagebox.showwarning(
                            "Let op",
                            f"Kon bundelmap niet openen:\n{exc}",
                            parent=self,
                        )

                self.after(0, on_done)
            threading.Thread(target=work, daemon=True).start()

        def _copy_per_prod(self):
            from tkinter import messagebox

            if not self._ensure_bom_loaded():
                return
            bom_df = self.bom_df
            attrs = getattr(bom_df, "attrs", {}) or {}
            missing_production = bool(attrs.get("production_column_missing"))
            if missing_production:
                messagebox.showwarning(
                    "Let op",
                    "De geladen BOM mist de kolom 'Production'. "
                    "Vul de productie in de BOM in om bestelbonnen per productie te exporteren.",
                    parent=self,
                )
                return
            exts = self._selected_exts()
            if not exts or not self.source_folder or not self.dest_folder:
                messagebox.showwarning("Let op", "Selecteer bron, bestemming en extensies."); return

            prods = sorted(
                set(
                    (str(r.get("Production") or "").strip() or "_Onbekend")
                    for _, r in bom_df.iterrows()
                )
            )
            finish_meta_map: Dict[str, Dict[str, str]] = {}
            finish_part_numbers: Dict[str, set[str]] = defaultdict(set)
            for _, row in bom_df.iterrows():
                finish_text = _to_str(row.get("Finish")).strip()
                if not finish_text:
                    continue
                meta = describe_finish_combo(row.get("Finish"), row.get("RAL color"))
                key = meta["key"]
                if key not in finish_meta_map:
                    finish_meta_map[key] = meta
                pn = _to_str(row.get("PartNumber")).strip()
                if pn:
                    finish_part_numbers[key].add(pn)
            finish_entries = []
            for key, meta in finish_meta_map.items():
                if not finish_part_numbers.get(key):
                    continue
                entry = meta.copy()
                entry["key"] = key
                finish_entries.append(entry)
            finish_entries.sort(
                key=lambda e: (
                    (_to_str(e.get("label")) or "").lower(),
                    (_to_str(e.get("key")) or "").lower(),
                )
            )
            finish_label_lookup = {
                entry["key"]: _to_str(entry.get("label")) or entry["key"]
                for entry in finish_entries
            }
            sel_frame = None

            self._refresh_opticutter_table()
            opticutter_analysis = getattr(self, "opticutter_last_analysis", None)
            scenarios_ready = bool(self.opticutter_profile_selection_scenarios)

            if (
                opticutter_analysis is not None
                and opticutter_analysis.profiles
                and not scenarios_ready
            ):
                use_auto = messagebox.askyesno(
                    "Opticutter niet ingevuld",
                    (
                        "De zaagplanning in Opticutter is nog niet ingevuld. "
                        "Wil je automatisch de beste scenario's gebruiken?\n"
                        "Kies 'Nee' om eerst naar de Opticutter-tab te gaan."
                    ),
                    parent=self,
                )
                if not use_auto:
                    self.nb.select(self.opticutter_frame)
                    return
                for profile in opticutter_analysis.profiles:
                    self.opticutter_profile_selection_choice[profile.key] = (
                        profile.best_choice
                    )
                    self.opticutter_profile_selection_scenarios[profile.key] = (
                        profile.scenarios
                    )
                self._refresh_opticutter_table()
                opticutter_analysis = getattr(self, "opticutter_last_analysis", None)
                scenarios_ready = bool(self.opticutter_profile_selection_scenarios)

            opticutter_notice_message = ""
            if (
                opticutter_analysis is not None
                and not opticutter_analysis.profiles
                and opticutter_analysis.error
            ):
                opticutter_notice_message = opticutter_analysis.error

            opticutter_context = None
            opticutter_details: Dict[str, OpticutterOrderComputation] = {}
            if (
                opticutter_analysis is not None
                and opticutter_analysis.profiles
                and scenarios_ready
            ):
                try:
                    opticutter_context = prepare_opticutter_export(
                        opticutter_analysis,
                        dict(self.opticutter_profile_selection_choice),
                    )
                except Exception:
                    opticutter_context = None
            if opticutter_context is not None:
                try:
                    opticutter_details = compute_opticutter_order_details(
                        bom_df, opticutter_context
                    )
                except Exception:
                    opticutter_details = {}

            def on_sel(
                sel_map: Dict[str, str],
                doc_map: Dict[str, str],
                doc_num_map: Dict[str, str],
                delivery_map_raw: Dict[str, str],
                remarks_map_raw: Dict[str, str],
                project_number: str,
                project_name: str,
                remember: bool,
            ):
                if not self._ensure_bom_loaded():
                    return
                current_bom = self.bom_df
                self._refresh_opticutter_table()
                opticutter_analysis_current = getattr(
                    self, "opticutter_last_analysis", None
                )

                if sel_frame is not None:
                    try:
                        if sel_frame.winfo_exists():
                            self._last_supplier_selection_state = sel_frame.serialize_state()
                    except Exception:
                        pass

                prod_override_map: Dict[str, str] = {}
                finish_override_map: Dict[str, str] = {}
                opticutter_override_map: Dict[str, str] = {}
                for key, value in sel_map.items():
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_override_map[identifier] = value
                    elif kind == "opticutter":
                        opticutter_override_map[identifier] = value
                    else:
                        prod_override_map[identifier] = value

                doc_type_map: Dict[str, str] = {}
                finish_doc_type_map: Dict[str, str] = {}
                opticutter_doc_type_map: Dict[str, str] = {}
                for key, value in doc_map.items():
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_doc_type_map[identifier] = value
                    elif kind == "opticutter":
                        opticutter_doc_type_map[identifier] = value
                    else:
                        doc_type_map[identifier] = value

                prod_doc_num_map: Dict[str, str] = {}
                finish_doc_num_map: Dict[str, str] = {}
                opticutter_doc_num_map: Dict[str, str] = {}
                for key, value in doc_num_map.items():
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_doc_num_map[identifier] = value
                    elif kind == "opticutter":
                        opticutter_doc_num_map[identifier] = value
                    else:
                        prod_doc_num_map[identifier] = value

                production_delivery_map: Dict[str, DeliveryAddress | None] = {}
                finish_delivery_map: Dict[str, DeliveryAddress | None] = {}
                opticutter_delivery_map: Dict[str, DeliveryAddress | None] = {}
                for key, name in delivery_map_raw.items():
                    clean = name.replace("★ ", "", 1)
                    if clean == "Geen":
                        resolved = None
                    elif clean in (
                        "Bestelling wordt opgehaald",
                        "Leveradres wordt nog meegedeeld",
                    ):
                        resolved = DeliveryAddress(name=clean)
                    else:
                        resolved = self.delivery_db.get(clean)
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_delivery_map[identifier] = resolved
                    elif kind == "opticutter":
                        opticutter_delivery_map[identifier] = resolved
                    else:
                        production_delivery_map[identifier] = resolved

                production_remarks_map: Dict[str, str] = {}
                finish_remarks_map: Dict[str, str] = {}
                opticutter_remarks_map: Dict[str, str] = {}
                for key, text in remarks_map_raw.items():
                    clean_text = text.strip()
                    if not clean_text:
                        continue
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_remarks_map[identifier] = clean_text
                    elif kind == "opticutter":
                        opticutter_remarks_map[identifier] = clean_text
                    else:
                        production_remarks_map[identifier] = clean_text

                custom_prefix_text = self.export_name_custom_prefix_text.get().strip()
                custom_prefix_enabled = bool(
                    self.export_name_custom_prefix_enabled_var.get()
                )
                custom_suffix_text = self.export_name_custom_suffix_text.get().strip()
                custom_suffix_enabled = bool(
                    self.export_name_custom_suffix_enabled_var.get()
                )

                def update_status(message: str) -> None:
                    def apply() -> None:
                        self.status_var.set(message)
                        if sel_frame is not None:
                            try:
                                if sel_frame.winfo_exists():
                                    sel_frame.update_status(message)
                            except tk.TclError:
                                pass

                    self.after(0, apply)

                def set_busy_state(active: bool, message: Optional[str] = None) -> None:
                    def apply() -> None:
                        btn = getattr(self, "copy_per_prod_button", None)
                        if btn is not None:
                            try:
                                btn.configure(state="disabled" if active else "normal")
                            except tk.TclError:
                                pass
                        if sel_frame is not None:
                            try:
                                if sel_frame.winfo_exists():
                                    sel_frame.set_busy(active, message)
                            except tk.TclError:
                                pass

                    self.after(0, apply)
                    if message is not None:
                        update_status(message)

                def work(
                    token_prefix_text=custom_prefix_text,
                    token_suffix_text=custom_suffix_text,
                    token_prefix_enabled=custom_prefix_enabled,
                    token_suffix_enabled=custom_suffix_enabled,
                    opticutter_analysis_snapshot=opticutter_analysis_current,
                    opticutter_choices_snapshot=None,
                ):
                    update_status("Bundelmap voorbereiden...")
                    try:
                        bundle = create_export_bundle(
                            self.dest_folder,
                            project_number or None,
                            project_name or None,
                            latest_symlink="latest" if self.bundle_latest_var.get() else False,
                            dry_run=bool(self.bundle_dry_run_var.get()),
                        )
                    except Exception as exc:
                        def on_error():
                            messagebox.showerror(
                                "Fout",
                                f"Kon bundelmap niet maken:\n{exc}",
                                parent=self,
                            )
                            update_status("Bundelmap maken mislukt.")
                            set_busy_state(False)

                        self.after(0, on_error)
                        return

                    self.last_bundle_result = bundle
                    bundle_dest = bundle.bundle_dir

                    if bundle.warnings:
                        warnings = list(bundle.warnings)

                        def show_warnings():
                            messagebox.showwarning("Let op", "\n".join(warnings), parent=self)

                        self.after(0, show_warnings)

                    if bundle.dry_run:
                        def on_dry():
                            lines = ["Testrun - doelmap:", bundle_dest]
                            if bundle.latest_symlink:
                                lines.append(f"Snelkoppeling: {bundle.latest_symlink}")
                            messagebox.showinfo("Testrun", "\n".join(lines), parent=self)
                            update_status(f"Testrun - doelmap: {bundle_dest}")
                            set_busy_state(False)

                        self.after(0, on_dry)
                        return

                    update_status("Kopiëren & bestelbonnen maken...")
                    client = self.client_db.get(
                        self.client_var.get().replace("★ ", "", 1)
                    )
                    path_limit_messages: List[str] = []
                    try:
                        if opticutter_choices_snapshot is None:
                            opticutter_choices_snapshot = dict(
                                self.opticutter_profile_selection_choice
                            )
                        cnt, chosen = copy_per_production_and_orders(
                            self.source_folder,
                            bundle_dest,
                            current_bom,
                            exts,
                            self.db,
                            prod_override_map,
                            doc_type_map,
                            prod_doc_num_map,
                            remember,
                            client=client,
                            delivery_map=production_delivery_map,
                            footer_note=self.footer_note_var.get(),
                            zip_parts=bool(self.zip_var.get()),
                            date_prefix_exports=bool(self.export_date_prefix_var.get()),
                            date_suffix_exports=bool(self.export_date_suffix_var.get()),
                            project_number=project_number,
                            project_name=project_name,
                            copy_finish_exports=bool(self.finish_export_var.get()),
                            zip_finish_exports=bool(self.zip_finish_var.get()),
                            export_bom=bool(self.export_bom_var.get()),
                            export_related_files=bool(
                                self.export_related_files_var.get()
                            ),
                            export_name_prefix_text=token_prefix_text,
                            export_name_prefix_enabled=token_prefix_enabled,
                            export_name_suffix_text=token_suffix_text,
                            export_name_suffix_enabled=token_suffix_enabled,
                            finish_override_map=finish_override_map,
                            finish_doc_type_map=finish_doc_type_map,
                            finish_doc_num_map=finish_doc_num_map,
                            finish_delivery_map=finish_delivery_map,
                            remarks_map=production_remarks_map,
                            finish_remarks_map=finish_remarks_map,
                            bom_source_path=self.bom_source_path,
                            path_limit_warnings=path_limit_messages,
                            opticutter_analysis=opticutter_analysis_snapshot,
                            opticutter_choices=opticutter_choices_snapshot,
                            opticutter_override_map=opticutter_override_map,
                            opticutter_doc_type_map=opticutter_doc_type_map,
                            opticutter_doc_num_map=opticutter_doc_num_map,
                            opticutter_delivery_map=opticutter_delivery_map,
                            opticutter_remarks_map=opticutter_remarks_map,
                        )
                    except Exception as exc:
                        error_message = str(exc)

                        def on_error():
                            messagebox.showerror(
                                "Fout",
                                f"Bestelbonnen exporteren mislukt:\n{error_message}",
                                parent=self,
                            )
                            update_status("Export mislukt.")
                            set_busy_state(False)

                        self.after(0, on_error)
                        return

                    def on_done():
                        friendly_pairs = []
                        for key, value in chosen.items():
                            kind, identifier = parse_selection_key(key)
                            if kind == "finish":
                                label = finish_label_lookup.get(identifier, identifier)
                                prefix = "Afwerking"
                            elif kind == "opticutter":
                                label = identifier
                                prefix = "Brutemateriaal"
                            else:
                                label = identifier
                                prefix = "Productie"
                            friendly_pairs.append(f"{prefix} {label}: {value}")
                        suppliers_text = (
                            "; ".join(friendly_pairs)
                            if friendly_pairs
                            else str(chosen)
                        )
                        final_status = (
                            f"Klaar. Gekopieerd: {cnt}. Leveranciers: {suppliers_text}. → {bundle_dest}"
                        )
                        update_status(final_status)
                        try:
                            info_lines = ["Bestelbonnen aangemaakt in:", bundle_dest]
                            if bundle.latest_symlink:
                                info_lines.append(f"Symlink: {bundle.latest_symlink}")
                            messagebox.showinfo("Klaar", "\n".join(info_lines), parent=self)
                            if path_limit_messages:
                                warning_lines = [
                                    "Sommige exportbestanden kregen een kortere naam omdat het pad te lang werd.",
                                    f"Windows laat maximaal {_WINDOWS_MAX_PATH} tekens per pad toe; Filehopper voegt dan automatisch een korte code toe.",
                                    "",
                                ]
                                warning_lines.extend(f"• {msg}" for msg in path_limit_messages)
                                warning_lines.extend(
                                    [
                                        "",
                                        "Kort de doelmap of de bestandsnaam in om dit te vermijden.",
                                    ]
                                )
                                messagebox.showwarning(
                                    "Padlimiet bereikt",
                                    "\n".join(warning_lines),
                                    parent=self,
                                )
                            try:
                                if sys.platform.startswith("win"):
                                    os.startfile(bundle_dest)
                                elif sys.platform == "darwin":
                                    subprocess.run(["open", bundle_dest], check=False)
                                else:
                                    subprocess.run(["xdg-open", bundle_dest], check=False)
                            except Exception as exc:
                                messagebox.showwarning(
                                    "Let op",
                                    f"Kon bundelmap niet openen:\n{exc}",
                                    parent=self,
                                )
                        finally:
                            current_sel = getattr(self, "sel_frame", None)
                            if current_sel is not None:
                                try:
                                    if current_sel.winfo_exists():
                                        self._last_supplier_selection_state = current_sel.serialize_state()
                                except Exception:
                                    pass
                            self.nb.select(self.main_frame)
                            set_busy_state(False)

                    self.after(0, on_done)

                set_busy_state(True, "Bundelmap voorbereiden...")

                choices_snapshot = dict(self.opticutter_profile_selection_choice)
                threading.Thread(
                    target=work,
                    kwargs={"opticutter_choices_snapshot": choices_snapshot},
                    daemon=True,
                ).start()

            sup_search_restore = ""
            sup_frame = getattr(self, "suppliers_frame", None)
            if sup_frame is not None and hasattr(sup_frame, "suspend_search_filter"):
                try:
                    sup_search_restore = sup_frame.suspend_search_filter()
                except Exception:
                    sup_search_restore = ""

            previous_state = getattr(self, "_last_supplier_selection_state", None)
            existing_frame = getattr(self, "sel_frame", None)
            if existing_frame is not None:
                try:
                    if existing_frame.winfo_exists():
                        previous_state = existing_frame.serialize_state()
                except Exception:
                    pass
                try:
                    self.nb.forget(existing_frame)
                except Exception:
                    pass
                try:
                    existing_frame.destroy()
                except Exception:
                    pass
                self.sel_frame = None

            try:
                sel_frame = SupplierSelectionFrame(
                    self.nb,
                    prods,
                    finish_entries,
                    self.db,
                    self.delivery_db,
                    on_sel,
                    self.project_number_var,
                    self.project_name_var,
                    opticutter_details=opticutter_details,
                    initial_state=previous_state,
                )
            except Exception:
                if sup_search_restore and hasattr(sup_frame, "restore_search_filter"):
                    try:
                        sup_frame.restore_search_filter(sup_search_restore)
                    except Exception:
                        pass
                raise
            self.sel_frame = sel_frame
            try:
                sel_frame.set_opticutter_notice(opticutter_notice_message)
            except Exception:
                pass
            try:
                self._last_supplier_selection_state = sel_frame.serialize_state()
            except Exception:
                pass
            self.nb.add(sel_frame, text="Bestelbonnen")
            self.nb.select(sel_frame)

            if sup_search_restore and hasattr(sup_frame, "restore_search_filter"):
                def _restore_search(_event=None, frame=sup_frame, value=sup_search_restore):
                    try:
                        frame.restore_search_filter(value)
                    except Exception:
                        pass

                sel_frame.bind("<Destroy>", _restore_search, add="+")

        def _combine_pdf(self):
            from tkinter import messagebox
            if not self._ensure_bom_loaded():
                return
            bom_df = self.bom_df
            if self.source_folder and bom_df is not None:
                def work():
                    self.status_var.set("PDF's combineren...")
                    try:
                        out_dir = self.dest_folder or self.source_folder
                        pn = self.project_number_var.get().strip() if self.project_number_var else ""
                        pname = self.project_name_var.get().strip() if self.project_name_var else ""
                        result = combine_pdfs_from_source(
                            self.source_folder,
                            bom_df,
                            out_dir,
                            project_number=pn or None,
                            project_name=pname or None,
                            combine_per_production=bool(
                                self.combine_pdf_per_production_var.get()
                            ),
                        )
                    except ModuleNotFoundError:
                        self.status_var.set("PyPDF2 ontbreekt")
                        messagebox.showwarning(
                            "PyPDF2 ontbreekt",
                            "Installeer PyPDF2 om PDF's te combineren.",
                        )
                        return
                    self.status_var.set(
                        f"Gecombineerde pdf's: {result.count} → {result.output_dir}"
                    )
                    messagebox.showinfo(
                        "Klaar",
                        "PDF's gecombineerd.\n\n" f"Map: {result.output_dir}",
                    )
                    try:
                        if sys.platform.startswith("win"):
                            os.startfile(result.output_dir)
                        elif sys.platform == "darwin":
                            subprocess.run(["open", result.output_dir], check=False)
                        else:
                            subprocess.run(["xdg-open", result.output_dir], check=False)
                    except Exception as exc:
                        messagebox.showwarning(
                            "Let op",
                            f"Kon exportmap niet openen:\n{exc}",
                        )
                threading.Thread(target=work, daemon=True).start()
            else:
                messagebox.showwarning(
                    "Let op", "Selecteer bronmap en laad een BOM."
                )

    App().mainloop()

