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
from dataclasses import dataclass, field
from copy import deepcopy
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple, TYPE_CHECKING

import pandas as pd

from app_settings import AppSettings, FileExtensionSetting, FILE_EXTENSION_PRESETS
from app_diagnostics import (
    backup_root,
    build_diagnostic_report,
    create_data_file_backups,
    format_report_for_clipboard,
)
from helpers import (
    _to_str,
    _build_file_index,
    create_export_bundle,
    ExportBundleResult,
    favorite_marker,
    favorite_prefix,
    strip_favorite_marker,
)
from models import Supplier, Client, DeliveryAddress, color_to_rgb, normalize_rgb_color
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from clients_db import ClientsDB, CLIENTS_DB_FILE
from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE
from app_paths import (
    APP_NAME,
    APP_VERSION,
    bundle_root,
    ensure_runtime_files,
    resolve_runtime_path,
    runtime_asset_dir,
    to_runtime_relative_path,
)
from order_presets_db import (
    ORDER_PRESETS_DB_FILE,
    OrderPresetRule,
    OrderPresetsDB,
)
from pdf_workdossier_presets import (
    PDF_WORKDOSSIER_PRESETS_DB_FILE,
    PdfWorkDossierPreset,
    PdfWorkDossierPresetsDB,
    PdfWorkDossierSection,
    default_pdf_workdossier_preset,
    tecno_art_pdf_workdossier_preset,
)
from bom import read_csv_flex, load_bom
from bom_custom_tab import BOMCustomTab
from manual_order_tab import ManualOrderTab
from bom_sync import prepare_custom_bom_for_main
from orders import (
    copy_per_production_and_orders,
    DEFAULT_FOOTER_NOTE,
    MIAMI_PINK,
    ORDER_MUTED_TEXT_COLOR,
    ORDER_RULE_COLOR,
    ORDER_TABLE_ALT_ROW_COLOR,
    ORDER_TABLE_GRID_COLOR,
    ORDER_TABLE_OUTLINE_COLOR,
    ORDER_TEXT_COLOR,
    combine_pdfs_from_source,
    combine_workdossier_pdf_from_source,
    build_pdf_workdossier_plan,
    PdfWorkDossierPlanItem,
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
    _fit_filename_within_path,
    _sanitize_component,
    _normalize_crop_box,
    build_document_export_basename,
    build_order_pricing_item_key,
    format_document_number_for_display,
    write_order_excel,
    generate_pdf_order_platypus,
)
from export_session_log import (
    EXPORT_SESSION_LOG_FILENAME,
    build_export_session_log,
    convert_offers_to_orders,
    find_export_session_logs,
    format_export_log_compatibility_message,
    load_export_session_log,
    merge_order_state_sections,
    resolve_export_document_path,
    state_keys_for_import_sections,
    summarize_export_log_compatibility,
    write_export_session_log,
)
from en1090 import EN1090_NOTE_TEXT, default_en1090_enabled, normalize_en1090_key
from opticutter import (
    StockScenarioResult,
    analyse_profiles,
    parse_length_to_mm,
    prepare_opticutter_export,
)
from changelog_viewer import (
    format_changelog_for_display,
    get_latest_release_notes,
    load_changelog,
)
from help_content import FAQ_ENTRIES, QUICK_START_STEPS

if TYPE_CHECKING:
    from orders import OpticutterOrderComputation


CLIENT_LOGO_DIR = "client_logos"
# A softer brand accent for manufacturing-focused actions.
MANUFACT_BRAND_COLOR = "#FADFA8"
RUNTIME_DATA_FILES = [
    "clients_db.json",
    "suppliers_db.json",
    "delivery_addresses_db.json",
    "app_settings.json",
    "order_presets.json",
    "pdf_workdossier_presets.json",
]
SUPPLIERS_TEMPLATE_FILE = "suppliers_template.csv"


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


def _crop_logo_preview_image(img, crop: object):
    """Return ``img`` cropped with validated logo crop data.

    The client editor preview should remain usable even when legacy or invalid
    crop metadata is stored on the client record.
    """

    if img is None:
        return None

    try:
        width = int(getattr(img, "width", 0) or 0)
        height = int(getattr(img, "height", 0) or 0)
    except Exception:
        return img

    crop_box = _normalize_crop_box(crop, width, height)
    if crop_box is None:
        return img

    try:
        cropped = img.crop(crop_box)
    except Exception:
        return img

    try:
        if int(getattr(cropped, "width", 0) or 0) <= 0:
            return img
        if int(getattr(cropped, "height", 0) or 0) <= 0:
            return img
    except Exception:
        return img
    return cropped


def _safe_make_logo_photo(img, image_tk, resample, max_size: tuple[int, int]):
    """Return a resized ``PhotoImage`` or ``None`` when rendering fails."""

    if img is None or image_tk is None:
        return None

    try:
        thumb = img.copy()
        if resample is not None:
            thumb.thumbnail(max_size, resample)
        else:  # pragma: no cover - fallback without Pillow resampling enum
            thumb.thumbnail(max_size)
        return image_tk.PhotoImage(thumb)
    except Exception:
        return None


def _resolve_file_dialog_initial_dir(preferred_path: object) -> str:
    """Return a usable start directory for native file dialogs."""

    fallback = os.getcwd()
    raw_path = _to_str(preferred_path).strip()
    if not raw_path:
        return fallback

    try:
        candidate = os.path.abspath(os.path.expanduser(raw_path))
    except (OSError, TypeError, ValueError):
        return fallback

    try:
        if os.path.isdir(candidate):
            return candidate
    except OSError:
        pass
    return fallback


def start_gui():
    ensure_runtime_files(RUNTIME_DATA_FILES)
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
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

    class _HelpTooltip:
        """Show a short help tooltip for any widget."""

        def __init__(self, widget: "tk.Widget", text: str, *, delay_ms: int = 350):
            self.widget = widget
            self.text = text
            self.delay_ms = delay_ms
            self._tipwindow: Optional["tk.Toplevel"] = None
            self._after_id: Optional[str] = None
            widget.bind("<Enter>", self._schedule_show, add="+")
            widget.bind("<Leave>", self._hide, add="+")
            widget.bind("<Destroy>", self._hide, add="+")

        def _schedule_show(self, _event=None) -> None:
            self._cancel_scheduled()
            if not self.text:
                return
            try:
                if not self.widget.winfo_viewable():
                    return
                self._after_id = self.widget.after(self.delay_ms, self._show)
            except tk.TclError:
                self._after_id = None

        def _show(self) -> None:
            self._after_id = None
            if self._tipwindow is not None:
                return
            try:
                x = self.widget.winfo_pointerx() + 12
                y = self.widget.winfo_pointery() + 12
            except tk.TclError:
                return
            tip = tk.Toplevel(self.widget)
            tip.wm_overrideredirect(True)
            try:
                tip.wm_attributes("-topmost", True)
            except tk.TclError:
                pass
            label = tk.Label(
                tip,
                text=self.text,
                background="#ffffe0",
                foreground="#444444",
                relief="solid",
                borderwidth=1,
                justify="left",
                padx=6,
                pady=4,
                wraplength=360,
            )
            label.pack()
            tip.wm_geometry(f"+{x}+{y}")
            self._tipwindow = tip

        def _cancel_scheduled(self) -> None:
            if self._after_id is not None:
                try:
                    self.widget.after_cancel(self._after_id)
                except tk.TclError:
                    pass
                self._after_id = None

        def _hide(self, _event=None) -> None:
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

            cols = ("Naam", "Adres", "BTW", "E-mail", "Website")
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
            fav_label = f"Favoriet {favorite_marker()}"
            tk.Button(btns, text=fav_label, command=self.toggle_fav_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Importeer CSV", command=self.import_csv).pack(side="left", padx=4)
            self.refresh()

        def refresh(self):
            for it in self.tree.get_children():
                self.tree.delete(it)
            for idx, c in enumerate(self.db.clients_sorted()):
                name = self.db.display_name(c)
                vals = (name, c.address or "", c.vat or "", c.email or "", c.website or "")
                tag = "odd" if idx % 2 == 0 else "even"
                self.tree.insert("", "end", values=vals, tags=(tag,))
            self.tree.tag_configure("odd", background=TREE_ODD_BG)
            self.tree.tag_configure("even", background=TREE_EVEN_BG)

        def _sel_name(self):
            sel = self.tree.selection()
            if not sel:
                return None
            vals = self.tree.item(sel[0], "values")
            return strip_favorite_marker(vals[0])

        def _open_edit_dialog(self, client: Optional[Client] = None):
            win = tk.Toplevel(self)
            win.title("Opdrachtgever")
            win.columnconfigure(1, weight=1)
            fields = [
                ("Naam", "name"),
                ("Adres", "address"),
                ("BTW", "vat"),
                ("E-mail", "email"),
                ("Website", "website"),
            ]
            entries: Dict[str, tk.Entry] = {}
            for i, (lbl, key) in enumerate(fields):
                tk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=4, pady=2)
                ent = tk.Entry(win, width=40)
                ent.grid(row=i, column=1, padx=4, pady=2, sticky="ew")
                if client:
                    ent.insert(0, _to_str(getattr(client, key)))
                entries[key] = ent
            accent_row = len(fields)
            tk.Label(win, text="Kleur RGB:").grid(
                row=accent_row, column=0, sticky="e", padx=4, pady=(2, 0)
            )
            accent_color_var = tk.StringVar(
                value=(client.accent_color if client and client.accent_color else "")
            )
            accent_frame = tk.Frame(win)
            accent_frame.grid(row=accent_row, column=1, padx=4, pady=(2, 0), sticky="ew")
            accent_frame.columnconfigure(0, weight=1)
            accent_entry = tk.Entry(accent_frame, textvariable=accent_color_var, width=28)
            accent_entry.grid(row=0, column=0, sticky="ew")
            accent_swatch = tk.Label(
                accent_frame,
                width=3,
                relief="solid",
                bd=1,
                anchor="center",
                cursor="hand2",
            )
            accent_swatch.grid(row=0, column=1, padx=(6, 0))
            accent_help_row = accent_row + 1
            tk.Label(
                win,
                text="bijv. 255,119,255 of #FF77FF",
                anchor="w",
                justify="left",
                fg="#666666",
            ).grid(row=accent_help_row, column=1, sticky="w", padx=4, pady=(2, 0))

            def current_accent_color() -> str:
                return normalize_rgb_color(accent_color_var.get()) or MIAMI_PINK

            def current_accent_text_color() -> str:
                rgb = color_to_rgb(current_accent_color())
                if rgb is None:
                    return ORDER_TEXT_COLOR
                luminance = (
                    (0.299 * rgb[0]) + (0.587 * rgb[1]) + (0.114 * rgb[2])
                ) / 255.0
                return ORDER_TEXT_COLOR if luminance >= 0.68 else "#FFFFFF"

            def update_accent_swatch(*_args) -> None:
                raw = accent_color_var.get().strip()
                normalized = normalize_rgb_color(raw)
                if raw and not normalized:
                    accent_swatch.configure(bg="#F8D7DA", text="!")
                    return
                accent_swatch.configure(bg=(normalized or MIAMI_PINK), text="")

            def open_native_color_chooser() -> None:
                _rgb, chosen = colorchooser.askcolor(
                    color=current_accent_color(),
                    parent=win,
                    title="Kies klantkleur",
                )
                if chosen:
                    accent_color_var.set(normalize_rgb_color(chosen) or chosen)

            accent_picker_menu = tk.Menu(win, tearoff=0)
            accent_picker_menu.add_command(
                label="Kleurenkiezer openen", command=open_native_color_chooser
            )
            accent_picker_menu.add_separator()
            accent_picker_menu.add_command(
                label="Kleur wissen", command=lambda: accent_color_var.set("")
            )

            def open_accent_picker_menu(event=None) -> None:
                try:
                    accent_picker_menu.tk_popup(
                        (event.x_root if event else accent_swatch.winfo_rootx()),
                        (
                            event.y_root
                            if event
                            else accent_swatch.winfo_rooty() + accent_swatch.winfo_height()
                        ),
                    )
                finally:
                    accent_picker_menu.grab_release()

            accent_color_var.trace_add("write", update_accent_swatch)
            update_accent_swatch()
            accent_swatch.bind("<Button-1>", lambda _event: open_native_color_chooser())
            accent_swatch.bind("<Button-3>", open_accent_picker_menu)

            fav_var = tk.BooleanVar(value=client.favorite if client else False)
            tk.Checkbutton(win, text="Favoriet", variable=fav_var).grid(
                row=len(fields) + 2, column=1, sticky="w", padx=4, pady=2
            )

            logo_path_var = tk.StringVar(
                value=(client.logo_path if client and client.logo_path else "")
            )
            logo_crop_state = (
                dict(client.logo_crop) if client and client.logo_crop else None
            )

            logo_frame = tk.LabelFrame(win, text="Logo")
            logo_frame.grid(
                row=len(fields) + 3,
                column=0,
                columnspan=2,
                sticky="ew",
                padx=4,
                pady=(6, 2),
            )
            logo_frame.columnconfigure(0, weight=1, minsize=280)
            logo_frame.columnconfigure(1, weight=0, minsize=240)

            preview_label = tk.Label(
                logo_frame,
                text="Geen logo",
                relief="sunken",
                width=40,
                height=10,
                anchor="center",
                justify="center",
                cursor="hand2",
            )
            preview_label.grid(row=0, column=0, rowspan=6, sticky="nsew", padx=4, pady=4)

            def resolve_logo_path(path_str: str) -> Optional[Path]:
                if not path_str:
                    return None
                return resolve_runtime_path(path_str)

            def load_logo_image(*, apply_crop: bool = True):
                path_str = logo_path_var.get().strip()
                if not path_str or Image is None:
                    return None
                abs_path = resolve_logo_path(path_str)
                if not abs_path or not abs_path.exists():
                    return None
                try:
                    with Image.open(abs_path) as src:  # type: ignore[union-attr]
                        img = src.convert("RGBA")
                except Exception:
                    return None
                crop = logo_crop_state if apply_crop else None
                return _crop_logo_preview_image(img, crop)

            def make_logo_photo(img, max_size: tuple[int, int]):
                return _safe_make_logo_photo(img, ImageTk, RESAMPLE, max_size)

            def update_preview() -> None:
                img = load_logo_image(apply_crop=True)
                if img is None:
                    preview_label.configure(text="Geen logo", image="")
                    preview_label.image = None  # type: ignore[attr-defined]
                    return
                photo = make_logo_photo(img, (280, 160))
                if photo is None:
                    preview_label.configure(text="Kan logo niet laden", image="")
                    preview_label.image = None  # type: ignore[attr-defined]
                    return
                preview_label.configure(image=photo, text="")
                preview_label.image = photo  # type: ignore[attr-defined]

            def open_logo_preview() -> None:
                original_img = load_logo_image(apply_crop=False)
                if original_img is None:
                    messagebox.showinfo(
                        "Geen logo",
                        "Upload eerst een logo voordat je het volledig bekijkt.",
                        parent=win,
                    )
                    return

                current_img = load_logo_image(apply_crop=True) or original_img
                preview_win = tk.Toplevel(win)
                preview_win.title("Logo voorbeeld")
                preview_win.transient(win)
                preview_win.resizable(False, False)

                tk.Label(
                    preview_win,
                    text=(
                        "Bovenaan zie je het volledige originele bestand. "
                        "Onderaan zie je exact welk deel op de bon terechtkomt."
                    ),
                    anchor="w",
                    justify="left",
                    wraplength=760,
                ).pack(fill="x", padx=10, pady=(10, 8))

                panels = tk.Frame(preview_win)
                panels.pack(fill="both", expand=True, padx=10, pady=(0, 8))

                def add_logo_panel(title: str, img, max_size: tuple[int, int]) -> None:
                    frame = tk.LabelFrame(panels, text=title)
                    frame.pack(fill="both", expand=True, pady=6)
                    photo = make_logo_photo(img, max_size)
                    if photo is None:
                        tk.Label(frame, text="Preview niet beschikbaar").pack(
                            padx=12, pady=12
                        )
                        return
                    label = tk.Label(frame, image=photo, bg="white", relief="sunken")
                    label.image = photo  # type: ignore[attr-defined]
                    label.pack(padx=10, pady=(10, 6))
                    tk.Label(
                        frame,
                        text=f"{img.width} x {img.height} px",
                        anchor="w",
                    ).pack(fill="x", padx=10, pady=(0, 10))

                add_logo_panel("Origineel logo", original_img, (720, 260))
                add_logo_panel("Resultaat op bon", current_img, (720, 180))

                buttons = tk.Frame(preview_win)
                buttons.pack(fill="x", padx=10, pady=(0, 10))
                tk.Button(
                    buttons,
                    text="Voorbeeld bon",
                    command=lambda: open_order_preview(parent=preview_win),
                ).pack(side="left")
                tk.Button(
                    buttons, text="Sluiten", command=preview_win.destroy
                ).pack(side="right")

                _place_window_near_parent(preview_win, win)
                preview_win.grab_set()
                preview_win.focus_set()

            def open_order_preview(parent=None) -> None:
                preview_win = tk.Toplevel(parent or win)
                preview_win.title("Voorbeeld bon")
                preview_win.transient(parent or win)
                preview_win.geometry("860x920")

                tk.Label(
                    preview_win,
                    text=(
                        "Visuele benadering van de PDF-bestelbon. "
                        "Gebruik dit om snel te controleren of het logo mooi in de kopzone staat."
                    ),
                    anchor="w",
                    justify="left",
                    wraplength=820,
                ).pack(fill="x", padx=12, pady=(10, 8))

                canvas_wrap = tk.Frame(preview_win)
                canvas_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 8))
                canvas_wrap.grid_columnconfigure(0, weight=1)
                canvas_wrap.grid_rowconfigure(0, weight=1)

                preview_canvas = tk.Canvas(
                    canvas_wrap,
                    width=800,
                    height=760,
                    bg="#d7dbe0",
                    highlightthickness=0,
                )
                preview_scroll = tk.Scrollbar(
                    canvas_wrap, orient="vertical", command=preview_canvas.yview
                )
                preview_canvas.configure(yscrollcommand=preview_scroll.set)
                preview_canvas.grid(row=0, column=0, sticky="nsew")
                preview_scroll.grid(row=0, column=1, sticky="ns")

                sheet_x = 36
                sheet_y = 24
                sheet_w = 720
                sheet_h = 980

                def render_order_preview() -> None:
                    preview_canvas.delete("all")
                    preview_canvas.preview_images = []  # type: ignore[attr-defined]
                    preview_canvas.configure(
                        scrollregion=(0, 0, sheet_x + sheet_w + 36, sheet_y + sheet_h + 24)
                    )
                    accent_color = current_accent_color()
                    accent_text_color = current_accent_text_color()

                    preview_canvas.create_rectangle(
                        0,
                        0,
                        sheet_x + sheet_w + 36,
                        sheet_y + sheet_h + 24,
                        fill="#d7dbe0",
                        outline="",
                    )
                    preview_canvas.create_rectangle(
                        sheet_x + 6,
                        sheet_y + 6,
                        sheet_x + sheet_w + 6,
                        sheet_y + sheet_h + 6,
                        fill="#c6cad0",
                        outline="",
                    )
                    preview_canvas.create_rectangle(
                        sheet_x,
                        sheet_y,
                        sheet_x + sheet_w,
                        sheet_y + sheet_h,
                        fill="white",
                        outline="#c5c5c5",
                    )

                    title_x = sheet_x + 28
                    top_y = sheet_y + 28
                    preview_canvas.create_text(
                        title_x,
                        top_y,
                        anchor="nw",
                        text="Bestelbon productie: Laser",
                        font=("TkDefaultFont", 18, "bold"),
                        fill=accent_color,
                    )
                    rule_y = top_y + 34
                    preview_canvas.create_line(
                        title_x,
                        rule_y,
                        sheet_x + sheet_w - 28,
                        rule_y,
                        fill=ORDER_RULE_COLOR,
                        width=1,
                    )

                    client_name = entries["name"].get().strip() or "Opdrachtgever"
                    client_address = entries["address"].get().strip() or "Voorbeeldstraat 1, 2000 Antwerpen"
                    client_vat = entries["vat"].get().strip() or "BE0123456789"
                    client_email = entries["email"].get().strip() or "info@example.com"
                    client_website = entries["website"].get().strip()
                    today_text = datetime.date.today().strftime("%Y-%m-%d")
                    preview_doc_number = format_document_number_for_display(
                        "BB-123",
                        "Bestelbon",
                    )
                    doc_top_y = rule_y + 10
                    doc_lines = [
                        ("Nummer:", preview_doc_number or "BB-123"),
                        ("Datum:", today_text),
                        ("Productie:", "Laser"),
                    ]
                    label_font = tkfont.Font(font=("TkDefaultFont", 10, "bold"))
                    value_font = tkfont.Font(font=("TkDefaultFont", 10))
                    line_height = max(
                        label_font.metrics("linespace"),
                        value_font.metrics("linespace"),
                    ) + 3
                    doc_right = title_x
                    for idx, (label_text, value_text) in enumerate(doc_lines):
                        line_y = doc_top_y + idx * line_height
                        preview_canvas.create_text(
                            title_x,
                            line_y,
                            anchor="nw",
                            text=label_text,
                            font=label_font,
                            fill=ORDER_TEXT_COLOR,
                        )
                        value_x = title_x + label_font.measure(label_text) + 4
                        value_item = preview_canvas.create_text(
                            value_x,
                            line_y,
                            anchor="nw",
                            text=value_text,
                            font=value_font,
                            fill=ORDER_TEXT_COLOR,
                        )
                        value_bbox = preview_canvas.bbox(value_item) or (
                            value_x,
                            line_y,
                            value_x + value_font.measure(value_text),
                            line_y + line_height,
                        )
                        doc_right = max(doc_right, value_bbox[2])
                    doc_bbox = (
                        title_x,
                        doc_top_y,
                        doc_right,
                        doc_top_y + (len(doc_lines) * line_height),
                    )

                    right_block_y = doc_bbox[1]
                    header_y = doc_bbox[3] + 8
                    inner_x = sheet_x + 28
                    inner_w = sheet_w - 56
                    col_gap = 18
                    left_col_w = inner_w * 0.58
                    right_col_w = inner_w - left_col_w - col_gap
                    left_x = inner_x
                    right_x = inner_x + left_col_w + col_gap

                    client_x = right_x
                    logo_x = right_x
                    logo_y = right_block_y
                    logo_box_w = max(96, int(right_col_w))
                    logo_box_h = 72
                    logo_img = load_logo_image(apply_crop=True)
                    logo_render_h = 0
                    if logo_img is not None and ImageTk is not None:
                        logo_photo = make_logo_photo(logo_img, (logo_box_w, logo_box_h))
                        if logo_photo is not None:
                            logo_render_h = logo_photo.height()
                            preview_canvas.preview_images.append(logo_photo)  # type: ignore[attr-defined]
                            preview_canvas.create_image(
                                logo_x,
                                logo_y,
                                anchor="nw",
                                image=logo_photo,
                            )
                    elif logo_path_var.get().strip():
                        preview_canvas.create_text(
                            logo_x,
                            logo_y + 10,
                            anchor="nw",
                            text="Logo",
                            fill="#9a9a9a",
                            font=("TkDefaultFont", 11, "italic"),
                        )
                        logo_render_h = 28

                    supplier_y = header_y
                    client_y = right_block_y + (logo_render_h + 6 if logo_render_h else 0)

                    supplier_text = (
                        "Besteld bij: Leverancier BV\n"
                        "Leveranciersstraat 12, 9000 Gent\n"
                        "BTW: BE0000000000\n"
                        "Contact sales: Voorbeeld Contact\n"
                        "E-mail: leverancier@example.com\n"
                        "Tel: +32 3 123 45 67"
                    )
                    supplier_item = preview_canvas.create_text(
                        left_x,
                        supplier_y,
                        anchor="nw",
                        text=supplier_text,
                        font=("TkDefaultFont", 10),
                        fill=ORDER_TEXT_COLOR,
                        width=left_col_w,
                    )
                    supplier_bbox = preview_canvas.bbox(supplier_item) or (
                        left_x,
                        supplier_y,
                        left_x,
                        supplier_y + 110,
                    )

                    client_lines = [
                        client_name,
                        client_address,
                        f"BTW: {client_vat}",
                        f"E-mail: {client_email}",
                    ]
                    if client_website:
                        client_lines.append(f"Website: {client_website}")
                    client_text = "\n".join(client_lines)
                    client_item = preview_canvas.create_text(
                        client_x,
                        client_y,
                        anchor="nw",
                        text=client_text,
                        font=("TkDefaultFont", 10),
                        fill=ORDER_TEXT_COLOR,
                        width=right_col_w,
                    )
                    client_bbox = preview_canvas.bbox(client_item) or (
                        client_x,
                        client_y,
                        client_x,
                        client_y + 90,
                    )

                    info_bottom = max(supplier_bbox[3], client_bbox[3])
                    delivery_y = info_bottom + 8
                    delivery_text = (
                        "Leveradres: Voorbeeld levering | Werfweg 25, 1000 Brussel"
                    )
                    delivery_item = preview_canvas.create_text(
                        inner_x,
                        delivery_y,
                        anchor="nw",
                        text=delivery_text,
                        font=("TkDefaultFont", 10),
                        fill=ORDER_TEXT_COLOR,
                        width=inner_w,
                    )
                    delivery_bbox = preview_canvas.bbox(delivery_item) or (
                        inner_x,
                        delivery_y,
                        inner_x,
                        delivery_y + 18,
                    )
                    delivery_bottom = delivery_bbox[3]

                    header_bottom = max(
                        right_block_y + (logo_render_h if logo_render_h else 0),
                        client_bbox[3],
                        delivery_bottom,
                    )
                    table_y = header_bottom + 24
                    table_x = inner_x
                    table_w = inner_w
                    col_fracs = [0.22, 0.40, 0.14, 0.06, 0.09, 0.09]
                    headers = ["PartNumber", "Omschrijving", "Materiaal", "St.", "m²", "kg"]
                    rows = [
                        ["PN-001", "Voetplaat voor voorbeeldbon", "S235JR", "2", "1,25", "4,80"],
                        ["PN-002", "Tweede regel om de layout te tonen", "RAL9005", "1", "", "2,10"],
                    ]
                    col_fracs = [0.07, 0.18, 0.37, 0.17, 0.07, 0.07, 0.07]
                    headers = [
                        "Nr.",
                        "PartNumber",
                        "Omschrijving",
                        "Materiaal",
                        "St.",
                        "m²",
                        "kg",
                    ]
                    rows = [
                        ["1", "PN-001", "Voetplaat voor voorbeeldbon", "S235JR", "2", "1,25", "4,80"],
                        ["2", "PN-002", "Tweede regel om de layout te tonen", "RAL9005", "1", "", "2,10"],
                    ]
                    row_h = 42
                    x_positions = [table_x]
                    for frac in col_fracs:
                        x_positions.append(x_positions[-1] + table_w * frac)

                    preview_canvas.create_rectangle(
                        table_x,
                        table_y,
                        table_x + table_w,
                        table_y + row_h,
                        fill=accent_color,
                        outline=ORDER_TABLE_OUTLINE_COLOR,
                    )
                    for idx, header in enumerate(headers):
                        x0 = x_positions[idx]
                        x1 = x_positions[idx + 1]
                        if idx:
                            preview_canvas.create_line(
                                x0,
                                table_y,
                                x0,
                                table_y + row_h,
                                fill=ORDER_TABLE_OUTLINE_COLOR,
                            )
                        if idx == 0:
                            preview_canvas.create_text(
                                (x0 + x1) / 2,
                                table_y + 12,
                                anchor="n",
                                text=header,
                                font=("TkDefaultFont", 10, "bold"),
                                fill=accent_text_color,
                            )
                        else:
                            preview_canvas.create_text(
                                x0 + 6,
                                table_y + 12,
                                anchor="nw",
                                text=header,
                                font=("TkDefaultFont", 10, "bold"),
                                fill=accent_text_color,
                                width=max(20, (x1 - x0) - 12),
                            )

                    for row_index, row_values in enumerate(rows, start=1):
                        y0 = table_y + row_h * row_index
                        y1 = y0 + row_h
                        fill = "#ffffff" if row_index % 2 else ORDER_TABLE_ALT_ROW_COLOR
                        preview_canvas.create_rectangle(
                            table_x,
                            y0,
                            table_x + table_w,
                            y1,
                            fill=fill,
                            outline=ORDER_TABLE_GRID_COLOR,
                        )
                        for idx, value in enumerate(row_values):
                            x0 = x_positions[idx]
                            x1 = x_positions[idx + 1]
                            if idx:
                                preview_canvas.create_line(
                                    x0,
                                    y0,
                                    x0,
                                    y1,
                                    fill=ORDER_TABLE_GRID_COLOR,
                                )
                            if idx == 0:
                                preview_canvas.create_text(
                                    (x0 + x1) / 2,
                                    y0 + 11,
                                    anchor="n",
                                    text=value,
                                    font=("TkDefaultFont", 9),
                                    fill=ORDER_TEXT_COLOR,
                                )
                            else:
                                is_numeric = idx >= 4
                                anchor = "ne" if is_numeric else "nw"
                                text_x = x1 - 6 if is_numeric else x0 + 6
                                preview_canvas.create_text(
                                    text_x,
                                    y0 + 11,
                                    anchor=anchor,
                                    text=value,
                                    font=("TkDefaultFont", 9),
                                    fill=ORDER_TEXT_COLOR,
                                    width=max(20, (x1 - x0) - 12),
                                )

                    footer_y = table_y + row_h * (len(rows) + 1) + 26
                    preview_canvas.create_text(
                        inner_x,
                        footer_y,
                        anchor="nw",
                        text=DEFAULT_FOOTER_NOTE,
                        font=("TkDefaultFont", 8),
                        fill=ORDER_MUTED_TEXT_COLOR,
                        width=inner_w,
                    )

                render_order_preview()

                buttons = tk.Frame(preview_win)
                buttons.pack(fill="x", padx=12, pady=(0, 12))
                tk.Button(
                    buttons, text="Vernieuwen", command=render_order_preview
                ).pack(side="left")
                tk.Button(
                    buttons, text="Sluiten", command=preview_win.destroy
                ).pack(side="right")

                _place_window_near_parent(preview_win, parent or win)
                preview_win.grab_set()
                preview_win.focus_set()

            def upload_logo() -> None:
                path = filedialog.askopenfilename(
                    filetypes=[
                        ("Afbeeldingen", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
                        ("Alle bestanden", "*.*"),
                    ]
                )
                if not path:
                    return
                dest_dir = runtime_asset_dir(CLIENT_LOGO_DIR)
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
                logo_path_var.set(to_runtime_relative_path(dest))
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

                max_w, max_h = 720, 420
                if base_img.width == 0 or base_img.height == 0:
                    messagebox.showerror(
                        "Fout", "Afbeelding heeft ongeldige afmetingen.", parent=win
                    )
                    crop_win.destroy()
                    return

                fit_scale = min(max_w / base_img.width, max_h / base_img.height)
                fit_scale = max(0.25, min(4.0, fit_scale if fit_scale > 0 else 1.0))
                zoom_var = tk.IntVar(value=int(round(fit_scale * 100)))
                zoom_label_var = tk.StringVar(value=f"{zoom_var.get()}%")

                current_box = [0.0, 0.0, float(base_img.width), float(base_img.height)]
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
                    current_box = [float(left), float(top), float(right), float(bottom)]

                canvas_frame = tk.Frame(crop_win)
                canvas_frame.pack(fill="both", expand=True, padx=8, pady=(8, 4))
                canvas_frame.grid_columnconfigure(0, weight=1)
                canvas_frame.grid_rowconfigure(0, weight=1)

                viewport_w = min(max_w, max(320, int(round(base_img.width * fit_scale))))
                viewport_h = min(max_h, max(180, int(round(base_img.height * fit_scale))))
                canvas = tk.Canvas(
                    canvas_frame,
                    width=viewport_w,
                    height=viewport_h,
                    bg="#1f1f1f",
                    highlightthickness=1,
                    highlightbackground="#c8c8c8",
                    cursor="crosshair",
                )
                x_scroll = tk.Scrollbar(
                    canvas_frame, orient="horizontal", command=canvas.xview
                )
                y_scroll = tk.Scrollbar(
                    canvas_frame, orient="vertical", command=canvas.yview
                )
                canvas.configure(
                    xscrollcommand=x_scroll.set,
                    yscrollcommand=y_scroll.set,
                )
                canvas.grid(row=0, column=0, sticky="nsew")
                y_scroll.grid(row=0, column=1, sticky="ns")
                x_scroll.grid(row=1, column=0, sticky="ew")

                status_var = tk.StringVar(value="")
                tk.Label(
                    crop_win,
                    text=(
                        "Klik en sleep om een selectie te maken. Gebruik het "
                        "muiswiel of de zoomknoppen voor fijnafstelling. "
                        "De PDF schaalt het gekozen deel automatisch in een vaste logozone."
                    ),
                    anchor="w",
                    justify="left",
                    wraplength=720,
                ).pack(fill="x", padx=8, pady=(0, 2))
                tk.Label(
                    crop_win,
                    textvariable=status_var,
                    anchor="w",
                    justify="left",
                    wraplength=720,
                ).pack(fill="x", padx=8, pady=(0, 6))

                controls = tk.Frame(crop_win)
                controls.pack(fill="x", padx=8, pady=(0, 4))
                tk.Button(
                    controls,
                    text="-",
                    width=3,
                    command=lambda: set_zoom(zoom_var.get() - 10),
                ).pack(side="left")
                tk.Button(
                    controls,
                    text="+",
                    width=3,
                    command=lambda: set_zoom(zoom_var.get() + 10),
                ).pack(side="left", padx=(4, 8))
                tk.Label(controls, text="Zoom").pack(side="left")
                zoom_scale = tk.Scale(
                    controls,
                    from_=25,
                    to=400,
                    orient="horizontal",
                    variable=zoom_var,
                    showvalue=False,
                    resolution=5,
                    length=220,
                )
                zoom_scale.pack(side="left", padx=(6, 6))
                tk.Label(controls, textvariable=zoom_label_var, width=6, anchor="w").pack(
                    side="left"
                )
                tk.Button(
                    controls,
                    text="Passend",
                    command=lambda: set_zoom(int(round(fit_scale * 100))),
                ).pack(side="left", padx=(8, 0))

                image_id = None
                rect_id = None
                start_point = [0.0, 0.0]

                def current_scale() -> float:
                    return max(0.25, min(4.0, zoom_var.get() / 100.0))

                def update_status() -> None:
                    width_px = max(0, int(round(current_box[2] - current_box[0])))
                    height_px = max(0, int(round(current_box[3] - current_box[1])))
                    if (
                        current_box[0] <= 0
                        and current_box[1] <= 0
                        and current_box[2] >= base_img.width
                        and current_box[3] >= base_img.height
                    ):
                        status_var.set(
                            "Volledige afbeelding geselecteerd. Gebruik zoom om nauwkeuriger te kaderen."
                        )
                    else:
                        status_var.set(
                            f"Geselecteerd: {width_px} x {height_px} px. "
                            "De PDF schaalt dit kader automatisch."
                        )

                def draw_rect() -> None:
                    nonlocal rect_id
                    scale = current_scale()
                    x0 = current_box[0] * scale
                    y0 = current_box[1] * scale
                    x1 = current_box[2] * scale
                    y1 = current_box[3] * scale
                    if rect_id is None:
                        rect_id = canvas.create_rectangle(
                            x0,
                            y0,
                            x1,
                            y1,
                            outline="#ff007f",
                            width=2,
                        )
                    else:
                        canvas.coords(rect_id, x0, y0, x1, y1)
                    canvas.tag_raise(rect_id)
                    update_status()

                def render_image(*_ignored) -> None:
                    nonlocal image_id
                    scale = current_scale()
                    disp_w = max(1, int(round(base_img.width * scale)))
                    disp_h = max(1, int(round(base_img.height * scale)))
                    if scale != 1.0:
                        disp_img = base_img.resize(
                            (disp_w, disp_h),
                            RESAMPLE or Image.BICUBIC,  # type: ignore[union-attr]
                        )
                    else:
                        disp_img = base_img.copy()
                    photo = ImageTk.PhotoImage(disp_img)  # type: ignore[union-attr]
                    canvas.image = photo  # type: ignore[attr-defined]
                    if image_id is None:
                        image_id = canvas.create_image(0, 0, anchor="nw", image=photo)
                    else:
                        canvas.itemconfigure(image_id, image=photo)
                    canvas.configure(scrollregion=(0, 0, disp_w, disp_h))
                    zoom_label_var.set(f"{zoom_var.get()}%")
                    draw_rect()

                def set_zoom(value: int) -> None:
                    zoom_var.set(max(25, min(400, int(value))))
                    render_image()

                def clamp_to_image(x: float, y: float) -> tuple[float, float]:
                    return (
                        max(0.0, min(float(base_img.width), x)),
                        max(0.0, min(float(base_img.height), y)),
                    )

                def image_coords(evt) -> tuple[float, float]:
                    scale = current_scale()
                    return clamp_to_image(
                        canvas.canvasx(evt.x) / scale,
                        canvas.canvasy(evt.y) / scale,
                    )

                def update_box(x0: float, y0: float, x1: float, y1: float) -> None:
                    left = min(x0, x1)
                    top = min(y0, y1)
                    right = max(x0, x1)
                    bottom = max(y0, y1)
                    if right - left < 1 or bottom - top < 1:
                        return
                    current_box[0] = left
                    current_box[1] = top
                    current_box[2] = right
                    current_box[3] = bottom
                    draw_rect()

                def on_press(evt):
                    start_point[0], start_point[1] = image_coords(evt)
                    update_box(
                        start_point[0],
                        start_point[1],
                        start_point[0] + 1,
                        start_point[1] + 1,
                    )

                def on_drag(evt):
                    x, y = image_coords(evt)
                    update_box(start_point[0], start_point[1], x, y)

                canvas.bind("<Button-1>", on_press)
                canvas.bind("<B1-Motion>", on_drag)
                canvas.bind("<ButtonRelease-1>", on_drag)

                btns = tk.Frame(crop_win)
                btns.pack(pady=6)

                def reset_full() -> None:
                    current_box[0] = 0.0
                    current_box[1] = 0.0
                    current_box[2] = float(base_img.width)
                    current_box[3] = float(base_img.height)
                    draw_rect()

                def apply_crop() -> None:
                    nonlocal logo_crop_state
                    left = int(round(current_box[0]))
                    top = int(round(current_box[1]))
                    right = int(round(current_box[2]))
                    bottom = int(round(current_box[3]))
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

                def on_mousewheel(evt):
                    step = 10 if evt.delta > 0 else -10
                    set_zoom(zoom_var.get() + step)
                    return "break"

                canvas.bind("<MouseWheel>", on_mousewheel)
                zoom_scale.configure(command=lambda _value: render_image())
                render_image()

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
            tk.Button(logo_frame, text="Volledig logo", command=open_logo_preview).grid(
                row=2, column=1, sticky="ew", padx=4, pady=2
            )
            tk.Button(logo_frame, text="Voorbeeld bon", command=open_order_preview).grid(
                row=3, column=1, sticky="ew", padx=4, pady=2
            )
            tk.Button(logo_frame, text="Verwijder", command=clear_logo).grid(
                row=4, column=1, sticky="ew", padx=4, pady=2
            )
            tk.Label(
                logo_frame,
                text=(
                    "De PDF schaalt het logo automatisch binnen een vaste kopzone. "
                    "Gebruik bijsnijden om alleen het juiste deel te tonen."
                ),
                anchor="w",
                justify="left",
                wraplength=240,
            ).grid(row=5, column=1, sticky="nw", padx=4, pady=(6, 4))

            preview_label.bind("<Button-1>", lambda _e: open_logo_preview())

            update_preview()

            def _save():
                rec = {k: e.get().strip() for k, e in entries.items()}
                accent_raw = accent_color_var.get().strip()
                accent_normalized = normalize_rgb_color(accent_raw)
                if accent_raw and not accent_normalized:
                    messagebox.showwarning(
                        "Let op",
                        "Gebruik RGB als 255,119,255 of hex als #FF77FF.",
                        parent=win,
                    )
                    return
                rec["favorite"] = fav_var.get()
                rec["accent_color"] = accent_normalized
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
            btnf.grid(row=len(fields) + 4, column=0, columnspan=2, pady=6)
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
            fav_label = f"Favoriet {favorite_marker()}"
            tk.Button(btns, text=fav_label, command=self.toggle_fav_sel).pack(side="left", padx=4)
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
            return strip_favorite_marker(vals[0])

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
            
            # Search bar
            search = tk.Frame(self)
            search.pack(fill="x", padx=8, pady=(8, 0))
            tk.Label(search, text="Zoek:").pack(side="left")
            self.search_var = tk.StringVar()
            entry = tk.Entry(search, textvariable=self.search_var)
            entry.pack(side="left", fill="x", expand=True)
            self.search_var.trace_add("write", lambda *_: self.refresh())
            
            # Filter frame
            filter_frame = tk.Frame(self)
            filter_frame.pack(fill="x", padx=8, pady=(4, 0))
            
            # Product type filter
            tk.Label(filter_frame, text="Product type:").pack(side="left", padx=(0, 4))
            self.product_type_var = tk.StringVar()
            self.product_type_combo = ttk.Combobox(filter_frame, textvariable=self.product_type_var, width=20, state="readonly")
            self.product_type_combo.pack(side="left", padx=(0, 12))
            # When product type changes, update description dropdown and refresh
            self.product_type_var.trace_add("write", lambda *_: self._on_product_type_changed())
            
            # Product description filter
            tk.Label(filter_frame, text="Beschrijving:").pack(side="left", padx=(0, 4))
            self.product_desc_var = tk.StringVar()
            self.product_desc_combo = ttk.Combobox(filter_frame, textvariable=self.product_desc_var, width=30, state="readonly")
            self.product_desc_combo.pack(side="left", padx=(0, 12))
            # When description changes, only refresh the table (not filter options)
            self.product_desc_var.trace_add("write", lambda *_: self._refresh_table_only())
            
            # Clear filters button
            tk.Button(filter_frame, text="Wis filters", command=self._clear_filters).pack(side="left")
            
            # Treeview with new columns
            cols = ("Supplier", "Product type", "Beschrijving", "BTW", "E-mail", "Tel", "Adres 1", "Adres 2")
            self.tree = ttk.Treeview(self, columns=cols, show="headings")
            for c in cols:
                self.tree.heading(c, text=c)
                self.tree.column(c, anchor="w")
            self.tree.pack(fill="both", expand=True, padx=8, pady=8)
            # Bind double-click to edit
            self.tree.bind("<Double-Button-1>", lambda e: self.edit_sel())
            # Bind Delete key to remove selected supplier when tree has focus
            self.tree.bind("<Delete>", lambda e: self.remove_sel())
            
            btns = tk.Frame(self)
            btns.pack(fill="x")
            tk.Button(btns, text="Toevoegen", command=self.add_supplier).pack(side="left", padx=4)
            tk.Button(btns, text="Bewerken", command=self.edit_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Verwijderen", command=self.remove_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Update uit CSV (merge)", command=self.merge_csv).pack(side="left", padx=4)
            tk.Button(
                btns,
                text="Laad voorbeeldlijst",
                command=self.load_example_template,
            ).pack(side="left", padx=4)
            fav_label = f"Favoriet {favorite_marker()}"
            tk.Button(btns, text=fav_label, command=self.toggle_fav_sel).pack(side="left", padx=4)
            # If DB is empty, load the built-in example supplier template silently
            if not self.db.suppliers:
                try:
                    template_path = bundle_root() / SUPPLIERS_TEMPLATE_FILE
                    if template_path.exists():
                        self._merge_suppliers_from_csv_path(template_path)
                except Exception:
                    # Fail silently; leave DB empty if loading template fails
                    pass

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

        def _clear_filters(self):
            """Clear all filter selections."""
            self.product_type_var.set("")
            self.product_desc_var.set("")
            self._update_filter_options()

        def _on_product_type_changed(self):
            """Called when product type selection changes.
            Updates the description dropdown to show only descriptions for this type."""
            # Get the newly selected product type
            product_type = self.product_type_var.get()
            
            # Clear the description filter when product type changes
            self.product_desc_var.set("")
            
            # Update the description dropdown options for this product type
            product_descs = self.db.get_product_descriptions_for_type(product_type)
            self.product_desc_combo['values'] = [""] + product_descs
            
            # Refresh only the table results (not the dropdown options)
            self._refresh_table_only()

        def _update_filter_options(self):
            """Update filter dropdown options based on current data."""
            product_types = self.db.get_unique_product_types()
            
            # Keep current selection if still valid
            current_type = self.product_type_var.get()
            
            self.product_type_combo['values'] = [""] + product_types
            
            # Restore selection if still valid
            if current_type in ([""] + product_types):
                self.product_type_var.set(current_type)
            
            # Also update the description dropdown for the current product type
            product_type = self.product_type_var.get()
            product_descs = self.db.get_product_descriptions_for_type(product_type)
            current_desc = self.product_desc_var.get()
            self.product_desc_combo['values'] = [""] + product_descs
            
            # Restore description selection if still valid
            if current_desc in ([""] + product_descs):
                self.product_desc_var.set(current_desc)
            else:
                self.product_desc_var.set("")

        def _refresh_table_only(self):
            """Refresh only the table results without updating dropdown options."""
            for r in self.tree.get_children():
                self.tree.delete(r)
            q = self.search_var.get()
            product_type_filter = self.product_type_var.get()
            product_desc_filter = self.product_desc_var.get()
            sups = self.db.find(q, product_type_filter, product_desc_filter)
            fav_prefix = favorite_prefix()
            for i, s in enumerate(sups):
                vals = (
                    (fav_prefix if s.favorite else "") + (s.supplier or ""),
                    s.product_type or "",
                    s.product_description or "",
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

        def refresh(self):
            self._update_filter_options()
            self._refresh_table_only()

        def _sel_name(self):
            sel = self.tree.selection()
            if not sel:
                return None
            name = self.tree.item(sel[0], "values")[0]
            return strip_favorite_marker(name)

        def _sel_supplier(self) -> Optional[Supplier]:
            n = self._sel_name()
            if not n:
                return None
            for s in self.db.suppliers:
                if s.supplier == n:
                    return s
            return None

        def add_supplier(self):
            dlg = self._EditDialog(
                self,
                Supplier(supplier=""),
                title="Nieuwe leverancier",
            )
            self.wait_window(dlg)
            if dlg.result:
                self.db.upsert(dlg.result)
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

        def _merge_suppliers_from_csv_path(self, path) -> int:
            df = read_csv_flex(str(path))
            changed = 0
            for rec in df.to_dict(orient="records"):
                try:
                    sup = Supplier.from_any(rec)
                    self.db.upsert(sup)
                    changed += 1
                except Exception:
                    pass
            self.db.save(SUPPLIERS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()
            return changed

        def merge_csv(self):
            path = filedialog.askopenfilename(
                parent=self,
                title="CSV bestand",
                filetypes=[("CSV", "*.csv"), ("Alle bestanden", "*.*")],
            )
            if not path:
                return
            try:
                changed = self._merge_suppliers_from_csv_path(path)
                messagebox.showinfo(
                    "Import",
                    f"Verwerkt (merge/upsert): {changed}",
                    parent=self,
                )
            except Exception as e:
                messagebox.showerror("Fout", str(e), parent=self)

        def load_example_template(self):
            template_path = bundle_root() / SUPPLIERS_TEMPLATE_FILE
            if not template_path.exists():
                messagebox.showerror(
                    "Template ontbreekt",
                    "De ingebouwde leveranciers-template werd niet gevonden.",
                    parent=self,
                )
                return
            if not messagebox.askyesno(
                "Voorbeeldlijst laden",
                (
                    "Laad de ingebouwde voorbeeldlijst met leveranciers?\n\n"
                    "Bestaande leveranciers met dezelfde naam worden bijgewerkt."
                ),
                parent=self,
            ):
                return
            try:
                changed = self._merge_suppliers_from_csv_path(template_path)
                messagebox.showinfo(
                    "Voorbeeldlijst geladen",
                    f"Verwerkt (merge/upsert): {changed}",
                    parent=self,
                )
            except Exception as e:
                messagebox.showerror("Fout", str(e), parent=self)

        class _EditDialog(tk.Toplevel):
            def __init__(
                self,
                master,
                supplier: Supplier,
                *,
                title: str = "Leverancier bewerken",
            ):
                super().__init__(master)
                self.title(title)
                self.result = None
                fields = [
                    ("supplier", "Naam"),
                    ("description", "Beschrijving"),
                    ("product_type", "Product type"),
                    ("product_description", "Product beschrijving"),
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
                first_entry = None
                for i, (f, lbl) in enumerate(fields):
                    tk.Label(self, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=4, pady=2)
                    var = tk.StringVar(value=getattr(supplier, f) or "")
                    entry = tk.Entry(self, textvariable=var, width=40)
                    entry.grid(row=i, column=1, padx=4, pady=2)
                    self.vars[f] = var
                    if first_entry is None:
                        first_entry = entry
                btn = tk.Frame(self)
                btn.grid(row=len(fields), column=0, columnspan=2, pady=4)
                tk.Button(btn, text="Opslaan", command=self._ok).pack(side="left", padx=4)
                tk.Button(btn, text="Annuleer", command=self.destroy).pack(side="left", padx=4)
                self.transient(master)
                _place_window_near_parent(self, master)
                self.grab_set()
                if first_entry is not None:
                    self.after_idle(first_entry.focus_set)

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

    class PresetRulesManagerFrame(tk.Frame):
        KIND_LABELS = {
            "production": "Productie",
            "finish": "Afwerking",
            "opticutter": "Brutemateriaal",
        }
        KIND_VALUES = {
            "Productie": "production",
            "Afwerking": "finish",
            "Brutemateriaal": "opticutter",
        }
        CLIENT_ALL_LABEL = "(alle opdrachtgevers)"
        FILTER_ALL_LABEL = "(alle klanten)"
        FILTER_GLOBAL_LABEL = "(algemene regels)"
        SUPPLIER_UNCHANGED_LABEL = "(ongewijzigd)"
        DELIVERY_UNCHANGED_LABEL = "(ongewijzigd)"
        DOC_TYPE_UNCHANGED_LABEL = "(ongewijzigd)"
        EN1090_UNCHANGED_LABEL = "(ongewijzigd)"

        def __init__(
            self,
            master,
            db,
            clients_db,
            suppliers_db,
            delivery_db,
            on_change=None,
        ):
            super().__init__(master)
            self.configure(padx=12, pady=12)
            self.db = db
            self.clients_db = clients_db
            self.suppliers_db = suppliers_db
            self.delivery_db = delivery_db
            self.on_change = on_change

            intro = tk.Label(
                self,
                text=(
                    "Presetregels vullen leveranciers- en bestelbonvelden automatisch in "
                    "op basis van opdrachtgever en productie/afwerking. "
                    "Ze blijven altijd handmatig aanpasbaar."
                ),
                justify="left",
                anchor="w",
                wraplength=760,
            )
            intro.pack(fill="x", padx=8, pady=(0, 8))

            filter_row = tk.Frame(self)
            filter_row.pack(fill="x", padx=8, pady=(0, 8))
            tk.Label(filter_row, text="Klantfilter:").pack(side="left")
            self.filter_client_var = tk.StringVar(value=self.FILTER_ALL_LABEL)
            self._filter_display_to_value: Dict[str, Optional[str]] = {}
            self.filter_client_combo = ttk.Combobox(
                filter_row,
                textvariable=self.filter_client_var,
                state="readonly",
                width=34,
            )
            self.filter_client_combo.pack(side="left", padx=(6, 8))
            self.filter_client_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh())
            _HelpTooltip(
                self.filter_client_combo,
                "Filter de lijst per opdrachtgever. Algemene regels staan apart onder '(algemene regels)'.",
            )
            self.filter_count_var = tk.StringVar()
            tk.Label(filter_row, textvariable=self.filter_count_var, anchor="w").pack(side="left")
            self.advanced_view_var = tk.IntVar(value=0)
            advanced_view_check = tk.Checkbutton(
                filter_row,
                text="Geavanceerde weergave",
                variable=self.advanced_view_var,
                command=self.refresh,
            )
            advanced_view_check.pack(side="right")
            _HelpTooltip(
                advanced_view_check,
                "Toon ook prioriteit en automatisch toepassen. Normaal hoef je deze kolommen niet te gebruiken.",
            )

            list_area = tk.Frame(self)
            list_area.pack(fill="both", expand=True, padx=8, pady=(0, 8))

            cols = (
                "Naam",
                "Klant",
                "Type",
                "Selecties",
                "Acties",
                "Status",
                "Prioriteit",
                "Auto",
                "Actief",
            )
            self._tree_columns = cols
            self._tree_widths = {
                "Naam": 210,
                "Klant": 160,
                "Type": 110,
                "Selecties": 240,
                "Acties": 380,
                "Status": 120,
                "Prioriteit": 80,
                "Auto": 60,
                "Actief": 70,
            }
            self._advanced_tree_columns = {"Prioriteit", "Auto"}
            self.tree = ttk.Treeview(list_area, columns=cols, show="headings", selectmode="browse")
            list_area.rowconfigure(0, weight=1)
            list_area.columnconfigure(0, weight=1)
            for col in cols:
                anchor = "center" if col in {"Status", "Prioriteit", "Auto", "Actief"} else "w"
                self.tree.heading(col, text=col, anchor=anchor)
                self.tree.column(col, width=self._tree_widths.get(col, 140), anchor=anchor)
            self._configure_tree_columns()
            tree_y_scroll = ttk.Scrollbar(list_area, orient="vertical", command=self.tree.yview)
            tree_x_scroll = ttk.Scrollbar(list_area, orient="horizontal", command=self.tree.xview)
            self.tree.configure(
                yscrollcommand=tree_y_scroll.set,
                xscrollcommand=tree_x_scroll.set,
            )
            self.tree.grid(row=0, column=0, sticky="nsew")
            tree_y_scroll.grid(row=0, column=1, sticky="ns")
            tree_x_scroll.grid(row=1, column=0, sticky="ew")
            self.tree_tooltips = _TreeTooltipManager(self.tree)
            self.tree.bind("<Double-Button-1>", lambda _e: self.edit_sel())
            self.tree.bind("<<TreeviewSelect>>", lambda _e: self._on_tree_selection_changed())

            btns = tk.Frame(self)
            btns.pack(fill="x", padx=8, pady=(0, 4))
            add_btn = tk.Button(btns, text="Toevoegen", command=self.add_rule)
            add_btn.pack(side="left", padx=4)
            _HelpTooltip(add_btn, "Maak een nieuwe presetregel.")
            edit_btn = tk.Button(btns, text="Bewerken", command=self.edit_sel)
            edit_btn.pack(side="left", padx=4)
            _HelpTooltip(edit_btn, "Open de geselecteerde regel om velden aan te passen.")
            duplicate_btn = tk.Button(btns, text="Dupliceren", command=self.duplicate_sel)
            duplicate_btn.pack(side="left", padx=4)
            _HelpTooltip(
                duplicate_btn,
                "Maak een kopie van de geselecteerde regel. Handig als alleen selectie of leverancier verschilt.",
            )
            remove_btn = tk.Button(btns, text="Verwijderen", command=self.remove_sel)
            remove_btn.pack(side="left", padx=4)
            _HelpTooltip(remove_btn, "Verwijder de geselecteerde presetregel.")
            toggle_btn = tk.Button(btns, text="Actief wisselen", command=self.toggle_sel)
            toggle_btn.pack(side="left", padx=4)
            _HelpTooltip(toggle_btn, "Zet de geselecteerde regel tijdelijk actief of inactief.")
            check_btn = tk.Button(btns, text="Controleer regels", command=self.check_rules)
            check_btn.pack(side="left", padx=4)
            _HelpTooltip(
                check_btn,
                "Controleer alle presetregels op overlap, ontbrekende leveranciers en ontbrekende invulacties.",
            )
            copy_client_btn = tk.Button(
                btns,
                text="Kopieer naar klant",
                command=self.copy_client_rules,
            )
            copy_client_btn.pack(side="left", padx=4)
            _HelpTooltip(
                copy_client_btn,
                "Kopieer alle regels van een opdrachtgever naar een andere opdrachtgever.",
            )

            detail = tk.LabelFrame(self, text="Regeldetails", labelanchor="n")
            detail.pack(fill="x", padx=8, pady=(4, 8))
            self.detail_var = tk.StringVar(value="Selecteer een presetregel voor details.")
            tk.Label(
                detail,
                textvariable=self.detail_var,
                anchor="w",
                justify="left",
                wraplength=980,
            ).pack(fill="x", padx=6, pady=6)

            preview = tk.LabelFrame(self, text="Preset preview / test", labelanchor="n")
            preview.pack(fill="x", padx=8, pady=(4, 0))
            preview.columnconfigure(1, weight=1)
            preview.columnconfigure(3, weight=1)

            client_values = [self.CLIENT_ALL_LABEL]
            client_values.extend(
                self.clients_db.display_name(client)
                for client in self.clients_db.clients_sorted()
            )
            self.preview_client_var = tk.StringVar(value=self.CLIENT_ALL_LABEL)
            self.preview_kind_var = tk.StringVar(value="Productie")
            self.preview_identifier_var = tk.StringVar()
            self.preview_auto_only_var = tk.IntVar(value=0)
            self.preview_result_var = tk.StringVar(
                value="Kies een context en klik op 'Test presetmatch'."
            )

            tk.Label(preview, text="Opdrachtgever:").grid(
                row=0, column=0, sticky="w", padx=6, pady=(6, 2)
            )
            self.preview_client_combo = ttk.Combobox(
                preview,
                textvariable=self.preview_client_var,
                values=client_values,
                state="readonly",
                width=32,
            )
            self.preview_client_combo.grid(row=0, column=1, sticky="ew", padx=6, pady=(6, 2))
            _HelpTooltip(self.preview_client_combo, "Kies de opdrachtgever om een presetmatch te testen.")

            tk.Label(preview, text="Type selectie:").grid(
                row=0, column=2, sticky="w", padx=6, pady=(6, 2)
            )
            self.preview_kind_combo = ttk.Combobox(
                preview,
                textvariable=self.preview_kind_var,
                values=list(self.KIND_VALUES.keys()),
                state="readonly",
                width=18,
            )
            self.preview_kind_combo.grid(row=0, column=3, sticky="w", padx=6, pady=(6, 2))
            _HelpTooltip(self.preview_kind_combo, "Kies of je een productie-, afwerkings- of brutemateriaalregel wil testen.")

            tk.Label(preview, text="Selectie:").grid(
                row=1, column=0, sticky="w", padx=6, pady=2
            )
            self.preview_identifier_entry = tk.Entry(
                preview,
                textvariable=self.preview_identifier_var,
                width=36,
            )
            self.preview_identifier_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=2)
            _HelpTooltip(self.preview_identifier_entry, "Vul exact de selectie in, bijvoorbeeld Cutting of Finish-Galvanised.")
            tk.Label(
                preview,
                text="Voorbeeld: Laser cutting, Tube laser cutting, Poedercoaten",
                anchor="w",
                justify="left",
            ).grid(row=2, column=1, columnspan=3, sticky="w", padx=6, pady=(0, 6))

            preview_actions = tk.Frame(preview)
            preview_actions.grid(row=1, column=2, columnspan=2, sticky="w", padx=6, pady=2)
            tk.Checkbutton(
                preview_actions,
                text="Alleen auto-regels",
                variable=self.preview_auto_only_var,
            ).pack(side="left")
            preview_test_btn = tk.Button(
                preview_actions,
                text="Test presetmatch",
                command=self._run_preview,
            )
            preview_test_btn.pack(side="left", padx=(12, 0))
            _HelpTooltip(preview_test_btn, "Toon welke regels matchen en welke velden ingevuld worden.")

            tk.Label(
                preview,
                textvariable=self.preview_result_var,
                anchor="w",
                justify="left",
                wraplength=760,
            ).grid(row=3, column=0, columnspan=4, sticky="ew", padx=6, pady=(2, 8))

            self.refresh()
            self.after_idle(self._seed_preview_from_selection)

        def refresh_data(self) -> None:
            self.refresh()

        @staticmethod
        def _key(value: object) -> str:
            return _to_str(value).strip().casefold()

        def _refresh_client_filter_options(self) -> None:
            current = self.filter_client_var.get().strip()
            mapping: Dict[str, Optional[str]] = {
                self.FILTER_ALL_LABEL: None,
                self.FILTER_GLOBAL_LABEL: "",
            }
            values = [self.FILTER_ALL_LABEL, self.FILTER_GLOBAL_LABEL]
            for client in self.clients_db.clients_sorted():
                display = self.clients_db.display_name(client)
                values.append(display)
                mapping[display] = strip_favorite_marker(display).strip()
            self._filter_display_to_value = mapping
            self.filter_client_combo.configure(values=values)
            if current not in mapping:
                self.filter_client_var.set(self.FILTER_ALL_LABEL)

        def _current_filter_client(self) -> Optional[str]:
            value = self.filter_client_var.get().strip()
            return self._filter_display_to_value.get(value, None)

        def _rule_visible_for_filter(self, rule: OrderPresetRule) -> bool:
            filter_client = self._current_filter_client()
            if filter_client is None:
                return True
            return self._key(rule.client) == self._key(filter_client)

        def _default_client_for_new_rule(self) -> str:
            filter_client = self._current_filter_client()
            return filter_client or ""

        def _configure_tree_columns(self) -> None:
            if not hasattr(self, "tree"):
                return
            advanced = bool(self.advanced_view_var.get()) if hasattr(self, "advanced_view_var") else False
            for col in self._tree_columns:
                anchor = "center" if col in {"Status", "Prioriteit", "Auto", "Actief"} else "w"
                if col in self._advanced_tree_columns and not advanced:
                    self.tree.heading(col, text="", anchor=anchor)
                    self.tree.column(
                        col,
                        width=0,
                        minwidth=0,
                        stretch=False,
                        anchor=anchor,
                    )
                    continue
                self.tree.heading(col, text=col, anchor=anchor)
                self.tree.column(
                    col,
                    width=self._tree_widths.get(col, 140),
                    minwidth=35,
                    stretch=col not in {"Status", "Prioriteit", "Auto", "Actief"},
                    anchor=anchor,
                )

        def _selection_suggestions_by_kind(self) -> Dict[str, List[str]]:
            suggestions: Dict[str, set[str]] = {
                "production": set(),
                "finish": set(),
                "opticutter": set(),
            }
            display_values: Dict[str, str] = {}
            for rule in self.db.rules:
                kind = rule.selection_kind or "production"
                if kind not in suggestions:
                    suggestions[kind] = set()
                for identifier in rule.identifiers:
                    text = _to_str(identifier).strip()
                    if not text:
                        continue
                    key = self._key(text)
                    suggestions[kind].add(key)
                    display_values.setdefault(f"{kind}:{key}", text)
            return {
                kind: sorted(
                    (display_values.get(f"{kind}:{key}", key) for key in keys),
                    key=lambda value: value.casefold(),
                )
                for kind, keys in suggestions.items()
            }

        def _known_supplier_names(self) -> set[str]:
            return {
                self._key(supplier.supplier)
                for supplier in self.suppliers_db.suppliers
                if _to_str(supplier.supplier).strip()
            }

        def _known_delivery_values(self) -> set[str]:
            values = {
                self._key("Geen"),
                self._key("Klantadres"),
                self._key("Opdrachtgeveradres"),
                self._key("Bestelling wordt opgehaald"),
                self._key("Leveradres wordt nog meegedeeld"),
            }
            for delivery in self.delivery_db.addresses:
                name = _to_str(delivery.name).strip()
                if name:
                    values.add(self._key(name))
                display = _to_str(self.delivery_db.display_name(delivery)).strip()
                if display:
                    values.add(self._key(strip_favorite_marker(display).strip()))
            for client in self.clients_db.clients:
                name = _to_str(client.name).strip()
                if name:
                    values.add(self._key(name))
            return values

        def _unknown_supplier(self, rule: OrderPresetRule) -> bool:
            supplier = _to_str(rule.supplier).strip()
            return bool(supplier and self._key(supplier) not in self._known_supplier_names())

        def _unknown_delivery(self, rule: OrderPresetRule) -> bool:
            delivery = _to_str(rule.delivery).strip()
            return bool(delivery and self._key(delivery) not in self._known_delivery_values())

        def _rule_status_issues(
            self,
            rule: OrderPresetRule,
            *,
            old_name: Optional[str] = None,
        ) -> List[str]:
            if not rule.enabled:
                return ["Uitgeschakeld"]

            issues: List[str] = []
            if not self._rule_action_fields(rule):
                issues.append("Geen invulactie")
            if self._unknown_supplier(rule):
                issues.append(f"Leverancier niet in lijst: {rule.supplier}")
            if self._unknown_delivery(rule):
                issues.append(f"Leveradres niet in lijst: {rule.delivery}")

            conflicts = self._conflict_messages_for_rule(rule, old_name=old_name or rule.name)
            if conflicts:
                issues.append(f"Overlap met {len(conflicts)} regel(s)")
            return issues

        def _rule_status_text(self, rule: OrderPresetRule) -> str:
            issues = self._rule_status_issues(rule, old_name=rule.name)
            if not issues:
                return "OK"
            if "Uitgeschakeld" in issues:
                return "Uit"
            for issue, label in (
                ("Overlap", "Overlap"),
                ("Leverancier", "Leverancier?"),
                ("Leveradres", "Leveradres?"),
                ("Geen invulactie", "Geen actie"),
            ):
                if any(text.startswith(issue) for text in issues):
                    return label
            return "Controleer"

        def _rule_status_detail(self, rule: OrderPresetRule) -> str:
            issues = self._rule_status_issues(rule, old_name=rule.name)
            if not issues:
                return "OK: geen directe problemen gevonden."
            lines = list(issues)
            conflicts = self._conflict_messages_for_rule(rule, old_name=rule.name)
            if conflicts:
                shown = conflicts[:5]
                lines.extend(f"Overlap: {message}" for message in shown)
                if len(conflicts) > len(shown):
                    lines.append(f"Overlap: ... en {len(conflicts) - len(shown)} andere")
            return "\n".join(lines)

        def refresh(self) -> None:
            self._refresh_client_filter_options()
            self._configure_tree_columns()
            for item in self.tree.get_children():
                self.tree.delete(item)
            tooltip_manager = getattr(self, "tree_tooltips", None)
            if tooltip_manager is not None:
                tooltip_manager.clear()
            visible_rules = [
                rule for rule in self.db.rules_sorted() if self._rule_visible_for_filter(rule)
            ]
            for idx, rule in enumerate(visible_rules):
                tag = "odd" if idx % 2 == 0 else "even"
                kind_label = self.KIND_LABELS.get(rule.selection_kind, rule.selection_kind)
                client_label = rule.client or self.CLIENT_ALL_LABEL
                selection_text = rule.selection_summary()
                action_text = rule.action_summary()
                status_text = self._rule_status_text(rule)
                values = (
                    rule.name,
                    client_label,
                    kind_label,
                    selection_text,
                    action_text,
                    status_text,
                    str(rule.priority),
                    "Ja" if rule.auto_apply else "Nee",
                    "Ja" if rule.enabled else "Nee",
                )
                self.tree.insert("", "end", iid=rule.name, values=values, tags=(tag,))
                if tooltip_manager is not None:
                    tooltip_manager.set(rule.name, "#1", rule.name)
                    tooltip_manager.set(rule.name, "#4", selection_text)
                    tooltip_manager.set(rule.name, "#5", action_text)
                    tooltip_manager.set(rule.name, "#6", self._rule_status_detail(rule))
            total = len(self.db.rules)
            visible = len(visible_rules)
            self.filter_count_var.set(f"{visible} van {total} regel(s)")
            self.tree.tag_configure("odd", background=TREE_ODD_BG)
            self.tree.tag_configure("even", background=TREE_EVEN_BG)
            if hasattr(self, "preview_client_combo"):
                client_values = [self.CLIENT_ALL_LABEL]
                client_values.extend(
                    self.clients_db.display_name(client)
                    for client in self.clients_db.clients_sorted()
                )
                current = self.preview_client_var.get().strip()
                self.preview_client_combo.configure(values=client_values)
                if current not in client_values:
                    self.preview_client_var.set(self.CLIENT_ALL_LABEL)
            self._update_detail_panel()

        def _selected_rule_name(self):
            selection = self.tree.selection()
            if not selection:
                return None
            return self.tree.item(selection[0], "values")[0]

        def _selected_rule(self):
            name = self._selected_rule_name()
            if not name:
                return None
            return self.db.get(name)

        def _on_tree_selection_changed(self) -> None:
            self._seed_preview_from_selection()
            self._update_detail_panel()

        def _update_detail_panel(self) -> None:
            if not hasattr(self, "detail_var"):
                return
            rule = self._selected_rule()
            if rule is None:
                self.detail_var.set("Selecteer een presetregel voor details.")
                return

            client = rule.client or self.CLIENT_ALL_LABEL
            kind = self.KIND_LABELS.get(rule.selection_kind, rule.selection_kind)
            selection = rule.selection_summary()
            actions = rule.action_summary() or "(geen invulactie)"
            status = self._rule_status_detail(rule).replace("\n", "; ")
            details = [
                f"Wanneer: opdrachtgever={client}, type={kind}, selectie={selection}",
                f"Invullen: {actions}",
                f"Status: {status}",
            ]
            if bool(self.advanced_view_var.get()):
                details.append(
                    "Geavanceerd: "
                    f"prioriteit={rule.priority}, "
                    f"auto={'ja' if rule.auto_apply else 'nee'}, "
                    f"actief={'ja' if rule.enabled else 'nee'}"
                )
            self.detail_var.set("\n".join(details))

        def _save_and_refresh(self) -> None:
            self.db.save(ORDER_PRESETS_DB_FILE)
            self.refresh()
            if self.on_change:
                self.on_change()

        def _rule_action_fields(self, rule: OrderPresetRule) -> set[str]:
            fields: set[str] = set()
            if _to_str(rule.supplier).strip():
                fields.add("leverancier")
            if _to_str(rule.doc_type).strip():
                fields.add("documenttype")
            if _to_str(rule.delivery).strip():
                fields.add("leveradres")
            if _to_str(rule.remark).strip():
                fields.add("opmerking")
            if rule.en1090 is not None:
                fields.add("EN1090")
            return fields

        def _rules_overlap(self, left: OrderPresetRule, right: OrderPresetRule) -> bool:
            if left.selection_kind != right.selection_kind:
                return False
            if left.client and right.client and self._key(left.client) != self._key(right.client):
                return False
            if not left.identifiers or not right.identifiers:
                return True
            left_ids = {self._key(identifier) for identifier in left.identifiers}
            right_ids = {self._key(identifier) for identifier in right.identifiers}
            return bool(left_ids & right_ids)

        def _conflict_messages_for_rule(
            self,
            rule: OrderPresetRule,
            *,
            old_name: Optional[str] = None,
        ) -> List[str]:
            if not rule.enabled:
                return []
            fields = self._rule_action_fields(rule)
            if not fields:
                return []
            messages: List[str] = []
            old_key = self._key(old_name)
            for other in self.db.rules:
                if old_key and self._key(other.name) == old_key:
                    continue
                if not other.enabled:
                    continue
                overlap_fields = fields & self._rule_action_fields(other)
                if not overlap_fields:
                    continue
                if not self._rules_overlap(rule, other):
                    continue
                field_text = ", ".join(sorted(overlap_fields))
                messages.append(f"{other.name} ({field_text})")
            return messages

        def _confirm_rule_conflicts(
            self,
            rule: OrderPresetRule,
            *,
            old_name: Optional[str] = None,
            parent=None,
        ) -> bool:
            conflicts = self._conflict_messages_for_rule(rule, old_name=old_name)
            if not conflicts:
                return True
            shown = conflicts[:6]
            extra = len(conflicts) - len(shown)
            lines = "\n".join(f"- {message}" for message in shown)
            if extra > 0:
                lines = f"{lines}\n- ... en {extra} andere"
            return messagebox.askyesno(
                "Overlappende presetregels",
                (
                    "Deze regel overlapt met bestaande actieve regels die dezelfde velden invullen.\n\n"
                    f"{lines}\n\n"
                    "Toch opslaan?"
                ),
                parent=parent or self,
            )

        def add_rule(self):
            initial_rule = OrderPresetRule(
                name="",
                client=self._default_client_for_new_rule(),
            )
            dlg = self._EditDialog(
                self,
                initial_rule,
                self.clients_db,
                self.suppliers_db,
                self.delivery_db,
                title="Nieuwe presetregel",
                selection_suggestions=self._selection_suggestions_by_kind(),
                conflict_checker=lambda rule, _old_name=None, parent=None: self._confirm_rule_conflicts(
                    rule,
                    old_name=None,
                    parent=parent,
                ),
            )
            self.wait_window(dlg)
            if dlg.result is not None:
                self.db.upsert(dlg.result)
                self._save_and_refresh()

        def edit_sel(self):
            rule = self._selected_rule()
            if rule is None:
                return
            dlg = self._EditDialog(
                self,
                rule,
                self.clients_db,
                self.suppliers_db,
                self.delivery_db,
                title="Presetregel bewerken",
                selection_suggestions=self._selection_suggestions_by_kind(),
                conflict_checker=lambda candidate, _old_name=rule.name, parent=None: self._confirm_rule_conflicts(
                    candidate,
                    old_name=_old_name,
                    parent=parent,
                ),
            )
            self.wait_window(dlg)
            if dlg.result is not None:
                self.db.upsert(dlg.result, old_name=rule.name)
                self._save_and_refresh()

        def _duplicate_name(self, name: str) -> str:
            base = f"{name} kopie"
            candidate = base
            index = 2
            while self.db.get(candidate) is not None:
                candidate = f"{base} {index}"
                index += 1
            return candidate

        def duplicate_sel(self):
            rule = self._selected_rule()
            if rule is None:
                return
            duplicate = OrderPresetRule.from_any(rule)
            duplicate.name = self._duplicate_name(rule.name)
            dlg = self._EditDialog(
                self,
                duplicate,
                self.clients_db,
                self.suppliers_db,
                self.delivery_db,
                title="Presetregel dupliceren",
                selection_suggestions=self._selection_suggestions_by_kind(),
                conflict_checker=lambda candidate, _old_name=None, parent=None: self._confirm_rule_conflicts(
                    candidate,
                    old_name=None,
                    parent=parent,
                ),
            )
            self.wait_window(dlg)
            if dlg.result is not None:
                self.db.upsert(dlg.result)
                self._save_and_refresh()
                try:
                    self.tree.selection_set(dlg.result.name)
                    self.tree.see(dlg.result.name)
                except tk.TclError:
                    pass

        def remove_sel(self):
            name = self._selected_rule_name()
            if not name:
                return
            if not messagebox.askyesno(
                "Bevestigen",
                f"Verwijder presetregel '{name}'?",
                parent=self,
            ):
                return
            if self.db.remove(name):
                self._save_and_refresh()

        def toggle_sel(self):
            name = self._selected_rule_name()
            if not name:
                return
            if self.db.toggle_enabled(name):
                self._save_and_refresh()

        def check_rules(self) -> None:
            issue_lines: List[str] = []
            first_problem_name: Optional[str] = None
            seen_names: set[str] = set()

            for rule in self.db.rules_sorted():
                name_key = self._key(rule.name)
                issues = [
                    issue
                    for issue in self._rule_status_issues(rule, old_name=rule.name)
                    if issue != "Uitgeschakeld"
                ]
                if name_key in seen_names:
                    issues.append("Dubbele naam")
                seen_names.add(name_key)

                if not issues:
                    continue
                if first_problem_name is None:
                    first_problem_name = rule.name
                issue_lines.append(f"{rule.name}: {', '.join(issues)}")

            if not issue_lines:
                messagebox.showinfo(
                    "Presetregels controleren",
                    "Geen problemen gevonden in de presetregels.",
                    parent=self,
                )
                return

            if first_problem_name:
                try:
                    self.tree.selection_set(first_problem_name)
                    self.tree.see(first_problem_name)
                except tk.TclError:
                    pass

            shown = issue_lines[:14]
            extra = len(issue_lines) - len(shown)
            message = "\n".join(f"- {line}" for line in shown)
            if extra > 0:
                message = f"{message}\n- ... en {extra} andere"
            messagebox.showwarning(
                "Presetregels controleren",
                f"Controleer deze presetregels:\n\n{message}",
                parent=self,
            )

        def _client_choice_values(
            self,
            *,
            include_global: bool = False,
        ) -> Tuple[List[str], Dict[str, str]]:
            values: List[str] = []
            mapping: Dict[str, str] = {}
            if include_global:
                values.append(self.FILTER_GLOBAL_LABEL)
                mapping[self.FILTER_GLOBAL_LABEL] = ""
            for client in self.clients_db.clients_sorted():
                display = self.clients_db.display_name(client)
                clean = strip_favorite_marker(display).strip()
                values.append(display)
                mapping[display] = clean
            return values, mapping

        def _client_display_for_value(self, value: str, *, include_global: bool = False) -> str:
            clean_value = _to_str(value).strip()
            if not clean_value:
                return self.FILTER_GLOBAL_LABEL if include_global else ""
            clean_key = self._key(clean_value)
            for client in self.clients_db.clients_sorted():
                display = self.clients_db.display_name(client)
                display_value = strip_favorite_marker(display).strip()
                if self._key(display_value) == clean_key:
                    return display
            return clean_value

        def _target_client_default(self, source: str, values: List[str], mapping: Dict[str, str]) -> str:
            source_key = self._key(source)
            for display in values:
                if self._key(mapping.get(display, "")) != source_key:
                    return display
            return values[0] if values else ""

        def _rule_copy_signature(self, rule: OrderPresetRule) -> Tuple[str, Tuple[str, ...]]:
            return (
                rule.selection_kind,
                tuple(self._key(identifier) for identifier in rule.identifiers),
            )

        def _copy_rule_name(self, rule: OrderPresetRule, target: str) -> str:
            target_label = _to_str(target).strip() or "Algemeen"
            base = _to_str(rule.name).strip() or "Presetregel"
            if self._key(base).startswith(self._key(target_label)):
                return base
            return f"{target_label} - {base}"

        def copy_client_rules(self) -> None:
            source_values, source_mapping = self._client_choice_values(include_global=True)
            target_values, target_mapping = self._client_choice_values(include_global=False)
            if not target_values:
                messagebox.showwarning(
                    "Kopieer naar klant",
                    "Er zijn nog geen opdrachtgevers om regels naartoe te kopieren.",
                    parent=self,
                )
                return

            selected_rule = self._selected_rule()
            source_client = self._current_filter_client()
            if source_client is None and selected_rule is not None:
                source_client = selected_rule.client
            source_default = self._client_display_for_value(
                source_client or "",
                include_global=True,
            )
            if source_default not in source_mapping:
                source_default = source_values[0] if source_values else ""

            result: Dict[str, object] = {}
            dialog = tk.Toplevel(self)
            dialog.title("Presetregels kopieren")
            dialog.columnconfigure(1, weight=1)

            source_var = tk.StringVar(value=source_default)
            target_default = self._target_client_default(
                source_mapping.get(source_default, ""),
                target_values,
                target_mapping,
            )
            target_var = tk.StringVar(value=target_default)
            skip_existing_var = tk.IntVar(value=1)

            tk.Label(dialog, text="Van klant:").grid(row=0, column=0, sticky="e", padx=8, pady=(10, 4))
            source_combo = ttk.Combobox(
                dialog,
                textvariable=source_var,
                values=source_values,
                state="readonly",
                width=42,
            )
            source_combo.grid(row=0, column=1, sticky="ew", padx=8, pady=(10, 4))

            tk.Label(dialog, text="Naar klant:").grid(row=1, column=0, sticky="e", padx=8, pady=4)
            target_combo = ttk.Combobox(
                dialog,
                textvariable=target_var,
                values=target_values,
                state="readonly",
                width=42,
            )
            target_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=4)

            tk.Checkbutton(
                dialog,
                text="Bestaande doelregels met dezelfde selectie overslaan",
                variable=skip_existing_var,
            ).grid(row=2, column=1, sticky="w", padx=8, pady=(2, 8))

            def _ok() -> None:
                source = source_mapping.get(source_var.get(), "")
                target = target_mapping.get(target_var.get(), "")
                if self._key(source) == self._key(target):
                    messagebox.showwarning(
                        "Kopieer naar klant",
                        "Kies twee verschillende opdrachtgevers.",
                        parent=dialog,
                    )
                    return
                result["source"] = source
                result["target"] = target
                result["skip_existing"] = bool(skip_existing_var.get())
                dialog.destroy()

            buttons = tk.Frame(dialog)
            buttons.grid(row=3, column=0, columnspan=2, pady=(2, 10))
            tk.Button(buttons, text="Kopieren", command=_ok).pack(side="left", padx=4)
            tk.Button(buttons, text="Annuleer", command=dialog.destroy).pack(side="left", padx=4)

            dialog.transient(self)
            _place_window_near_parent(dialog, self)
            dialog.grab_set()
            self.wait_window(dialog)

            if not result:
                return

            source = _to_str(result.get("source")).strip()
            target = _to_str(result.get("target")).strip()
            skip_existing = bool(result.get("skip_existing"))
            source_rules = [
                rule
                for rule in self.db.rules_sorted()
                if self._key(rule.client) == self._key(source)
            ]
            if not source_rules:
                messagebox.showinfo(
                    "Kopieer naar klant",
                    "Geen presetregels gevonden voor de gekozen bronklant.",
                    parent=self,
                )
                return

            existing_signatures = {
                self._rule_copy_signature(rule)
                for rule in self.db.rules
                if self._key(rule.client) == self._key(target)
            }
            copied = 0
            skipped = 0
            for rule in source_rules:
                signature = self._rule_copy_signature(rule)
                if skip_existing and signature in existing_signatures:
                    skipped += 1
                    continue
                clone = OrderPresetRule.from_any(rule)
                clone.client = target
                clone.name = self._duplicate_name(self._copy_rule_name(rule, target))
                self.db.upsert(clone)
                existing_signatures.add(signature)
                copied += 1

            if copied:
                self.filter_client_var.set(self._client_display_for_value(target, include_global=True))
                self._save_and_refresh()
            messagebox.showinfo(
                "Kopieer naar klant",
                f"Gekopieerd: {copied}\nOvergeslagen: {skipped}",
                parent=self,
            )

        def _preview_context(self) -> Dict[str, str]:
            client_value = self.preview_client_var.get().strip()
            if client_value == self.CLIENT_ALL_LABEL:
                client_name = ""
            else:
                client_name = strip_favorite_marker(client_value).strip()
            return {
                "client": client_name,
                "selection_kind": self.KIND_VALUES.get(
                    self.preview_kind_var.get().strip(),
                    "production",
                ),
                "identifier": self.preview_identifier_var.get().strip(),
            }

        def _run_preview(self) -> None:
            context = self._preview_context()
            if not context["identifier"]:
                self.preview_result_var.set(
                    "Geef een productie, afwerking of brutemateriaal op om de presetregels te testen."
                )
                return
            try:
                evaluation = self.db.evaluate(
                    context,
                    auto_apply_only=bool(self.preview_auto_only_var.get()),
                )
            except Exception as exc:
                self.preview_result_var.set(f"Kon preset preview niet berekenen: {exc}")
                return

            header = (
                f"Context: opdrachtgever={context['client'] or '(alle)'} | "
                f"type={self.KIND_LABELS.get(context['selection_kind'], context['selection_kind'])} | "
                f"selectie={context['identifier']}"
            )
            matched = ", ".join(evaluation.matched_rule_names) or "(geen)"
            applied = ", ".join(evaluation.applied_rule_names) or "(geen)"
            actions: List[str] = []
            if evaluation.supplier:
                actions.append(f"leverancier={evaluation.supplier}")
            if evaluation.doc_type:
                actions.append(f"documenttype={evaluation.doc_type}")
            if evaluation.delivery:
                actions.append(f"leveradres={evaluation.delivery}")
            if evaluation.remark:
                actions.append(f"opmerking={evaluation.remark}")
            if evaluation.en1090 is not None:
                actions.append(f"EN1090={'aan' if evaluation.en1090 else 'uit'}")
            action_text = ", ".join(actions) or "(geen)"
            self.preview_result_var.set(
                f"{header}\nGematchte regels: {matched}\nToegepaste regels: {applied}\nInvulling: {action_text}"
            )

        def _seed_preview_from_selection(self) -> None:
            rule = self._selected_rule()
            if rule is None:
                return
            if rule.client:
                for client in self.clients_db.clients_sorted():
                    display = self.clients_db.display_name(client)
                    if strip_favorite_marker(display).strip().lower() == rule.client.strip().lower():
                        self.preview_client_var.set(display)
                        break
                else:
                    self.preview_client_var.set(rule.client)
            else:
                self.preview_client_var.set(self.CLIENT_ALL_LABEL)
            self.preview_kind_var.set(
                self.KIND_LABELS.get(rule.selection_kind, "Productie")
            )
            self.preview_identifier_var.set(rule.identifiers[0] if rule.identifiers else "")
            self._run_preview()

        class _EditDialog(tk.Toplevel):
            DOC_TYPE_OPTIONS = (
                "(ongewijzigd)",
                "Geen",
                "Bestelbon",
                "Standaard bon",
                "Offerteaanvraag",
            )
            EN1090_OPTIONS = (
                "(ongewijzigd)",
                "Aan",
                "Uit",
            )

            def __init__(
                self,
                master,
                rule,
                clients_db,
                suppliers_db,
                delivery_db,
                *,
                title="Presetregel",
                selection_suggestions=None,
                conflict_checker=None,
            ):
                super().__init__(master)
                self.title(title)
                self.result = None
                self.clients_db = clients_db
                self.suppliers_db = suppliers_db
                self.delivery_db = delivery_db
                self.conflict_checker = conflict_checker
                self.selection_suggestions = {
                    kind: list(values or [])
                    for kind, values in (selection_suggestions or {}).items()
                }

                rule = OrderPresetRule.from_any(rule) if rule is not None else OrderPresetRule(name="")
                self.columnconfigure(0, weight=1)

                self._client_display_to_value = {
                    PresetRulesManagerFrame.CLIENT_ALL_LABEL: ""
                }
                self._supplier_display_to_value = {
                    PresetRulesManagerFrame.SUPPLIER_UNCHANGED_LABEL: ""
                }
                self._delivery_display_to_value = {
                    PresetRulesManagerFrame.DELIVERY_UNCHANGED_LABEL: ""
                }

                client_values = [PresetRulesManagerFrame.CLIENT_ALL_LABEL]
                for client in self.clients_db.clients_sorted():
                    display = self.clients_db.display_name(client)
                    client_values.append(display)
                    self._client_display_to_value[display] = strip_favorite_marker(display).strip()

                supplier_values = [PresetRulesManagerFrame.SUPPLIER_UNCHANGED_LABEL]
                for supplier in self.suppliers_db.suppliers_sorted():
                    display = self.suppliers_db.display_name(supplier)
                    supplier_values.append(display)
                    self._supplier_display_to_value[display] = supplier.supplier

                delivery_values = [PresetRulesManagerFrame.DELIVERY_UNCHANGED_LABEL]
                delivery_values.extend(
                    [
                        "Geen",
                        "Opdrachtgeveradres",
                        "Bestelling wordt opgehaald",
                        "Leveradres wordt nog meegedeeld",
                    ]
                )
                for delivery in self.delivery_db.addresses_sorted():
                    display = self.delivery_db.display_name(delivery)
                    if display not in self._delivery_display_to_value:
                        delivery_values.append(display)
                        self._delivery_display_to_value[display] = strip_favorite_marker(display).strip()
                self._delivery_display_to_value["Geen"] = "Geen"
                self._delivery_display_to_value["Opdrachtgeveradres"] = "Klantadres"
                self._delivery_display_to_value["Klantadres"] = "Klantadres"
                self._delivery_display_to_value["Bestelling wordt opgehaald"] = "Bestelling wordt opgehaald"
                self._delivery_display_to_value["Leveradres wordt nog meegedeeld"] = "Leveradres wordt nog meegedeeld"

                def _display_for_value(mapping, value, fallback):
                    clean = _to_str(value).strip()
                    if not clean:
                        return fallback
                    for display, raw_value in mapping.items():
                        if _to_str(raw_value).strip().lower() == clean.lower():
                            return display
                    return clean

                self.name_var = tk.StringVar(value=rule.name)
                self.enabled_var = tk.IntVar(value=1 if rule.enabled else 0)
                self.auto_apply_var = tk.IntVar(value=1 if rule.auto_apply else 0)
                self.priority_var = tk.StringVar(value=str(rule.priority))
                self.client_var = tk.StringVar(
                    value=_display_for_value(
                        self._client_display_to_value,
                        rule.client,
                        PresetRulesManagerFrame.CLIENT_ALL_LABEL,
                    )
                )
                self.kind_var = tk.StringVar(
                    value=PresetRulesManagerFrame.KIND_LABELS.get(
                        rule.selection_kind, "Productie"
                    )
                )
                self.supplier_var = tk.StringVar(
                    value=_display_for_value(
                        self._supplier_display_to_value,
                        rule.supplier,
                        PresetRulesManagerFrame.SUPPLIER_UNCHANGED_LABEL,
                    )
                )
                self.doc_type_var = tk.StringVar(
                    value=rule.doc_type or PresetRulesManagerFrame.DOC_TYPE_UNCHANGED_LABEL
                )
                self.delivery_var = tk.StringVar(
                    value=_display_for_value(
                        self._delivery_display_to_value,
                        rule.delivery,
                        PresetRulesManagerFrame.DELIVERY_UNCHANGED_LABEL,
                    )
                )
                self.remark_var = tk.StringVar(value=rule.remark or "")
                if rule.en1090 is None:
                    en1090_text = PresetRulesManagerFrame.EN1090_UNCHANGED_LABEL
                else:
                    en1090_text = "Aan" if rule.en1090 else "Uit"
                self.en1090_var = tk.StringVar(value=en1090_text)

                form = tk.Frame(self)
                form.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
                form.columnconfigure(1, weight=1)

                tk.Label(form, text="Naam:").grid(row=0, column=0, sticky="e", padx=4, pady=2)
                name_entry = tk.Entry(form, textvariable=self.name_var, width=42)
                name_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
                _HelpTooltip(
                    name_entry,
                    "Herkenbare naam voor deze regel. Bijvoorbeeld: Tecno Art - Tube laser.",
                )

                enabled_check = tk.Checkbutton(form, text="Actief", variable=self.enabled_var)
                enabled_check.grid(row=1, column=1, sticky="w", padx=4, pady=(2, 6))
                _HelpTooltip(enabled_check, "Uitgeschakelde regels blijven bewaard, maar worden niet toegepast.")

                when_frame = tk.LabelFrame(form, text="Wanneer toepassen?")
                when_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=0, pady=(2, 8))
                when_frame.columnconfigure(1, weight=1)

                tk.Label(when_frame, text="Opdrachtgever:").grid(row=0, column=0, sticky="e", padx=4, pady=2)
                client_combo = ttk.Combobox(
                    when_frame,
                    textvariable=self.client_var,
                    values=client_values,
                    state="readonly",
                    width=40,
                )
                client_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
                _HelpTooltip(
                    client_combo,
                    "Beperk de regel tot een opdrachtgever. Kies '(alle opdrachtgevers)' voor een algemene regel.",
                )

                tk.Label(when_frame, text="Type selectie:").grid(row=1, column=0, sticky="e", padx=4, pady=2)
                kind_combo = ttk.Combobox(
                    when_frame,
                    textvariable=self.kind_var,
                    values=list(PresetRulesManagerFrame.KIND_VALUES.keys()),
                    state="readonly",
                    width=18,
                )
                kind_combo.grid(row=1, column=1, sticky="w", padx=4, pady=2)
                _HelpTooltip(
                    kind_combo,
                    "Productie matcht de BOM-kolom Production. Afwerking matcht Finish-sleutels. Brutemateriaal matcht de Opticutter-brutemateriaalrij.",
                )
                kind_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_selection_suggestions())

                tk.Label(when_frame, text="Bestaande selectie:").grid(row=2, column=0, sticky="e", padx=4, pady=2)
                suggestion_row = tk.Frame(when_frame)
                suggestion_row.grid(row=2, column=1, sticky="ew", padx=4, pady=2)
                suggestion_row.columnconfigure(0, weight=1)
                self.selection_suggestion_var = tk.StringVar()
                self.selection_suggestion_combo = ttk.Combobox(
                    suggestion_row,
                    textvariable=self.selection_suggestion_var,
                    state="readonly",
                    width=34,
                )
                self.selection_suggestion_combo.grid(row=0, column=0, sticky="ew")
                suggestion_add_btn = tk.Button(
                    suggestion_row,
                    text="Toevoegen",
                    command=self._append_selection_suggestion,
                )
                suggestion_add_btn.grid(row=0, column=1, padx=(6, 0))
                _HelpTooltip(
                    self.selection_suggestion_combo,
                    "Kies een selectie die al in andere presetregels gebruikt wordt.",
                )
                _HelpTooltip(
                    suggestion_add_btn,
                    "Voeg de gekozen waarde toe aan de selecties van deze regel.",
                )

                tk.Label(when_frame, text="Selecties:").grid(row=3, column=0, sticky="ne", padx=4, pady=2)
                self.identifiers_text = tk.Text(
                    when_frame,
                    width=42,
                    height=4,
                    wrap="word",
                    font=tkfont.nametofont("TkDefaultFont"),
                )
                self.identifiers_text.grid(row=3, column=1, sticky="ew", padx=4, pady=2)
                self.identifiers_text.insert("1.0", "\n".join(rule.identifiers))
                _HelpTooltip(
                    self.identifiers_text,
                    "Exacte selectie(s) waarop de regel moet matchen. Gebruik dezelfde tekst als in de bestelbonrij, bijvoorbeeld Cutting of Finish-Galvanised.",
                )
                selection_help = tk.Label(
                    when_frame,
                    text="Meerdere waarden gescheiden door komma of nieuwe regel. Beweeg over velden voor uitleg.",
                    anchor="w",
                    justify="left",
                )
                selection_help.grid(row=4, column=1, sticky="w", padx=4, pady=(0, 6))
                _HelpTooltip(
                    selection_help,
                    "Voor Productie en Brutemateriaal gebruik je bv. Cutting, Tube laser of Roof. Voor Afwerking gebruik je de Finish-... sleutel uit de afwerkingsrij.",
                )
                self._refresh_selection_suggestions()

                what_frame = tk.LabelFrame(form, text="Wat invullen?")
                what_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=0, pady=(0, 8))
                what_frame.columnconfigure(1, weight=1)

                tk.Label(what_frame, text="Leverancier:").grid(row=0, column=0, sticky="e", padx=4, pady=2)
                supplier_combo = ttk.Combobox(
                    what_frame,
                    textvariable=self.supplier_var,
                    values=supplier_values,
                    state="readonly",
                    width=40,
                )
                supplier_combo.grid(row=0, column=1, sticky="ew", padx=4, pady=2)
                _HelpTooltip(
                    supplier_combo,
                    "Laat op '(ongewijzigd)' als deze regel geen leverancier moet invullen.",
                )

                tk.Label(what_frame, text="Documenttype:").grid(row=1, column=0, sticky="e", padx=4, pady=2)
                doc_combo = ttk.Combobox(
                    what_frame,
                    textvariable=self.doc_type_var,
                    values=self.DOC_TYPE_OPTIONS,
                    state="readonly",
                    width=24,
                )
                doc_combo.grid(row=1, column=1, sticky="w", padx=4, pady=2)
                _HelpTooltip(
                    doc_combo,
                    "Kies welk documenttype automatisch wordt gezet. '(ongewijzigd)' laat de bestaande keuze staan.",
                )

                tk.Label(what_frame, text="Leveradres:").grid(row=2, column=0, sticky="e", padx=4, pady=2)
                delivery_combo = ttk.Combobox(
                    what_frame,
                    textvariable=self.delivery_var,
                    values=delivery_values,
                    state="readonly",
                    width=40,
                )
                delivery_combo.grid(row=2, column=1, sticky="ew", padx=4, pady=2)
                _HelpTooltip(
                    delivery_combo,
                    "Opdrachtgeveradres gebruikt het adres van de gekozen opdrachtgever. Een specifiek leveradres kan ook gekozen worden.",
                )

                tk.Label(what_frame, text="EN 1090:").grid(row=3, column=0, sticky="e", padx=4, pady=2)
                en1090_combo = ttk.Combobox(
                    what_frame,
                    textvariable=self.en1090_var,
                    values=self.EN1090_OPTIONS,
                    state="readonly",
                    width=16,
                )
                en1090_combo.grid(row=3, column=1, sticky="w", padx=4, pady=2)
                _HelpTooltip(
                    en1090_combo,
                    "Aan of Uit vult de EN1090-keuze op de bestelbonrij in. '(ongewijzigd)' past dit veld niet aan.",
                )

                tk.Label(what_frame, text="Opmerking:").grid(row=4, column=0, sticky="e", padx=4, pady=2)
                remark_entry = tk.Entry(what_frame, textvariable=self.remark_var, width=42)
                remark_entry.grid(row=4, column=1, sticky="ew", padx=4, pady=2)
                _HelpTooltip(
                    remark_entry,
                    "Optionele opmerking die op de bestelbonrij wordt ingevuld.",
                )

                advanced_toggle = tk.Button(form)
                advanced_toggle.grid(row=4, column=1, sticky="w", padx=4, pady=(0, 4))
                advanced_frame = tk.LabelFrame(form, text="Geavanceerd")
                advanced_frame.columnconfigure(1, weight=1)
                show_advanced = tk.IntVar(
                    value=1 if rule.priority != 100 or not rule.auto_apply else 0
                )

                tk.Label(advanced_frame, text="Prioriteit:").grid(row=0, column=0, sticky="e", padx=4, pady=2)
                priority_entry = tk.Entry(advanced_frame, textvariable=self.priority_var, width=10)
                priority_entry.grid(row=0, column=1, sticky="w", padx=4, pady=2)
                _HelpTooltip(
                    priority_entry,
                    "Meestal 100 laten staan. Een hogere prioriteit wint wanneer meerdere regels dezelfde selectie matchen.",
                )
                auto_check = tk.Checkbutton(
                    advanced_frame,
                    text="Automatisch toepassen bij openen",
                    variable=self.auto_apply_var,
                )
                auto_check.grid(row=1, column=1, sticky="w", padx=4, pady=2)
                _HelpTooltip(
                    auto_check,
                    "Aan: de regel vult de bestelbonvelden automatisch in zodra de keuzelijst opent.",
                )

                def _sync_advanced() -> None:
                    if show_advanced.get():
                        advanced_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=0, pady=(0, 8))
                        advanced_toggle.configure(text="Geavanceerd verbergen")
                    else:
                        advanced_frame.grid_remove()
                        advanced_toggle.configure(text="Geavanceerd tonen")

                def _toggle_advanced() -> None:
                    show_advanced.set(0 if show_advanced.get() else 1)
                    _sync_advanced()

                advanced_toggle.configure(command=_toggle_advanced)
                _HelpTooltip(
                    advanced_toggle,
                    "Toon prioriteit en automatisch toepassen. Deze velden hoef je meestal niet te wijzigen.",
                )
                _sync_advanced()

                btn = tk.Frame(form)
                btn.grid(row=6, column=0, columnspan=2, pady=8)
                save_btn = tk.Button(btn, text="Opslaan", command=self._ok)
                save_btn.pack(side="left", padx=4)
                cancel_btn = tk.Button(btn, text="Annuleer", command=self.destroy)
                cancel_btn.pack(side="left", padx=4)

                self.transient(master)
                _place_window_near_parent(self, master)
                self.grab_set()
                self.after_idle(name_entry.focus_set)

            def _current_selection_kind(self) -> str:
                return PresetRulesManagerFrame.KIND_VALUES.get(
                    self.kind_var.get(),
                    "production",
                )

            def _refresh_selection_suggestions(self) -> None:
                combo = getattr(self, "selection_suggestion_combo", None)
                if combo is None:
                    return
                values = self.selection_suggestions.get(self._current_selection_kind(), [])
                combo.configure(values=values)
                current = self.selection_suggestion_var.get().strip()
                if current not in values:
                    self.selection_suggestion_var.set("")

            def _append_selection_suggestion(self) -> None:
                value = self.selection_suggestion_var.get().strip()
                if not value:
                    return
                raw = self.identifiers_text.get("1.0", "end").strip()
                existing = {
                    _to_str(part).strip().casefold()
                    for part in re.split(r"[\r\n,;]+", raw)
                    if _to_str(part).strip()
                }
                if value.casefold() in existing:
                    return
                if raw:
                    self.identifiers_text.insert("end", "\n")
                self.identifiers_text.insert("end", value)

            def _ok(self):
                name = self.name_var.get().strip()
                if not name:
                    messagebox.showwarning(
                        "Let op",
                        "Geef de presetregel een naam.",
                        parent=self,
                    )
                    return

                try:
                    priority = int(self.priority_var.get().strip())
                except Exception:
                    messagebox.showwarning(
                        "Let op",
                        "Prioriteit moet een getal zijn.",
                        parent=self,
                    )
                    return

                client = self._client_display_to_value.get(self.client_var.get(), "")
                selection_kind = PresetRulesManagerFrame.KIND_VALUES.get(
                    self.kind_var.get(),
                    "production",
                )
                identifiers_raw = self.identifiers_text.get("1.0", "end").strip()
                supplier = self._supplier_display_to_value.get(self.supplier_var.get(), "")
                doc_type = self.doc_type_var.get().strip()
                if doc_type == PresetRulesManagerFrame.DOC_TYPE_UNCHANGED_LABEL:
                    doc_type = ""
                delivery = self._delivery_display_to_value.get(self.delivery_var.get(), "")
                if delivery == PresetRulesManagerFrame.DELIVERY_UNCHANGED_LABEL:
                    delivery = ""
                remark = self.remark_var.get().strip()

                en1090_text = self.en1090_var.get().strip()
                if en1090_text == "Aan":
                    en1090_value = True
                elif en1090_text == "Uit":
                    en1090_value = False
                else:
                    en1090_value = None

                if not any((supplier, doc_type, delivery, remark, en1090_value is not None)):
                    messagebox.showwarning(
                        "Let op",
                        "Geef minstens één invulactie op.",
                        parent=self,
                    )
                    return

                try:
                    self.result = OrderPresetRule.from_any(
                        {
                            "name": name,
                            "enabled": bool(self.enabled_var.get()),
                            "priority": priority,
                            "auto_apply": bool(self.auto_apply_var.get()),
                            "client": client,
                            "selection_kind": selection_kind,
                            "identifiers": identifiers_raw,
                            "supplier": supplier,
                            "doc_type": doc_type,
                            "delivery": delivery,
                            "remark": remark,
                            "en1090": en1090_value,
                        }
                    )
                except Exception as exc:
                    messagebox.showerror("Fout", str(exc), parent=self)
                    return

                if callable(self.conflict_checker):
                    try:
                        if not self.conflict_checker(self.result, parent=self):
                            return
                    except Exception as exc:
                        messagebox.showerror("Fout", str(exc), parent=self)
                        return
                self.destroy()

    def _clean_supplier_pricing_value(value: object) -> Dict[str, object]:
        if not isinstance(value, Mapping):
            return {}
        cleaned: Dict[str, object] = {
            "unit_price": _to_str(value.get("unit_price")).strip(),
            "total_price": _to_str(value.get("total_price")).strip(),
            "quote_ref": _to_str(value.get("quote_ref")).strip(),
            "note": _to_str(value.get("note")).strip(),
        }
        line_items: Dict[str, Dict[str, str]] = {}
        raw_items = value.get("items")
        if isinstance(raw_items, Mapping):
            for raw_key, raw_item in raw_items.items():
                item_key = _to_str(raw_key).strip()
                if not item_key or not isinstance(raw_item, Mapping):
                    continue
                item_entry = {
                    "unit_price": _to_str(raw_item.get("unit_price")).strip(),
                    "total_price": _to_str(raw_item.get("total_price")).strip(),
                    "quote_ref": _to_str(raw_item.get("quote_ref")).strip(),
                    "note": _to_str(raw_item.get("note")).strip(),
                }
                if any(item_entry.values()):
                    line_items[item_key] = item_entry
        if line_items:
            cleaned["items"] = line_items
        if any(
            _to_str(cleaned.get(name)).strip()
            for name in ("unit_price", "total_price", "quote_ref", "note")
        ) or line_items:
            return cleaned
        return {}

    def _parse_supplier_decimal(value: object) -> Optional[Decimal]:
        text = _to_str(value).strip()
        if not text:
            return None
        text = (
            text.replace("\u00a0", "")
            .replace(" ", "")
            .replace("€", "")
            .replace("%", "")
        )
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
        try:
            return Decimal(text)
        except (InvalidOperation, ValueError):
            return None

    def _format_supplier_decimal(value: Decimal) -> str:
        rounded = value.quantize(Decimal("0.01"))
        return f"{rounded:.2f}"

    def _format_supplier_quantity(value: Decimal) -> str:
        rounded = value.quantize(Decimal("0.001"))
        if rounded == rounded.to_integral_value():
            return str(int(rounded))
        return f"{rounded:.3f}".rstrip("0").rstrip(".")

    def _clean_supplier_vat_rate(value: object, default: str = "21") -> str:
        rate = _parse_supplier_decimal(value)
        if rate is None or rate < 0 or rate > 100:
            return default
        rounded = rate.quantize(Decimal("0.01"))
        if rounded == rounded.to_integral_value():
            return str(int(rounded))
        return f"{rounded:.2f}".rstrip("0").rstrip(".")

    @dataclass
    class SupplierSelectionState:
        selections: Dict[str, str]
        groups: Dict[str, str]
        doc_types: Dict[str, str]
        doc_numbers: Dict[str, str]
        remarks: Dict[str, str]
        deliveries: Dict[str, str]
        exports: Dict[str, bool] = field(default_factory=dict)
        en1090: Dict[str, bool] = field(default_factory=dict)
        vat_rates: Dict[str, str] = field(default_factory=dict)
        pricing: Dict[str, Dict[str, object]] = field(default_factory=dict)
        remember: bool = True

        @classmethod
        def from_mapping(cls, data: Mapping[str, object]) -> "SupplierSelectionState":
            def _dict(name: str) -> Dict[str, object]:
                value = data.get(name, {}) if isinstance(data, Mapping) else {}
                return dict(value) if isinstance(value, Mapping) else {}

            pricing_raw = _dict("pricing")
            pricing: Dict[str, Dict[str, object]] = {}
            for key, value in pricing_raw.items():
                clean_value = _clean_supplier_pricing_value(value)
                if clean_value:
                    pricing[_to_str(key)] = clean_value

            return cls(
                selections={_to_str(k): _to_str(v) for k, v in _dict("selections").items()},
                groups={_to_str(k): _to_str(v) for k, v in _dict("groups").items()},
                doc_types={_to_str(k): _to_str(v) for k, v in _dict("doc_types").items()},
                doc_numbers={_to_str(k): _to_str(v) for k, v in _dict("doc_numbers").items()},
                remarks={_to_str(k): _to_str(v) for k, v in _dict("remarks").items()},
                deliveries={_to_str(k): _to_str(v) for k, v in _dict("deliveries").items()},
                exports={_to_str(k): bool(v) for k, v in _dict("exports").items()},
                en1090={_to_str(k): bool(v) for k, v in _dict("en1090").items()},
                vat_rates={
                    _to_str(k): _clean_supplier_vat_rate(v)
                    for k, v in _dict("vat_rates").items()
                },
                pricing=pricing,
                remember=bool(data.get("remember", True)) if isinstance(data, Mapping) else True,
            )

    class SupplierSelectionFrame(tk.Frame):
        """Per productie: type-to-filter of dropdown; rechts detailkaart (klik = selecteer).
           Knoppen altijd zichtbaar onderaan.
        """

        LABEL_COLUMN_WIDTH = 30
        EN1090_COLUMN_WIDTH = 64
        EN1090_MIN_COLUMN_WIDTH = 32
        EN1090_COLUMN_PADDING = 12
        EN1090_HEADER_TEXT = "EN 1090"
        GROUP_APART_LABEL = "Apart"
        GROUP_INDICATOR_WIDTH = 10
        GROUP_INDICATOR_GAP = 6
        GROUP_ACCENT_COLORS = (
            "#D55E00",
            "#009E73",
            "#0072B2",
            "#CC79A7",
            "#E69F00",
            "#56B4E9",
            "#6C5CE7",
            "#7A5230",
        )
        PRICE_COLUMN_KEYS = (
            "unit_price_entry",
            "total_price_entry",
            "vat_combo",
            "line_price_button",
        )
        ROW_ALT_BACKGROUND = "#E8E7E2"

        @staticmethod
        def _state_has_price_values(state: Optional["SupplierSelectionState"]) -> bool:
            if state is None:
                return False
            pricing = getattr(state, "pricing", {}) or {}
            if isinstance(pricing, Mapping):
                for value in pricing.values():
                    if not isinstance(value, Mapping):
                        continue
                    if (
                        _to_str(value.get("unit_price")).strip()
                        or _to_str(value.get("total_price")).strip()
                    ):
                        return True
                    items = value.get("items")
                    if isinstance(items, Mapping) and items:
                        return True
            vat_rates = getattr(state, "vat_rates", {}) or {}
            if isinstance(vat_rates, Mapping):
                for value in vat_rates.values():
                    text = _to_str(value).strip().replace("%", "")
                    if text and text != "21":
                        return True
            return False

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

        CLIENT_DELIVERY_PRESET = "Opdrachtgeveradres"
        LEGACY_CLIENT_DELIVERY_PRESET = "Klantadres"
        DELIVERY_PRESETS = (
            "Geen",
            CLIENT_DELIVERY_PRESET,
            "Bestelling wordt opgehaald",
            "Leveradres wordt nog meegedeeld",
        )

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
            clients_db: Optional["ClientsDB"] = None,
            client_var: Optional[tk.StringVar] = None,
            presets_db=None,
            on_manage_presets=None,
            opticutter_details: Dict[str, "OpticutterOrderComputation"] | None = None,
            selection_items: Optional[Dict[str, List[Dict[str, object]]]] = None,
            initial_state: Optional["SupplierSelectionState"] = None,
            en1090_enabled: bool = True,
            en1090_getter=None,
            en1090_setter=None,
            pdf_dossier_context: bool = False,
        ):
            super().__init__(master)
            self.configure(padx=12, pady=12)
            self.db = db
            self.delivery_db = delivery_db
            self.callback = callback
            self.project_number_var = project_number_var
            self.project_name_var = project_name_var
            self.clients_db = clients_db
            self.client_var = client_var
            self.presets_db = presets_db
            self.on_manage_presets = on_manage_presets
            self.opticutter_details = opticutter_details or {}
            self.selection_items = selection_items or {}
            self._en1090_enabled = bool(en1090_enabled)
            self._preview_supplier: Optional[Supplier] = None
            self._active_key: Optional[str] = None  # laatst gefocuste rij
            self._type_filter_by_key: Dict[str, str] = {}
            self.sel_vars: Dict[str, tk.StringVar] = {}
            self.group_vars: Dict[str, tk.StringVar] = {}
            self.doc_vars: Dict[str, tk.StringVar] = {}
            self.doc_num_vars: Dict[str, tk.StringVar] = {}
            self.remark_vars: Dict[str, tk.StringVar] = {}
            self.price_unit_vars: Dict[str, tk.StringVar] = {}
            self.price_total_vars: Dict[str, tk.StringVar] = {}
            self.vat_vars: Dict[str, tk.StringVar] = {}
            self.line_pricing: Dict[str, Dict[str, Dict[str, str]]] = {}
            self.delivery_vars: Dict[str, tk.StringVar] = {}
            self.group_combos: Dict[str, ttk.Combobox] = {}
            self.group_value_to_display: Dict[str, Dict[str, str]] = {}
            self.group_display_to_value: Dict[str, Dict[str, str]] = {}
            self.delivery_combos: Dict[str, ttk.Combobox] = {}
            self.row_meta: Dict[str, Dict[str, str]] = {}
            self.en1090_vars: Dict[str, tk.IntVar] = {}
            self.export_vars: Dict[str, tk.IntVar] = {}
            self._group_sync_in_progress = False
            self._en1090_getter = en1090_getter
            self._en1090_setter = en1090_setter
            self._last_preset_status = ""
            self._preset_state_by_key: Dict[str, Dict[str, object]] = {}
            self._export_log_state_by_key: Dict[str, Dict[str, object]] = {}
            self._loaded_export_log_path = ""
            self._loaded_export_log_export_info: Dict[str, object] = {}
            self._price_link_in_progress: set[str] = set()
            self._price_auto_fields: Dict[str, str] = {}
            self._price_columns_visible = self._state_has_price_values(initial_state)
            self.pdf_dossier_context = bool(pdf_dossier_context)
            self.finish_entries = finishes
            self._row_widget_maps: List[Dict[str, tk.Misc]] = []
            self._row_widgets_by_key: Dict[str, Dict[str, tk.Misc]] = {}

            # Grid layout: content (row=0, weight=1), buttons (row=1)
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)

            content = tk.Frame(self)
            content.grid(row=0, column=0, sticky="nsew", padx=10, pady=6)
            content.grid_columnconfigure(0, weight=1)
            content.grid_rowconfigure(0, weight=0)
            content.grid_rowconfigure(1, weight=1)
            content.grid_rowconfigure(2, weight=0)
            content.grid_rowconfigure(3, weight=0)

            top_area = tk.Frame(content)
            top_area.grid(row=0, column=0, sticky="ew")
            top_area.grid_columnconfigure(0, weight=1)

            if self.pdf_dossier_context:
                context_label = tk.Label(
                    top_area,
                    text=(
                        "PDF dossier voorbereiden - maak hier alleen de bon-PDF's "
                        "die straks in het werkdossier worden ingevoegd."
                    ),
                    anchor="w",
                    justify="left",
                    foreground="#1F4F82",
                    background="#EAF3FF",
                    padx=10,
                    pady=6,
                )
                context_label.pack(fill="x", pady=(0, 8))

            rows_scroll_container = tk.Frame(content)
            rows_scroll_container.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
            rows_scroll_container.grid_columnconfigure(0, weight=1)
            rows_scroll_container.grid_rowconfigure(0, weight=0)
            rows_scroll_container.grid_rowconfigure(1, weight=1)

            self.selection_header_canvas = tk.Canvas(
                rows_scroll_container,
                height=1,
                highlightthickness=0,
                borderwidth=0,
            )
            self.selection_header_canvas.grid(row=0, column=0, sticky="ew")

            def _sync_selection_xview(*args) -> None:
                self.selection_rows_canvas.xview(*args)
                self.selection_header_canvas.xview(*args)

            self.selection_rows_canvas = tk.Canvas(
                rows_scroll_container,
                highlightthickness=0,
                borderwidth=0,
            )
            self.selection_rows_canvas.grid(row=1, column=0, sticky="nsew")
            self.selection_rows_scrollbar = ttk.Scrollbar(
                rows_scroll_container,
                orient="vertical",
                command=self.selection_rows_canvas.yview,
            )
            self.selection_rows_scrollbar.grid(row=1, column=1, sticky="ns")
            self.selection_rows_x_scrollbar = ttk.Scrollbar(
                rows_scroll_container,
                orient="horizontal",
                command=_sync_selection_xview,
            )
            self.selection_rows_x_scrollbar.grid(row=2, column=0, sticky="ew")

            def _on_selection_rows_xscroll(first: str, last: str) -> None:
                self.selection_rows_x_scrollbar.set(first, last)
                try:
                    self.selection_header_canvas.xview_moveto(float(first))
                except (tk.TclError, ValueError):
                    pass

            self.selection_rows_canvas.configure(
                yscrollcommand=self.selection_rows_scrollbar.set,
                xscrollcommand=_on_selection_rows_xscroll,
            )

            left = tk.Frame(self.selection_rows_canvas)
            self._selection_rows_window = self.selection_rows_canvas.create_window(
                (0, 0),
                window=left,
                anchor="nw",
            )

            def _sync_selection_rows_window_width(viewport_width: Optional[int] = None) -> None:
                try:
                    requested_width = left.winfo_reqwidth()
                    canvas_width = (
                        int(viewport_width)
                        if viewport_width is not None
                        else self.selection_rows_canvas.winfo_width()
                    )
                    width = max(requested_width, canvas_width)
                    self.selection_rows_canvas.itemconfigure(
                        self._selection_rows_window,
                        width=width,
                    )
                except tk.TclError:
                    return

            def _update_selection_scroll_region(_event=None) -> None:
                _sync_selection_rows_window_width()
                try:
                    bbox = self.selection_rows_canvas.bbox("all")
                except tk.TclError:
                    return
                if bbox:
                    self.selection_rows_canvas.configure(scrollregion=bbox)

            left.bind("<Configure>", _update_selection_scroll_region)

            def _resize_selection_rows_content(event) -> None:
                _sync_selection_rows_window_width(event.width)
                _update_selection_scroll_region()

            self.selection_rows_canvas.bind("<Configure>", _resize_selection_rows_content)

            def _on_selection_rows_mousewheel(event):
                delta = getattr(event, "delta", 0)
                if delta:
                    step = -1 if delta > 0 else 1
                elif getattr(event, "num", 0) == 4:
                    step = -1
                elif getattr(event, "num", 0) == 5:
                    step = 1
                else:
                    return None
                try:
                    self.selection_rows_canvas.yview_scroll(step, "units")
                except tk.TclError:
                    return None
                return "break"

            def _bind_selection_rows_mousewheel(widget: tk.Misc) -> None:
                if getattr(widget, "_selection_rows_mousewheel_bound", False):
                    return
                for sequence in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                    widget.bind(sequence, _on_selection_rows_mousewheel, add="+")
                widget._selection_rows_mousewheel_bound = True

            def _bind_selection_rows_mousewheel_tree(widget: tk.Misc) -> None:
                _bind_selection_rows_mousewheel(widget)
                try:
                    children = widget.winfo_children()
                except tk.TclError:
                    return
                for child in children:
                    _bind_selection_rows_mousewheel_tree(child)

            self._bind_selection_rows_mousewheel_tree = _bind_selection_rows_mousewheel_tree

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
            proj_container = tk.Frame(top_area)
            proj_container.pack(fill="x", pady=(0, 6))
            proj_container.grid_columnconfigure(0, weight=0)
            proj_container.grid_columnconfigure(1, weight=0)

            proj_frame = tk.LabelFrame(
                proj_container,
                text="Projectgegevens",
                labelanchor="n",
                padx=12,
                pady=10,
            )
            proj_frame.grid(row=0, column=0, sticky="nw")

            clear_btn_container = tk.Frame(proj_container)
            clear_btn_container.grid(row=0, column=1, sticky="nw", padx=(12, 0))
            clear_btn_container.grid_columnconfigure(0, weight=1)

            secondary_actions = tk.LabelFrame(
                clear_btn_container,
                text="Beheer en log",
                padx=6,
                pady=4,
            )
            secondary_actions.grid(row=0, column=0, sticky="nw", pady=(2, 0))

            top_button_kwargs = dict(padx=10, pady=1)

            def _top_action_button(
                parent,
                *,
                text: str,
                command,
                tooltip: str = "",
                state: Optional[str] = None,
            ) -> tk.Button:
                kwargs = dict(top_button_kwargs)
                if state is not None:
                    kwargs["state"] = state
                button = tk.Button(parent, text=text, command=command, **kwargs)
                if tooltip:
                    _HelpTooltip(button, tooltip)
                return button

            _top_action_button(
                secondary_actions,
                text="Nieuw leveradres",
                command=self._add_delivery_address,
                tooltip="Maak een nieuw leveradres aan en selecteer het voor de actieve rij.",
            ).pack(side="left", padx=(0, 4))
            _top_action_button(
                secondary_actions,
                text="Presetregels",
                command=self._open_preset_manager,
                tooltip=(
                    "Beheer regels die leveranciers, documenttypes en opmerkingen "
                    "automatisch invullen."
                ),
            ).pack(side="left", padx=(0, 4))
            _top_action_button(
                secondary_actions,
                text="Exportlog laden",
                command=self._load_export_log_from_file,
                tooltip="Laad een eerder opgeslagen exportlog voor deze bestelbon.",
            ).pack(side="left", padx=(0, 4))
            _top_action_button(
                secondary_actions,
                text="Laatste log",
                command=self._load_latest_export_log,
                tooltip="Laad automatisch de meest recente exportlog voor dit project.",
            ).pack(side="left", padx=(0, 4))
            self.export_log_documents_button = _top_action_button(
                secondary_actions,
                text="Vorige docs",
                command=self._open_loaded_export_documents_dialog,
                state="disabled",
                tooltip="Open de vorige exportmap of documenten uit de laatst geladen exportlog.",
            )
            self.export_log_documents_button.pack(side="left", padx=(0, 4))
            _top_action_button(
                secondary_actions,
                text="Leegmaken",
                command=self._clear_saved_suppliers,
                tooltip="Wis opgeslagen leverancierkeuzes en reset de velden op deze pagina.",
            ).pack(side="left")

            action_row = tk.Frame(clear_btn_container)
            action_row.grid(row=1, column=0, sticky="nw", pady=(4, 0))
            action_row.grid_rowconfigure(0, weight=1)
            export_group = tk.LabelFrame(
                action_row,
                text="Export selectie",
                padx=6,
                pady=4,
            )
            export_group.grid(row=0, column=0, sticky="s", padx=(0, 8))
            _top_action_button(
                export_group,
                text="Alles aan",
                command=lambda: self._set_all_exports(True),
                tooltip="Selecteer alle bestelbonregels voor export.",
            ).pack(side="left")
            _top_action_button(
                export_group,
                text="Alles uit",
                command=lambda: self._set_all_exports(False),
                tooltip="Deselecteer alle bestelbonregels voor export.",
            ).pack(side="left", padx=(6, 0))
            preset_group = tk.LabelFrame(
                action_row,
                text="Presets",
                padx=6,
                pady=4,
            )
            preset_group.grid(row=0, column=1, sticky="s", padx=(0, 8))
            _top_action_button(
                preset_group,
                text="Invullen",
                command=self._apply_presets_from_button,
                tooltip="Pas de actieve presetregels toe op de bestelbonregels.",
            ).pack(side="left")
            view_group = tk.LabelFrame(
                action_row,
                text="Weergave",
                padx=6,
                pady=4,
            )
            view_group.grid(row=0, column=2, sticky="s")
            self.price_columns_visible_var = tk.BooleanVar(
                value=self._price_columns_visible
            )
            price_columns_toggle = tk.Checkbutton(
                view_group,
                text="Prijsvelden",
                variable=self.price_columns_visible_var,
                padx=4,
                pady=1,
                command=lambda: self.set_price_columns_visible(
                    bool(self.price_columns_visible_var.get())
                ),
            )
            price_columns_toggle.pack(side="left")
            _HelpTooltip(
                price_columns_toggle,
                "Toon of verberg de prijsgerelateerde kolommen in de bestelbonregels.",
            )
            # Option to include brutomateriaal note on generated PDFs
            self.include_bruto_note_var = tk.BooleanVar(value=True)
            bruto_note_toggle = tk.Checkbutton(
                view_group,
                text="Bruto-opm.",
                variable=self.include_bruto_note_var,
                padx=4,
                pady=1,
            )
            bruto_note_toggle.pack(side="left", padx=(10, 0))
            _HelpTooltip(
                bruto_note_toggle,
                "Voeg de bruto materiaal-opmerking toe aan de gegenereerde PDF's.",
            )

            readonly_bg = "#f0f0f0"
            copy_button_text = "⧉"
            copied_button_text = "✓"

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

            copy_reset_job: Optional[str] = None
            name_copy_reset_job: Optional[str] = None

            def copy_project_number() -> None:
                nonlocal copy_reset_job
                project_number = self.project_number_var.get().strip()
                if not project_number:
                    self.bell()
                    return
                try:
                    self.clipboard_clear()
                    self.clipboard_append(project_number)
                    self.update_idletasks()
                except tk.TclError:
                    messagebox.showerror(
                        "Kopieren mislukt",
                        "Het projectnummer kon niet naar het klembord worden gekopieerd.",
                        parent=self,
                    )
                    return
                if copy_reset_job is not None:
                    try:
                        self.after_cancel(copy_reset_job)
                    except Exception:
                        pass
                project_number_copy_btn.configure(text=copied_button_text)
                copy_reset_job = self.after(
                    1200,
                    lambda: project_number_copy_btn.configure(text=copy_button_text),
                )

            def copy_project_name() -> None:
                nonlocal name_copy_reset_job
                project_name = self.project_name_var.get().strip()
                if not project_name:
                    self.bell()
                    return
                try:
                    self.clipboard_clear()
                    self.clipboard_append(project_name)
                    self.update_idletasks()
                except tk.TclError:
                    messagebox.showerror(
                        "Kopieren mislukt",
                        "De projectnaam kon niet naar het klembord worden gekopieerd.",
                        parent=self,
                    )
                    return
                if name_copy_reset_job is not None:
                    try:
                        self.after_cancel(name_copy_reset_job)
                    except Exception:
                        pass
                project_name_copy_btn.configure(text=copied_button_text)
                name_copy_reset_job = self.after(
                    1200,
                    lambda: project_name_copy_btn.configure(text=copy_button_text),
                )

            project_number_value = tk.Label(
                pn_row,
                textvariable=self.project_number_var,
                **field_kwargs,
            )
            project_number_value.pack(side="left", padx=(6, 0))
            project_number_copy_btn = tk.Button(
                pn_row,
                text=copy_button_text,
                width=3,
                command=copy_project_number,
            )
            project_number_copy_btn.pack(side="left", padx=(6, 0))
            self._project_number_label = project_number_label
            self._project_number_value = project_number_value
            self._project_number_copy_btn = project_number_copy_btn

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
            project_name_copy_btn = tk.Button(
                name_row,
                text=copy_button_text,
                width=3,
                command=copy_project_name,
            )
            project_name_copy_btn.pack(side="left", padx=(6, 0))
            self._project_name_label = project_name_label
            self._project_name_value = project_name_value
            self._project_name_copy_btn = project_name_copy_btn

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
                + project_number_value.winfo_reqwidth()
                + project_number_copy_btn.winfo_reqwidth()
                + 6,
                project_name_label.winfo_reqwidth()
                + project_name_value.winfo_reqwidth()
                + project_name_copy_btn.winfo_reqwidth()
                + 6,
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

            ttk.Separator(top_area, orient="horizontal").pack(fill="x", pady=(0, 6))

            delivery_opts = self._delivery_options()

            doc_type_opts = [
                "Geen",
                "Bestelbon",
                "Standaard bon",
                "Offerteaanvraag",
            ]
            self._doc_type_prefixes = {
                _prefix_for_doc_type(t) for t in doc_type_opts
            }

            header_row = tk.Frame(
                self.selection_header_canvas,
                background=left.cget("bg"),
            )
            self._selection_header_window = self.selection_header_canvas.create_window(
                (0, 0),
                window=header_row,
                anchor="nw",
            )

            def _sync_selection_header_window_width(
                viewport_width: Optional[int] = None,
            ) -> None:
                try:
                    requested_width = header_row.winfo_reqwidth()
                    canvas_width = (
                        int(viewport_width)
                        if viewport_width is not None
                        else self.selection_header_canvas.winfo_width()
                    )
                    width = max(requested_width, canvas_width)
                    self.selection_header_canvas.itemconfigure(
                        self._selection_header_window,
                        width=width,
                    )
                    requested_height = header_row.winfo_reqheight()
                    if requested_height:
                        self.selection_header_canvas.configure(height=requested_height + 3)
                except tk.TclError:
                    return

            def _update_selection_header_scroll_region(_event=None) -> None:
                _sync_selection_header_window_width()
                try:
                    bbox = self.selection_header_canvas.bbox("all")
                except tk.TclError:
                    return
                if bbox:
                    self.selection_header_canvas.configure(scrollregion=bbox)

            def _resize_selection_header_content(event) -> None:
                _sync_selection_header_window_width(event.width)
                _update_selection_header_scroll_region()

            header_row.bind("<Configure>", _update_selection_header_scroll_region)
            self.selection_header_canvas.bind(
                "<Configure>",
                _resize_selection_header_content,
            )
            self._group_header_spacer = tk.Frame(
                header_row,
                width=self.GROUP_INDICATOR_WIDTH,
                background=left.cget("bg"),
            )
            self._group_header_spacer.pack(
                side="left",
                padx=(0, self.GROUP_INDICATOR_GAP),
            )
            header_label_kwargs = dict(
                anchor=tk.W,
                justify=tk.LEFT,
                background=left.cget("bg"),
            )
            header_font = ("TkDefaultFont", 10, "bold")
            self._en1090_header_font = tkfont.Font(font=header_font)
            self._en1090_column_width_px = self._compute_initial_en1090_width()

            self._column_specs: List[Tuple[str, str, int, Optional[Tuple[str, int, str]]]] = [
                ("export_check", "Export", 8, header_font),
                ("label", "Producttype", self.LABEL_COLUMN_WIDTH, header_font),
                ("group_combo", "Bon", 24, None),
                ("supplier_combo", "Leverancier", 50, None),
                ("doc_combo", "Documenttype", 18, None),
                (
                    "en1090_widget",
                    self.EN1090_HEADER_TEXT,
                    self._compute_en1090_header_char_width(),
                    None,
                ),
                ("doc_entry", "Nr.", 12, None),
                ("unit_price_entry", "Prijs/st.", 12, None),
                ("total_price_entry", "Totaalprijs", 12, None),
                ("vat_combo", "BTW %", 7, None),
                ("line_price_button", "Regelprijzen", 10, None),
                ("remark_entry", "Opmerking", 24, None),
                ("delivery_combo", "Leveradres", 50, None),
            ]
            self._column_keys = [key for key, *_ in self._column_specs]
            self._visible_column_keys: List[str] = []
            self._refresh_visible_column_keys()

            self._header_column_frames_map: Dict[str, tk.Frame] = {}
            self._header_labels_map: Dict[str, tk.Label] = {}
            self._header_aligned = False
            self._header_alignment_pending = False

            for key, text, width, font in self._column_specs:
                label_kwargs = dict(header_label_kwargs)
                if font is not None:
                    label_kwargs["font"] = font
                column_frame = tk.Frame(header_row, background=left.cget("bg"))
                label = tk.Label(
                    column_frame,
                    text=text,
                    width=width,
                    **label_kwargs,
                )
                label.pack(fill="x", anchor="w")
                self._header_column_frames_map[key] = column_frame
                self._header_labels_map[key] = label

            self._repack_header_columns()
            self._refresh_en1090_header_width()

            self.finish_label_by_key: Dict[str, str] = {
                entry.get("key", ""): _to_str(entry.get("label")) or _to_str(entry.get("key"))
                for entry in finishes
            }

            self.rows = []
            self.combo_by_key: Dict[str, ttk.Combobox] = {}
            self._en1090_frames: List[tk.Frame] = []
            self._rows_background = left.cget("bg")

            def add_row(display_text: str, sel_key: str, metadata: Dict[str, str]):
                row_index = len(self._row_widget_maps)
                row_bg = self._row_background_for_index(row_index)
                row = tk.Frame(left, background=row_bg)
                row.pack(fill="x", pady=3)
                group_stripe = tk.Frame(
                    row,
                    width=self.GROUP_INDICATOR_WIDTH,
                    background=row_bg,
                )
                group_stripe.pack(
                    side="left",
                    fill="y",
                    padx=(0, self.GROUP_INDICATOR_GAP),
                )
                export_var = tk.IntVar(value=1)
                self.export_vars[sel_key] = export_var
                export_check = tk.Checkbutton(
                    row,
                    variable=export_var,
                    takefocus=False,
                    background=row_bg,
                    activebackground=row_bg,
                )
                row_label = tk.Label(
                    row,
                    text=display_text,
                    width=self.LABEL_COLUMN_WIDTH,
                    anchor="w",
                    background=row_bg,
                )
                _OverflowTooltip(
                    row_label,
                    lambda label=row_label: _to_str(label.cget("text")).strip(),
                )
                var = tk.StringVar()
                self.sel_vars[sel_key] = var
                combo = ttk.Combobox(row, textvariable=var, state="normal", width=50)
                combo.bind("<<ComboboxSelected>>", self._on_combo_change)
                combo.bind(
                    "<FocusIn>",
                    lambda _e, key=sel_key: self._on_supplier_combo_focus_in(key),
                )
                combo.bind("<FocusOut>", self._on_supplier_combo_focus_out)
                self._install_supplier_focus_behavior(combo)
                combo.bind(
                    "<KeyRelease>",
                    lambda ev, key=sel_key, c=combo: self._on_combo_type(ev, key, c),
                )

                group_var = tk.StringVar(value="Apart")
                self.group_vars[sel_key] = group_var
                group_combo = ttk.Combobox(
                    row,
                    textvariable=group_var,
                    values=("Apart",),
                    state="readonly",
                    width=24,
                )
                group_combo.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, key=sel_key: self._on_group_change(key),
                )
                group_combo.bind(
                    "<FocusIn>", lambda _e, key=sel_key: self._on_focus_key(key)
                )
                self.group_combos[sel_key] = group_combo

                doc_var = tk.StringVar(value="Bestelbon")
                self.doc_vars[sel_key] = doc_var
                doc_combo = ttk.Combobox(
                    row,
                    textvariable=doc_var,
                    values=doc_type_opts,
                    state="readonly",
                    width=18,
                )
                doc_combo.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, key=sel_key: self._on_doc_type_change(key),
                )

                meta_kind = metadata.get("kind")
                en1090_var = tk.IntVar(value=0)
                self.en1090_vars[sel_key] = en1090_var
                en1090_frame = tk.Frame(
                    row,
                    width=self._en1090_column_width_px,
                    background=row_bg,
                )
                en1090_frame.pack_propagate(False)
                self._en1090_frames.append(en1090_frame)
                en1090_widget: tk.Misc = en1090_frame

                if meta_kind in {"production", "opticutter"}:
                    default_flag = False
                    if self._en1090_getter is not None:
                        try:
                            default_flag = bool(self._en1090_getter(metadata))
                        except Exception:
                            default_flag = False
                    en1090_var.set(1 if default_flag else 0)
                    checkbutton = tk.Checkbutton(
                        en1090_frame,
                        variable=en1090_var,
                        command=lambda key=sel_key: self._on_en1090_toggle(key),
                        takefocus=False,
                        background=row_bg,
                        activebackground=row_bg,
                    )
                    checkbutton.pack(anchor="w", padx=(2, 0))
                    en1090_frame.update_idletasks()
                    needed_width = checkbutton.winfo_reqwidth() + 4
                    self._ensure_en1090_width(needed_width)
                    en1090_frame.configure(
                        width=self._en1090_column_width_px,
                        height=checkbutton.winfo_reqheight(),
                    )
                else:
                    placeholder = tk.Label(en1090_frame, text="", background=row_bg)
                    placeholder.pack(anchor="w", fill="x")
                    en1090_frame.update_idletasks()
                    en1090_frame.configure(
                        width=self._en1090_column_width_px,
                        height=placeholder.winfo_reqheight(),
                    )

                doc_num_var = tk.StringVar()
                self.doc_num_vars[sel_key] = doc_num_var
                doc_entry = tk.Entry(row, textvariable=doc_num_var, width=12)

                unit_price_var = tk.StringVar()
                self.price_unit_vars[sel_key] = unit_price_var
                unit_price_entry = tk.Entry(row, textvariable=unit_price_var, width=12)
                _HelpTooltip(
                    unit_price_entry,
                    "Eenheidsprijs excl. BTW uit de offerte. Als je dit invult, wordt de totaalprijs automatisch berekend.",
                )

                total_price_var = tk.StringVar()
                self.price_total_vars[sel_key] = total_price_var
                total_price_entry = tk.Entry(row, textvariable=total_price_var, width=12)
                _HelpTooltip(
                    total_price_entry,
                    "Totaalprijs excl. BTW voor deze bon/selectie. Als je dit invult, wordt de prijs/st. automatisch berekend.",
                )

                vat_var = tk.StringVar(value="21")
                self.vat_vars[sel_key] = vat_var
                vat_combo = ttk.Combobox(
                    row,
                    textvariable=vat_var,
                    values=("0", "6", "12", "21"),
                    state="normal",
                    width=7,
                )
                _HelpTooltip(
                    vat_combo,
                    "BTW-percentage voor deze bon. Standaard 21%; pas dit aan voor 0%, 6% of 12% indien nodig.",
                )

                line_price_button = tk.Button(
                    row,
                    text="Regels",
                    width=8,
                    command=lambda key=sel_key: self._open_line_pricing_dialog(key),
                )
                _HelpTooltip(
                    line_price_button,
                    "Vul prijzen per onderdeel of brutomateriaalregel in. Deze waarden worden in de exportlog bewaard.",
                )
                if not self._selection_items_for_key(sel_key):
                    line_price_button.configure(state="disabled")

                remark_var = tk.StringVar()
                self.remark_vars[sel_key] = remark_var
                remark_entry = tk.Entry(row, textvariable=remark_var, width=24)
                _scroll_entry_to_end(remark_entry, remark_var)
                _OverflowTooltip(remark_entry, lambda v=remark_var: v.get().strip())

                dvar = tk.StringVar(value=self._default_delivery_value())
                self.delivery_vars[sel_key] = dvar
                dcombo = ttk.Combobox(
                    row,
                    textvariable=dvar,
                    values=delivery_opts,
                    state="readonly",
                    width=50,
                )
                dcombo.bind(
                    "<FocusIn>", lambda _e, key=sel_key: self._on_focus_key(key)
                )
                dcombo.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, key=sel_key: self._on_focus_key(key),
                )
                self.delivery_combos[sel_key] = dcombo

                self.rows.append((sel_key, combo))
                self.combo_by_key[sel_key] = combo
                metadata_record = dict(metadata)
                metadata_record["base_display"] = display_text
                self.row_meta[sel_key] = metadata_record

                for trace_var in (
                    var,
                    doc_var,
                    doc_num_var,
                    unit_price_var,
                    total_price_var,
                    vat_var,
                    remark_var,
                    dvar,
                    export_var,
                    en1090_var,
                ):
                    trace_var.trace_add("write", lambda *_args: self._sync_grouped_rows())
                unit_price_var.trace_add(
                    "write",
                    lambda *_args, key=sel_key: self._on_price_field_change(key, "unit"),
                )
                total_price_var.trace_add(
                    "write",
                    lambda *_args, key=sel_key: self._on_price_field_change(key, "total"),
                )

                row_widgets = {
                    "row": row,
                    "group_stripe": group_stripe,
                    "export_check": export_check,
                    "label": row_label,
                    "group_combo": group_combo,
                    "supplier_combo": combo,
                    "doc_combo": doc_combo,
                    "en1090_widget": en1090_frame,
                    "doc_entry": doc_entry,
                    "unit_price_entry": unit_price_entry,
                    "total_price_entry": total_price_entry,
                    "vat_combo": vat_combo,
                    "line_price_button": line_price_button,
                    "remark_entry": remark_entry,
                    "delivery_combo": dcombo,
                    "row_background": row_bg,
                    "label_default_fg": row_label.cget("fg"),
                    "label_default_font": row_label.cget("font"),
                }
                self._row_widget_maps.append(row_widgets)
                self._row_widgets_by_key[sel_key] = row_widgets
                self._pack_row_widgets(row_widgets)

                self._schedule_header_alignment(
                    {
                        "label": row_label,
                        "group_combo": group_combo,
                        "supplier_combo": combo,
                        "doc_combo": doc_combo,
                        "en1090_widget": en1090_widget,
                        "doc_entry": doc_entry,
                        "unit_price_entry": unit_price_entry,
                        "total_price_entry": total_price_entry,
                        "vat_combo": vat_combo,
                        "line_price_button": line_price_button,
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

            # Zorg dat de kolomhoofden overeenkomen met de zichtbare kolommen bij opstart
            self.set_en1090_enabled(self._en1090_enabled)

            supplier_details_bar = tk.Frame(content)
            supplier_details_bar.grid(row=2, column=0, sticky="ew", pady=(6, 0))
            supplier_details_bar.grid_columnconfigure(1, weight=1)
            self._supplier_details_visible = False
            self._supplier_details_button = tk.Button(
                supplier_details_bar,
                text="Leverancier details tonen",
                command=self._toggle_supplier_details,
                padx=10,
            )
            self._supplier_details_button.grid(row=0, column=0, sticky="w")
            self._supplier_details_hint_var = tk.StringVar(
                value="Klik om leverancierkaarten tijdelijk te tonen."
            )
            tk.Label(
                supplier_details_bar,
                textvariable=self._supplier_details_hint_var,
                anchor="w",
                foreground="#666666",
            ).grid(row=0, column=1, sticky="ew", padx=(8, 0))

            # Container voor kaarten
            preview_frame = tk.LabelFrame(
                content,
                text="Leverancier details\n(klik om te selecteren)",
                labelanchor="n",
            )
            self._supplier_preview_frame = preview_frame
            self._supplier_preview_grid_options = {
                "row": 3,
                "column": 0,
                "sticky": "nsew",
                "pady": (4, 0),
            }
            preview_frame.grid(**self._supplier_preview_grid_options)
            preview_frame.grid_rowconfigure(0, weight=1)
            preview_frame.grid_columnconfigure(0, weight=1)

            self.cards_frame = tk.Frame(preview_frame)
            self.cards_frame.grid(row=0, column=0, sticky="nsew", pady=(8, 0))
            self._supplier_details_hide_after_id: Optional[str] = None
            self._supplier_details_auto_open = False
            self._set_supplier_details_visible(False)

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
            confirm_text = (
                "Bon-PDF's aanmaken en terug naar PDF dossier"
                if self.pdf_dossier_context
                else "Bevestig"
            )
            self.confirm_button = tk.Button(btns, text=confirm_text, command=self._confirm)
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
            _bind_selection_rows_mousewheel_tree(content)

        def _row_background_for_index(self, row_index: int) -> str:
            base_bg = getattr(self, "_rows_background", None) or self.cget("bg")
            return self.ROW_ALT_BACKGROUND if row_index % 2 else base_bg

        def _toggle_supplier_details(self) -> None:
            self._cancel_supplier_details_auto_hide()
            self._set_supplier_details_visible(
                not bool(getattr(self, "_supplier_details_visible", False)),
                automatic=False,
            )

        def _set_supplier_details_visible(
            self,
            visible: bool,
            *,
            automatic: bool = False,
        ) -> None:
            self._supplier_details_visible = bool(visible)
            self._supplier_details_auto_open = bool(automatic) if visible else False
            frame = getattr(self, "_supplier_preview_frame", None)
            if frame is not None:
                try:
                    if self._supplier_details_visible:
                        frame.grid(**getattr(self, "_supplier_preview_grid_options", {}))
                    else:
                        frame.grid_remove()
                except tk.TclError:
                    pass

            button = getattr(self, "_supplier_details_button", None)
            if button is not None:
                try:
                    button.configure(
                        text=(
                            "Leverancier details verbergen"
                            if self._supplier_details_visible
                            else "Leverancier details tonen"
                        )
                    )
                except tk.TclError:
                    pass

            hint_var = getattr(self, "_supplier_details_hint_var", None)
            if hint_var is not None:
                hint_var.set(
                    "Klik op een leverancierkaart om die leverancier te selecteren."
                    if self._supplier_details_visible
                    else "Klik om leverancierkaarten tijdelijk te tonen."
                )

        def _cancel_supplier_details_auto_hide(self) -> None:
            after_id = getattr(self, "_supplier_details_hide_after_id", None)
            if after_id:
                try:
                    self.after_cancel(after_id)
                except Exception:
                    pass
            self._supplier_details_hide_after_id = None

        def _show_supplier_details_for_supplier_search(self) -> None:
            self._cancel_supplier_details_auto_hide()
            self._set_supplier_details_visible(True, automatic=True)

        def _widget_is_or_contains(
            self,
            container: Optional["tk.Misc"],
            widget: Optional["tk.Misc"],
        ) -> bool:
            if container is None or widget is None:
                return False
            current = widget
            while current is not None:
                if current is container or str(current) == str(container):
                    return True
                try:
                    current = current.master
                except Exception:
                    return False
            return False

        def _supplier_details_focus_inside(self) -> bool:
            try:
                focused = self.focus_get()
            except tk.TclError:
                focused = None
            if focused is None:
                return False

            for combo in getattr(self, "combo_by_key", {}).values():
                if self._widget_is_or_contains(combo, focused):
                    return True
            for container_name in ("_supplier_preview_frame", "cards_frame"):
                if self._widget_is_or_contains(
                    getattr(self, container_name, None),
                    focused,
                ):
                    return True
            return False

        def _hide_supplier_details_if_focus_left(self) -> None:
            self._supplier_details_hide_after_id = None
            if not bool(getattr(self, "_supplier_details_auto_open", False)):
                return
            if self._supplier_details_focus_inside():
                return
            self._set_supplier_details_visible(False)

        def _schedule_supplier_details_auto_hide(self) -> None:
            if not bool(getattr(self, "_supplier_details_auto_open", False)):
                return
            self._cancel_supplier_details_auto_hide()
            try:
                self._supplier_details_hide_after_id = self.after(
                    180,
                    self._hide_supplier_details_if_focus_left,
                )
            except tk.TclError:
                self._hide_supplier_details_if_focus_left()

        def _on_supplier_combo_focus_in(self, sel_key: str) -> None:
            self._on_focus_key(sel_key)
            self._cancel_supplier_details_auto_hide()
            combo = self.combo_by_key.get(sel_key)
            text = ""
            if combo is not None:
                try:
                    text = combo.get().strip()
                except tk.TclError:
                    text = ""
            if text and text.lower() not in {"(geen)", "geen"}:
                self._show_supplier_details_for_supplier_search()

        def _on_supplier_combo_focus_out(self, _event=None) -> None:
            self._schedule_supplier_details_auto_hide()

        def _resolve_current_client(self) -> Optional[Client]:
            clients_db = getattr(self, "clients_db", None)
            client_var = getattr(self, "client_var", None)
            if clients_db is None or client_var is None:
                return None

            try:
                raw_value = str(client_var.get() or "").strip()
            except Exception:
                raw_value = ""
            if not raw_value:
                return None

            lowered = raw_value.lower()
            for client in getattr(clients_db, "clients", []):
                name = str(getattr(client, "name", "") or "").strip()
                if name and lowered == name.lower():
                    return client
                try:
                    display_name = str(clients_db.display_name(client) or "").strip()
                except Exception:
                    display_name = ""
                if display_name and lowered == display_name.lower():
                    return client

            getter = getattr(clients_db, "get", None)
            if callable(getter):
                try:
                    return getter(raw_value)
                except Exception:
                    return None
            return None

        def _default_delivery_value(self) -> str:
            return self.DELIVERY_PRESETS[0]

        @classmethod
        def _is_client_delivery_choice(cls, value: object) -> bool:
            clean = strip_favorite_marker(_to_str(value)).strip().casefold()
            return clean in {
                cls.CLIENT_DELIVERY_PRESET.casefold(),
                cls.LEGACY_CLIENT_DELIVERY_PRESET.casefold(),
            }

        def _delivery_options(self) -> List[str]:
            options = list(self.DELIVERY_PRESETS)
            options.extend(
                self.delivery_db.display_name(a)
                for a in self.delivery_db.addresses_sorted()
            )
            return options

        def _open_delivery_address_dialog(
            self, addr: Optional[DeliveryAddress] = None
        ) -> Optional[DeliveryAddress]:
            win = tk.Toplevel(self)
            win.title("Leveradres")
            fields = [
                ("Naam", "name"),
                ("Adres", "address"),
                ("Opmerkingen", "remarks"),
            ]
            entries: Dict[str, tk.Entry] = {}
            for i, (lbl, key) in enumerate(fields):
                tk.Label(win, text=lbl + ":").grid(
                    row=i, column=0, sticky="e", padx=4, pady=2
                )
                ent = tk.Entry(win, width=40)
                ent.grid(row=i, column=1, padx=4, pady=2)
                if addr:
                    ent.insert(0, _to_str(getattr(addr, key)))
                entries[key] = ent
            fav_var = tk.BooleanVar(value=addr.favorite if addr else False)
            tk.Checkbutton(win, text="Favoriet", variable=fav_var).grid(
                row=len(fields), column=1, sticky="w", padx=4, pady=2
            )

            result: Dict[str, Optional[DeliveryAddress]] = {"value": None}

            def _save() -> None:
                rec = {k: (e.get().strip() or None) for k, e in entries.items()}
                rec["favorite"] = fav_var.get()
                if not rec["name"]:
                    messagebox.showwarning(
                        "Let op", "Naam is verplicht.", parent=win
                    )
                    return
                result["value"] = DeliveryAddress.from_any(rec)
                win.destroy()

            btnf = tk.Frame(win)
            btnf.grid(row=len(fields) + 1, column=0, columnspan=2, pady=6)
            tk.Button(btnf, text="Opslaan", command=_save).pack(side="left", padx=4)
            tk.Button(btnf, text="Annuleer", command=win.destroy).pack(
                side="left", padx=4
            )
            win.transient(self)
            _place_window_near_parent(win, self)
            win.grab_set()
            entries["name"].focus_set()
            self.wait_window(win)
            return result["value"]

        def _add_delivery_address(self) -> None:
            created = self._open_delivery_address_dialog(None)
            if created is None:
                return
            self.delivery_db.upsert(created)
            self.delivery_db.save(DELIVERY_DB_FILE)
            self._refresh_options()
            display_name = self.delivery_db.display_name(created)
            active_key = self._active_key
            if not active_key and self.rows:
                active_key = self.rows[0][0]
            if active_key:
                dvar = self.delivery_vars.get(active_key)
                dcombo = self.delivery_combos.get(active_key)
                if dvar is not None:
                    dvar.set(display_name)
                if dcombo is not None:
                    dcombo.set(display_name)

        def _line_pricing_count(self, sel_key: str) -> int:
            return sum(
                1
                for value in self.line_pricing.get(sel_key, {}).values()
                if isinstance(value, Mapping)
                and (
                    _to_str(value.get("unit_price")).strip()
                    or _to_str(value.get("total_price")).strip()
                )
            )

        def _refresh_line_price_button(self, sel_key: str) -> None:
            widgets = self._row_widgets_by_key.get(sel_key, {})
            button = widgets.get("line_price_button")
            if button is None:
                return
            count = self._line_pricing_count(sel_key)
            text = f"Regels ({count})" if count else "Regels"
            try:
                button.configure(text=text)
            except tk.TclError:
                pass

        def _selection_items_for_key(self, sel_key: str) -> List[Dict[str, object]]:
            items = [
                dict(item)
                for item in (self.selection_items.get(sel_key) or [])
                if isinstance(item, Mapping)
            ]
            if items:
                return items

            kind, identifier = self._parse_selection_key(sel_key)
            if kind != "opticutter":
                return []
            comp = (self.opticutter_details or {}).get(identifier)
            raw_items = getattr(comp, "raw_items", None)
            if not raw_items:
                return []
            derived: List[Dict[str, object]] = []
            for raw_item in raw_items:
                if not isinstance(raw_item, Mapping):
                    continue
                item = dict(raw_item)
                key = build_order_pricing_item_key(
                    item,
                    context_kind="Brutemateriaal",
                )
                profile = _to_str(item.get("Profiel")).strip()
                material = _to_str(item.get("Materiaal")).strip()
                length = _to_str(item.get("Lengte")).strip()
                label = " - ".join(part for part in (profile, material, length) if part)
                item["key"] = key
                item["label"] = label or key
                item["quantity"] = _to_str(item.get("St.")).strip()
                derived.append(item)
            return derived

        def _pricing_quantity_for_key(self, sel_key: str) -> Optional[Decimal]:
            total = Decimal("0")
            seen_any = False
            for item in self._selection_items_for_key(sel_key):
                qty = _parse_supplier_decimal(
                    item.get("quantity")
                    or item.get("Aantal")
                    or item.get("St.")
                    or ""
                )
                if qty is None:
                    continue
                total += qty
                seen_any = True
            if not seen_any or total == 0:
                return None
            return total

        def _on_price_field_change(
            self,
            sel_key: str,
            source: str,
            *,
            force: bool = True,
        ) -> None:
            if sel_key in self._price_link_in_progress:
                return
            source = "total" if source == "total" else "unit"
            if self._price_auto_fields.get(sel_key) == source:
                self._price_auto_fields.pop(sel_key, None)

            quantity = self._pricing_quantity_for_key(sel_key)
            if quantity is None:
                return

            unit_var = self.price_unit_vars.get(sel_key)
            total_var = self.price_total_vars.get(sel_key)
            if unit_var is None or total_var is None:
                return

            source_var = unit_var if source == "unit" else total_var
            target = "total" if source == "unit" else "unit"
            target_var = total_var if target == "total" else unit_var
            source_text = source_var.get().strip()
            target_text = target_var.get().strip()
            target_is_auto = self._price_auto_fields.get(sel_key) == target

            if not source_text:
                if target_is_auto and target_text:
                    self._price_link_in_progress.add(sel_key)
                    try:
                        target_var.set("")
                    finally:
                        self._price_link_in_progress.discard(sel_key)
                    self._price_auto_fields.pop(sel_key, None)
                return

            source_value = _parse_supplier_decimal(source_text)
            if source_value is None:
                return
            if target_text and not target_is_auto and not force:
                return

            if source == "unit":
                calculated = source_value * quantity
            else:
                calculated = source_value / quantity
            calculated_text = _format_supplier_decimal(calculated)

            if target_text == calculated_text:
                self._price_auto_fields[sel_key] = target
                return

            self._price_link_in_progress.add(sel_key)
            try:
                target_var.set(calculated_text)
            finally:
                self._price_link_in_progress.discard(sel_key)
            self._price_auto_fields[sel_key] = target

        def _open_line_pricing_dialog(self, sel_key: str) -> None:
            items = self._selection_items_for_key(sel_key)
            if not items:
                messagebox.showinfo(
                    "Regelprijzen",
                    "Voor deze selectie zijn geen orderregels beschikbaar.",
                    parent=self,
                )
                return

            win = tk.Toplevel(self)
            win.title(f"Regelprijzen - {self._base_row_label(sel_key)}")
            win.transient(self)
            win.grid_columnconfigure(0, weight=1)
            win.grid_rowconfigure(1, weight=1)

            tk.Label(
                win,
                text=(
                    "Vul alleen de prijzen in die je van de leverancier hebt gekregen. "
                    "Lege velden blijven leeg; een bonbrede prijs blijft als fallback gelden."
                ),
                anchor="w",
                justify="left",
                wraplength=760,
            ).grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))

            container = tk.Frame(win)
            container.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))
            container.grid_columnconfigure(0, weight=1)
            container.grid_rowconfigure(0, weight=1)

            canvas = tk.Canvas(container, highlightthickness=0, borderwidth=0)
            canvas.grid(row=0, column=0, sticky="nsew")
            scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")
            canvas.configure(yscrollcommand=scrollbar.set)

            inner = tk.Frame(canvas)
            inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

            def _update_scroll_region(_event=None) -> None:
                try:
                    canvas.configure(scrollregion=canvas.bbox("all"))
                except tk.TclError:
                    pass

            def _resize_inner(event) -> None:
                try:
                    canvas.itemconfigure(inner_id, width=event.width)
                except tk.TclError:
                    pass

            inner.bind("<Configure>", _update_scroll_region)
            canvas.bind("<Configure>", _resize_inner)

            for col, weight in ((0, 1), (1, 0), (2, 0), (3, 0)):
                inner.grid_columnconfigure(col, weight=weight)

            header_font = ("TkDefaultFont", 10, "bold")
            headers = ("Onderdeel", "St.", "Prijs/st.", "Totaalprijs")
            for col, text in enumerate(headers):
                header_padx = (0, 18) if col == len(headers) - 1 else (0, 8)
                tk.Label(inner, text=text, font=header_font, anchor="w").grid(
                    row=0,
                    column=col,
                    sticky="ew",
                    padx=header_padx,
                    pady=(0, 4),
                )

            existing = self.line_pricing.get(sel_key, {})
            vars_by_key: Dict[str, Tuple[tk.StringVar, tk.StringVar]] = {}
            line_auto_fields: Dict[str, str] = {}
            line_link_in_progress: set[str] = set()

            def _on_line_price_change(
                item_key: str,
                source: str,
                quantity_text: str,
                *,
                force: bool = True,
            ) -> None:
                if item_key in line_link_in_progress:
                    return
                source = "total" if source == "total" else "unit"
                if line_auto_fields.get(item_key) == source:
                    line_auto_fields.pop(item_key, None)

                quantity_value = _parse_supplier_decimal(quantity_text)
                if quantity_value is None or quantity_value == 0:
                    return
                pair = vars_by_key.get(item_key)
                if pair is None:
                    return
                unit_var, total_var = pair
                source_var = unit_var if source == "unit" else total_var
                target = "total" if source == "unit" else "unit"
                target_var = total_var if target == "total" else unit_var
                source_text = source_var.get().strip()
                target_text = target_var.get().strip()
                target_is_auto = line_auto_fields.get(item_key) == target

                if not source_text:
                    if target_is_auto and target_text:
                        line_link_in_progress.add(item_key)
                        try:
                            target_var.set("")
                        finally:
                            line_link_in_progress.discard(item_key)
                        line_auto_fields.pop(item_key, None)
                    return

                source_value = _parse_supplier_decimal(source_text)
                if source_value is None:
                    return
                if target_text and not target_is_auto and not force:
                    return

                if source == "unit":
                    calculated = source_value * quantity_value
                else:
                    calculated = source_value / quantity_value
                calculated_text = _format_supplier_decimal(calculated)

                if target_text == calculated_text:
                    line_auto_fields[item_key] = target
                    return

                line_link_in_progress.add(item_key)
                try:
                    target_var.set(calculated_text)
                finally:
                    line_link_in_progress.discard(item_key)
                line_auto_fields[item_key] = target

            default_font = tkfont.nametofont("TkDefaultFont")
            for row_index, item in enumerate(items, start=1):
                item_key = _to_str(item.get("key")).strip()
                if not item_key:
                    continue
                label_text = _to_str(item.get("label")).strip()
                if not label_text:
                    part_number = _to_str(item.get("PartNumber")).strip()
                    description = _to_str(item.get("Description")).strip()
                    label_text = " - ".join(
                        part for part in (part_number, description) if part
                    )
                label_text = label_text or item_key
                quantity = _to_str(
                    item.get("quantity")
                    or item.get("Aantal")
                    or item.get("St.")
                    or ""
                ).strip()
                price_info = existing.get(item_key, {})
                if not isinstance(price_info, Mapping):
                    price_info = {}
                unit_var = tk.StringVar(
                    value=_to_str(price_info.get("unit_price")).strip()
                )
                total_var = tk.StringVar(
                    value=_to_str(price_info.get("total_price")).strip()
                )
                vars_by_key[item_key] = (unit_var, total_var)

                label = tk.Label(
                    inner,
                    text=label_text,
                    anchor="w",
                    justify="left",
                    wraplength=420,
                )
                label.grid(row=row_index, column=0, sticky="ew", padx=(0, 8), pady=2)
                tk.Label(inner, text=quantity, anchor="e", width=8).grid(
                    row=row_index,
                    column=1,
                    sticky="ew",
                    padx=(0, 8),
                    pady=2,
                )
                unit_entry = tk.Entry(
                    inner,
                    textvariable=unit_var,
                    width=14,
                    font=default_font,
                )
                unit_entry.grid(row=row_index, column=2, sticky="ew", padx=(0, 8), pady=2)
                total_entry = tk.Entry(
                    inner,
                    textvariable=total_var,
                    width=14,
                    font=default_font,
                )
                total_entry.grid(
                    row=row_index,
                    column=3,
                    sticky="ew",
                    padx=(0, 18),
                    pady=2,
                )
                unit_var.trace_add(
                    "write",
                    lambda *_args, key=item_key, qty=quantity: _on_line_price_change(
                        key, "unit", qty
                    ),
                )
                total_var.trace_add(
                    "write",
                    lambda *_args, key=item_key, qty=quantity: _on_line_price_change(
                        key, "total", qty
                    ),
                )
                if unit_var.get().strip():
                    _on_line_price_change(item_key, "unit", quantity, force=False)
                elif total_var.get().strip():
                    _on_line_price_change(item_key, "total", quantity, force=False)

            btns = tk.Frame(win)
            btns.grid(row=2, column=0, sticky="e", padx=12, pady=(0, 12))

            def _save() -> None:
                new_values: Dict[str, Dict[str, str]] = {}
                for item_key, (unit_var, total_var) in vars_by_key.items():
                    old = existing.get(item_key, {})
                    if not isinstance(old, Mapping):
                        old = {}
                    entry = {
                        "unit_price": unit_var.get().strip(),
                        "total_price": total_var.get().strip(),
                        "quote_ref": _to_str(old.get("quote_ref")).strip(),
                        "note": _to_str(old.get("note")).strip(),
                    }
                    if any(entry.values()):
                        new_values[item_key] = entry
                if new_values:
                    self.line_pricing[sel_key] = new_values
                else:
                    self.line_pricing.pop(sel_key, None)
                self._refresh_line_price_button(sel_key)
                sync_grouped_rows = getattr(self, "_sync_grouped_rows", None)
                if callable(sync_grouped_rows):
                    sync_grouped_rows()
                win.destroy()

            tk.Button(btns, text="Opslaan", command=_save).pack(side="left", padx=(0, 6))
            tk.Button(btns, text="Annuleer", command=win.destroy).pack(side="left")
            _place_window_near_parent(win, self)
            win.grab_set()
            win.focus_set()

        def _open_preset_manager(self) -> None:
            callback = getattr(self, "on_manage_presets", None)
            if callable(callback):
                callback()

        @staticmethod
        def _clean_display_value(value: object) -> str:
            return strip_favorite_marker(_to_str(value)).strip()

        def _store_preset_state(self, sel_key: str, evaluation) -> None:
            preset_state_by_key = getattr(self, "_preset_state_by_key", None)
            if not isinstance(preset_state_by_key, dict):
                preset_state_by_key = {}
                setattr(self, "_preset_state_by_key", preset_state_by_key)
            applied_rule_names = [
                _to_str(name).strip()
                for name in getattr(evaluation, "applied_rule_names", [])
                if _to_str(name).strip()
            ]
            if not applied_rule_names:
                preset_state_by_key.pop(sel_key, None)
                return

            field_names: List[str] = []
            if _to_str(getattr(evaluation, "supplier", "")).strip():
                field_names.append("leverancier")
            if _to_str(getattr(evaluation, "doc_type", "")).strip():
                field_names.append("documenttype")
            if _to_str(getattr(evaluation, "delivery", "")).strip():
                field_names.append("leveradres")
            if _to_str(getattr(evaluation, "remark", "")).strip():
                field_names.append("opmerking")
            if getattr(evaluation, "en1090", None) is not None:
                field_names.append("EN1090")

            preset_state_by_key[sel_key] = {
                "applied_rule_names": applied_rule_names,
                "field_names": field_names,
            }

        def _preset_indicator_suffix(self, sel_key: str) -> str:
            preset_state_by_key = getattr(self, "_preset_state_by_key", {})
            if sel_key not in preset_state_by_key:
                return ""
            return " [Preset]"

        def _export_log_indicator_suffix(self, sel_key: str) -> str:
            export_log_state_by_key = getattr(self, "_export_log_state_by_key", {})
            if sel_key not in export_log_state_by_key:
                return ""
            return " [Log]"

        def _store_export_log_state(
            self,
            applied_keys: Iterable[str],
            *,
            source_label: str,
            converted: bool,
        ) -> None:
            state_by_key: Dict[str, Dict[str, object]] = {}
            for key in applied_keys:
                clean_key = _to_str(key).strip()
                if not clean_key:
                    continue
                state_by_key[clean_key] = {
                    "source": source_label,
                    "converted": bool(converted),
                }
            self._export_log_state_by_key = state_by_key

        def _preset_status_message(self, sel_key: str) -> str:
            preset_state_by_key = getattr(self, "_preset_state_by_key", {})
            state = preset_state_by_key.get(sel_key)
            if not state:
                return ""
            base_label = self._base_row_label(sel_key)
            applied_rule_names = state.get("applied_rule_names") or []
            field_names = state.get("field_names") or []
            rules_text = ", ".join(
                _to_str(name).strip()
                for name in applied_rule_names
                if _to_str(name).strip()
            )
            fields_text = ", ".join(
                _to_str(name).strip()
                for name in field_names
                if _to_str(name).strip()
            )
            if rules_text and fields_text:
                return f"{base_label}: ingevuld via presetregel(s) {rules_text} ({fields_text})."
            if rules_text:
                return f"{base_label}: ingevuld via presetregel(s) {rules_text}."
            return f"{base_label}: preset toegepast."

        def _export_log_status_message(self, sel_key: str) -> str:
            export_log_state_by_key = getattr(self, "_export_log_state_by_key", {})
            state = export_log_state_by_key.get(sel_key)
            if not state:
                return ""
            base_label = self._base_row_label(sel_key)
            source = _to_str(state.get("source")).strip()
            converted = bool(state.get("converted"))
            message = f"{base_label}: waarden geladen uit exportlog"
            if source:
                message += f" ({source})"
            if converted:
                message += "; offerteaanvragen omgezet naar bestelbonnen"
            return message + "."

        def _find_supplier_display_for_name(self, supplier_name: str) -> str:
            clean = self._clean_display_value(supplier_name)
            if not clean:
                return ""
            for display, raw_name in getattr(self, "_disp_to_name", {}).items():
                if _to_str(raw_name).strip().lower() == clean.lower():
                    return display
            return clean

        def _find_delivery_display_for_name(self, delivery_name: str) -> str:
            clean = self._clean_display_value(delivery_name)
            if not clean:
                return ""
            if SupplierSelectionFrame._is_client_delivery_choice(clean):
                return self.CLIENT_DELIVERY_PRESET
            options = []
            delivery_options = getattr(self, "_delivery_options", None)
            if callable(delivery_options):
                try:
                    options = list(delivery_options())
                except Exception:
                    options = []
            for option in options:
                if self._clean_display_value(option).lower() == clean.lower():
                    return option
            return clean

        def _preset_context_for_row(self, sel_key: str) -> Dict[str, str]:
            parser = getattr(
                self,
                "_parse_selection_key",
                SupplierSelectionFrame._parse_selection_key,
            )
            kind, identifier = parser(sel_key)
            client_name = ""
            client = self._resolve_current_client()
            if client is not None:
                client_name = _to_str(getattr(client, "name", "")).strip()
            elif getattr(self, "client_var", None) is not None:
                try:
                    client_name = self._clean_display_value(self.client_var.get())
                except Exception:
                    client_name = ""
            return {
                "client": client_name,
                "selection_kind": kind,
                "identifier": _to_str(identifier).strip(),
            }

        def _apply_preset_evaluation_to_row(self, sel_key: str, evaluation) -> bool:
            changed = False

            supplier_name = _to_str(getattr(evaluation, "supplier", "")).strip()
            if supplier_name:
                combo = self.combo_by_key.get(sel_key)
                if combo is not None:
                    display_value = self._find_supplier_display_for_name(supplier_name)
                    if combo.get().strip() != display_value:
                        self._set_combo_value(combo, display_value)
                        changed = True

            doc_type = _to_str(getattr(evaluation, "doc_type", "")).strip()
            if doc_type:
                doc_var = self.doc_vars.get(sel_key)
                if doc_var is not None and doc_var.get().strip() != doc_type:
                    doc_var.set(doc_type)
                    changed = True

            delivery = _to_str(getattr(evaluation, "delivery", "")).strip()
            if delivery:
                dvar = self.delivery_vars.get(sel_key)
                if dvar is not None:
                    display_value = self._find_delivery_display_for_name(delivery)
                    if dvar.get().strip() != display_value:
                        dvar.set(display_value)
                        dcombo = self.delivery_combos.get(sel_key)
                        if dcombo is not None:
                            dcombo.set(display_value)
                        changed = True

            remark = _to_str(getattr(evaluation, "remark", "")).strip()
            if remark:
                remark_var = self.remark_vars.get(sel_key)
                if remark_var is not None and remark_var.get().strip() != remark:
                    remark_var.set(remark)
                    changed = True

            en1090_value = getattr(evaluation, "en1090", None)
            if en1090_value is not None:
                en1090_var = self.en1090_vars.get(sel_key)
                if en1090_var is not None and bool(en1090_var.get()) != bool(en1090_value):
                    en1090_var.set(1 if en1090_value else 0)
                    changed = True

            if doc_type:
                try:
                    self._on_doc_type_change(sel_key)
                except Exception:
                    pass

            return changed

        def _apply_presets(self, *, auto_apply_only: bool, status_when_idle: bool) -> int:
            presets_db = getattr(self, "presets_db", None)
            if presets_db is None:
                if status_when_idle:
                    self.update_status("Geen presetregels beschikbaar.")
                return 0

            changed_rows = 0
            applied_rule_names: set[str] = set()

            for sel_key, _combo in self.rows:
                context = self._preset_context_for_row(sel_key)
                try:
                    evaluation = presets_db.evaluate(
                        context,
                        auto_apply_only=bool(auto_apply_only),
                    )
                except Exception:
                    continue
                if not evaluation.has_changes():
                    preset_state_by_key = getattr(self, "_preset_state_by_key", None)
                    if isinstance(preset_state_by_key, dict):
                        preset_state_by_key.pop(sel_key, None)
                    continue
                self._store_preset_state(sel_key, evaluation)
                applied_rule_names.update(
                    name for name in getattr(evaluation, "applied_rule_names", []) if name
                )
                if self._apply_preset_evaluation_to_row(sel_key, evaluation):
                    changed_rows += 1

            sync_grouped_rows = getattr(self, "_sync_grouped_rows", None)
            if callable(sync_grouped_rows):
                sync_grouped_rows()
            self._update_preview_from_any_combo()

            if status_when_idle or changed_rows:
                if changed_rows:
                    message = (
                        f"{changed_rows} rij(en) ingevuld via "
                        f"{len(applied_rule_names) or changed_rows} presetregel(s)."
                    )
                else:
                    message = "Geen passende presetregels gevonden."
                self._last_preset_status = message
                self.update_status(message)
            return changed_rows

        def _apply_presets_from_button(self) -> None:
            self._apply_presets(auto_apply_only=False, status_when_idle=True)

        def _state_selection_keys(self, state_dict: Mapping[str, object]) -> set[str]:
            keys: set[str] = set()
            for name in (
                "selections",
                "groups",
                "doc_types",
                "doc_numbers",
                "remarks",
                "deliveries",
                "exports",
                "en1090",
                "vat_rates",
                "pricing",
            ):
                value = state_dict.get(name, {}) if isinstance(state_dict, Mapping) else {}
                if isinstance(value, Mapping):
                    keys.update(_to_str(key) for key in value.keys() if _to_str(key).strip())
            return keys

        def _open_filesystem_path(self, path: str, *, parent=None) -> None:
            clean_path = _to_str(path).strip()
            if not clean_path:
                return
            if not os.path.exists(clean_path):
                messagebox.showwarning(
                    "Niet gevonden",
                    f"Kan dit pad niet vinden:\n{clean_path}",
                    parent=parent or self,
                )
                return
            try:
                if sys.platform.startswith("win"):
                    os.startfile(clean_path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", clean_path], check=False)
                else:
                    subprocess.run(["xdg-open", clean_path], check=False)
            except Exception as exc:
                messagebox.showwarning(
                    "Openen mislukt",
                    f"Kon dit pad niet openen:\n{exc}",
                    parent=parent or self,
                )

        def _open_loaded_export_documents_dialog(self) -> None:
            log_path = _to_str(getattr(self, "_loaded_export_log_path", "")).strip()
            if not log_path:
                messagebox.showinfo(
                    "Vorige documenten",
                    "Laad eerst een exportlog.",
                    parent=self,
                )
                return
            export_dir = os.path.dirname(log_path)
            export_info = getattr(self, "_loaded_export_log_export_info", {}) or {}
            documents = export_info.get("generated_documents", [])
            documents = [doc for doc in documents if isinstance(doc, Mapping)]
            if not documents:
                if messagebox.askyesno(
                    "Vorige documenten",
                    "Deze exportlog bevat geen documentlijst. Wil je de exportmap openen?",
                    parent=self,
                ):
                    self._open_filesystem_path(export_dir, parent=self)
                return

            win = tk.Toplevel(self)
            win.title("Vorige exportdocumenten")
            win.transient(self)
            win.geometry("980x460")
            win.grid_columnconfigure(0, weight=1)
            win.grid_rowconfigure(0, weight=1)

            columns = ("type", "context", "document", "leverancier", "pad")
            tree = ttk.Treeview(win, columns=columns, show="headings", selectmode="browse")
            headings = {
                "type": "Type",
                "context": "Context",
                "document": "Document",
                "leverancier": "Leverancier",
                "pad": "Pad",
            }
            widths = {
                "type": 90,
                "context": 180,
                "document": 190,
                "leverancier": 150,
                "pad": 330,
            }
            for column in columns:
                tree.heading(column, text=headings[column], anchor="w")
                tree.column(column, width=widths[column], anchor="w", stretch=True)
            tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
            scroll = ttk.Scrollbar(win, orient="vertical", command=tree.yview)
            scroll.grid(row=0, column=1, sticky="ns", pady=12)
            tree.configure(yscrollcommand=scroll.set)

            path_by_iid: Dict[str, str] = {}
            for index, record in enumerate(documents):
                resolved_path = resolve_export_document_path(log_path, record)
                context = " - ".join(
                    part
                    for part in (
                        _to_str(record.get("context_kind")).strip(),
                        _to_str(record.get("context_label")).strip(),
                    )
                    if part
                )
                document_label = " ".join(
                    part
                    for part in (
                        _to_str(record.get("doc_type")).strip(),
                        _to_str(record.get("doc_number")).strip(),
                    )
                    if part
                )
                if not document_label:
                    document_label = _to_str(record.get("kind")).strip()
                iid = f"doc_{index}"
                tree.insert(
                    "",
                    "end",
                    iid=iid,
                    values=(
                        _to_str(record.get("format")).upper(),
                        context,
                        document_label,
                        _to_str(record.get("supplier")).strip(),
                        _to_str(record.get("path")).strip(),
                    ),
                )
                path_by_iid[iid] = resolved_path

            if documents:
                first = tree.get_children()[0]
                tree.selection_set(first)
                tree.focus(first)

            buttons = tk.Frame(win)
            buttons.grid(row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 12))
            buttons.grid_columnconfigure(0, weight=1)

            def _selected_path() -> str:
                selected = tree.selection()
                if not selected:
                    return ""
                return path_by_iid.get(selected[0], "")

            def _open_selected() -> None:
                self._open_filesystem_path(_selected_path(), parent=win)

            def _open_selected_folder() -> None:
                selected_path = _selected_path()
                if selected_path:
                    self._open_filesystem_path(os.path.dirname(selected_path), parent=win)

            tree.bind("<Double-1>", lambda _event: _open_selected())
            tk.Button(buttons, text="Open document", command=_open_selected).grid(
                row=0, column=1, padx=(6, 0), sticky="e"
            )
            tk.Button(buttons, text="Open map", command=_open_selected_folder).grid(
                row=0, column=2, padx=(6, 0), sticky="e"
            )
            tk.Button(
                buttons,
                text="Open exportmap",
                command=lambda: self._open_filesystem_path(export_dir, parent=win),
            ).grid(row=0, column=3, padx=(6, 0), sticky="e")
            tk.Button(buttons, text="Sluiten", command=win.destroy).grid(
                row=0, column=4, padx=(6, 0), sticky="e"
            )
            _place_window_near_parent(win, self)

        def _prompt_export_log_import_sections(self, *, converted: bool) -> Optional[set[str]]:
            win = tk.Toplevel(self)
            win.title("Exportlog toepassen")
            win.transient(self)
            win.resizable(False, False)

            tk.Label(
                win,
                text=(
                    "Kies welke waarden uit de exportlog de huidige "
                    "bestelbonpagina mogen overschrijven."
                ),
                anchor="w",
                justify="left",
                wraplength=520,
            ).grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
            if converted:
                tk.Label(
                    win,
                    text="Offerteaanvragen zijn in de importwaarden omgezet naar Bestelbon.",
                    anchor="w",
                    justify="left",
                    foreground="#7A4A00",
                    wraplength=520,
                ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 8))

            options = [
                ("suppliers", "Leveranciers", "Leverancierkeuze per selectie."),
                ("groups", "Bonkoppelingen", "Welke selecties samen op een bon staan."),
                ("documents", "Documenttype en nummers", "Bestelbon/offerte en documentnummers."),
                ("deliveries", "Leveradressen", "Leveradreskeuzes."),
                ("remarks", "Opmerkingen", "Opmerkingen op bonnen."),
                ("exports", "Export aan/uit", "Welke regels worden mee geexporteerd."),
                ("en1090", "EN 1090", "EN 1090-vlaggen."),
                ("pricing", "Prijzen", "Bonprijzen en regelprijzen."),
                ("vat", "BTW", "BTW-percentages per bon."),
            ]
            vars_by_section: Dict[str, tk.BooleanVar] = {}
            rows_frame = tk.Frame(win)
            rows_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
            for index, (key, label, help_text) in enumerate(options):
                var = tk.BooleanVar(value=True)
                vars_by_section[key] = var
                row = tk.Frame(rows_frame)
                row.grid(row=index, column=0, sticky="ew", pady=2)
                check = tk.Checkbutton(row, text=label, variable=var, anchor="w")
                check.pack(side="left")
                _HelpTooltip(check, help_text)

            result: Dict[str, Optional[set[str]]] = {"sections": None}

            def _set_all(value: bool) -> None:
                for var in vars_by_section.values():
                    var.set(bool(value))

            def _only_pricing() -> None:
                for section, var in vars_by_section.items():
                    var.set(section in {"pricing", "vat"})

            def _apply() -> None:
                selected = {
                    section
                    for section, var in vars_by_section.items()
                    if bool(var.get())
                }
                if not selected:
                    messagebox.showwarning(
                        "Exportlog toepassen",
                        "Kies minstens een onderdeel om toe te passen.",
                        parent=win,
                    )
                    return
                result["sections"] = selected
                win.destroy()

            button_row = tk.Frame(win)
            button_row.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 12))
            button_row.grid_columnconfigure(0, weight=1)
            tk.Button(button_row, text="Alles", command=lambda: _set_all(True)).grid(
                row=0, column=0, sticky="w"
            )
            tk.Button(button_row, text="Alleen prijzen", command=_only_pricing).grid(
                row=0, column=1, padx=(6, 0), sticky="e"
            )
            tk.Button(button_row, text="Toepassen", command=_apply).grid(
                row=0, column=2, padx=(12, 0), sticky="e"
            )
            tk.Button(button_row, text="Annuleer", command=win.destroy).grid(
                row=0, column=3, padx=(6, 0), sticky="e"
            )
            _place_window_near_parent(win, self)
            win.grab_set()
            win.focus_set()
            self.wait_window(win)
            return result["sections"]

        def _load_export_log_from_file(self) -> None:
            parent_app = getattr(self.master, "master", None)
            initial_dir = _to_str(getattr(parent_app, "dest_folder", "")).strip()
            if not initial_dir or not os.path.isdir(initial_dir):
                initial_dir = os.getcwd()
            path = filedialog.askopenfilename(
                parent=self,
                title="Exportlog laden",
                initialdir=initial_dir,
                filetypes=[
                    ("Filehopper exportlog", EXPORT_SESSION_LOG_FILENAME),
                    ("JSON", "*.json"),
                    ("Alle bestanden", "*.*"),
                ],
            )
            if not path:
                return

            self._load_export_log_path(path)

        def _load_latest_export_log(self) -> None:
            parent_app = getattr(self.master, "master", None)
            root_dir = _to_str(getattr(parent_app, "dest_folder", "")).strip()
            if not root_dir or not os.path.isdir(root_dir):
                messagebox.showwarning(
                    "Laatste exportlog",
                    "Selecteer eerst een geldige exportbestemming.",
                    parent=self,
                )
                return
            logs = find_export_session_logs(root_dir, limit=1)
            if not logs:
                messagebox.showinfo(
                    "Laatste exportlog",
                    f"Geen {EXPORT_SESSION_LOG_FILENAME} gevonden onder:\n{root_dir}",
                    parent=self,
                )
                return
            self._load_export_log_path(logs[0])

        def _load_export_log_path(self, path: str) -> None:
            parent_app = getattr(self.master, "master", None)
            try:
                payload = load_export_session_log(path)
            except Exception as exc:
                messagebox.showerror(
                    "Exportlog laden",
                    f"Kon exportlog niet laden:\n{exc}",
                    parent=self,
                )
                return

            state_dict = dict(payload.get("order_state", {}) or {})
            doc_types = state_dict.get("doc_types", {})
            offer_count = 0
            if isinstance(doc_types, Mapping):
                offer_count = sum(
                    1
                    for value in doc_types.values()
                    if _to_str(value).strip().lower().startswith("offerte")
                )
            converted = False
            if offer_count:
                answer = messagebox.askyesnocancel(
                    "Offerte omzetten",
                    (
                        f"Deze exportlog bevat {offer_count} offertebon(nen).\n\n"
                        "Ja: omzetten naar Bestelbon en OFF-nummers leegmaken.\n"
                        "Nee: documenttypes behouden.\n"
                        "Annuleer: niets inladen."
                    ),
                    parent=self,
                )
                if answer is None:
                    return
                if answer:
                    state_dict = convert_offers_to_orders(state_dict)
                    converted = True

            current_keys = {key for key, _combo in self.rows}
            payload_for_review = dict(payload)
            payload_for_review["order_state"] = state_dict
            try:
                compatibility = summarize_export_log_compatibility(
                    payload_for_review,
                    current_keys,
                    current_bom_df=getattr(parent_app, "bom_df", None),
                )
            except Exception:
                compatibility = {
                    "incoming_keys": sorted(self._state_selection_keys(state_dict)),
                    "matched_keys": [],
                    "missing_keys": [],
                    "new_keys": [],
                    "bom_changed": False,
                }
            review_message = format_export_log_compatibility_message(compatibility)
            if review_message:
                proceed = messagebox.askyesno(
                    "Exportlog controle",
                    (
                        f"{review_message}\n\n"
                        "Wil je de gevonden waarden toch toepassen op de huidige bestelbonpagina?"
                    ),
                    parent=self,
                )
                if not proceed:
                    return

            selected_sections = self._prompt_export_log_import_sections(
                converted=converted
            )
            if selected_sections is None:
                self.update_status("Exportlog laden geannuleerd.")
                return

            current_state = self.serialize_state()
            state_to_apply = merge_order_state_sections(
                current_state,
                state_dict,
                selected_sections,
            )
            selected_state_keys = state_keys_for_import_sections(selected_sections)
            selected_incoming_state = {
                key: state_dict.get(key, {})
                for key in selected_state_keys
                if key in state_dict
            }
            incoming_keys = self._state_selection_keys(selected_incoming_state)
            if not incoming_keys:
                incoming_keys = (
                    set(compatibility.get("incoming_keys", []))
                    or self._state_selection_keys(state_dict)
                )
            applied_keys = incoming_keys & current_keys
            matched = len(applied_keys)
            missing = len(incoming_keys - current_keys)
            try:
                self.apply_state(SupplierSelectionState.from_mapping(state_to_apply))
            except Exception as exc:
                messagebox.showerror(
                    "Exportlog laden",
                    f"Kon exportlog niet toepassen:\n{exc}",
                    parent=self,
                )
                return

            source_label = os.path.basename(os.path.dirname(path)) or os.path.basename(path)
            self._store_export_log_state(
                applied_keys,
                source_label=source_label,
                converted=converted,
            )
            self._loaded_export_log_path = path
            self._loaded_export_log_export_info = dict(payload.get("export", {}) or {})
            docs_button = getattr(self, "export_log_documents_button", None)
            if docs_button is not None:
                try:
                    docs_button.configure(state="normal")
                except tk.TclError:
                    pass
            self._update_preview_from_any_combo()
            sync_grouped_rows = getattr(self, "_sync_grouped_rows", None)
            if callable(sync_grouped_rows):
                sync_grouped_rows()
            if parent_app is not None:
                try:
                    parent_app._last_supplier_selection_state = self.serialize_state()
                except Exception:
                    pass
            project = payload.get("project", {}) if isinstance(payload, Mapping) else {}
            project_label = ""
            if isinstance(project, Mapping):
                project_number = _to_str(project.get("number")).strip()
                project_name = _to_str(project.get("name")).strip()
                project_label = " - ".join(part for part in (project_number, project_name) if part)
            suffix = " Offertes omgezet naar bestelbonnen." if converted else ""
            self.update_status(
                f"Exportlog geladen{f' voor {project_label}' if project_label else ''}: "
                f"{matched} regel(s) toegepast, {missing} niet gevonden.{suffix} "
                f"Bron: {source_label}"
            )

        def serialize_state(self) -> "SupplierSelectionState":
            selections: Dict[str, str] = {}
            for sel_key, combo in self.rows:
                try:
                    selections[sel_key] = combo.get()
                except tk.TclError:
                    continue

            exports = {key: bool(var.get()) for key, var in self.export_vars.items()}
            current_group_links = getattr(self, "_current_group_links", None)
            groups = current_group_links() if callable(current_group_links) else {}
            doc_types = {key: var.get() for key, var in self.doc_vars.items()}
            doc_numbers = {
                key: var.get() for key, var in self.doc_num_vars.items()
            }
            remarks = {key: var.get() for key, var in self.remark_vars.items()}
            deliveries = {
                key: var.get() for key, var in self.delivery_vars.items()
            }
            en1090 = {key: bool(var.get()) for key, var in self.en1090_vars.items()}
            vat_rates = {
                key: _clean_supplier_vat_rate(var.get())
                for key, var in self.vat_vars.items()
            }
            pricing: Dict[str, Dict[str, object]] = {}
            all_price_keys = (
                set(self.price_unit_vars)
                | set(self.price_total_vars)
                | set(self.line_pricing)
            )
            for key in all_price_keys:
                unit_price = self.price_unit_vars.get(key)
                total_price = self.price_total_vars.get(key)
                entry: Dict[str, object] = {
                    "unit_price": unit_price.get().strip() if unit_price else "",
                    "total_price": total_price.get().strip() if total_price else "",
                    "quote_ref": "",
                    "note": "",
                }
                line_items = self.line_pricing.get(key, {})
                if line_items:
                    entry["items"] = deepcopy(line_items)
                if _clean_supplier_pricing_value(entry):
                    pricing[key] = entry

            remember = bool(self.remember_var.get()) if hasattr(self, "remember_var") else True

            return SupplierSelectionState(
                selections=selections,
                groups=groups,
                exports=exports,
                doc_types=doc_types,
                doc_numbers=doc_numbers,
                remarks=remarks,
                deliveries=deliveries,
                en1090=en1090,
                vat_rates=vat_rates,
                pricing=pricing,
                remember=remember,
            )

        def apply_state(self, state: "SupplierSelectionState") -> None:
            if SupplierSelectionFrame._state_has_price_values(state):
                set_price_columns_visible = getattr(
                    self, "set_price_columns_visible", None
                )
                if callable(set_price_columns_visible):
                    set_price_columns_visible(True)

            set_combo_value = getattr(
                type(self), "_set_combo_value", SupplierSelectionFrame._set_combo_value
            )
            for sel_key, combo in self.rows:
                if sel_key in state.selections:
                    set_combo_value(combo, state.selections[sel_key])

            for sel_key, enabled in state.exports.items():
                var = self.export_vars.get(sel_key)
                if var is not None:
                    try:
                        var.set(1 if enabled else 0)
                    except tk.TclError:
                        pass

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
                        dcombo.set(self._find_delivery_display_for_name(value))
                    except tk.TclError:
                        pass

            for sel_key, enabled in state.en1090.items():
                var = self.en1090_vars.get(sel_key)
                if var is not None:
                    try:
                        var.set(1 if enabled else 0)
                    except tk.TclError:
                        pass

            for sel_key, value in getattr(state, "vat_rates", {}).items():
                var = self.vat_vars.get(sel_key)
                if var is not None:
                    try:
                        var.set(_clean_supplier_vat_rate(value))
                    except tk.TclError:
                        pass

            for sel_key, price_info in getattr(state, "pricing", {}).items():
                if not isinstance(price_info, Mapping):
                    continue
                unit_var = self.price_unit_vars.get(sel_key)
                total_var = self.price_total_vars.get(sel_key)
                self._price_link_in_progress.add(sel_key)
                try:
                    if unit_var is not None:
                        unit_var.set(_to_str(price_info.get("unit_price")).strip())
                    if total_var is not None:
                        total_var.set(_to_str(price_info.get("total_price")).strip())
                finally:
                    self._price_link_in_progress.discard(sel_key)
                self._price_auto_fields.pop(sel_key, None)
                if unit_var is not None and unit_var.get().strip():
                    self._on_price_field_change(sel_key, "unit", force=False)
                elif total_var is not None and total_var.get().strip():
                    self._on_price_field_change(sel_key, "total", force=False)
                items = _clean_supplier_pricing_value(price_info).get("items")
                if isinstance(items, Mapping) and items:
                    self.line_pricing[sel_key] = {
                        _to_str(item_key): dict(item_value)
                        for item_key, item_value in items.items()
                        if _to_str(item_key).strip()
                        and isinstance(item_value, Mapping)
                    }
                else:
                    self.line_pricing.pop(sel_key, None)
                self._refresh_line_price_button(sel_key)

            if hasattr(self, "remember_var"):
                try:
                    self.remember_var.set(1 if state.remember else 0)
                except tk.TclError:
                    pass
            refresh_group_options = getattr(self, "_refresh_group_options", None)
            if callable(refresh_group_options):
                refresh_group_options(getattr(state, "groups", {}))

        def _compute_initial_en1090_width(self) -> int:
            width = self.EN1090_COLUMN_WIDTH
            font = getattr(self, "_en1090_header_font", None)
            if font is not None:
                try:
                    text_width = font.measure(self.EN1090_HEADER_TEXT)
                except tk.TclError:
                    text_width = 0
                if text_width:
                    width = min(
                        width,
                        text_width + self.EN1090_COLUMN_PADDING,
                    )
            return max(self.EN1090_MIN_COLUMN_WIDTH, int(width))

        def _compute_en1090_header_char_width(self) -> int:
            text = self.EN1090_HEADER_TEXT
            font = getattr(self, "_en1090_header_font", None)
            if font is None:
                return len(text)
            try:
                zero_width = font.measure("0")
            except tk.TclError:
                zero_width = 0
            if not zero_width:
                return len(text)
            return max(len(text), math.ceil(self._en1090_column_width_px / zero_width))

        def _refresh_en1090_header_width(self) -> None:
            label = self._header_labels_map.get("en1090_widget")
            if label is None:
                return
            try:
                label.configure(width=self._compute_en1090_header_char_width())
            except tk.TclError:
                pass

        def _set_en1090_column_width(self, width: int) -> None:
            width = int(max(width, 0))
            if width == self._en1090_column_width_px:
                for frame in self._en1090_frames:
                    try:
                        frame.configure(width=width)
                    except tk.TclError:
                        continue
                return
            self._en1090_column_width_px = width
            for frame in self._en1090_frames:
                try:
                    frame.configure(width=width)
                except tk.TclError:
                    continue
            self._refresh_en1090_header_width()
            self._header_aligned = False
            self._header_alignment_pending = False
            if self._row_widget_maps:
                self._schedule_header_alignment(self._row_widget_maps[0])

        def _ensure_en1090_width(self, min_width: int) -> None:
            if min_width > self._en1090_column_width_px:
                self._set_en1090_column_width(min_width)
            else:
                self._set_en1090_column_width(self._en1090_column_width_px)

        def _schedule_header_alignment(self, row_widgets: Dict[str, "tk.Misc"]) -> None:
            if not getattr(self, "_header_column_frames_map", None):
                return
            if self._header_aligned or self._header_alignment_pending:
                return

            def _do_align() -> None:
                self._align_header_columns(row_widgets)

            self._header_alignment_pending = True
            self.after_idle(_do_align)

        def _on_en1090_toggle(self, sel_key: str) -> None:
            setter = getattr(self, "_en1090_setter", None)
            if setter is None:
                return
            metadata = self.row_meta.get(sel_key, {})
            if metadata.get("kind") not in {"production", "opticutter"}:
                return
            var = self.en1090_vars.get(sel_key)
            if var is None:
                return
            try:
                setter(metadata, bool(var.get()))
            except Exception:
                pass
            sync_grouped_rows = getattr(self, "_sync_grouped_rows", None)
            if callable(sync_grouped_rows):
                sync_grouped_rows()

        def _align_header_columns(self, row_widgets: Dict[str, "tk.Misc"]) -> None:
            try:
                column_widgets = [row_widgets[key] for key in self._visible_column_keys]
            except KeyError:
                self._header_alignment_pending = False
                return

            self.update_idletasks()

            header_frames = [
                self._header_column_frames_map[key] for key in self._visible_column_keys
            ]
            header_labels = [
                self._header_labels_map[key] for key in self._visible_column_keys
            ]

            for frame, label, widget in zip(header_frames, header_labels, column_widgets):
                frame.pack_propagate(False)
                width = widget.winfo_width() or widget.winfo_reqwidth()
                height = label.winfo_reqheight() or widget.winfo_reqheight()
                frame.configure(width=width, height=height)
                label.configure(anchor="w", width=0)
                label.pack_configure(fill="x")

            self._header_aligned = True
            self._header_alignment_pending = False

        def _repack_header_columns(self) -> None:
            for frame in self._header_column_frames_map.values():
                frame.pack_forget()
            for key in self._visible_column_keys:
                frame = self._header_column_frames_map.get(key)
                if frame is not None:
                    frame.pack(side="left", padx=(0, 6))

        def _refresh_visible_column_keys(self) -> None:
            show_prices = bool(getattr(self, "_price_columns_visible", False))
            self._visible_column_keys = [
                key
                for key in self._column_keys
                if (key != "en1090_widget" or self._en1090_enabled)
                and (show_prices or key not in self.PRICE_COLUMN_KEYS)
            ]

        def _pack_row_widgets(self, widgets: Dict[str, "tk.Misc"]) -> None:
            for key in self._column_keys:
                widget = widgets.get(key)
                if widget is None:
                    continue
                try:
                    widget.pack_forget()
                except Exception:
                    pass
            for key in self._visible_column_keys:
                widget = widgets.get(key)
                if widget is None:
                    continue
                widget.pack(side="left", padx=(0, 6))

        def _repack_all_rows(self) -> None:
            for widgets in self._row_widget_maps:
                self._pack_row_widgets(widgets)

        def _repack_visible_columns(self) -> None:
            self._refresh_visible_column_keys()
            self._repack_header_columns()
            self._repack_all_rows()
            self._header_aligned = False
            self._header_alignment_pending = False
            if self._row_widget_maps:
                self._schedule_header_alignment(self._row_widget_maps[0])

        def set_price_columns_visible(self, visible: bool) -> None:
            self._price_columns_visible = bool(visible)
            var = getattr(self, "price_columns_visible_var", None)
            if var is not None:
                try:
                    if bool(var.get()) != self._price_columns_visible:
                        var.set(1 if self._price_columns_visible else 0)
                except Exception:
                    pass
            self._repack_visible_columns()

        def set_en1090_enabled(self, enabled: bool) -> None:
            self._en1090_enabled = bool(enabled)
            self._repack_visible_columns()

        def _clear_saved_suppliers(self) -> None:
            self.db.defaults_by_production.clear()
            self.db.defaults_by_finish.clear()
            self.db.save()
            preset_state_by_key = getattr(self, "_preset_state_by_key", None)
            if isinstance(preset_state_by_key, dict):
                preset_state_by_key.clear()
            export_log_state_by_key = getattr(self, "_export_log_state_by_key", None)
            if isinstance(export_log_state_by_key, dict):
                export_log_state_by_key.clear()
            self._loaded_export_log_path = ""
            self._loaded_export_log_export_info = {}
            docs_button = getattr(self, "export_log_documents_button", None)
            if docs_button is not None:
                try:
                    docs_button.configure(state="disabled")
                except tk.TclError:
                    pass

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

            refresh_group_options = getattr(self, "_refresh_group_options", None)
            if callable(refresh_group_options):
                refresh_group_options({})
            self._on_combo_change()

        def _on_focus_key(self, sel_key: str):
            self._active_key = sel_key
            SupplierSelectionFrame._get_type_filter_map(self).pop(sel_key, None)
            messages = [
                message
                for message in (
                    self._preset_status_message(sel_key),
                    self._export_log_status_message(sel_key),
                )
                if message
            ]
            if messages:
                self.update_status(" ".join(messages))

        def _set_all_exports(self, enabled: bool) -> None:
            value = 1 if enabled else 0
            for var in self.export_vars.values():
                try:
                    var.set(value)
                except tk.TclError:
                    continue
            sync_grouped_rows = getattr(self, "_sync_grouped_rows", None)
            if callable(sync_grouped_rows):
                sync_grouped_rows()

        @staticmethod
        def _get_type_filter_map(instance) -> Dict[str, str]:
            type_map = getattr(instance, "_type_filter_by_key", None)
            if type_map is None:
                type_map = {}
                setattr(instance, "_type_filter_by_key", type_map)
            return type_map

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

        @staticmethod
        def _is_groupable_kind(kind: str) -> bool:
            return kind in {"production", "finish"}

        @staticmethod
        def _group_code_from_index(index):
            try:
                remaining = int(index)
            except (TypeError, ValueError):
                remaining = 0
            remaining = max(remaining, 0)
            letters = ""
            while True:
                remaining, offset = divmod(remaining, 26)
                letters = chr(ord("A") + offset) + letters
                if remaining == 0:
                    return letters
                remaining -= 1

        def _base_row_label(self, sel_key: str) -> str:
            metadata = self.row_meta.get(sel_key, {})
            label = _to_str(
                metadata.get("base_display")
                or metadata.get("display")
                or metadata.get("identifier")
            ).strip()
            return label or sel_key

        def _group_root_code_map(self, group_links) -> Dict[str, str]:
            root_codes: Dict[str, str] = {}
            seen_roots: set[str] = set()
            counters_by_kind: Dict[str, int] = {}
            for row_key, _combo in self.rows:
                kind, _identifier = self._parse_selection_key(row_key)
                if not self._is_groupable_kind(kind):
                    continue
                root_key = self._resolve_group_root(row_key, group_links)
                if root_key in seen_roots:
                    continue
                seen_roots.add(root_key)
                index = counters_by_kind.get(kind, 0)
                counters_by_kind[kind] = index + 1
                root_codes[root_key] = self._group_code_from_index(index)
            return root_codes

        def _group_root_color_map(self, group_links) -> Dict[str, str]:
            color_map: Dict[str, str] = {}
            root_codes = self._group_root_code_map(group_links)
            roots_by_kind: Dict[str, List[str]] = {"production": [], "finish": []}
            for root_key in root_codes:
                kind, _identifier = self._parse_selection_key(root_key)
                if kind not in roots_by_kind:
                    roots_by_kind[kind] = []
                roots_by_kind[kind].append(root_key)

            for roots in roots_by_kind.values():
                for index, root_key in enumerate(roots):
                    color_map[root_key] = self.GROUP_ACCENT_COLORS[
                        index % len(self.GROUP_ACCENT_COLORS)
                    ]
            return color_map

        def _group_followers_by_root(self, group_links) -> Dict[str, List[str]]:
            followers: Dict[str, List[str]] = {}
            for row_key, _combo in self.rows:
                kind, _identifier = self._parse_selection_key(row_key)
                if not self._is_groupable_kind(kind):
                    continue
                root_key = self._resolve_group_root(row_key, group_links)
                if root_key == row_key:
                    followers.setdefault(root_key, [])
                    continue
                followers.setdefault(root_key, []).append(row_key)
            return followers

        def _group_visual_spec(self, sel_key: str, group_links):
            base_label = self._base_row_label(sel_key)
            kind, _identifier = self._parse_selection_key(sel_key)
            if not self._is_groupable_kind(kind):
                return {
                    "text": base_label,
                    "accent": "",
                    "grouped": False,
                    "is_root": False,
                }

            root_key = self._resolve_group_root(sel_key, group_links)
            followers_by_root = self._group_followers_by_root(group_links)
            root_codes = self._group_root_code_map(group_links)
            color_map = self._group_root_color_map(group_links)
            followers = followers_by_root.get(root_key, [])
            has_group = bool(followers)
            if root_key == sel_key and not has_group:
                return {
                    "text": base_label,
                    "accent": "",
                    "grouped": False,
                    "is_root": True,
                }

            code = root_codes.get(root_key, "")
            accent = color_map.get(root_key, "")
            if root_key == sel_key:
                suffix = f"[Bon {code}]" if code else "[Bon]"
                text = f"{base_label} {suffix}"
                return {
                    "text": text,
                    "accent": accent,
                    "grouped": True,
                    "is_root": True,
                }

            suffix = f"[Volgt {code}]" if code else "[Volgt]"
            text = f"{base_label} {suffix}"
            return {
                "text": text,
                "accent": accent,
                "grouped": True,
                "is_root": False,
            }

        def _apply_group_visuals(self, group_links) -> None:
            base_bg = getattr(self, "_rows_background", None) or self.cget("bg")
            preset_fg = "#2E7D32"
            export_log_fg = "#1565C0"
            preset_state_by_key = getattr(self, "_preset_state_by_key", {})
            export_log_state_by_key = getattr(self, "_export_log_state_by_key", {})
            for sel_key, _combo in self.rows:
                widgets = self._row_widgets_by_key.get(sel_key, {})
                row = widgets.get("row")
                row_bg = widgets.get("row_background", base_bg)
                row_label = widgets.get("label")
                group_stripe = widgets.get("group_stripe")
                default_fg = widgets.get("label_default_fg", "")
                default_font = widgets.get("label_default_font", "")
                if row_label is None:
                    continue
                if row is not None:
                    try:
                        row.configure(background=row_bg)
                    except tk.TclError:
                        pass

                spec = self._group_visual_spec(sel_key, group_links)
                accent = _to_str(spec.get("accent")).strip()
                grouped = bool(spec.get("grouped"))
                is_root = bool(spec.get("is_root"))
                base_text = _to_str(spec.get("text")).strip() or self._base_row_label(sel_key)
                label_text = (
                    f"{base_text}"
                    f"{self._preset_indicator_suffix(sel_key)}"
                    f"{self._export_log_indicator_suffix(sel_key)}"
                )
                label_fg = accent if accent else (
                    preset_fg if sel_key in preset_state_by_key else (
                        export_log_fg if sel_key in export_log_state_by_key else default_fg
                    )
                )
                try:
                    row_label.configure(
                        text=label_text,
                        fg=label_fg,
                        font=("TkDefaultFont", 10, "bold") if grouped and is_root else default_font,
                        background=row_bg,
                    )
                except tk.TclError:
                    pass

                if group_stripe is not None:
                    try:
                        group_stripe.configure(background=accent if accent else row_bg)
                    except tk.TclError:
                        pass

        @staticmethod
        def _resolve_group_root(sel_key: str, group_map: Dict[str, str]) -> str:
            current = sel_key
            seen = {sel_key}
            while True:
                parent = _to_str(group_map.get(current)).strip()
                if not parent or parent == current:
                    return current
                if parent in seen:
                    return sel_key
                seen.add(parent)
                current = parent

        def _group_row_label(self, sel_key: str, group_links=None) -> str:
            base_label = self._base_row_label(sel_key)
            links = (
                group_links
                if group_links is not None
                else self._sanitize_group_links(self._current_group_links_raw())
            )
            code = self._group_root_code_map(links).get(sel_key, "")
            if code:
                return f"Bon {code} - {base_label}"
            return base_label

        def _current_group_links_raw(self) -> Dict[str, str]:
            raw_links: Dict[str, str] = {}
            for sel_key, var in self.group_vars.items():
                display_value = var.get().strip()
                actual_value = self.group_display_to_value.get(sel_key, {}).get(
                    display_value, ""
                )
                if actual_value:
                    raw_links[sel_key] = actual_value
            return raw_links

        def _sanitize_group_links(
            self, raw_group_links: Optional[Dict[str, str]] = None
        ) -> Dict[str, str]:
            raw_links = dict(raw_group_links or {})
            sanitized: Dict[str, str] = {}
            seen_by_kind: Dict[str, List[str]] = {"production": [], "finish": []}

            for sel_key, _combo in self.rows:
                kind, _identifier = self._parse_selection_key(sel_key)
                if not self._is_groupable_kind(kind):
                    continue
                allowed_masters = seen_by_kind[kind]
                raw_master = _to_str(raw_links.get(sel_key)).strip()
                if raw_master and raw_master in allowed_masters:
                    root = self._resolve_group_root(raw_master, sanitized)
                    if root in allowed_masters:
                        sanitized[sel_key] = root
                seen_by_kind[kind].append(sel_key)
            return sanitized

        def _available_group_roots(
            self, sel_key: str, group_links: Dict[str, str]
        ) -> List[str]:
            kind, _identifier = self._parse_selection_key(sel_key)
            if not self._is_groupable_kind(kind):
                return []

            roots: List[str] = []
            seen_roots: set[str] = set()
            for row_key, _combo in self.rows:
                if row_key == sel_key:
                    break
                row_kind, _row_identifier = self._parse_selection_key(row_key)
                if row_kind != kind or not self._is_groupable_kind(row_kind):
                    continue
                root = self._resolve_group_root(row_key, group_links)
                if root == sel_key or root in seen_roots:
                    continue
                seen_roots.add(root)
                roots.append(root)
            return roots

        def _set_row_grouped_state(self, sel_key: str, grouped: bool) -> None:
            widgets = self._row_widgets_by_key.get(sel_key, {})
            supplier_combo = widgets.get("supplier_combo")
            if supplier_combo is not None:
                supplier_combo.configure(state="disabled" if grouped else "normal")

            doc_combo = widgets.get("doc_combo")
            if doc_combo is not None:
                doc_combo.configure(state="disabled" if grouped else "readonly")

            doc_entry = widgets.get("doc_entry")
            if doc_entry is not None:
                doc_entry.configure(state="disabled" if grouped else "normal")

            unit_price_entry = widgets.get("unit_price_entry")
            if unit_price_entry is not None:
                unit_price_entry.configure(state="disabled" if grouped else "normal")

            total_price_entry = widgets.get("total_price_entry")
            if total_price_entry is not None:
                total_price_entry.configure(state="disabled" if grouped else "normal")

            vat_combo = widgets.get("vat_combo")
            if vat_combo is not None:
                vat_combo.configure(state="disabled" if grouped else "normal")

            line_price_button = widgets.get("line_price_button")
            if line_price_button is not None:
                has_items = bool(self._selection_items_for_key(sel_key))
                state = "normal" if has_items and not grouped else "disabled"
                line_price_button.configure(state=state)

            remark_entry = widgets.get("remark_entry")
            if remark_entry is not None:
                remark_entry.configure(state="disabled" if grouped else "normal")

            delivery_combo = widgets.get("delivery_combo")
            if delivery_combo is not None:
                delivery_combo.configure(state="disabled" if grouped else "readonly")

            export_check = widgets.get("export_check")
            if export_check is not None:
                export_check.configure(state="disabled" if grouped else "normal")

            en1090_widget = widgets.get("en1090_widget")
            if en1090_widget is not None:
                for child in en1090_widget.winfo_children():
                    try:
                        child.configure(state="disabled" if grouped else "normal")
                    except tk.TclError:
                        continue

        def _copy_group_master_values(self, follower_key: str, master_key: str) -> None:
            follower_combo = self.combo_by_key.get(follower_key)
            master_combo = self.combo_by_key.get(master_key)
            if follower_combo is not None and master_combo is not None:
                self._set_combo_value(follower_combo, master_combo.get())

            self._price_link_in_progress.add(follower_key)
            try:
                for source_map in (
                    self.doc_vars,
                    self.doc_num_vars,
                    self.price_unit_vars,
                    self.price_total_vars,
                    self.vat_vars,
                    self.remark_vars,
                ):
                    follower_var = source_map.get(follower_key)
                    master_var = source_map.get(master_key)
                    if follower_var is not None and master_var is not None:
                        follower_var.set(master_var.get())
            finally:
                self._price_link_in_progress.discard(follower_key)

            master_line_pricing = self.line_pricing.get(master_key)
            if master_line_pricing:
                self.line_pricing[follower_key] = deepcopy(master_line_pricing)
            else:
                self.line_pricing.pop(follower_key, None)
            self._refresh_line_price_button(follower_key)

            follower_delivery = self.delivery_combos.get(follower_key)
            master_delivery = self.delivery_combos.get(master_key)
            if follower_delivery is not None and master_delivery is not None:
                follower_delivery.set(master_delivery.get())

            follower_export = self.export_vars.get(follower_key)
            master_export = self.export_vars.get(master_key)
            if follower_export is not None and master_export is not None:
                follower_export.set(master_export.get())

            follower_en1090 = self.en1090_vars.get(follower_key)
            master_en1090 = self.en1090_vars.get(master_key)
            if follower_en1090 is not None and master_en1090 is not None:
                follower_en1090.set(master_en1090.get())

        def _current_group_links(self) -> Dict[str, str]:
            return self._sanitize_group_links(self._current_group_links_raw())

        def _refresh_group_options(
            self, desired_group_links: Optional[Dict[str, str]] = None
        ) -> None:
            group_links = self._sanitize_group_links(
                desired_group_links
                if desired_group_links is not None
                else self._current_group_links_raw()
            )

            for sel_key, _combo in self.rows:
                group_combo = self.group_combos.get(sel_key)
                group_var = self.group_vars.get(sel_key)
                if group_combo is None or group_var is None:
                    continue

                kind, _identifier = self._parse_selection_key(sel_key)
                value_to_display = {"": self.GROUP_APART_LABEL}
                if self._is_groupable_kind(kind):
                    for master_key in self._available_group_roots(sel_key, group_links):
                        value_to_display[master_key] = self._group_row_label(
                            master_key, group_links
                        )
                    group_combo.configure(
                        state="readonly" if len(value_to_display) > 1 else "disabled"
                    )
                else:
                    group_combo.configure(state="disabled")

                display_to_value = {
                    display: value for value, display in value_to_display.items()
                }
                self.group_value_to_display[sel_key] = value_to_display
                self.group_display_to_value[sel_key] = display_to_value
                group_combo["values"] = list(value_to_display.values())
                group_var.set(
                    value_to_display.get(
                        group_links.get(sel_key, ""),
                        self.GROUP_APART_LABEL,
                    )
                )

            self._sync_grouped_rows(group_links)

        def _on_group_change(self, sel_key: str) -> None:
            from tkinter import messagebox

            desired_links = self._sanitize_group_links(self._current_group_links_raw())
            master_key = desired_links.get(sel_key, "")
            kind, _identifier = self._parse_selection_key(sel_key)
            if kind == "production" and master_key:
                follower_var = self.en1090_vars.get(sel_key)
                master_var = self.en1090_vars.get(master_key)
                follower_en1090 = bool(follower_var.get()) if follower_var is not None else False
                master_en1090 = bool(master_var.get()) if master_var is not None else False
                if follower_en1090 != master_en1090:
                    desired_links.pop(sel_key, None)
                    self._refresh_group_options(desired_links)
                    messagebox.showwarning(
                        "EN 1090 ongeldig",
                        "Gekoppelde producties moeten dezelfde EN 1090-instelling hebben.",
                        parent=self,
                    )
                    return
            self._refresh_group_options(desired_links)

        def _sync_grouped_rows(
            self, group_links: Optional[Dict[str, str]] = None
        ) -> None:
            if self._group_sync_in_progress:
                return

            sanitized_links = self._sanitize_group_links(
                group_links if group_links is not None else self._current_group_links_raw()
            )
            self._group_sync_in_progress = True
            try:
                for sel_key, _combo in self.rows:
                    master_key = sanitized_links.get(sel_key, "")
                    self._set_row_grouped_state(sel_key, bool(master_key))
                    if master_key:
                        self._copy_group_master_values(sel_key, master_key)
                self._apply_group_visuals(sanitized_links)
            finally:
                self._group_sync_in_progress = False

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
                disp = None
                for k, v in self._disp_to_name.items():
                    if v and name and v.lower() == name.lower():
                        disp = k
                        break
                if disp:
                    set_combo_value(combo, disp)
                elif self._base_options:
                    set_combo_value(combo, self._base_options[0])

            delivery_opts = self._delivery_options()
            default_delivery = self._default_delivery_value()
            for sel_key, dcombo in self.delivery_combos.items():
                cur = dcombo.get()
                dcombo["values"] = delivery_opts
                if cur in delivery_opts:
                    dcombo.set(cur)
                elif SupplierSelectionFrame._is_client_delivery_choice(cur):
                    dcombo.set(self.CLIENT_DELIVERY_PRESET)
                else:
                    dcombo.set(default_delivery)
            refresh_group_options = getattr(self, "_refresh_group_options", None)
            if callable(refresh_group_options):
                refresh_group_options()
            if initial:
                apply_presets = getattr(self, "_apply_presets", None)
                if callable(apply_presets):
                    apply_presets(auto_apply_only=True, status_when_idle=False)

        def _on_combo_change(self, _evt=None):
            for sel_key, combo in self.rows:
                SupplierSelectionFrame._get_type_filter_map(self).pop(sel_key, None)
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
            sync_grouped_rows = getattr(self, "_sync_grouped_rows", None)
            if callable(sync_grouped_rows):
                sync_grouped_rows()

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
            sync_grouped_rows = getattr(self, "_sync_grouped_rows", None)
            if callable(sync_grouped_rows):
                sync_grouped_rows()

        def _on_combo_type(self, evt, sel_key: str, combo):
            self._active_key = sel_key
            base_options = getattr(self, "_base_options", None)
            if not base_options:
                return

            keysym = getattr(evt, "keysym", "")
            if keysym in ("Up", "Down", "Prior", "Next", "Tab"):
                return

            type_map = SupplierSelectionFrame._get_type_filter_map(self)
            text_so_far = type_map.get(sel_key, "")

            if keysym == "Escape":
                type_map[sel_key] = ""
                combo["values"] = base_options
                self._populate_cards([], sel_key)
                self._update_preview_for_text("")
                self._set_supplier_details_visible(False)
                return

            if keysym == "BackSpace":
                text_so_far = text_so_far[:-1]
            elif keysym == "Delete":
                text_so_far = ""
            elif keysym in ("Return", "KP_Enter"):
                pass
            else:
                char = getattr(evt, "char", "")
                if not char:
                    # ``event.char`` is sometimes empty on newer Tk versions. Fall back
                    # to the keysym so that plain letter keys still count as typed
                    # characters, e.g. when Tk returns ``keysym="A"`` and
                    # ``char=""``.
                    if len(keysym) == 1:
                        char = keysym
                    elif keysym == "space":
                        char = " "

                state = getattr(evt, "state", 0)

                if not char or state & 0x4 or state & 0x8:
                    # If we don't have a reliable character (e.g. Control+key or when
                    # Tk skips ``event.char`` altogether) we fall back to whatever text
                    # Tk already placed in the widget so the search keeps working.
                    try:
                        text_so_far = combo.get()
                    except tk.TclError:
                        text_so_far = ""
                else:
                    try:
                        if hasattr(combo, "selection_present") and combo.selection_present():
                            text_so_far = ""
                    except tk.TclError:
                        pass
                    text_so_far += char

            type_map[sel_key] = text_so_far
            self._show_supplier_details_for_supplier_search()

            if keysym not in ("Return", "KP_Enter"):
                try:
                    combo.delete(0, tk.END)
                    combo.insert(0, text_so_far)
                    if hasattr(combo, "selection_clear"):
                        combo.selection_clear()
                    combo.icursor(tk.END)
                except tk.TclError:
                    pass

            norm_text = _norm(text_so_far.strip())
            if not norm_text:
                combo["values"] = base_options
                self._populate_cards([], sel_key)
                self._update_preview_for_text("")
                if bool(getattr(self, "_supplier_details_auto_open", False)):
                    self._set_supplier_details_visible(False)
                return

            disp_to_name = getattr(self, "_disp_to_name", {})

            def _option_norm(opt: str) -> str:
                name = disp_to_name.get(opt, opt)
                cleaned = strip_favorite_marker(name if name else opt)
                cleaned = cleaned.replace("(", "").replace(")", "")
                return _norm(cleaned)

            filtered = [
                opt for opt in base_options if _option_norm(opt).startswith(norm_text)
            ]
            if not filtered:
                filtered = [
                    opt for opt in base_options if norm_text in _option_norm(opt)
                ]

            filtered = sort_supplier_options(
                filtered, self.db.suppliers, disp_to_name
            )
            combo["values"] = filtered
            self._populate_cards(filtered, sel_key)

            if keysym in ("Return", "KP_Enter") and len(filtered) == 1:
                set_combo_value = getattr(
                    type(self),
                    "_set_combo_value",
                    SupplierSelectionFrame._set_combo_value,
                )
                set_combo_value(combo, filtered[0])
                type_map.pop(sel_key, None)
                self._update_preview_for_text(filtered[0])
            else:
                preview_text = filtered[0] if filtered else text_so_far
                self._update_preview_for_text(preview_text)

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
                def handler(_e, o=opt, key=sel_key):
                    self._on_card_click(o, key)

                card.bind("<Button-1>", handler)
                for w in widgets:
                    w.bind("<Button-1>", handler)
            bind_mousewheel_tree = getattr(
                self, "_bind_selection_rows_mousewheel_tree", None
            )
            if callable(bind_mousewheel_tree):
                bind_mousewheel_tree(self.cards_frame)

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
                app = getattr(self.master, "master", None)
                if hasattr(self.master, "select") and app is not None:
                    target = (
                        getattr(app, "pdf_workdossier_frame", None)
                        if self.pdf_dossier_context
                        else getattr(app, "main_frame", None)
                    )
                    if target is not None:
                        self.master.select(target)
                    if self.pdf_dossier_context:
                        options_frame = getattr(
                            app, "pdf_workdossier_options_frame", None
                        )
                        if options_frame is not None:
                            try:
                                options_frame.select_work_tab()
                            except tk.TclError:
                                pass
                        refresh = getattr(app, "_refresh_pdf_workdossier_preview", None)
                        if callable(refresh):
                            refresh()
                if app is not None and hasattr(app, "sel_frame"):
                    app.sel_frame = None
            self.destroy()

        def _confirm(self):
            """Collect selected suppliers per production and return via callback."""
            import inspect

            sel_map: Dict[str, str] = {}
            export_map: Dict[str, bool] = {}
            doc_map: Dict[str, str] = {}
            current_group_links = getattr(self, "_current_group_links", None)
            group_map = current_group_links() if callable(current_group_links) else {}
            export_vars = getattr(self, "export_vars", {}) or {}
            for sel_key, combo in self.rows:
                typed = combo.get().strip()
                if not typed or typed.lower() in ("(geen)", "geen"):
                    sel_map[sel_key] = ""
                else:
                    s = self._resolve_text_to_supplier(typed)
                    if s:
                        sel_map[sel_key] = s.supplier
                export_var = export_vars.get(sel_key)
                export_map[sel_key] = bool(export_var.get()) if export_var else True
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

            en1090_vars = getattr(self, "en1090_vars", {}) or {}
            en1090_map = {key: bool(var.get()) for key, var in en1090_vars.items()}

            for follower_key, master_key in group_map.items():
                sel_map[follower_key] = sel_map.get(master_key, "")
                export_map[follower_key] = export_map.get(master_key, True)
                doc_map[follower_key] = doc_map.get(master_key, "Bestelbon")
                doc_num_map[follower_key] = doc_num_map.get(master_key, "")
                default_delivery_value = getattr(
                    self, "_default_delivery_value", lambda: "Geen"
                )
                delivery_map[follower_key] = delivery_map.get(
                    master_key, default_delivery_value()
                )
                remarks_map[follower_key] = remarks_map.get(master_key, "")
                en1090_map[follower_key] = en1090_map.get(master_key, False)

            project_number = self.project_number_var.get().strip()
            project_name = self.project_name_var.get().strip()

            remember_flag = bool(self.remember_var.get())
            callback = self.callback
            use_new_signature = False
            supports_exports = False
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
                    supports_exports = True
                elif len(params) >= 11:
                    use_new_signature = True
                    supports_exports = True
                elif len(params) >= 10:
                    use_new_signature = True

            if use_new_signature:
                base_args = [
                    sel_map,
                    doc_map,
                    doc_num_map,
                    delivery_map,
                    remarks_map,
                    group_map,
                    en1090_map,
                    project_number,
                    project_name,
                    remember_flag,
                ]
                attempt_exports = supports_exports
                while True:
                    try:
                        if attempt_exports:
                            callback(*base_args, export_map)
                        else:
                            callback(*base_args)
                        return
                    except TypeError as exc:
                        msg = str(exc)
                        if attempt_exports:
                            attempt_exports = False
                            continue
                        if not (
                            "positional" in msg
                            or "keyword" in msg
                            or (sig_params is not None and len(sig_params) >= 8)
                        ):
                            raise
                        use_new_signature = False
                        break

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
            self._settings_bg = "#F5F7FA"
            self._settings_card_bg = "#FFFFFF"
            self._settings_border = "#D8DEE6"
            self._settings_text = "#1F2933"
            self._settings_muted = "#5F6B7A"
            self._settings_accent = "#2E7D6B"
            self.configure(background=self._settings_bg)

            self.columnconfigure(0, weight=1)
            self.rowconfigure(0, weight=1)

            scroll_container = tk.Frame(self, background=self._settings_bg)
            scroll_container.grid(row=0, column=0, sticky="nsew")

            settings_canvas = tk.Canvas(
                scroll_container,
                highlightthickness=0,
                background=self._settings_bg,
            )
            settings_canvas.pack(side="left", fill="both", expand=True)

            settings_scrollbar = ttk.Scrollbar(
                scroll_container, orient="vertical", command=settings_canvas.yview
            )
            settings_scrollbar.pack(side="right", fill="y")
            settings_canvas.configure(yscrollcommand=settings_scrollbar.set)

            content = tk.Frame(settings_canvas, background=self._settings_bg)
            content.configure(padx=18, pady=18)
            content_id = settings_canvas.create_window(
                (0, 0), window=content, anchor="nw"
            )

            def _update_scroll_region(_event: "tk.Event") -> None:
                try:
                    bbox = settings_canvas.bbox("all")
                except tk.TclError:
                    bbox = None
                if bbox:
                    settings_canvas.configure(scrollregion=bbox)

            content.bind("<Configure>", _update_scroll_region)

            def _resize_content(event: "tk.Event") -> None:
                try:
                    settings_canvas.itemconfigure(content_id, width=event.width)
                except tk.TclError:
                    pass

            settings_canvas.bind("<Configure>", _resize_content)

            def _on_mousewheel(event: "tk.Event") -> None:
                if getattr(event, "delta", 0):
                    step = -1 if event.delta > 0 else 1
                elif getattr(event, "num", None) == 4:
                    step = -1
                elif getattr(event, "num", None) == 5:
                    step = 1
                else:
                    step = 0
                if step:
                    settings_canvas.yview_scroll(step, "units")

            settings_canvas.bind("<Enter>", lambda _e: settings_canvas.focus_set())
            settings_canvas.bind("<MouseWheel>", _on_mousewheel)
            settings_canvas.bind("<Button-4>", _on_mousewheel)
            settings_canvas.bind("<Button-5>", _on_mousewheel)

            content.columnconfigure(0, weight=1)
            content.rowconfigure(6, weight=1)

            def _section(row: int, title: str, description: str = "") -> tk.Frame:
                section = tk.Frame(
                    content,
                    background=self._settings_card_bg,
                    highlightthickness=1,
                    highlightbackground=self._settings_border,
                    highlightcolor=self._settings_border,
                )
                section.grid(row=row, column=0, sticky="nsew", pady=(0, 12))
                section.columnconfigure(0, weight=1)

                header_frame = tk.Frame(section, background=self._settings_card_bg)
                header_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
                header_frame.columnconfigure(0, weight=1)
                tk.Label(
                    header_frame,
                    text=title,
                    background=self._settings_card_bg,
                    foreground=self._settings_text,
                    font=("TkDefaultFont", 11, "bold"),
                    anchor="w",
                ).grid(row=0, column=0, sticky="ew")
                if description:
                    tk.Label(
                        header_frame,
                        text=description,
                        background=self._settings_card_bg,
                        foreground=self._settings_muted,
                        justify="left",
                        anchor="w",
                        wraplength=620,
                    ).grid(row=1, column=0, sticky="ew", pady=(2, 0))

                body_frame = tk.Frame(section, background=self._settings_card_bg)
                body_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 14))
                body_frame.columnconfigure(0, weight=1)
                return body_frame

            header = tk.Frame(
                content,
                background=self._settings_card_bg,
                highlightthickness=1,
                highlightbackground=self._settings_border,
                highlightcolor=self._settings_border,
            )
            header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
            header.columnconfigure(0, weight=1)
            title_block = tk.Frame(header, background=self._settings_card_bg)
            title_block.grid(row=0, column=0, sticky="ew", padx=18, pady=14)
            title_block.columnconfigure(0, weight=1)
            tk.Label(
                title_block,
                text="Instellingen",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                font=("TkDefaultFont", 16, "bold"),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            tk.Label(
                title_block,
                text="Beheer templates, documentteksten, bestandstypen, data en diagnose.",
                background=self._settings_card_bg,
                foreground=self._settings_muted,
                anchor="w",
            ).grid(row=1, column=0, sticky="ew", pady=(2, 0))
            tk.Label(
                header,
                text=f"{APP_NAME} {APP_VERSION}",
                background="#E7F3EF",
                foreground=self._settings_accent,
                padx=10,
                pady=4,
                font=("TkDefaultFont", 9, "bold"),
            ).grid(row=0, column=1, sticky="ne", padx=18, pady=16)

            export_options = _section(
                1,
                "Exportgedrag",
                "Standaardgedrag voor extra exportbestanden en hulpmappen.",
            )
            export_options.columnconfigure(0, weight=1)

            def _add_option(
                parent: tk.Widget,
                text: str,
                description: str,
                variable: "tk.IntVar",
            ) -> None:
                row = parent.grid_size()[1]
                container = tk.Frame(parent, background=self._settings_card_bg)
                container.grid(row=row, column=0, sticky="ew", pady=(5, 5))
                container.columnconfigure(0, weight=1)
                tk.Label(
                    container,
                    text=text,
                    background=self._settings_card_bg,
                    foreground=self._settings_text,
                    font=("TkDefaultFont", 10, "bold"),
                    anchor="w",
                    justify="left",
                ).grid(row=0, column=0, sticky="ew")
                tk.Label(
                    container,
                    text=description,
                    justify="left",
                    anchor="w",
                    wraplength=560,
                    foreground=self._settings_muted,
                    background=self._settings_card_bg,
                ).grid(row=1, column=0, sticky="ew", pady=(1, 0))
                tk.Checkbutton(
                    container,
                    variable=variable,
                    background=self._settings_card_bg,
                    activebackground=self._settings_card_bg,
                    anchor="w",
                    justify="left",
                ).grid(row=0, column=1, rowspan=2, sticky="e", padx=(12, 0))

            _add_option(
                export_options,
                "Exporteer bewerkte BOM naar exportmap",
                "Bewaar de actuele BOM mee in de hoofdmap van elke export.",
                self.app.export_bom_var,
            )

            _add_option(
                export_options,
                "Exporteer gerelateerde exportbestanden naar exportmap",
                "Kopieer extra bestanden die bij de geopende BOM horen.",
                self.app.export_related_files_var,
            )

            _add_option(
                export_options,
                "Maak snelkoppeling naar nieuwste exportmap",
                "Maak een 'latest'-koppeling naar de meest recente export.",
                self.app.bundle_latest_var,
            )
            _add_option(
                export_options,
                "Testrun: toon alleen doelmap (niets wordt gekopieerd)",
                "Toon de exportdoelmap zonder bestanden aan te maken.",
                self.app.bundle_dry_run_var,
            )
            _add_option(
                export_options,
                "Vul Custom BOM automatisch na het laden van de hoofd-BOM",
                "Zet de geopende BOM ook klaar in de Custom BOM-tab.",
                self.app.autofill_custom_bom_var,
            )

            template_frame = _section(
                2,
                "BOM-template",
                "Download een leeg Excel-sjabloon met alle BOM-kolommen.",
            )
            template_frame.columnconfigure(0, weight=1)

            tk.Button(
                template_frame,
                text="Download BOM template",
                command=self._download_bom_template,
            ).grid(row=0, column=0, sticky="w")

            document_texts_frame = _section(
                3,
                "Documentteksten",
                "Beheer vaste teksten die op bestelbonnen en offerteaanvragen verschijnen.",
            )
            document_texts_frame.columnconfigure(0, weight=1)
            document_texts_frame.rowconfigure(0, weight=1)
            document_texts_frame.rowconfigure(1, weight=1)

            en1090_frame = tk.Frame(
                document_texts_frame,
                background=self._settings_card_bg,
                highlightthickness=1,
                highlightbackground="#EEF1F5",
                highlightcolor="#EEF1F5",
            )
            en1090_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
            en1090_frame.columnconfigure(0, weight=1)
            en1090_frame.rowconfigure(2, weight=1)

            tk.Label(
                en1090_frame,
                text="EN 1090",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                font=("TkDefaultFont", 10, "bold"),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))

            tk.Checkbutton(
                en1090_frame,
                text="Toon EN 1090-kolom bij bestelbonnen",
                variable=self.app.en1090_enabled_var,
                anchor="w",
                command=self._toggle_en1090_setting,
                background=self._settings_card_bg,
                activebackground=self._settings_card_bg,
            ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 8))

            en1090_note_container = tk.Frame(en1090_frame, background=self._settings_card_bg)
            en1090_note_container.grid(
                row=2,
                column=0,
                sticky="nsew",
                padx=12,
                pady=(0, 4),
            )
            en1090_note_container.columnconfigure(0, weight=1)
            en1090_note_container.rowconfigure(0, weight=1)

            self.en1090_note_text = tk.Text(
                en1090_note_container,
                height=5,
                wrap="word",
                font=("TkDefaultFont", 10),
            )
            self.en1090_note_text.grid(row=0, column=0, sticky="nsew")

            en1090_scrollbar = tk.Scrollbar(
                en1090_note_container,
                orient="vertical",
                command=self.en1090_note_text.yview,
            )
            en1090_scrollbar.grid(row=0, column=1, sticky="ns")
            self.en1090_note_text.configure(yscrollcommand=en1090_scrollbar.set)
            self._reload_en1090_note()

            en1090_btns = tk.Frame(en1090_frame, background=self._settings_card_bg)
            en1090_btns.grid(row=3, column=0, sticky="e", padx=12, pady=(4, 10))
            tk.Button(
                en1090_btns, text="Opslaan", command=self._save_en1090_note
            ).pack(side="left", padx=4)
            tk.Button(
                en1090_btns,
                text="Reset naar standaard",
                command=self._reset_en1090_note,
            ).pack(side="left", padx=4)

            footer_frame = tk.Frame(
                document_texts_frame,
                background=self._settings_card_bg,
                highlightthickness=1,
                highlightbackground="#EEF1F5",
                highlightcolor="#EEF1F5",
            )
            footer_frame.grid(row=1, column=0, sticky="nsew")
            footer_frame.columnconfigure(0, weight=1)
            footer_frame.rowconfigure(1, weight=1)

            tk.Label(
                footer_frame,
                text="Bestelbon/offerte onderschrift",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                font=("TkDefaultFont", 10, "bold"),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))

            footer_text_container = tk.Frame(footer_frame, background=self._settings_card_bg)
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
                font=("TkDefaultFont", 10),
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

            footer_btns = tk.Frame(footer_frame, background=self._settings_card_bg)
            footer_btns.grid(row=2, column=0, sticky="e", padx=12, pady=(4, 10))
            tk.Button(footer_btns, text="Opslaan", command=self._save_footer_note).pack(
                side="left", padx=4
            )
            tk.Button(
                footer_btns,
                text="Reset naar standaard",
                command=self._reset_footer_note,
            ).pack(side="left", padx=4)

            extensions_frame = _section(
                5,
                "Bestandstypen",
                "Beheer de extensies die op het hoofdscherm beschikbaar zijn.",
            )
            extensions_frame.columnconfigure(0, weight=1)
            extensions_frame.rowconfigure(0, weight=1)

            list_container = tk.Frame(extensions_frame, background=self._settings_card_bg)
            list_container.grid(row=0, column=0, columnspan=2, sticky="nsew")
            list_container.columnconfigure(0, weight=1)
            list_container.rowconfigure(0, weight=1)

            list_frame = tk.Frame(list_container, background=self._settings_card_bg)
            list_frame.grid(row=0, column=0, sticky="nsew")
            list_frame.columnconfigure(0, weight=1)

            self.extensions_tree = ttk.Treeview(
                list_frame,
                columns=("status", "label", "patterns"),
                show="headings",
                height=7,
                selectmode="browse",
            )
            self.extensions_tree.heading("status", text="Status")
            self.extensions_tree.heading("label", text="Naam")
            self.extensions_tree.heading("patterns", text="Extensies")
            self.extensions_tree.column("status", width=72, minwidth=60, anchor="center", stretch=False)
            self.extensions_tree.column("label", width=210, minwidth=150, stretch=True)
            self.extensions_tree.column("patterns", width=260, minwidth=160, stretch=True)
            self.extensions_tree.grid(row=0, column=0, sticky="nsew")
            scrollbar = tk.Scrollbar(list_frame, command=self.extensions_tree.yview)
            scrollbar.grid(row=0, column=1, sticky="ns")
            self.extensions_tree.configure(yscrollcommand=scrollbar.set)
            self.extensions_tree.bind("<Double-Button-1>", lambda _e: self._edit_selected())

            move_btns = tk.Frame(list_container, background=self._settings_card_bg)
            move_btns.grid(row=0, column=1, sticky="ns", padx=(8, 0))
            move_btns.grid_rowconfigure(0, weight=1)
            move_btns.grid_rowconfigure(3, weight=1)
            move_btns.grid_columnconfigure(0, weight=1)
            tk.Button(
                move_btns,
                text="Omhoog",
                width=8,
                command=lambda: self._move_selected(-1),
            ).grid(row=1, column=0, pady=2, sticky="nsew")
            tk.Button(
                move_btns,
                text="Omlaag",
                width=8,
                command=lambda: self._move_selected(1),
            ).grid(row=2, column=0, pady=2, sticky="nsew")

            btns = tk.Frame(extensions_frame, background=self._settings_card_bg)
            btns.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
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
            self._build_release_notes_section(content, row=6)
            self._build_help_section(content, row=7)
            self._build_diagnostics_section(content, row=8)
            self._refresh_diagnostics()

        def _build_release_notes_section(self, content: tk.Widget, *, row: int) -> None:
            section = tk.Frame(
                content,
                background=self._settings_card_bg,
                highlightthickness=1,
                highlightbackground=self._settings_border,
                highlightcolor=self._settings_border,
            )
            section.grid(row=row, column=0, sticky="nsew", pady=(0, 12))
            section.columnconfigure(0, weight=1)

            header_frame = tk.Frame(section, background=self._settings_card_bg)
            header_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
            header_frame.columnconfigure(0, weight=1)
            tk.Label(
                header_frame,
                text="Release Notes & Info",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                font=("TkDefaultFont", 11, "bold"),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            tk.Label(
                header_frame,
                text="Bekijk de meest recente release notes of informatie over Filehopper.",
                background=self._settings_card_bg,
                foreground=self._settings_muted,
                justify="left",
                anchor="w",
                wraplength=620,
            ).grid(row=1, column=0, sticky="ew", pady=(2, 0))

            body_frame = tk.Frame(section, background=self._settings_card_bg)
            body_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 14))
            body_frame.columnconfigure(0, weight=1)

            tk.Label(
                body_frame,
                text=f"Huidige versie: {APP_VERSION}",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                anchor="w",
            ).grid(row=0, column=0, sticky="w", pady=(0, 8))

            latest_notes = get_latest_release_notes(load_changelog())
            tk.Label(
                body_frame,
                text="Laatste release notes:",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                font=("TkDefaultFont", 10, "bold"),
                anchor="w",
            ).grid(row=2, column=0, sticky="w")
            tk.Label(
                body_frame,
                text=latest_notes,
                background=self._settings_card_bg,
                foreground=self._settings_text,
                justify="left",
                anchor="w",
                wraplength=620,
            ).grid(row=3, column=0, sticky="ew", pady=(2, 10))

            btn_frame = tk.Frame(body_frame, background=self._settings_card_bg)
            btn_frame.grid(row=4, column=0, sticky="w")
            tk.Button(
                btn_frame,
                text="Bekijk release notes",
                command=self.app._show_release_notes,
            ).pack(side="left", padx=(0, 8), pady=4)
            tk.Button(
                btn_frame,
                text="Over Filehopper",
                command=self.app._show_about_dialog,
            ).pack(side="left", padx=(0, 8), pady=4)

        def _build_help_section(self, content: tk.Widget, *, row: int) -> None:
            section = tk.Frame(
                content,
                background=self._settings_card_bg,
                highlightthickness=1,
                highlightbackground=self._settings_border,
                highlightcolor=self._settings_border,
            )
            section.grid(row=row, column=0, sticky="nsew", pady=(0, 12))
            section.columnconfigure(0, weight=1)

            header_frame = tk.Frame(section, background=self._settings_card_bg)
            header_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
            header_frame.columnconfigure(0, weight=1)
            tk.Label(
                header_frame,
                text="Hulp & Quick Start",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                font=("TkDefaultFont", 11, "bold"),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            tk.Label(
                header_frame,
                text=(
                    "Volg een korte opstartgids of bekijk de meest voorkomende vragen en oplossingen."
                ),
                background=self._settings_card_bg,
                foreground=self._settings_muted,
                justify="left",
                anchor="w",
                wraplength=620,
            ).grid(row=1, column=0, sticky="ew", pady=(2, 0))

            body_frame = tk.Frame(section, background=self._settings_card_bg)
            body_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 14))
            body_frame.columnconfigure(0, weight=1)

            tk.Label(
                body_frame,
                text="Snel aan de slag:",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                anchor="w",
            ).grid(row=0, column=0, sticky="w", pady=(0, 4))

            preview_text = "\n".join(
                f"{index + 1}. {step['title']}" for index, step in enumerate(QUICK_START_STEPS[:3])
            )
            tk.Label(
                body_frame,
                text=preview_text,
                background=self._settings_card_bg,
                foreground=self._settings_text,
                justify="left",
                anchor="w",
                wraplength=620,
            ).grid(row=1, column=0, sticky="ew")

            btn_frame = tk.Frame(body_frame, background=self._settings_card_bg)
            btn_frame.grid(row=2, column=0, sticky="w", pady=(10, 0))
            tk.Button(
                btn_frame,
                text="Start Quick Start",
                command=self.app._show_quick_start,
            ).pack(side="left", padx=(0, 8), pady=4)
            tk.Button(
                btn_frame,
                text="Visuele quick manual",
                command=self.app._open_visual_quick_manual,
            ).pack(side="left", padx=(0, 8), pady=4)
            tk.Button(
                btn_frame,
                text="Bekijk FAQ",
                command=self.app._show_faq,
            ).pack(side="left", padx=(0, 8), pady=4)

        def _build_diagnostics_section(self, content: tk.Widget, *, row: int) -> None:
            section = tk.Frame(
                content,
                background=self._settings_card_bg,
                highlightthickness=1,
                highlightbackground=self._settings_border,
                highlightcolor=self._settings_border,
            )
            section.grid(row=row, column=0, sticky="nsew", pady=(0, 12))
            section.columnconfigure(0, weight=1)

            header_frame = tk.Frame(section, background=self._settings_card_bg)
            header_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 6))
            header_frame.columnconfigure(0, weight=1)
            tk.Label(
                header_frame,
                text="Data en diagnose",
                background=self._settings_card_bg,
                foreground=self._settings_text,
                font=("TkDefaultFont", 11, "bold"),
                anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            tk.Label(
                header_frame,
                text="Controleer databestanden, backups, schrijfrechten en waarschuwingen.",
                background=self._settings_card_bg,
                foreground=self._settings_muted,
                justify="left",
                anchor="w",
                wraplength=620,
            ).grid(row=1, column=0, sticky="ew", pady=(2, 0))

            diagnostics_frame = tk.Frame(section, background=self._settings_card_bg)
            diagnostics_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 14))
            diagnostics_frame.columnconfigure(0, weight=1)
            diagnostics_frame.rowconfigure(2, weight=1)

            cards_frame = tk.Frame(diagnostics_frame, background=self._settings_card_bg)
            cards_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
            for col in range(3):
                cards_frame.columnconfigure(col, weight=1)

            self.diagnostics_files_var = tk.StringVar(value="0")
            self.diagnostics_warnings_var = tk.StringVar(value="0")
            self.diagnostics_problems_var = tk.StringVar(value="0")

            def _status_card(column: int, title: str, variable: "tk.StringVar", color: str) -> tk.Frame:
                card = tk.Frame(
                    cards_frame,
                    background=color,
                    highlightthickness=1,
                    highlightbackground=self._settings_border,
                )
                card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 6, 0))
                tk.Label(
                    card,
                    text=title,
                    background=color,
                    foreground=self._settings_muted,
                    anchor="w",
                ).grid(row=0, column=0, sticky="ew", padx=10, pady=(7, 0))
                tk.Label(
                    card,
                    textvariable=variable,
                    background=color,
                    foreground=self._settings_text,
                    font=("TkDefaultFont", 14, "bold"),
                    anchor="w",
                ).grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 7))
                return card

            self.diagnostics_files_card = _status_card(
                0, "Databestanden", self.diagnostics_files_var, "#F7FAFC"
            )
            self.diagnostics_warnings_card = _status_card(
                1, "Waarschuwingen", self.diagnostics_warnings_var, "#FFF7E6"
            )
            self.diagnostics_problems_card = _status_card(
                2, "Problemen", self.diagnostics_problems_var, "#FDECEC"
            )

            self.diagnostics_summary_var = tk.StringVar(value="")
            tk.Label(
                diagnostics_frame,
                textvariable=self.diagnostics_summary_var,
                justify="left",
                anchor="w",
                background=self._settings_card_bg,
                foreground=self._settings_muted,
                wraplength=560,
            ).grid(row=1, column=0, sticky="ew", pady=(0, 8))

            columns = ("status", "aantal", "gewijzigd", "schrijfbaar", "backups")
            self.diagnostics_tree = ttk.Treeview(
                diagnostics_frame,
                columns=columns,
                show="tree headings",
                height=6,
                selectmode="browse",
            )
            self.diagnostics_tree.heading("#0", text="Bestand")
            self.diagnostics_tree.heading("status", text="Status")
            self.diagnostics_tree.heading("aantal", text="Aantal")
            self.diagnostics_tree.heading("gewijzigd", text="Gewijzigd")
            self.diagnostics_tree.heading("schrijfbaar", text="Schrijfbaar")
            self.diagnostics_tree.heading("backups", text="Backups")
            self.diagnostics_tree.column("#0", width=155, minwidth=130, stretch=True)
            self.diagnostics_tree.column("status", width=90, minwidth=80, anchor="center")
            self.diagnostics_tree.column("aantal", width=70, minwidth=60, anchor="center")
            self.diagnostics_tree.column("gewijzigd", width=120, minwidth=110)
            self.diagnostics_tree.column("schrijfbaar", width=80, minwidth=70, anchor="center")
            self.diagnostics_tree.column("backups", width=85, minwidth=75, anchor="center")
            self.diagnostics_tree.grid(row=2, column=0, sticky="nsew", pady=(0, 8))

            warning_frame = tk.Frame(diagnostics_frame, background=self._settings_card_bg)
            warning_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 8))
            warning_frame.columnconfigure(0, weight=1)
            warning_frame.rowconfigure(0, weight=1)

            self.diagnostics_warning_text = tk.Text(
                warning_frame,
                height=4,
                wrap="word",
                font=("TkDefaultFont", 9),
                state="disabled",
            )
            self.diagnostics_warning_text.grid(row=0, column=0, sticky="nsew")
            warning_scroll = tk.Scrollbar(
                warning_frame,
                orient="vertical",
                command=self.diagnostics_warning_text.yview,
            )
            warning_scroll.grid(row=0, column=1, sticky="ns")
            self.diagnostics_warning_text.configure(yscrollcommand=warning_scroll.set)

            btns = tk.Frame(diagnostics_frame, background=self._settings_card_bg)
            btns.grid(row=4, column=0, sticky="ew")
            tk.Button(btns, text="Vernieuwen", command=self._refresh_diagnostics).pack(
                side="left", padx=(0, 4)
            )
            tk.Button(btns, text="Backup maken", command=self._create_diagnostics_backup).pack(
                side="left", padx=4
            )
            tk.Button(btns, text="Datamap openen", command=self._open_diagnostics_data_folder).pack(
                side="left", padx=4
            )
            tk.Button(btns, text="Backupmap openen", command=self._open_diagnostics_backup_folder).pack(
                side="left", padx=4
            )
            tk.Button(btns, text="Kopieer diagnose", command=self._copy_diagnostics).pack(
                side="left", padx=4
            )

        def _refresh_diagnostics(self) -> None:
            try:
                report = build_diagnostic_report(self.app.settings)
            except Exception as exc:
                messagebox.showerror("Diagnose mislukt", str(exc), parent=self)
                return
            self._diagnostic_report = report
            problem_count = sum(1 for item in report.data_files if item.status != "OK")
            warning_count = len(report.warnings)
            self.diagnostics_files_var.set(str(len(report.data_files)))
            self.diagnostics_warnings_var.set(str(warning_count))
            self.diagnostics_problems_var.set(str(problem_count))
            self.diagnostics_warnings_card.configure(
                background="#FFF7E6" if warning_count else "#F0F7F4"
            )
            self.diagnostics_problems_card.configure(
                background="#FDECEC" if problem_count else "#F0F7F4"
            )
            for child in self.diagnostics_warnings_card.winfo_children():
                child.configure(background=self.diagnostics_warnings_card.cget("background"))
            for child in self.diagnostics_problems_card.winfo_children():
                child.configure(background=self.diagnostics_problems_card.cget("background"))
            self.diagnostics_summary_var.set(
                (
                    f"{report.app_name} {report.app_version} | Runtime: {report.runtime_mode} | "
                    f"Datamap: {report.storage_path}"
                )
            )
            existing_items = self.diagnostics_tree.get_children()
            if existing_items:
                self.diagnostics_tree.delete(*existing_items)
            for item in report.data_files:
                backups_label = (
                    f"{item.backup_count} ({item.latest_backup_label})"
                    if item.backup_count
                    else "0"
                )
                self.diagnostics_tree.insert(
                    "",
                    tk.END,
                    text=item.filename,
                    values=(
                        item.status,
                        item.count_label,
                        item.modified_label,
                        item.writable_label,
                        backups_label,
                    ),
                )
            self._set_diagnostic_warnings(report.warnings)

        def _set_diagnostic_warnings(self, warnings: List[str]) -> None:
            self.diagnostics_warning_text.configure(state="normal")
            self.diagnostics_warning_text.delete("1.0", "end")
            if warnings:
                self.diagnostics_warning_text.insert(
                    "1.0",
                    "\n".join(f"- {warning}" for warning in warnings),
                )
            else:
                self.diagnostics_warning_text.insert("1.0", "Geen waarschuwingen.")
            self.diagnostics_warning_text.configure(state="disabled")

        def _create_diagnostics_backup(self) -> None:
            try:
                backups = create_data_file_backups()
            except Exception as exc:
                messagebox.showerror("Backup mislukt", str(exc), parent=self)
                return
            self._refresh_diagnostics()
            if backups:
                messagebox.showinfo(
                    "Backup gemaakt",
                    f"{len(backups)} backupbestand(en) gemaakt.",
                    parent=self,
                )
            else:
                messagebox.showinfo(
                    "Geen backup gemaakt",
                    "Er zijn geen bestaande databestanden gevonden om te backuppen.",
                    parent=self,
                )

        def _open_diagnostics_data_folder(self) -> None:
            report = getattr(self, "_diagnostic_report", None)
            folder = report.storage_path if report is not None else Path.cwd()
            self._open_folder(folder)

        def _open_diagnostics_backup_folder(self) -> None:
            folder = backup_root()
            folder.mkdir(parents=True, exist_ok=True)
            self._open_folder(folder)

        def _copy_diagnostics(self) -> None:
            report = getattr(self, "_diagnostic_report", None)
            if report is None:
                report = build_diagnostic_report(self.app.settings)
                self._diagnostic_report = report
            self.clipboard_clear()
            self.clipboard_append(format_report_for_clipboard(report))
            messagebox.showinfo("Diagnose gekopieerd", "Diagnose staat op het klembord.", parent=self)

        def _open_folder(self, folder: Path) -> None:
            try:
                if sys.platform.startswith("win"):
                    os.startfile(str(folder))  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", str(folder)])
                else:
                    subprocess.Popen(["xdg-open", str(folder)])
            except Exception as exc:
                messagebox.showerror("Map openen mislukt", str(exc), parent=self)

        def _refresh_list(self) -> None:
            existing_items = self.extensions_tree.get_children()
            if existing_items:
                self.extensions_tree.delete(*existing_items)
            if not self.extensions:
                self.extensions_tree.insert(
                    "",
                    tk.END,
                    iid="empty",
                    values=("-", "Geen bestandstypen gedefinieerd", "-"),
                )
                return
            for idx, ext in enumerate(self.extensions):
                status = "Aan" if ext.enabled else "Uit"
                patterns = ", ".join(ext.patterns)
                self.extensions_tree.insert(
                    "",
                    tk.END,
                    iid=str(idx),
                    values=(status, ext.label, patterns),
                )

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

        def _reload_en1090_note(self) -> None:
            text = self.app.en1090_note_var.get()
            self.en1090_note_text.delete("1.0", "end")
            if text:
                self.en1090_note_text.insert("1.0", text)

        def _current_en1090_note_text(self) -> str:
            raw = self.en1090_note_text.get("1.0", "end-1c")
            return raw.replace("\r\n", "\n")

        def _save_en1090_note(self) -> None:
            note = self._current_en1090_note_text()
            self.app.update_en1090_note(note)
            self._reload_en1090_note()

        def _reset_en1090_note(self) -> None:
            self.app.update_en1090_note(EN1090_NOTE_TEXT)
            self._reload_en1090_note()

        def _toggle_en1090_setting(self) -> None:
            self.app.set_en1090_enabled(bool(self.app.en1090_enabled_var.get()))

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
            return

        def _update_listbox_width(self) -> None:
            return

        def _selected_index(self) -> Optional[int]:
            if not self.extensions:
                return None
            sel = self.extensions_tree.selection()
            if not sel:
                return None
            try:
                idx = int(sel[0])
            except (TypeError, ValueError):
                return None
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
                item_id = str(new_idx)
                self.extensions_tree.selection_set(item_id)
                self.extensions_tree.focus(item_id)
                self.extensions_tree.see(item_id)

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

    class PdfWorkDossierSectionsEditor(tk.Frame):
        UNMATCHED_DISPLAY_NAME = "Overige producties"
        LEGACY_UNMATCHED_NAMES = {"other names", "overige"}

        def __init__(self, master) -> None:
            super().__init__(master)
            self.sections: List[PdfWorkDossierSection] = []
            self._row_controls: List[Dict[str, object]] = []
            self._selected_index: Optional[int] = None
            self._drag_start_index: Optional[int] = None
            self._drag_over_index: Optional[int] = None
            self._drag_widget = None
            self._enabled = True

            body = tk.Frame(self)
            body.pack(fill="both", expand=True)
            body.rowconfigure(0, weight=1)
            body.columnconfigure(0, weight=1)
            body.columnconfigure(1, minsize=160)

            content_frame = tk.Frame(body)
            content_frame.grid(row=0, column=0, sticky="nsew")
            content_frame.rowconfigure(0, weight=1)
            content_frame.columnconfigure(0, weight=1)

            self.canvas = tk.Canvas(content_frame, highlightthickness=0, borderwidth=0, height=210)
            self.canvas.grid(row=0, column=0, sticky="nsew")
            self.scrollbar = ttk.Scrollbar(
                content_frame,
                orient="vertical",
                command=self.canvas.yview,
            )
            self.scrollbar.grid(row=0, column=1, sticky="ns", padx=(8, 0))
            self.canvas.configure(yscrollcommand=self.scrollbar.set)

            actions = tk.Frame(body)
            actions.grid(row=0, column=1, sticky="n", padx=(12, 0), pady=(0, 6))
            self._add_button = ttk.Button(
                actions,
                text="Blok toevoegen",
                command=self.add_blank_section,
            )
            self._add_button.pack(fill="x", pady=(0, 4))
            self._up_button = ttk.Button(actions, text="Omhoog", command=self.move_selected_up)
            self._up_button.pack(fill="x", pady=(0, 4))
            self._down_button = ttk.Button(actions, text="Omlaag", command=self.move_selected_down)
            self._down_button.pack(fill="x", pady=(0, 4))
            self._delete_button = ttk.Button(
                actions,
                text="Verwijder blok",
                command=self.delete_selected,
            )
            self._delete_button.pack(fill="x", pady=(0, 4))
            self._clear_button = ttk.Button(
                actions,
                text="Leegmaken",
                command=self.clear_sections,
            )
            self._clear_button.pack(fill="x", pady=(0, 4))
            self.inner = tk.Frame(self.canvas)
            self.inner_window = self.canvas.create_window(
                (0, 0),
                window=self.inner,
                anchor="nw",
            )
            self.inner.bind(
                "<Configure>",
                lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
            )
            self.canvas.bind(
                "<Configure>",
                lambda event: self.canvas.itemconfigure(
                    self.inner_window,
                    width=event.width,
                ),
            )
            self._render_sections()

        def set_enabled(self, enabled: bool) -> None:
            self._enabled = bool(enabled)
            state = "normal" if self._enabled else "disabled"
            for button in (
                self._add_button,
                self._up_button,
                self._down_button,
                self._delete_button,
                self._clear_button,
            ):
                button.configure(state=state)
            for controls in self._row_controls:
                for widget in controls.get("stateful_widgets", []):
                    try:
                        widget.configure(state=state)
                    except Exception:
                        pass
                identifiers_entry = controls.get("identifiers_entry")
                identifiers_var = controls.get("identifiers_var")
                unmatched_var = controls.get("unmatched_var")
                if (
                    identifiers_entry is not None
                    and identifiers_var is not None
                    and unmatched_var is not None
                ):
                    self._set_unmatched_state(
                        identifiers_entry,
                        identifiers_var,
                        unmatched_var,
                    )

        def _sync_sections_from_rows(self) -> None:
            if not self._row_controls:
                return
            sections: List[PdfWorkDossierSection] = []
            for controls in self._row_controls:
                name_var = controls.get("name_var")
                identifiers_var = controls.get("identifiers_var")
                unmatched_var = controls.get("unmatched_var")
                name = name_var.get().strip() if name_var is not None else ""
                include_unmatched = bool(unmatched_var.get()) if unmatched_var is not None else False
                if include_unmatched and not name:
                    name = self.UNMATCHED_DISPLAY_NAME
                if not name:
                    continue
                identifiers: List[str] = []
                if not include_unmatched and identifiers_var is not None:
                    identifiers = [
                        part.strip()
                        for part in re.split(r"[,;]+", identifiers_var.get())
                        if part.strip()
                    ]
                sections.append(
                    PdfWorkDossierSection(
                        name,
                        identifiers=identifiers,
                        include_unmatched=include_unmatched,
                    )
                )
            self.sections = sections
            if self._selected_index is not None and self._selected_index >= len(self.sections):
                self._selected_index = len(self.sections) - 1 if self.sections else None

        def _select_index(self, index: int) -> None:
            if 0 <= index < len(self.sections):
                self._selected_index = index
                self._highlight_selected_rows()

        def _set_unmatched_state(self, identifiers_entry, identifiers_var, unmatched_var) -> None:
            if bool(unmatched_var.get()):
                identifiers_var.set("")
                identifiers_entry.configure(state="disabled")
            else:
                identifiers_entry.configure(state="normal" if self._enabled else "disabled")

        def _display_section_name(self, section: PdfWorkDossierSection) -> str:
            name = _to_str(section.name).strip()
            if section.include_unmatched and name.casefold() in self.LEGACY_UNMATCHED_NAMES:
                return self.UNMATCHED_DISPLAY_NAME
            return name

        def _bind_drag_handle(self, widget, index: int) -> None:
            widget.bind(
                "<ButtonPress-1>",
                lambda event, i=index: self._start_drag(i, event),
                add="+",
            )
            widget.bind("<B1-Motion>", self._drag_motion, add="+")
            widget.bind("<ButtonRelease-1>", self._finish_drag, add="+")

        def _start_drag(self, index: int, event=None) -> str:
            if not self._enabled:
                return "break"
            self._sync_sections_from_rows()
            self._selected_index = index
            self._drag_start_index = index
            self._drag_over_index = index
            self._drag_widget = getattr(event, "widget", None)
            if self._drag_widget is not None:
                try:
                    self._drag_widget.grab_set()
                except tk.TclError:
                    pass
            self._highlight_selected_rows()
            return "break"

        def _target_index_from_root_y(self, y_root: int) -> Optional[int]:
            if not self._row_controls:
                return None
            last_index = len(self._row_controls) - 1
            for index, controls in enumerate(self._row_controls):
                row = controls.get("row")
                if row is None:
                    continue
                try:
                    top = row.winfo_rooty()
                    height = max(1, row.winfo_height())
                except tk.TclError:
                    continue
                if y_root < top + (height // 2):
                    return index
            return last_index

        def _drag_motion(self, event) -> str:
            if self._drag_start_index is None:
                return "break"
            target = self._target_index_from_root_y(int(event.y_root))
            if target is not None:
                self._drag_over_index = target
            return "break"

        def _finish_drag(self, event=None) -> str:
            start = self._drag_start_index
            target = None
            if event is not None:
                try:
                    target = self._target_index_from_root_y(int(event.y_root))
                except Exception:
                    target = None
            if target is None:
                target = self._drag_over_index
            widget = self._drag_widget
            self._drag_start_index = None
            self._drag_over_index = None
            self._drag_widget = None
            if widget is not None:
                try:
                    widget.grab_release()
                except tk.TclError:
                    pass
            if start is None or target is None or start == target:
                return "break"
            self._move_index(start, target)
            return "break"

        def _highlight_selected_rows(self) -> None:
            for index, controls in enumerate(self._row_controls):
                row = controls.get("row")
                if row is None:
                    continue
                color = "#EAF2FF" if index == self._selected_index else "#FFFFFF"
                try:
                    row.configure(background=color)
                    for child in row.winfo_children():
                        if not isinstance(child, ttk.Entry) and not isinstance(child, ttk.Button):
                            try:
                                child.configure(background=color)
                            except tk.TclError:
                                pass
                except tk.TclError:
                    pass

        def _convert_example_to_real_section(self, _event=None) -> None:
            if self.sections:
                return
            self.sections = [PdfWorkDossierSection("", [])]
            self._selected_index = 0
            self._render_sections()
            first_row = self._row_controls[0] if self._row_controls else {}
            name_entry = first_row.get("name_entry")
            if name_entry is not None:
                try:
                    name_entry.focus_set()
                except tk.TclError:
                    pass

        def _render_sections(self) -> None:
            for child in self.inner.winfo_children():
                child.destroy()
            self._row_controls = []

            if not self.sections:
                bg = "#FFFFFF"
                placeholder_entry_style = "Placeholder.TEntry"
                try:
                    style = ttk.Style(self)
                    style.configure(placeholder_entry_style, foreground="#8A8F98")
                    style.map(
                        placeholder_entry_style,
                        foreground=[
                            ("readonly", "#8A8F98"),
                            ("disabled", "#8A8F98"),
                        ],
                    )
                except tk.TclError:
                    placeholder_entry_style = "TEntry"
                example = tk.Frame(
                    self.inner,
                    relief="solid",
                    borderwidth=1,
                    background=bg,
                )
                example.pack(fill="x", padx=(2, 10), pady=3)
                example.columnconfigure(3, weight=1)

                handle = tk.Label(
                    example,
                    text="::",
                    width=3,
                    cursor="sb_v_double_arrow",
                    background=bg,
                    foreground="#999999",
                )
                handle.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(4, 2), pady=4)
                number = tk.Label(
                    example,
                    text="1.",
                    width=4,
                    anchor="e",
                    background=bg,
                    foreground="#999999",
                )
                number.grid(row=0, column=1, rowspan=2, sticky="ns", padx=(0, 6), pady=4)

                cat_label = tk.Label(
                    example,
                    text="Categorie:",
                    background=bg,
                )
                cat_label.grid(row=0, column=2, sticky="e", padx=(0, 4), pady=(4, 2))
                cat_entry_var = tk.StringVar(value="Parts")
                cat_entry = ttk.Entry(
                    example,
                    textvariable=cat_entry_var,
                    width=26,
                    state="readonly",
                    style=placeholder_entry_style,
                )
                cat_entry.grid(row=0, column=3, sticky="ew", pady=(4, 2))

                prod_label = tk.Label(
                    example,
                    text="Producties:",
                    background=bg,
                )
                prod_label.grid(row=1, column=2, sticky="e", padx=(0, 4), pady=(2, 4))
                prod_entry_var = tk.StringVar(value="Lasercutting, Tube laser")
                prod_entry = ttk.Entry(
                    example,
                    textvariable=prod_entry_var,
                    state="readonly",
                    style=placeholder_entry_style,
                )
                prod_entry.grid(row=1, column=3, sticky="ew", pady=(2, 4))
                unmatched_var = tk.IntVar(value=0)
                unmatched_check = tk.Checkbutton(
                    example,
                    text="Overige producties",
                    variable=unmatched_var,
                    background=bg,
                    command=self._convert_example_to_real_section,
                )
                unmatched_check.grid(row=0, column=4, rowspan=2, sticky="w", padx=(8, 4))
                delete_btn = ttk.Button(
                    example,
                    text="X",
                    width=3,
                    command=self._convert_example_to_real_section,
                )
                delete_btn.grid(row=0, column=5, rowspan=2, sticky="ns", padx=(0, 4), pady=4)

                for widget in (
                    example,
                    handle,
                    number,
                    cat_label,
                    cat_entry,
                    prod_label,
                    prod_entry,
                    unmatched_check,
                ):
                    widget.bind("<Button-1>", self._convert_example_to_real_section, add="+")

                return

            for index, section in enumerate(self.sections):
                selected = index == self._selected_index
                row = tk.Frame(
                    self.inner,
                    relief="solid",
                    borderwidth=1,
                    background="#EAF2FF" if selected else "#FFFFFF",
                )
                row.pack(fill="x", padx=(2, 10), pady=3)
                row.columnconfigure(3, weight=1)

                handle = tk.Label(
                    row,
                    text="::",
                    width=3,
                    cursor="sb_v_double_arrow",
                    background=row.cget("background"),
                )
                handle.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(4, 2), pady=4)
                number = tk.Label(
                    row,
                    text=f"{index + 1}.",
                    width=4,
                    anchor="e",
                    background=row.cget("background"),
                )
                number.grid(row=0, column=1, rowspan=2, sticky="ns", padx=(0, 6), pady=4)
                name_var = tk.StringVar(value=self._display_section_name(section))
                identifiers_var = tk.StringVar(value=", ".join(section.identifiers))
                unmatched_var = tk.IntVar(value=1 if section.include_unmatched else 0)

                name_label = tk.Label(row, text="Categorie:", background=row.cget("background"))
                name_label.grid(
                    row=0,
                    column=2,
                    sticky="e",
                    padx=(0, 4),
                    pady=(4, 2),
                )
                _HelpTooltip(
                    name_label,
                    "Categorie naam voor het werkdossier. Gebruik dezelfde naam om meerdere producties te groeperen.",
                )
                name_entry = ttk.Entry(row, textvariable=name_var, width=26)
                name_entry.grid(row=0, column=3, sticky="ew", pady=(4, 2))
                _HelpTooltip(
                    name_entry,
                    "Categorie naam voor het werkdossier. Gebruik dezelfde naam om meerdere producties te groeperen.",
                )
                identifiers_label = tk.Label(
                    row,
                    text="Producties:",
                    background=row.cget("background"),
                )
                identifiers_label.grid(
                    row=1,
                    column=2,
                    sticky="e",
                    padx=(0, 4),
                    pady=(2, 4),
                )
                identifiers_entry = ttk.Entry(row, textvariable=identifiers_var)
                identifiers_entry.grid(row=1, column=3, sticky="ew", pady=(2, 4))
                unmatched_check = tk.Checkbutton(
                    row,
                    text="Overige producties",
                    variable=unmatched_var,
                    background=row.cget("background"),
                    command=lambda e=identifiers_entry, v=identifiers_var, u=unmatched_var: self._set_unmatched_state(e, v, u),
                )
                unmatched_check.grid(row=0, column=4, rowspan=2, sticky="w", padx=(8, 4))
                _HelpTooltip(
                    unmatched_check,
                    "Dit blok vangt producties op die niet overeenkomen met een ander blok.",
                )
                delete_btn = ttk.Button(
                    row,
                    text="X",
                    width=3,
                    command=lambda i=index: self.delete_index(i),
                )
                delete_btn.grid(row=0, column=5, rowspan=2, sticky="ns", padx=(0, 4), pady=4)

                self._bind_drag_handle(handle, index)
                self._bind_drag_handle(number, index)
                row.bind("<Button-1>", lambda _e, i=index: self._select_index(i), add="+")
                for widget in (name_label, identifiers_label, unmatched_check):
                    widget.bind(
                        "<Button-1>",
                        lambda _e, i=index: self._select_index(i),
                        add="+",
                    )
                for widget in (name_entry, identifiers_entry):
                    widget.bind(
                        "<FocusIn>",
                        lambda _e, i=index: self._select_index(i),
                        add="+",
                    )

                self._set_unmatched_state(identifiers_entry, identifiers_var, unmatched_var)
                stateful_widgets = [
                    name_entry,
                    identifiers_entry,
                    unmatched_check,
                    delete_btn,
                ]
                self._row_controls.append(
                    {
                        "name_var": name_var,
                        "identifiers_var": identifiers_var,
                        "unmatched_var": unmatched_var,
                        "row": row,
                        "name_entry": name_entry,
                        "identifiers_entry": identifiers_entry,
                        "stateful_widgets": stateful_widgets,
                    }
                )
            self.set_enabled(self._enabled)

        def set_sections(self, sections: Sequence[PdfWorkDossierSection]) -> None:
            self.sections = [
                PdfWorkDossierSection.from_any(section)
                for section in sections
                if not getattr(section, "include_bom_pdf", False)
            ]
            self._selected_index = 0 if self.sections else None
            self._render_sections()

        def get_sections(self) -> List[PdfWorkDossierSection]:
            self._sync_sections_from_rows()
            return [PdfWorkDossierSection.from_any(section) for section in self.sections]

        def add_section(
            self,
            name: str,
            identifiers: Optional[List[str]] = None,
            *,
            include_unmatched: bool = False,
        ) -> None:
            if not self._enabled:
                return
            self._sync_sections_from_rows()
            self.sections.append(
                PdfWorkDossierSection(
                    name or (self.UNMATCHED_DISPLAY_NAME if include_unmatched else ""),
                    identifiers=list(identifiers or []),
                    include_unmatched=include_unmatched,
                )
            )
            self._selected_index = len(self.sections) - 1
            self._render_sections()

        def add_blank_section(self) -> None:
            next_number = len(self.sections) + 1
            self.add_section(f"Nieuw blok {next_number}", [])

        def clear_sections(self) -> None:
            if not self._enabled:
                return
            self.sections = []
            self._selected_index = None
            self._render_sections()

        def delete_index(self, index: int) -> None:
            if not self._enabled:
                return
            self._sync_sections_from_rows()
            if not (0 <= index < len(self.sections)):
                return
            self.sections.pop(index)
            if not self.sections:
                self._selected_index = None
            else:
                self._selected_index = min(index, len(self.sections) - 1)
            self._render_sections()

        def delete_selected(self) -> None:
            if self._selected_index is not None:
                self.delete_index(self._selected_index)

        def _move_index(self, start: int, target: int) -> None:
            if not self._enabled:
                return
            self._sync_sections_from_rows()
            if not (0 <= start < len(self.sections) and 0 <= target < len(self.sections)):
                return
            section = self.sections.pop(start)
            self.sections.insert(target, section)
            self._selected_index = target
            self._render_sections()

        def move_selected_up(self) -> None:
            if self._selected_index is None or self._selected_index <= 0:
                return
            self._move_index(self._selected_index, self._selected_index - 1)

        def move_selected_down(self) -> None:
            if self._selected_index is None or self._selected_index >= len(self.sections) - 1:
                return
            self._move_index(self._selected_index, self._selected_index + 1)

    class PdfWorkDossierOptionsFrame(tk.Frame):
        MODE_WORKDOSSIER = "Werkdossier (een PDF)"
        MODE_PER_PRODUCTION = "Aparte PDF per productie"
        MODE_ALPHABETIC_SINGLE = "PDF alfabetisch (alle bestanden)"
        NO_PRESET_LABEL = "(Blanco template)"
        ROLE_LABELS = {
            "bom": "BOM",
            "drawing": "Tekening",
            "order": "Bestelbon",
        }

        def __init__(
            self,
            master,
            presets_db: PdfWorkDossierPresetsDB,
            *,
            on_presets_changed=None,
            on_options_changed=None,
            on_prepare_orders=None,
        ):
            super().__init__(master)
            self.presets_db = presets_db
            self.on_presets_changed = on_presets_changed
            self.on_options_changed = on_options_changed
            self.on_prepare_orders = on_prepare_orders
            self._preset_map: Dict[str, PdfWorkDossierPreset] = {}
            self._busy = False
            self._option_widgets: List[tk.Widget] = []
            self._preset_action_widgets: List[tk.Widget] = []

            self.mode_var = tk.StringVar(value=self.MODE_WORKDOSSIER)
            self.preset_var = tk.StringVar(value=self.NO_PRESET_LABEL)
            self.include_bom_var = tk.IntVar(value=0)
            self.include_order_docs_var = tk.IntVar(value=0)
            self.include_offers_var = tk.IntVar(value=0)
            self.open_pdf_var = tk.IntVar(value=1)
            self.export_filename_var = tk.StringVar(value="")
            self.export_folder_var = tk.StringVar(value="")
            self.preview_status_var = tk.StringVar(value="Geen voorbeeld beschikbaar.")
            self.order_flow_message_var = tk.StringVar(value="")

            self.columnconfigure(0, weight=1)
            self.rowconfigure(0, weight=1)

            self.subtabs = ttk.Notebook(self)
            self.subtabs.grid(row=0, column=0, sticky="nsew")
            self.work_tab = tk.Frame(self.subtabs, padx=12, pady=12)
            self.preset_tab = tk.Frame(self.subtabs, padx=12, pady=12)
            self.subtabs.add(self.work_tab, text="Dossier maken")
            self.subtabs.add(self.preset_tab, text="Presetregels")

            self._build_work_tab()
            self._build_preset_tab()

            for var in (
                self.mode_var,
                self.preset_var,
                self.include_bom_var,
                self.include_order_docs_var,
                self.include_offers_var,
                self.open_pdf_var,
            ):
                var.trace_add("write", lambda *_args: self._notify_options_changed())

            self._reload_presets()
            self._sync_mode()

        def _build_work_tab(self) -> None:
            form = self.work_tab
            form.columnconfigure(0, weight=0)
            form.columnconfigure(1, weight=0)
            form.columnconfigure(2, weight=1)
            form.rowconfigure(6, weight=1)

            tk.Label(form, text="Modus:").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=3)
            self.mode_combo = ttk.Combobox(
                form,
                textvariable=self.mode_var,
                values=[
                    self.MODE_WORKDOSSIER,
                    self.MODE_PER_PRODUCTION,
                    self.MODE_ALPHABETIC_SINGLE,
                ],
                state="readonly",
                width=26,
            )
            self.mode_combo.grid(row=0, column=1, sticky="w", pady=3)
            self.mode_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_mode())
            self._option_widgets.append(self.mode_combo)

            self.mode_info_frame = tk.Frame(form)
            self.mode_info_frame.grid(
                row=0,
                column=2,
                rowspan=2,
                sticky="w",
                padx=(18, 0),
                pady=3,
            )
            tk.Label(
                self.mode_info_frame,
                text="Kies hoe het dossier wordt opgebouwd.",
                anchor="w",
                justify="left",
                foreground="#263238",
            ).pack(anchor="w")
            tk.Label(
                self.mode_info_frame,
                text=(
                    "Werkdossier gebruikt de presetvolgorde. "
                    "Per productie maakt aparte PDF's. "
                    "Alfabetisch sorteert alle bestanden op naam."
                ),
                anchor="w",
                justify="left",
                foreground="#5D6670",
                wraplength=720,
            ).pack(anchor="w", pady=(2, 0))

            self.preset_label = tk.Label(form, text="Volgorde preset:")
            self.preset_label.grid(row=1, column=0, sticky="e", padx=(0, 6), pady=3)
            self.preset_combo = ttk.Combobox(
                form,
                textvariable=self.preset_var,
                state="readonly",
                width=26,
            )
            self.preset_combo.grid(row=1, column=1, sticky="w", pady=3)
            self.preset_combo.bind("<<ComboboxSelected>>", self._on_preset_selected)
            self._option_widgets.append(self.preset_combo)
            _HelpTooltip(
                self.preset_combo,
                "Blanco template start zonder categorieën. Elke categorie wordt alfabetisch gesorteerd, "
                "en de volgorde van de categorieën bepaalt de volgorde in het werkdossier.",
            )
            _HelpTooltip(
                self.mode_combo,
                "Werkdossier: gebruik een preset voor categorieën en volgorde. "
                "Aparte PDF per productie: maak per productie een PDF, met producties alfabetisch gerangschikt. "
                "PDF alfabetisch: geen preset, alle bestanden alfabetisch gesorteerd.",
            )

            options = tk.LabelFrame(form, text="Opties", labelanchor="n")
            options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0))
            options.columnconfigure(0, weight=1)
            self.include_order_docs_check = tk.Checkbutton(
                options,
                text="Bestelbonnen en standaardbonnen voor elke productie invoegen",
                variable=self.include_order_docs_var,
                anchor="w",
            )
            self.include_order_docs_check.grid(row=0, column=0, sticky="w", padx=6, pady=(4, 0))
            self.include_offers_check = tk.Checkbutton(
                options,
                text="Offerteaanvragen ook invoegen",
                variable=self.include_offers_var,
                anchor="w",
            )
            self.include_offers_check.grid(row=1, column=0, sticky="w", padx=6)
            self.open_pdf_check = tk.Checkbutton(
                options,
                text="PDF openen na combineren",
                variable=self.open_pdf_var,
                anchor="w",
            )
            self.open_pdf_check.grid(row=2, column=0, sticky="w", padx=6, pady=(0, 4))
            self._option_widgets.extend(
                [self.include_order_docs_check, self.include_offers_check, self.open_pdf_check]
            )

            export_box = tk.LabelFrame(form, text="Exportbestand", labelanchor="n")
            export_box.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 0))
            export_box.columnconfigure(1, weight=1)
            tk.Label(export_box, text="Naam:").grid(
                row=0, column=0, sticky="e", padx=(8, 6), pady=(6, 2)
            )
            ttk.Entry(
                export_box,
                textvariable=self.export_filename_var,
                state="readonly",
            ).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(6, 2))
            tk.Label(export_box, text="Map:").grid(
                row=1, column=0, sticky="e", padx=(8, 6), pady=(2, 6)
            )
            ttk.Entry(
                export_box,
                textvariable=self.export_folder_var,
                state="readonly",
            ).grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(2, 6))

            self.order_flow_frame = tk.Frame(form)
            self.order_flow_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(6, 0))
            self.order_flow_frame.columnconfigure(0, weight=1)
            self.order_flow_label = tk.Label(
                self.order_flow_frame,
                textvariable=self.order_flow_message_var,
                anchor="w",
                justify="left",
                foreground="#8A5A00",
                wraplength=740,
            )
            self.order_flow_label.grid(row=0, column=0, sticky="ew")
            self.prepare_orders_button = ttk.Button(
                self.order_flow_frame,
                text="Bestelbonnen klaarmaken",
                command=self._prepare_order_documents,
            )
            self.prepare_orders_button.grid(row=0, column=1, sticky="e", padx=(8, 0))
            self.order_flow_frame.grid_remove()

            preview = tk.LabelFrame(form, text="Voorbeeldvolgorde", labelanchor="n")
            preview.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
            preview.columnconfigure(0, weight=1)
            preview.rowconfigure(0, weight=1)
            columns = ("nr", "section", "production", "role", "filename")
            self.preview_tree = ttk.Treeview(
                preview,
                columns=columns,
                show="headings",
                height=12,
                selectmode="browse",
            )
            self.preview_tree.heading("nr", text="#")
            self.preview_tree.heading("section", text="Categorie")
            self.preview_tree.heading("production", text="Productie")
            self.preview_tree.heading("role", text="Type")
            self.preview_tree.heading("filename", text="Bestandsnaam")
            self.preview_tree.column("nr", width=44, minwidth=36, anchor="e", stretch=False)
            self.preview_tree.column("section", width=180, minwidth=120)
            self.preview_tree.column("production", width=180, minwidth=120)
            self.preview_tree.column("role", width=90, minwidth=80, stretch=False)
            self.preview_tree.column("filename", width=360, minwidth=180)
            self.preview_tree.grid(row=0, column=0, sticky="nsew", padx=(6, 0), pady=(6, 2))
            preview_scroll = ttk.Scrollbar(
                preview,
                orient="vertical",
                command=self.preview_tree.yview,
            )
            preview_scroll.grid(row=0, column=1, sticky="ns", padx=(6, 6), pady=(6, 2))
            self.preview_tree.configure(yscrollcommand=preview_scroll.set)
            tk.Label(
                preview,
                textvariable=self.preview_status_var,
                anchor="w",
                foreground="#5D6670",
            ).grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 6))

        def _build_preset_tab(self) -> None:
            form = self.preset_tab
            form.columnconfigure(1, weight=1)
            form.rowconfigure(3, weight=1)

            tk.Label(form, text="Preset:").grid(
                row=0, column=0, sticky="e", padx=(0, 6), pady=3
            )
            self.preset_editor_combo = ttk.Combobox(
                form,
                textvariable=self.preset_var,
                state="readonly",
                width=42,
            )
            self.preset_editor_combo.grid(row=0, column=1, sticky="ew", pady=3)
            self.preset_editor_combo.bind("<<ComboboxSelected>>", self._on_preset_selected)
            self._preset_action_widgets.append(self.preset_editor_combo)

            self.include_bom_check = tk.Checkbutton(
                form,
                text="Hoofdassembly PDF uit BOM/projectnaam eerst plaatsen",
                variable=self.include_bom_var,
                anchor="w",
            )
            self.include_bom_check.grid(row=1, column=1, sticky="w", pady=(4, 0))
            self._preset_action_widgets.append(self.include_bom_check)

            tk.Label(form, text="Secties:").grid(
                row=3, column=0, sticky="ne", padx=(0, 6), pady=(8, 3)
            )
            self.sections_editor = PdfWorkDossierSectionsEditor(form)
            self.sections_editor.grid(row=3, column=1, sticky="nsew", pady=(8, 3))

            help_label = tk.Label(
                form,
                text=(
                    "Sleep aan de ::-greep of gebruik Omhoog/Omlaag. "
                    "Overige producties vangt alles op dat nergens anders matcht."
                ),
                anchor="w",
                justify="left",
                foreground="#5D6670",
            )
            help_label.grid(row=4, column=1, sticky="ew", pady=(0, 6))

            actions = tk.Frame(form)
            actions.grid(row=5, column=1, sticky="ew", pady=(6, 0))
            save_button = tk.Button(
                actions,
                text="Preset bewaren als...",
                command=self.save_preset_as,
            )
            save_button.pack(side="left", padx=(0, 6))
            update_button = tk.Button(
                actions,
                text="Geselecteerde preset bijwerken",
                command=self.update_selected_preset,
            )
            update_button.pack(side="left", padx=(0, 6))
            delete_button = tk.Button(
                actions,
                text="Preset verwijderen",
                command=self.delete_selected_preset,
            )
            delete_button.pack(side="left", padx=(0, 6))
            self._preset_action_widgets.extend([save_button, update_button, delete_button])

        def _notify_options_changed(self) -> None:
            if callable(self.on_options_changed):
                self.on_options_changed()

        def _reload_presets(self, show_blank_template: bool = True) -> None:
            base_choices = [self.NO_PRESET_LABEL] if show_blank_template else []
            built_ins = [
                default_pdf_workdossier_preset(),
                tecno_art_pdf_workdossier_preset(),
            ]
            self._preset_map = {}
            for built_in in built_ins:
                self._preset_map[built_in.name] = built_in
                base_choices.append(built_in.name)
            for preset in self.presets_db.presets_sorted():
                label = preset.name
                if label in self._preset_map:
                    label = f"{preset.name} (opgeslagen)"
                self._preset_map[label] = preset
                base_choices.append(label)
            self.preset_combo.configure(values=base_choices)

            full_choices = [self.NO_PRESET_LABEL]
            for built_in in built_ins:
                full_choices.append(built_in.name)
            for preset in self.presets_db.presets_sorted():
                label = preset.name
                if label in self._preset_map:
                    label = f"{preset.name} (opgeslagen)"
                full_choices.append(label)
            self.preset_editor_combo.configure(values=full_choices)

            if self.preset_var.get() not in base_choices:
                if show_blank_template:
                    self.preset_var.set(self.NO_PRESET_LABEL)
                elif base_choices:
                    self.preset_var.set(base_choices[0])

        def _saved_preset_label(self, name: str) -> str:
            if name in {
                default_pdf_workdossier_preset().name,
                tecno_art_pdf_workdossier_preset().name,
            }:
                return f"{name} (opgeslagen)"
            return name

        def _sync_mode(self) -> None:
            is_workdossier = self.mode_var.get() == self.MODE_WORKDOSSIER
            readonly = "readonly" if is_workdossier and not self._busy else "disabled"
            if is_workdossier:
                self.preset_label.grid()
                self.preset_combo.grid()
                self._reload_presets(show_blank_template=False)
            else:
                self.preset_label.grid_remove()
                self.preset_combo.grid_remove()
                self._reload_presets(show_blank_template=True)
            self.preset_combo.configure(state=readonly)
            self.include_order_docs_check.configure(
                state="normal" if is_workdossier and not self._busy else "disabled"
            )
            self.include_offers_check.configure(
                state="normal"
                if is_workdossier and bool(self.include_order_docs_var.get()) and not self._busy
                else "disabled"
            )
            self.open_pdf_check.configure(state="normal" if not self._busy else "disabled")
            self.mode_combo.configure(state="readonly" if not self._busy else "disabled")
            self.preset_editor_combo.configure(state="readonly" if not self._busy else "disabled")
            self.include_bom_check.configure(state="normal" if not self._busy else "disabled")
            self.sections_editor.set_enabled(not self._busy)
            self._notify_options_changed()

        def _on_preset_selected(self, _event=None) -> None:
            preset = self._preset_map.get(self.preset_var.get())
            self.include_bom_var.set(0)
            section_blocks: List[PdfWorkDossierSection] = []
            if preset is not None:
                for section in preset.sections:
                    if section.include_bom_pdf:
                        self.include_bom_var.set(1)
                        continue
                    section_blocks.append(section)
            else:
                section_blocks = []
            self.sections_editor.set_sections(section_blocks)
            self._sync_mode()

        def _parse_preset_from_form(self, name: str = "Aangepast") -> Optional[PdfWorkDossierPreset]:
            sections: List[PdfWorkDossierSection] = []
            if self.include_bom_var.get():
                sections.append(PdfWorkDossierSection("Hoofdassembly", include_bom_pdf=True))
            sections.extend(self.sections_editor.get_sections())
            if not sections:
                return None
            return PdfWorkDossierPreset(
                name=name,
                sections=sections,
                include_unmatched=False,
                unmatched_section_name="Overige producties",
            )

        def save_preset_as(self) -> None:
            preset_name = simpledialog.askstring(
                "Preset bewaren",
                "Naam voor deze PDF-volgorde preset:",
                parent=self,
            )
            if not preset_name:
                return
            preset = self._parse_preset_from_form(preset_name.strip())
            if preset is None:
                messagebox.showwarning(
                    "Let op",
                    "Geef minstens een hoofdassembly-optie of een sectieregel op.",
                    parent=self,
                )
                return
            self.presets_db.upsert(preset)
            self.presets_db.save(PDF_WORKDOSSIER_PRESETS_DB_FILE)
            self._reload_presets()
            self.preset_var.set(self._saved_preset_label(preset.name))
            if callable(self.on_presets_changed):
                self.on_presets_changed()
                self.preset_var.set(self._saved_preset_label(preset.name))
            messagebox.showinfo("Preset bewaard", f"Preset '{preset.name}' is bewaard.", parent=self)

        def update_selected_preset(self) -> None:
            selected = self._preset_map.get(self.preset_var.get())
            if selected is None:
                messagebox.showwarning(
                    "Let op",
                    "Kies eerst een PDF-preset om bij te werken.",
                    parent=self,
                )
                return
            was_saved = self.presets_db.get(selected.name) is not None
            preset = self._parse_preset_from_form(selected.name)
            if preset is None:
                messagebox.showwarning(
                    "Let op",
                    "Geef minstens een hoofdassembly-optie of een sectieblok op.",
                    parent=self,
                )
                return
            self.presets_db.upsert(preset, old_name=selected.name)
            self.presets_db.save(PDF_WORKDOSSIER_PRESETS_DB_FILE)
            self._reload_presets()
            self.preset_var.set(self._saved_preset_label(preset.name))
            if callable(self.on_presets_changed):
                self.on_presets_changed()
                self.preset_var.set(self._saved_preset_label(preset.name))
            if was_saved:
                message = f"Preset '{preset.name}' is bijgewerkt."
            else:
                message = f"Preset '{preset.name}' is opgeslagen als bewerkbare preset."
            messagebox.showinfo("Preset bijgewerkt", message, parent=self)

        def delete_selected_preset(self) -> None:
            selected = self._preset_map.get(self.preset_var.get())
            if selected is None or self.presets_db.get(selected.name) is None:
                messagebox.showwarning(
                    "Let op",
                    "Kies eerst een opgeslagen PDF-preset om te verwijderen.",
                    parent=self,
                )
                return
            if not messagebox.askyesno(
                "Preset verwijderen",
                f"PDF-preset '{selected.name}' verwijderen?",
                parent=self,
            ):
                return
            if self.presets_db.remove(selected.name):
                self.presets_db.save(PDF_WORKDOSSIER_PRESETS_DB_FILE)
            self._reload_presets()
            self.preset_var.set(self.NO_PRESET_LABEL)
            self._on_preset_selected()
            if callable(self.on_presets_changed):
                self.on_presets_changed()

        def build_options(self) -> Dict[str, object]:
            preset = None
            if self.mode_var.get() == self.MODE_WORKDOSSIER:
                preset = self._parse_preset_from_form("Aangepast")
            return {
                "mode": self.mode_var.get(),
                "preset": preset,
                "include_order_documents": bool(self.include_order_docs_var.get()),
                "include_offers": bool(self.include_offers_var.get()),
                "open_pdf": bool(self.open_pdf_var.get()),
            }

        def set_busy(self, busy: bool) -> None:
            self._busy = bool(busy)
            state = "disabled" if self._busy else "normal"
            for widget in self._preset_action_widgets:
                try:
                    widget.configure(state=state)
                except tk.TclError:
                    pass
            self._sync_mode()

        def select_work_tab(self) -> None:
            try:
                self.subtabs.select(self.work_tab)
            except tk.TclError:
                pass

        def select_preset_tab(self) -> None:
            try:
                self.subtabs.select(self.preset_tab)
            except tk.TclError:
                pass

        def set_export_info(self, filename: str, folder: str) -> None:
            self.export_filename_var.set(filename)
            self.export_folder_var.set(folder)

        def clear_preview(self, message: str = "Geen voorbeeld beschikbaar.") -> None:
            for item_id in self.preview_tree.get_children():
                self.preview_tree.delete(item_id)
            self.preview_status_var.set(message)

        def set_preview_items(self, items: Sequence[object]) -> None:
            for item_id in self.preview_tree.get_children():
                self.preview_tree.delete(item_id)
            for index, item in enumerate(items, start=1):
                path = _to_str(getattr(item, "path", ""))
                section = _to_str(getattr(item, "section_name", "")).strip() or "-"
                production = _to_str(getattr(item, "production", "")).strip() or "-"
                role = _to_str(getattr(item, "role", "")).strip().lower()
                role_label = self.ROLE_LABELS.get(role, role or "-")
                filename = os.path.basename(path) if path else "-"
                filename_key = filename.casefold()
                if role == "order":
                    if filename_key.startswith("offerteaanvraag"):
                        role_label = "Offerte"
                    elif filename_key.startswith("standaard"):
                        role_label = "Standaardbon"
                    else:
                        role_label = "Bestelbon"
                self.preview_tree.insert(
                    "",
                    "end",
                    values=(index, section, production, role_label, filename),
                )
            count = len(items)
            suffix = "" if count == 1 else "en"
            self.preview_status_var.set(f"{count} bestand{suffix} in de mergevolgorde.")

        def set_order_flow_message(self, message: str, *, show_button: bool = False) -> None:
            clean = _to_str(message).strip()
            self.order_flow_message_var.set(clean)
            if clean:
                self.order_flow_frame.grid()
            else:
                self.order_flow_frame.grid_remove()
            if show_button:
                self.prepare_orders_button.grid()
            else:
                self.prepare_orders_button.grid_remove()

        def _prepare_order_documents(self) -> None:
            if callable(self.on_prepare_orders):
                self.on_prepare_orders()

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
            self.title(f"{APP_NAME} {APP_VERSION}")
            
            self.minsize(1024, 720)
            self._schedule_window_maximize()
            
            # Schedule icon loading after window is displayed (to avoid startup lag)
            self.after_idle(self._load_window_icon)

            self.db = SuppliersDB.load(SUPPLIERS_DB_FILE)
            self.client_db = ClientsDB.load(CLIENTS_DB_FILE)
            self.delivery_db = DeliveryAddressesDB.load(DELIVERY_DB_FILE)
            self.order_presets_db = OrderPresetsDB.load(ORDER_PRESETS_DB_FILE)
            self.pdf_workdossier_presets_db = PdfWorkDossierPresetsDB.load(
                PDF_WORKDOSSIER_PRESETS_DB_FILE
            )

            self.settings = AppSettings.load()
            self._suspend_save = False
            raw_en1090 = getattr(self.settings, "en1090_preferences", {}) or {}
            self.en1090_preferences: Dict[str, bool] = {}
            for key, value in raw_en1090.items():
                norm = normalize_en1090_key(key)
                if norm:
                    self.en1090_preferences[norm] = bool(value)

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
            self.document_filename_profile_var = tk.StringVar(
                master=self,
                value=getattr(self.settings, "document_filename_profile", "standard"),
            )
            self._document_filename_profile_choices: List[Tuple[str, str]] = [
                ("Standaard", "standard"),
                ("Kort (BB-123)", "short"),
                ("Compact (BB123)", "compact"),
                ("Aangepast", "custom"),
            ]
            self._document_filename_profile_value_to_display = {
                value: label for label, value in self._document_filename_profile_choices
            }
            self._document_filename_profile_display_to_value = {
                label: value for label, value in self._document_filename_profile_choices
            }
            self.document_filename_profile_display_var = tk.StringVar(
                master=self,
                value=self._document_filename_profile_value_to_display.get(
                    self.document_filename_profile_var.get().strip() or "standard",
                    "Standaard",
                ),
            )
            self.document_filename_show_doc_type_var = tk.IntVar(
                master=self,
                value=1 if getattr(self.settings, "document_filename_show_doc_type", True) else 0,
            )
            self.document_filename_show_doc_number_var = tk.IntVar(
                master=self,
                value=1
                if getattr(self.settings, "document_filename_show_doc_number", True)
                else 0,
            )
            self.document_filename_show_context_var = tk.IntVar(
                master=self,
                value=1 if getattr(self.settings, "document_filename_show_context", True) else 0,
            )
            self.document_filename_show_date_var = tk.IntVar(
                master=self,
                value=1 if getattr(self.settings, "document_filename_show_date", True) else 0,
            )
            self.document_filename_compact_doc_number_var = tk.IntVar(
                master=self,
                value=1
                if getattr(self.settings, "document_filename_compact_doc_number", False)
                else 0,
            )
            self.document_filename_separator_var = tk.StringVar(
                master=self,
                value=getattr(self.settings, "document_filename_separator", "underscore"),
            )
            self.document_filename_preview_var = tk.StringVar(master=self, value="")
            self.document_filename_details_open_var = tk.IntVar(
                master=self,
                value=1
                if (self.document_filename_profile_var.get().strip().lower() == "custom")
                else 0,
            )
            self.document_display_compact_doc_number_var = tk.IntVar(
                master=self,
                value=1
                if getattr(self.settings, "document_display_compact_doc_number", False)
                else 0,
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
            self.en1090_enabled_var = tk.IntVar(
                master=self,
                value=1 if getattr(self.settings, "en1090_enabled", True) else 0,
            )
            raw_en1090_note = getattr(self.settings, "en1090_note", EN1090_NOTE_TEXT)
            note_text = (
                EN1090_NOTE_TEXT
                if raw_en1090_note is None
                else _to_str(raw_en1090_note)
            ).replace("\r\n", "\n")
            self.en1090_note_var = tk.StringVar(master=self, value=note_text)
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
            self._pdf_action_running = False
            self._pdf_preview_after_id: Optional[str] = None
            self._pdf_order_document_root = ""
            self._pdf_generated_order_documents: List[Dict[str, object]] = []
            self._pdf_order_context_signature: Tuple[str, str, str, str, str] | None = None

            for var in (
                self.source_folder_var,
                self.dest_folder_var,
                self.project_number_var,
                self.project_name_var,
                self.export_name_custom_prefix_text,
                self.export_name_custom_suffix_text,
                self.document_filename_profile_var,
                self.document_filename_separator_var,
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
                self.document_filename_show_doc_type_var,
                self.document_filename_show_doc_number_var,
                self.document_filename_show_context_var,
                self.document_filename_show_date_var,
                self.document_filename_compact_doc_number_var,
                self.document_display_compact_doc_number_var,
                self.bundle_latest_var,
                self.bundle_dry_run_var,
                self.autofill_custom_bom_var,
            ):
                var.trace_add("write", self._save_settings)

            for var in (
                self.document_filename_profile_var,
                self.document_filename_show_doc_type_var,
                self.document_filename_show_doc_number_var,
                self.document_filename_show_context_var,
                self.document_filename_show_date_var,
                self.document_filename_compact_doc_number_var,
                self.document_filename_separator_var,
            ):
                var.trace_add("write", self._update_document_filename_preview)

            self.document_filename_profile_var.trace_add(
                "write", self._sync_document_filename_profile_display
            )
            self.document_filename_profile_var.trace_add(
                "write", self._refresh_document_filename_controls
            )

            self.zip_var.trace_add("write", self._update_zip_per_finish_var)
            self.zip_finish_var.trace_add("write", self._update_zip_per_finish_var)
            for var in (
                self.source_folder_var,
                self.dest_folder_var,
                self.project_number_var,
                self.project_name_var,
            ):
                var.trace_add("write", self._schedule_pdf_workdossier_preview)
            self._update_zip_per_finish_var()

            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)

            content = tk.Frame(self)
            content.grid(row=0, column=0, sticky="nsew")
            content.grid_columnconfigure(0, weight=1)

            content.rowconfigure(5, weight=1)

            tabs_wrapper = tk.Frame(content)
            tabs_wrapper.grid(row=5, column=0, sticky="nsew", padx=8, pady=(12, 0))

            tabs_background = (
                style.lookup("TNotebook", "background")
                or style.lookup("TFrame", "background")
                or self.cget("background")
            )

            tabs_container = tk.Frame(tabs_wrapper, background=tabs_background)
            tabs_container.pack(fill="both", expand=True)

            self.nb = ttk.Notebook(tabs_container)
            self.nb.pack(fill="both", expand=True)
            self._last_selected_notebook_tab = ""
            self.nb.bind("<<NotebookTabChanged>>", self._handle_tab_changed, add="+")
            self.custom_bom_tab: Optional[BOMCustomTab] = None
            self._custom_bom_needs_sync = False
            self._bom_load_in_progress = False
            self._custom_bom_placeholder = tk.Frame(
                self.nb, background=tabs_background
            )
            tk.Label(
                self._custom_bom_placeholder,
                text=(
                    "De Custom BOM-tab wordt geladen zodra u deze opent. Dit maakt"
                    " het opstarten sneller."
                ),
                wraplength=360,
                justify="left",
                background=tabs_background,
            ).pack(padx=24, pady=24, anchor="w")
            main = tk.Frame(self.nb)
            main.configure(padx=12, pady=12)
            self.nb.add(main, text="Main")
            self.nb.add(self._custom_bom_placeholder, text="Custom BOM")
            self.manual_order_tab = ManualOrderTab(
                self.nb,
                suppliers_db=self.db,
                delivery_db=self.delivery_db,
                clients_db=self.client_db,
                project_number_var=self.project_number_var,
                project_name_var=self.project_name_var,
                client_var=getattr(self, "client_var", None),
                dest_folder_var=self.dest_folder_var,
                on_export=self._export_manual_order,
                on_manage_clients=lambda: self.nb.select(self.clients_frame),
                on_manage_suppliers=lambda: self.nb.select(self.suppliers_frame),
                on_manage_deliveries=lambda: self.nb.select(self.delivery_frame),
                document_name_builder=self._build_document_export_basename,
            )
            self.nb.add(self.manual_order_tab, text="Bestelbon-editor")
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
            self._opticutter_dirty = False
            self._opticutter_needs_refresh = False
            self._opticutter_analysis_stale = True
            self._opticutter_refresh_generation = 0
            self._opticutter_analysis_refresh_running = False
            self._opticutter_toast_window = None
            self._opticutter_toast_after_id = None
            self.opticutter_kerf_var.trace_add(
                "write", self._on_opticutter_kerf_change
            )
            self.opticutter_custom_stock_var.trace_add(
                "write", self._on_opticutter_custom_stock_change
            )

            self.opticutter_update_status_var = tk.StringVar(
                master=self.opticutter_frame,
                value="Geen wijzigingen",
            )
            self.opticutter_update_status_label = tk.Label(
                controls_frame,
                textvariable=self.opticutter_update_status_var,
                anchor="e",
                fg="#6B7280",
            )
            self.opticutter_update_status_label.pack(anchor="e", pady=(8, 0))

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
                summary_container,
                text="Profiel, materiaal en productie",
                padx=8,
                pady=8,
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
                    stretch=True,
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
                section_frame = tk.LabelFrame(
                    scenarios_frame,
                    text=section_title,
                    padx=8,
                    pady=8,
                )
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
                        stretch=True,
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
            self.preset_rules_frame = PresetRulesManagerFrame(
                self.nb,
                self.order_presets_db,
                self.client_db,
                self.db,
                self.delivery_db,
                on_change=self._on_preset_rules_change,
            )
            self.preset_rules_frame.configure(padx=12, pady=12)
            self.nb.add(self.preset_rules_frame, text="Presetregels")

            self.pdf_workdossier_frame = tk.Frame(self.nb)
            self.pdf_workdossier_frame.configure(padx=12, pady=12)
            self._build_pdf_workdossier_tab()
            self.nb.add(self.pdf_workdossier_frame, text="PDF werkdossier")

            self.settings_frame = SettingsFrame(self.nb, self)
            self.settings_frame.configure(padx=12, pady=12)
            self.nb.add(self.settings_frame, text="⚙ Settings")

            main_body = tk.Frame(main)
            main_body.pack(fill="both", expand=True)
            main_footer = tk.Frame(main)
            main_footer.pack(fill="x", side="bottom")

            # Top folders
            top = tk.Frame(main_body)
            top.pack(fill="x", padx=8, pady=6)
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
            filters_row = tk.Frame(main_body)
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
            export_name_inner.grid(row=0, column=0, sticky="nsew", padx=8, pady=4)
            export_name_inner.grid_columnconfigure(0, weight=1, uniform="export_name")
            export_name_inner.grid_columnconfigure(1, weight=1, uniform="export_name")
            general_export_name_frame = tk.LabelFrame(
                export_name_inner,
                text="Andere exportbestanden",
                labelanchor="n",
            )
            general_export_name_frame.grid(
                row=0, column=0, sticky="nsew", padx=(0, 6)
            )
            general_export_name_inner = tk.Frame(general_export_name_frame)
            general_export_name_inner.pack(fill="x", padx=6, pady=6)
            document_name_frame = tk.LabelFrame(
                export_name_inner,
                text="Bestelbon / offerte",
                labelanchor="n",
            )
            document_name_frame.grid(
                row=0, column=1, sticky="nsew", padx=(6, 0)
            )
            document_name_inner = tk.Frame(document_name_frame)
            document_name_inner.pack(fill="x", padx=6, pady=6)

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
                general_export_name_inner,
                text="Datumprefix (YYYYMMDD-)",
                variable=self.export_date_prefix_var,
                anchor="w",
            ).pack(anchor="w", pady=2)
            tk.Checkbutton(
                general_export_name_inner,
                text="Datumsuffix (-YYYYMMDD)",
                variable=self.export_date_suffix_var,
                anchor="w",
            ).pack(anchor="w", pady=2)
            prefix_row = tk.Frame(general_export_name_inner)
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
            suffix_row = tk.Frame(general_export_name_inner)
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
            profile_row = tk.Frame(document_name_inner)
            profile_row.pack(fill="x")
            tk.Label(
                profile_row,
                text="Profiel:",
                anchor="w",
            ).pack(side="left")
            self.document_filename_profile_combo = ttk.Combobox(
                profile_row,
                textvariable=self.document_filename_profile_display_var,
                values=[
                    label for label, _value in self._document_filename_profile_choices
                ],
                state="readonly",
                width=18,
            )
            self.document_filename_profile_combo.pack(side="left", padx=(8, 0))
            self.document_filename_profile_combo.bind(
                "<<ComboboxSelected>>",
                self._on_document_filename_profile_selected,
            )

            preview_row = tk.Frame(document_name_inner)
            preview_row.pack(fill="x", pady=(8, 0))
            tk.Label(preview_row, text="Voorbeeld:").pack(anchor="w")
            tk.Entry(
                preview_row,
                textvariable=self.document_filename_preview_var,
                state="readonly",
            ).pack(fill="x", pady=(2, 0))
            tk.Checkbutton(
                document_name_inner,
                text="Compact documentnummer tonen in PDF/Excel",
                variable=self.document_display_compact_doc_number_var,
                anchor="w",
                justify="left",
                wraplength=280,
            ).pack(anchor="w", pady=(8, 0))

            self.document_filename_details_toggle_row = tk.Frame(document_name_inner)
            self.document_filename_details_toggle_text_var = tk.StringVar(master=self, value="")
            tk.Button(
                self.document_filename_details_toggle_row,
                textvariable=self.document_filename_details_toggle_text_var,
                command=self._toggle_document_filename_details,
                relief="flat",
                anchor="w",
                cursor="hand2",
            ).pack(anchor="w")

            self.document_filename_details_frame = tk.Frame(document_name_inner)
            custom_document_options = tk.Frame(self.document_filename_details_frame)
            custom_document_options.pack(fill="x")
            self._document_filename_custom_widgets = []

            def _add_doc_custom_checkbutton(text: str, variable: tk.IntVar) -> None:
                widget = tk.Checkbutton(
                    custom_document_options,
                    text=text,
                    variable=variable,
                    anchor="w",
                    justify="left",
                )
                widget.pack(anchor="w")
                self._document_filename_custom_widgets.append(widget)

            _add_doc_custom_checkbutton(
                "Documenttype tonen",
                self.document_filename_show_doc_type_var,
            )
            _add_doc_custom_checkbutton(
                "Documentnummer tonen",
                self.document_filename_show_doc_number_var,
            )
            _add_doc_custom_checkbutton(
                "Productie / afwerking tonen",
                self.document_filename_show_context_var,
            )
            _add_doc_custom_checkbutton(
                "Datum tonen",
                self.document_filename_show_date_var,
            )
            _add_doc_custom_checkbutton(
                "Compact documentnummer",
                self.document_filename_compact_doc_number_var,
            )

            separator_row = tk.Frame(custom_document_options)
            separator_row.pack(anchor="w", pady=(4, 0))
            tk.Label(separator_row, text="Scheiding:").pack(side="left")
            for text, value in (
                ("_", "underscore"),
                ("-", "dash"),
                ("geen", "none"),
            ):
                widget = tk.Radiobutton(
                    separator_row,
                    text=text,
                    value=value,
                    variable=self.document_filename_separator_var,
                    anchor="w",
                )
                widget.pack(side="left", padx=(8, 0))
                self._document_filename_custom_widgets.append(widget)
            # Legacy options moved to settings tab

            # BOM controls
            bf = tk.Frame(main_body)
            bf.pack(fill="x", padx=8, pady=6)
            tk.Button(bf, text="Laad BOM (CSV/Excel)", command=self._load_bom).pack(side="left", padx=6)
            tk.Button(
                bf,
                text="Custom BOM",
                command=self._select_custom_bom_tab,
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
            treef = tk.Frame(main_body)
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

            # Actions and status
            ttk.Separator(main_footer, orient="horizontal").pack(fill="x", padx=8)
            footer_row = tk.Frame(main_footer)
            footer_row.pack(fill="x", padx=8, pady=(8, 8))

            act = tk.Frame(footer_row)
            act.pack(side="left", fill="x", expand=True)
            button_style = dict(
                bg=MANUFACT_BRAND_COLOR,
                activebackground="#F4C46C",
                fg="black",
                activeforeground="black",
                highlightthickness=0,
                highlightbackground=MANUFACT_BRAND_COLOR,
                highlightcolor=MANUFACT_BRAND_COLOR,
                takefocus=False,
                padx=14,
                pady=6,
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
                act, text="PDF combineren", command=self._combine_pdf, **button_style
            ).pack(side="left", padx=6)

            self.status_var = tk.StringVar(value="Klaar voor gebruik")
            status_row = tk.Frame(footer_row)
            status_row.pack(side="right", fill="x")
            tk.Label(
                status_row,
                textvariable=self.status_var,
                anchor="e",
            ).pack(side="left", padx=(0, 20), fill="x", expand=True)
            tk.Label(
                status_row,
                text=f"Versie {APP_VERSION}",
                anchor="e",
            ).pack(side="right")
            self._refresh_document_filename_controls()
            self._sync_document_filename_profile_display()
            self._update_document_filename_preview()
            self._save_settings()

        def _current_client(self) -> Optional[Client]:
            if not hasattr(self, "client_var"):
                return None
            client_name = strip_favorite_marker(_to_str(self.client_var.get())).strip()
            if not client_name:
                return None
            return self.client_db.get(client_name)

        def _resolve_delivery_choice(
            self,
            delivery_display: str,
            client: Optional[Client] = None,
        ) -> Optional[DeliveryAddress]:
            clean = strip_favorite_marker(_to_str(delivery_display)).strip()
            if not clean or clean == "Geen":
                return None
            if SupplierSelectionFrame._is_client_delivery_choice(clean):
                selected_client = client if client is not None else self._current_client()
                if selected_client is None:
                    return None
                name = _to_str(selected_client.name).strip() or "Opdrachtgever"
                address = _to_str(selected_client.address).strip() or None
                if not name and not address:
                    return None
                return DeliveryAddress(name=name, address=address)
            if clean in (
                "Bestelling wordt opgehaald",
                "Leveradres wordt nog meegedeeld",
            ):
                return DeliveryAddress(name=clean)
            delivery = self.delivery_db.get(clean)
            if delivery is not None:
                return delivery
            return DeliveryAddress(name=clean)

        def _export_manual_order(self, payload: Dict[str, object]) -> None:
            from tkinter import messagebox, filedialog

            manual_tab = getattr(self, "manual_order_tab", None)
            if not payload:
                return

            doc_type = _to_str(payload.get("doc_type")).strip() or "Bestelbon"
            doc_number_raw = payload.get("doc_number", "")
            doc_number = _normalize_doc_number(doc_number_raw, doc_type)
            doc_number_display = self._format_document_display_number(
                doc_type,
                doc_number,
            )
            if manual_tab is not None:
                try:
                    manual_tab.set_doc_number(doc_number)
                except Exception:
                    pass

            supplier_display = _to_str(payload.get("supplier")).strip()
            supplier_name_clean = strip_favorite_marker(supplier_display).strip()
            supplier: Optional[Supplier] = None
            if supplier_name_clean and supplier_name_clean.lower() not in {"geen", "(geen)"}:
                for sup in getattr(self.db, "suppliers", []):
                    if _to_str(sup.supplier).strip().lower() == supplier_name_clean.lower():
                        supplier = sup
                        break

            client_display = _to_str(payload.get("client")).strip()
            client_name_clean = strip_favorite_marker(client_display).strip()
            client: Optional[Client] = None
            if client_name_clean:
                client = self.client_db.get(client_name_clean)
            if client is None:
                client = self._current_client()

            delivery_display = _to_str(payload.get("delivery")).strip()
            delivery = self._resolve_delivery_choice(delivery_display, client)

            column_layout_raw = payload.get("column_layout")
            column_layout: List[Dict[str, object]] = []
            if isinstance(column_layout_raw, list):
                for column in column_layout_raw:
                    if not isinstance(column, dict):
                        continue
                    key = _to_str(column.get("key")).strip()
                    if not key:
                        continue
                    label = _to_str(column.get("label")).strip() or key
                    justify = _to_str(column.get("justify")).strip().lower() or "left"
                    wrap = bool(column.get("wrap"))
                    numeric = bool(column.get("numeric"))
                    integer_flag = bool(column.get("integer"))
                    weight_raw = column.get("weight")
                    try:
                        weight = float(weight_raw)
                    except Exception:
                        weight = None
                    entry: Dict[str, object] = {
                        "key": key,
                        "label": label,
                        "justify": justify,
                        "wrap": wrap,
                        "numeric": numeric,
                    }
                    if integer_flag:
                        entry["integer"] = True
                    if weight is not None and math.isfinite(weight) and weight > 0:
                        entry["weight"] = weight
                    if column.get("total_weight"):
                        entry["total_weight"] = True
                    column_layout.append(entry)

            items = list(payload.get("items") or [])
            if not items:
                messagebox.showwarning(
                    "Geen gegevens",
                    "Voeg minstens één regel toe voordat je exporteert.",
                    parent=manual_tab or self,
                )
                return

            remark_text = _to_str(payload.get("remark")).strip()
            context_label = (
                _to_str(payload.get("context_label")).strip()
                or ManualOrderTab.DEFAULT_CONTEXT_LABEL
            )
            context_kind = _to_str(payload.get("context_kind")).strip() or "document"

            dest = self.dest_folder_var.get().strip()
            if not dest:
                dest = filedialog.askdirectory(
                    parent=manual_tab or self,
                    title="Kies exportmap voor handmatige bestelbon",
                )
                if not dest:
                    return
                self.dest_folder_var.set(dest)
                self.dest_folder = dest

            try:
                os.makedirs(dest, exist_ok=True)
            except Exception as exc:
                messagebox.showerror(
                    "Fout",
                    f"Kon exportmap niet maken:\n{exc}",
                    parent=manual_tab or self,
                )
                return

            project_number = self.project_number_var.get().strip()
            project_name = self.project_name_var.get().strip()
            filename_base = self._build_document_export_basename(
                doc_type,
                doc_number,
                context_label or project_name or ManualOrderTab.DEFAULT_CONTEXT_LABEL,
            )
            excel_requested = f"{filename_base}.xlsx"
            pdf_requested = f"{filename_base}.pdf"
            excel_filename = _fit_filename_within_path(dest, excel_requested)
            pdf_filename = _fit_filename_within_path(dest, pdf_requested)
            excel_path = os.path.join(dest, excel_filename)
            pdf_path = os.path.join(dest, pdf_filename)

            company_info = {
                "name": client.name if client else "",
                "address": client.address if client else "",
                "vat": client.vat if client else "",
                "email": client.email if client else "",
                "website": client.website if client else "",
                "accent_color": client.accent_color if client else "",
                "logo_path": client.logo_path if client else "",
                "logo_crop": client.logo_crop if client else None,
            }

            footer_note_text = self.footer_note_var.get().strip() or DEFAULT_FOOTER_NOTE
            total_weight = payload.get("total_weight")
            if not isinstance(total_weight, (int, float)):
                try:
                    total_weight = float(total_weight)
                except Exception:
                    total_weight = None

            try:
                write_order_excel(
                    excel_path,
                    items,
                    company_info,
                    supplier,
                    delivery,
                    doc_type,
                    doc_number_display or None,
                    project_number=project_number or None,
                    project_name=project_name or None,
                    context_label=context_label,
                    context_kind=context_kind,
                    order_remark=remark_text or None,
                    total_weight_kg=total_weight if isinstance(total_weight, (int, float)) else None,
                    en1090_required=False,
                    en1090_note=None,
                    column_layout=column_layout or None,
                )
                generate_pdf_order_platypus(
                    pdf_path,
                    company_info,
                    supplier,
                    context_label,
                    items,
                    doc_type=doc_type,
                    doc_number=doc_number_display or None,
                    footer_note=footer_note_text,
                    delivery=delivery,
                    project_number=project_number or None,
                    project_name=project_name or None,
                    label_kind=context_kind,
                    order_remark=remark_text or None,
                    total_weight_kg=total_weight if isinstance(total_weight, (int, float)) else None,
                    en1090_required=False,
                    en1090_note=None,
                    include_bruto_note=getattr(self, "include_bruto_note_var", None) and bool(self.include_bruto_note_var.get()),
                    column_layout=column_layout or None,
                )
            except Exception as exc:
                messagebox.showerror(
                    "Fout",
                    f"Kon handmatige {doc_type.lower()} niet opslaan:\n{exc}",
                    parent=manual_tab or self,
                )
                return

            self.status_var.set(
                f"Handmatige {doc_type.lower()} opgeslagen in {dest}"
            )
            messagebox.showinfo(
                "Klaar",
                "Bestanden opgeslagen:\n"
                f"- {excel_filename}\n"
                f"- {pdf_filename}",
                parent=manual_tab or self,
            )
            # Open the destination folder
            try:
                os.startfile(dest)
            except Exception:
                pass

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
            manual_tab = getattr(self, "manual_order_tab", None)
            if manual_tab is not None:
                try:
                    manual_tab.refresh_data()
                except Exception:
                    pass
            presets_tab = getattr(self, "preset_rules_frame", None)
            if presets_tab is not None:
                try:
                    presets_tab.refresh_data()
                except Exception:
                    pass

        def _on_preset_rules_change(self):
            self.status_var.set("Presetregels bijgewerkt.")
            sel = getattr(self, "sel_frame", None)
            if sel is not None:
                try:
                    if sel.winfo_exists():
                        sel.presets_db = self.order_presets_db
                except Exception:
                    pass

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
            self.settings.document_filename_profile = (
                self.document_filename_profile_var.get().strip() or "standard"
            )
            self.settings.document_filename_show_doc_type = bool(
                self.document_filename_show_doc_type_var.get()
            )
            self.settings.document_filename_show_doc_number = bool(
                self.document_filename_show_doc_number_var.get()
            )
            self.settings.document_filename_show_context = bool(
                self.document_filename_show_context_var.get()
            )
            self.settings.document_filename_show_date = bool(
                self.document_filename_show_date_var.get()
            )
            self.settings.document_filename_compact_doc_number = bool(
                self.document_filename_compact_doc_number_var.get()
            )
            self.settings.document_filename_separator = (
                self.document_filename_separator_var.get().strip() or "underscore"
            )
            self.settings.document_display_compact_doc_number = bool(
                self.document_display_compact_doc_number_var.get()
            )
            self.settings.bundle_latest = bool(self.bundle_latest_var.get())
            self.settings.bundle_dry_run = bool(self.bundle_dry_run_var.get())
            self.settings.autofill_custom_bom = bool(
                self.autofill_custom_bom_var.get()
            )
            self.settings.footer_note = self.footer_note_var.get().replace("\r\n", "\n")
            self.settings.en1090_enabled = bool(self.en1090_enabled_var.get())
            self.settings.en1090_note = self.en1090_note_var.get().replace("\r\n", "\n")
            for ext in self.settings.file_extensions:
                var = self.extension_vars.get(ext.key)
                if var is not None:
                    ext.enabled = bool(var.get())
            self._sync_en1090_preferences()
            try:
                self.settings.save()
            except Exception as exc:
                print(f"Kon instellingen niet opslaan: {exc}", file=sys.stderr)

        def _document_filename_settings_kwargs(self) -> Dict[str, object]:
            return {
                "profile": self.document_filename_profile_var.get().strip() or "standard",
                "show_doc_type": bool(self.document_filename_show_doc_type_var.get()),
                "show_doc_number": bool(self.document_filename_show_doc_number_var.get()),
                "show_context": bool(self.document_filename_show_context_var.get()),
                "show_date": bool(self.document_filename_show_date_var.get()),
                "compact_doc_number": bool(
                    self.document_filename_compact_doc_number_var.get()
                ),
                "separator": self.document_filename_separator_var.get().strip()
                or "underscore",
            }

        def _build_document_export_basename(
            self,
            doc_type: str,
            doc_number: str | None,
            context_label: str | None,
            export_date: str | None = None,
            *,
            extra_context_label: str | None = None,
        ) -> str:
            return build_document_export_basename(
                doc_type,
                doc_number,
                context_label,
                export_date,
                extra_context_label=extra_context_label,
                **self._document_filename_settings_kwargs(),
            )

        def _format_document_display_number(
            self,
            doc_type: str,
            doc_number: str | None,
        ) -> str:
            return format_document_number_for_display(
                doc_number,
                doc_type,
                compact=bool(self.document_display_compact_doc_number_var.get()),
            )

        def _sync_document_filename_profile_display(self, *_args) -> None:
            display_var = getattr(self, "document_filename_profile_display_var", None)
            if display_var is None:
                return
            value_to_display = getattr(
                self, "_document_filename_profile_value_to_display", {}
            )
            current_value = self.document_filename_profile_var.get().strip() or "standard"
            desired_display = value_to_display.get(current_value, "Standaard")
            if display_var.get() != desired_display:
                display_var.set(desired_display)

        def _on_document_filename_profile_selected(self, _evt=None) -> None:
            display_var = getattr(self, "document_filename_profile_display_var", None)
            display_to_value = getattr(
                self, "_document_filename_profile_display_to_value", {}
            )
            if display_var is None:
                return
            selected_display = display_var.get().strip()
            selected_value = display_to_value.get(selected_display, "standard")
            details_open_var = getattr(self, "document_filename_details_open_var", None)
            if details_open_var is not None and selected_value == "custom":
                details_open_var.set(1)
            if self.document_filename_profile_var.get().strip() != selected_value:
                self.document_filename_profile_var.set(selected_value)

        def _toggle_document_filename_details(self) -> None:
            details_open_var = getattr(self, "document_filename_details_open_var", None)
            if details_open_var is None:
                return
            current = 1 if bool(details_open_var.get()) else 0
            details_open_var.set(0 if current else 1)
            self._refresh_document_filename_controls()

        def _refresh_document_filename_controls(self, *_args) -> None:
            enabled = self.document_filename_profile_var.get().strip().lower() == "custom"
            state = "normal" if enabled else "disabled"
            for widget in getattr(self, "_document_filename_custom_widgets", []):
                try:
                    widget.configure(state=state)
                except Exception:
                    pass

            details_open_var = getattr(self, "document_filename_details_open_var", None)
            details_toggle_row = getattr(
                self, "document_filename_details_toggle_row", None
            )
            details_frame = getattr(self, "document_filename_details_frame", None)
            details_toggle_text_var = getattr(
                self, "document_filename_details_toggle_text_var", None
            )

            details_open = bool(details_open_var.get()) if details_open_var is not None else enabled

            if details_toggle_text_var is not None:
                details_toggle_text_var.set(
                    "Naamopties verbergen ▴" if enabled and details_open else "Naamopties tonen ▾"
                )

            if details_toggle_row is not None:
                if enabled:
                    if not details_toggle_row.winfo_manager():
                        details_toggle_row.pack(fill="x", pady=(8, 0))
                elif details_toggle_row.winfo_manager():
                    details_toggle_row.pack_forget()

            if details_frame is not None:
                if enabled and details_open:
                    if not details_frame.winfo_manager():
                        details_frame.pack(fill="x", pady=(6, 0))
                elif details_frame.winfo_manager():
                    details_frame.pack_forget()

        def _update_document_filename_preview(self, *_args) -> None:
            preview = self._build_document_export_basename(
                "Bestelbon",
                "BB-123",
                "Laser",
                datetime.date.today().strftime("%Y-%m-%d"),
            )
            self.document_filename_preview_var.set(f"{preview}.pdf")
            manual_tab = getattr(self, "manual_order_tab", None)
            if manual_tab is not None and hasattr(manual_tab, "_update_doc_name_preview"):
                try:
                    manual_tab._update_doc_name_preview()
                except Exception:
                    pass

        def _sync_en1090_preferences(self) -> None:
            try:
                mapping = {
                    key: bool(value) for key, value in self.en1090_preferences.items()
                }
            except AttributeError:
                mapping = {}
            self.settings.en1090_preferences = mapping

        def _get_en1090_preference(self, metadata: Dict[str, str]) -> bool:
            if metadata.get("kind") not in {"production", "opticutter"}:
                return False
            identifier = metadata.get("identifier") or metadata.get("display") or ""
            norm = normalize_en1090_key(identifier)
            if not norm:
                return False
            existing = self.en1090_preferences.get(norm)
            if existing is not None:
                return existing
            default_flag = default_en1090_enabled(identifier or metadata.get("display", ""))
            self.en1090_preferences[norm] = default_flag
            return default_flag

        def _set_en1090_preference(self, metadata: Dict[str, str], enabled: bool) -> None:
            if metadata.get("kind") not in {"production", "opticutter"}:
                return
            identifier = metadata.get("identifier") or metadata.get("display") or ""
            norm = normalize_en1090_key(identifier)
            if not norm:
                return
            self.en1090_preferences[norm] = bool(enabled)
            self._save_settings()

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

        def update_en1090_note(self, text: str) -> None:
            normalized = (text or "").replace("\r\n", "\n")
            prev = getattr(self, "_suspend_save", False)
            self._suspend_save = True
            try:
                self.en1090_note_var.set(normalized)
            finally:
                self._suspend_save = prev
            self._save_settings()

        def set_en1090_enabled(self, enabled: bool) -> None:
            desired = 1 if enabled else 0
            prev = getattr(self, "_suspend_save", False)
            self._suspend_save = True
            try:
                if self.en1090_enabled_var.get() != desired:
                    self.en1090_enabled_var.set(desired)
            finally:
                self._suspend_save = prev
            self._save_settings()
            sel_frame = getattr(self, "sel_frame", None)
            if sel_frame is not None:
                try:
                    if sel_frame.winfo_exists():
                        sel_frame.set_en1090_enabled(bool(desired))
                except Exception:
                    pass

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

        def _schedule_window_maximize(self) -> None:
            def _maximize() -> None:
                try:
                    self.state("zoomed")
                    return
                except tk.TclError:
                    pass
                try:
                    self.attributes("-zoomed", True)
                    return
                except tk.TclError:
                    pass
                try:
                    self.update_idletasks()
                    screen_w = self.winfo_screenwidth()
                    screen_h = self.winfo_screenheight()
                    self.geometry(f"{screen_w}x{screen_h}+0+0")
                except tk.TclError:
                    pass

            try:
                self.after_idle(_maximize)
            except tk.TclError:
                _maximize()

        def _load_window_icon(self) -> None:
            """Load window icon asynchronously to avoid startup lag."""
            try:
                import os
                if Image is None or ImageTk is None:
                    return
                icon_path = os.path.join(os.path.dirname(__file__), "app_icon.png")
                if not os.path.isfile(icon_path):
                    return
                icon_img = Image.open(icon_path)
                icon_photo = ImageTk.PhotoImage(icon_img)
                self.iconphoto(False, icon_photo)
                self._icon_photo = icon_photo  # Keep reference
            except Exception:
                pass

        def _show_release_notes(self) -> None:
            """Display release notes in a new window."""
            import tkinter as tk
            from tkinter import scrolledtext
            
            changelog = load_changelog()
            
            win = tk.Toplevel(self)
            win.title("Release Notes")
            win.geometry("800x600")
            
            # Title
            title_frame = tk.Frame(win, background="#f0f0f0", height=60)
            title_frame.pack(fill="x", side="top")
            title_frame.pack_propagate(False)
            
            title_label = tk.Label(
                title_frame,
                text="Release Notes - Wat is er nieuw?",
                font=("TkDefaultFont", 14, "bold"),
                background="#f0f0f0",
                foreground="#333333",
                anchor="w",
                justify="left",
            )
            title_label.pack(padx=16, pady=12, anchor="w")
            
            # Changelog content in scrolled text widget
            text_frame = tk.Frame(win)
            text_frame.pack(fill="both", expand=True, padx=8, pady=8)
            
            text_widget = scrolledtext.ScrolledText(
                text_frame,
                wrap="word",
                font=("Courier", 9),
                background="white",
                foreground="#333333",
                relief="solid",
                borderwidth=1,
                padx=8,
                pady=8,
            )
            text_widget.pack(fill="both", expand=True)
            text_widget.insert("1.0", format_changelog_for_display(changelog))
            text_widget.config(state="disabled")  # Read-only
            
            # Close button
            button_frame = tk.Frame(win)
            button_frame.pack(fill="x", padx=8, pady=8)
            close_btn = tk.Button(button_frame, text="Sluiten", command=win.destroy)
            close_btn.pack(side="right", padx=4)
            
            # Focus and center window
            win.transient(self)
            win.grab_set()

        def _show_quick_start(self) -> None:
            """Show the interactive quick-start guide."""
            import tkinter as tk
            from tkinter import ttk

            steps = QUICK_START_STEPS
            current = {"index": 0}

            win = tk.Toplevel(self)
            win.title("Quick Start Guide")
            win.geometry("680x420")

            def update_step() -> None:
                step = steps[current["index"]]
                title_label.config(text=step["title"])
                description_label.config(text=step["description"])
                prev_button.config(state="normal" if current["index"] > 0 else "disabled")
                next_button.config(text="Volgende" if current["index"] < len(steps) - 1 else "Sluiten")

            title_label = tk.Label(
                win,
                text="",
                font=("TkDefaultFont", 13, "bold"),
                anchor="w",
                justify="left",
                wraplength=640,
                pady=12,
            )
            title_label.pack(fill="x", padx=16)

            description_label = tk.Label(
                win,
                text="",
                anchor="nw",
                justify="left",
                wraplength=640,
                padx=4,
                pady=8,
            )
            description_label.pack(fill="both", expand=True, padx=16)

            control_frame = tk.Frame(win)
            control_frame.pack(fill="x", padx=16, pady=(0, 12))
            prev_button = tk.Button(control_frame, text="Vorige", command=lambda: navigate(-1))
            prev_button.pack(side="left")
            next_button = tk.Button(control_frame, text="Volgende", command=lambda: navigate(1))
            next_button.pack(side="right")

            def navigate(direction: int) -> None:
                if direction < 0 and current["index"] > 0:
                    current["index"] -= 1
                    update_step()
                elif direction > 0:
                    if current["index"] < len(steps) - 1:
                        current["index"] += 1
                        update_step()
                    else:
                        win.destroy()

            update_step()
            win.transient(self)
            win.grab_set()

        def _show_faq(self) -> None:
            """Show the FAQ dialog."""
            import tkinter as tk
            from tkinter import scrolledtext

            content = []
            for entry in FAQ_ENTRIES:
                content.append(f"Q: {entry['question']}")
                content.append(f"A: {entry['answer']}")
                content.append("")

            win = tk.Toplevel(self)
            win.title("FAQ")
            win.geometry("760x520")

            title_label = tk.Label(
                win,
                text="Veelgestelde vragen",
                font=("TkDefaultFont", 14, "bold"),
                anchor="w",
                padx=16,
                pady=12,
            )
            title_label.pack(fill="x")

            text_widget = scrolledtext.ScrolledText(
                win,
                wrap="word",
                font=("TkDefaultFont", 10),
                background="white",
                relief="solid",
                borderwidth=1,
                padx=8,
                pady=8,
            )
            text_widget.pack(fill="both", expand=True, padx=16, pady=(0, 8))

            text_widget.tag_configure(
                "question",
                font=("TkDefaultFont", 10, "bold"),
                foreground="#1F2937",
                lmargin1=4,
                lmargin2=4,
            )
            text_widget.tag_configure(
                "answer",
                font=("TkDefaultFont", 10),
                foreground="#374151",
                lmargin1=12,
                lmargin2=12,
            )

            for entry in FAQ_ENTRIES:
                text_widget.insert("end", f"Q: {entry['question']}\n", "question")
                text_widget.insert("end", f"A: {entry['answer']}\n\n", "answer")

            text_widget.config(state="disabled")

            close_btn = tk.Button(win, text="Sluiten", command=win.destroy)
            close_btn.pack(anchor="e", padx=16, pady=(0, 12))
            win.transient(self)
            win.grab_set()

        def _open_visual_quick_manual(self) -> None:
            """Open the visual quick-start markdown file."""
            import os
            import tkinter as tk
            from tkinter import messagebox
            from pathlib import Path

            doc_path = Path(__file__).resolve().parent / "docs" / "quick_start_visual.md"
            if doc_path.exists():
                try:
                    os.startfile(str(doc_path))
                    return
                except Exception:
                    pass
            try:
                import webbrowser
                webbrowser.open(doc_path.as_uri())
            except Exception as exc:
                messagebox.showerror(
                    "Fout",
                    f"Kon de visuele quick manual niet openen.\n{doc_path}\n{exc}",
                    parent=self,
                )

        def _show_about_dialog(self) -> None:
            """Show about dialog."""
            import tkinter as tk
            from tkinter import messagebox
            
            about_text = (
                f"Filehopper {APP_VERSION}\n\n"
                f"Een tool voor het efficiënt beheren van bestanden\n"
                f"en het genereren van bestelbon-documenten.\n\n"
                f"© 2024-2026"
            )
            messagebox.showinfo("Over Filehopper", about_text, parent=self)

        def _handle_tab_changed(self, event: "tk.Event") -> None:
            if event.widget is not self.nb:
                return
            selected = event.widget.select()
            previous = getattr(self, "_last_selected_notebook_tab", "")
            placeholder = getattr(self, "_custom_bom_placeholder", None)
            if placeholder is not None and selected and str(selected) == str(placeholder):
                self._handle_opticutter_tab_transition(previous, selected)
                tab = self._ensure_custom_bom_tab()
                self.nb.select(tab)
                self._last_selected_notebook_tab = str(tab)
                return

            self._handle_opticutter_tab_transition(previous, selected)
            self._refresh_opticutter_if_needed(selected)
            self._last_selected_notebook_tab = str(selected or "")

        def _is_opticutter_tab(self, tab_id: object) -> bool:
            opticutter_frame = getattr(self, "opticutter_frame", None)
            return bool(
                tab_id
                and opticutter_frame is not None
                and str(tab_id) == str(opticutter_frame)
            )

        def _apply_opticutter_analysis_state(self, analysis) -> None:
            self.opticutter_last_analysis = analysis
            if analysis is None or not getattr(analysis, "profiles", None):
                self.opticutter_profile_selection_scenarios = {}
                self.opticutter_profile_selection_choice.clear()
                return

            valid_keys = {profile.key for profile in analysis.profiles}
            for stored_key in list(self.opticutter_profile_custom_lengths.keys()):
                if stored_key not in valid_keys:
                    self.opticutter_profile_custom_lengths.pop(stored_key, None)
            for stored_key in list(self.opticutter_profile_selection_choice.keys()):
                if stored_key not in valid_keys:
                    self.opticutter_profile_selection_choice.pop(stored_key, None)

            selection_scenarios = {}
            for profile in analysis.profiles:
                selection_scenarios[profile.key] = profile.scenarios
                available_values = set(profile.scenarios.keys()) | {"input"}
                previous_choice = self.opticutter_profile_selection_choice.get(
                    profile.key
                )
                selected_value = (
                    previous_choice
                    if previous_choice in available_values
                    else profile.best_choice
                )
                self.opticutter_profile_selection_choice[profile.key] = selected_value
            self.opticutter_profile_selection_scenarios = selection_scenarios

        def _compute_opticutter_analysis_snapshot(
            self,
            bom_df_snapshot: Optional["pd.DataFrame"],
            kerf_mm: float,
            custom_stock_mm: Optional[int],
            manual_lengths: Dict[tuple[str, str, str], int],
        ):
            if bom_df_snapshot is None or bom_df_snapshot.empty:
                return None
            return analyse_profiles(
                bom_df_snapshot,
                kerf_mm=kerf_mm,
                custom_stock_mm=custom_stock_mm,
                manual_lengths=manual_lengths,
            )

        def _ensure_opticutter_analysis_current(self) -> None:
            if not getattr(self, "_opticutter_analysis_stale", True):
                return

            self._opticutter_refresh_generation += 1
            self._opticutter_analysis_refresh_running = False
            bom_df = self.bom_df
            bom_df_snapshot = (
                bom_df.copy(deep=True)
                if bom_df is not None and not bom_df.empty
                else None
            )
            analysis = self._compute_opticutter_analysis_snapshot(
                bom_df_snapshot,
                self._get_opticutter_kerf_mm(),
                self._get_opticutter_custom_stock_mm(),
                dict(self.opticutter_profile_custom_lengths),
            )
            self._apply_opticutter_analysis_state(analysis)
            self._opticutter_analysis_stale = False

        def _start_background_opticutter_analysis_refresh(self) -> None:
            self._opticutter_refresh_generation += 1
            generation = self._opticutter_refresh_generation
            self._opticutter_analysis_stale = True

            bom_df = self.bom_df
            bom_df_snapshot = (
                bom_df.copy(deep=True)
                if bom_df is not None and not bom_df.empty
                else None
            )
            kerf_mm = self._get_opticutter_kerf_mm()
            custom_stock_mm = self._get_opticutter_custom_stock_mm()
            manual_lengths = dict(self.opticutter_profile_custom_lengths)

            if bom_df_snapshot is None:
                self._apply_opticutter_analysis_state(None)
                self._opticutter_analysis_stale = False
                self._opticutter_analysis_refresh_running = False
                return

            self._opticutter_analysis_refresh_running = True

            def work() -> None:
                try:
                    analysis = self._compute_opticutter_analysis_snapshot(
                        bom_df_snapshot,
                        kerf_mm,
                        custom_stock_mm,
                        manual_lengths,
                    )
                except Exception as exc:
                    analysis = None
                    print(
                        f"Kon Opticutter-analyse niet bijwerken: {exc}",
                        file=sys.stderr,
                    )

                def apply() -> None:
                    if generation != getattr(
                        self, "_opticutter_refresh_generation", generation
                    ):
                        return
                    self._opticutter_analysis_refresh_running = False
                    self._apply_opticutter_analysis_state(analysis)
                    self._opticutter_analysis_stale = False
                    try:
                        selected = self.nb.select()
                    except Exception:
                        selected = None
                    if self._is_opticutter_tab(selected):
                        self._opticutter_needs_refresh = False
                        self._refresh_opticutter_table(analysis_override=analysis)
                    else:
                        self._opticutter_needs_refresh = True

                try:
                    self.after(0, apply)
                except tk.TclError:
                    pass

            threading.Thread(
                target=work,
                name="OpticutterAnalysisRefresh",
                daemon=True,
            ).start()

        def _refresh_opticutter_if_needed(self, selected_tab: object = None) -> None:
            if not getattr(self, "_opticutter_needs_refresh", False):
                return
            tab_id = selected_tab
            if tab_id is None:
                try:
                    tab_id = self.nb.select()
                except Exception:
                    tab_id = None
            if not self._is_opticutter_tab(tab_id):
                return
            self._opticutter_needs_refresh = False
            analysis = None
            if not getattr(self, "_opticutter_analysis_stale", True):
                analysis = getattr(self, "opticutter_last_analysis", None)
            self._refresh_opticutter_table(analysis_override=analysis)

        def _request_opticutter_refresh(self) -> None:
            self._opticutter_needs_refresh = True
            self._start_background_opticutter_analysis_refresh()

        def _handle_opticutter_tab_transition(
            self,
            previous_tab: object,
            selected_tab: object,
        ) -> None:
            if not self._is_opticutter_tab(previous_tab):
                return
            if self._is_opticutter_tab(selected_tab):
                return
            if not getattr(self, "_opticutter_dirty", False):
                return
            self._confirm_opticutter_update()

        def _ensure_custom_bom_tab(self) -> "BOMCustomTab":
            tab = getattr(self, "custom_bom_tab", None)
            if tab is not None:
                return tab
            tab = BOMCustomTab(
                self.nb,
                app_name="Filehopper",
                on_custom_bom_ready=self._on_custom_bom_ready,
                on_push_to_main=self._apply_custom_bom_to_main,
                event_target=self,
            )
            placeholder = getattr(self, "_custom_bom_placeholder", None)
            insert_index = None
            if placeholder is not None:
                try:
                    insert_index = self.nb.index(placeholder)
                except tk.TclError:
                    insert_index = None
                try:
                    self.nb.forget(placeholder)
                except tk.TclError:
                    pass
                try:
                    placeholder.destroy()
                except Exception:
                    pass
                self._custom_bom_placeholder = None
            if insert_index is None:
                self.nb.add(tab, text="Custom BOM")
            else:
                self.nb.insert(insert_index, tab, text="Custom BOM")
            self.custom_bom_tab = tab
            if self._autofill_custom_bom_enabled() and (
                getattr(self, "_custom_bom_needs_sync", False)
                or self.bom_df is not None
            ):
                self._load_current_bom_into_custom_tab(tab)
            return tab

        def _select_custom_bom_tab(self) -> None:
            tab = self._ensure_custom_bom_tab()
            self.nb.select(tab)

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

        def _load_current_bom_into_custom_tab(self, tab: "BOMCustomTab") -> None:
            df = self.bom_df
            if df is None:
                df = pd.DataFrame(columns=BOMCustomTab.MAIN_COLUMN_ORDER)

            try:
                tab.load_from_main_dataframe(df)
            except Exception as exc:
                print(
                    f"Kon custom BOM niet vullen vanuit hoofd-BOM: {exc}",
                    file=sys.stderr,
                )
            finally:
                self._custom_bom_needs_sync = False

        def _sync_custom_bom_from_main(self) -> None:
            """Update the Custom BOM tab so it mirrors the main BOM."""

            if not self._autofill_custom_bom_enabled():
                self._custom_bom_needs_sync = False
                return

            tab = getattr(self, "custom_bom_tab", None)
            if tab is None:
                self._custom_bom_needs_sync = self.bom_df is not None
                return
            self._load_current_bom_into_custom_tab(tab)

        def _apply_loaded_bom(
            self,
            path: str,
            df: "pd.DataFrame",
            *,
            mark_as_custom: bool = False,
        ) -> None:
            if "Bestanden gevonden" not in df.columns:
                df["Bestanden gevonden"] = ""
            if "Status" not in df.columns:
                df["Status"] = ""
            if "Link" not in df.columns:
                df["Link"] = ""
            self._store_custom_row_flags(df, [mark_as_custom] * len(df.index))
            self.bom_df = df
            previous_source_path = _to_str(getattr(self, "bom_source_path", "")).strip()
            if mark_as_custom and previous_source_path:
                self.bom_source_path = previous_source_path
            else:
                self.bom_source_path = os.path.abspath(path)
            self._refresh_tree()
            self.status_var.set(f"BOM geladen: {len(df)} rijen")
            self._sync_custom_bom_from_main()

        def _load_bom_from_path(self, path: str, *, mark_as_custom: bool = False) -> None:
            df = load_bom(path)
            self._apply_loaded_bom(path, df, mark_as_custom=mark_as_custom)

        def _load_bom(self):
            from tkinter import filedialog, messagebox

            start_dir = _resolve_file_dialog_initial_dir(self.source_folder)
            path = filedialog.askopenfilename(
                filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")],
                initialdir=start_dir,
            )
            if not path:
                return
            if getattr(self, "_bom_load_in_progress", False):
                messagebox.showinfo(
                    "BOM wordt geladen",
                    "Er wordt al een BOM geladen. Wacht tot die klaar is.",
                    parent=self,
                )
                return

            self._bom_load_in_progress = True
            self.status_var.set("BOM laden...")

            def work() -> None:
                try:
                    df = load_bom(path)
                except Exception as exc:
                    error_message = str(exc)

                    def on_error(message: str = error_message) -> None:
                        self._bom_load_in_progress = False
                        messagebox.showerror("Fout", message, parent=self)

                    try:
                        self.after(0, on_error)
                    except tk.TclError:
                        pass
                    return

                def on_loaded(loaded_df: "pd.DataFrame" = df) -> None:
                    self._bom_load_in_progress = False
                    try:
                        self._apply_loaded_bom(path, loaded_df)
                    except Exception as exc:
                        messagebox.showerror("Fout", str(exc), parent=self)

                try:
                    self.after(0, on_loaded)
                except tk.TclError:
                    pass

            threading.Thread(target=work, name="BOMLoader", daemon=True).start()

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


            parent_widget = self.custom_bom_tab if self.custom_bom_tab is not None else self
            if custom_df is None or custom_df.empty:
                messagebox.showwarning(
                    "Geen gegevens",
                    "Er zijn geen rijen met gegevens om naar de Main-tab te sturen.",
                    parent=parent_widget,
                )
                return

            try:
                normalized = prepare_custom_bom_for_main(custom_df, self.bom_df)
            except ValueError as exc:
                messagebox.showerror("Fout", str(exc), parent=parent_widget)
                return

            normalized.attrs.pop("production_column_missing", None)
            self._store_custom_row_flags(normalized, [True] * len(normalized.index))
            self.bom_df = normalized
            file_status_refreshed = False
            try:
                exts = self._selected_exts() if getattr(self, "source_folder", "") else []
                if exts:
                    self._update_bom_file_status(exts)
                    file_status_refreshed = True
            except Exception as exc:
                print(
                    f"Kon bestandsstatus niet herberekenen na Custom BOM-update: {exc}",
                    file=sys.stderr,
                )
            if not file_status_refreshed:
                self._refresh_tree()
            self._sync_custom_bom_from_main()
            self.nb.select(self.main_frame)
            selection_frame = getattr(self, "sel_frame", None)
            selection_refreshed = False
            if selection_frame is not None:
                try:
                    if selection_frame.winfo_exists():
                        selection_refreshed = (
                            self._show_supplier_selection_tab(
                                select_tab=False,
                                prompt_opticutter=False,
                            )
                            is not None
                        )
                except Exception:
                    selection_refreshed = False
            status_message = f"Custom BOM wijzigingen toegepast ({len(normalized)} rijen)."
            if file_status_refreshed:
                status_message += " Bestandscontrole bijgewerkt."
            if selection_refreshed:
                status_message += " Bestelbonnen bijgewerkt."
            self.status_var.set(status_message)

        def _collect_supplier_selection_payload(
            self,
            *,
            prompt_opticutter: bool = True,
        ) -> Optional[Dict[str, object]]:
            from tkinter import messagebox

            if not self._ensure_bom_loaded():
                return None

            bom_df = self.bom_df
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
            selection_items: Dict[str, List[Dict[str, object]]] = defaultdict(list)
            selection_item_keys: Dict[str, set[str]] = defaultdict(set)

            def _append_order_line_item(
                sel_key: str,
                row: Mapping[str, object],
                *,
                context_kind: str,
            ) -> None:
                item = {
                    "PartNumber": row.get("PartNumber", ""),
                    "Description": row.get("Description", ""),
                    "Materiaal": row.get("Materiaal", ""),
                    "Aantal": row.get("Aantal", ""),
                    "Oppervlakte": row.get("Oppervlakte", ""),
                    "Gewicht": row.get("Gewicht", ""),
                }
                item_key = build_order_pricing_item_key(
                    item,
                    context_kind=context_kind,
                )
                if not item_key:
                    return
                if item_key in selection_item_keys[sel_key]:
                    qty_to_add = _parse_supplier_decimal(item.get("Aantal", ""))
                    if qty_to_add is not None:
                        for existing_item in selection_items[sel_key]:
                            if _to_str(existing_item.get("key")).strip() != item_key:
                                continue
                            current_qty = _parse_supplier_decimal(
                                existing_item.get("quantity")
                            )
                            if current_qty is not None:
                                existing_item["quantity"] = _format_supplier_quantity(
                                    current_qty + qty_to_add
                                )
                            break
                    return
                part_number = _to_str(item.get("PartNumber")).strip()
                description = _to_str(item.get("Description")).strip()
                material = _to_str(item.get("Materiaal")).strip()
                label = " - ".join(
                    part for part in (part_number, description, material) if part
                )
                item["key"] = item_key
                item["label"] = label or item_key
                item["quantity"] = _to_str(item.get("Aantal")).strip()
                selection_item_keys[sel_key].add(item_key)
                selection_items[sel_key].append(item)

            for _, row in bom_df.iterrows():
                prod = _to_str(row.get("Production")).strip() or "_Onbekend"
                _append_order_line_item(
                    make_production_selection_key(prod),
                    row,
                    context_kind="Productie",
                )
                finish_text = _to_str(row.get("Finish")).strip()
                if finish_text:
                    meta = describe_finish_combo(row.get("Finish"), row.get("RAL color"))
                    finish_key = meta.get("key", "")
                    if finish_key:
                        _append_order_line_item(
                            make_finish_selection_key(finish_key),
                            row,
                            context_kind="Afwerking",
                        )

            self._ensure_opticutter_analysis_current()
            opticutter_analysis = getattr(self, "opticutter_last_analysis", None)
            scenarios_ready = bool(self.opticutter_profile_selection_scenarios)

            if (
                opticutter_analysis is not None
                and opticutter_analysis.profiles
                and not scenarios_ready
            ):
                if prompt_opticutter:
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
                        return None
                    for profile in opticutter_analysis.profiles:
                        self.opticutter_profile_selection_choice[profile.key] = (
                            profile.best_choice
                        )
                        self.opticutter_profile_selection_scenarios[profile.key] = (
                            profile.scenarios
                        )
                    self._ensure_opticutter_analysis_current()
                    opticutter_analysis = getattr(self, "opticutter_last_analysis", None)
                    scenarios_ready = bool(self.opticutter_profile_selection_scenarios)

            opticutter_notice_message = ""
            if (
                opticutter_analysis is not None
                and opticutter_analysis.profiles
                and not scenarios_ready
                and not prompt_opticutter
            ):
                opticutter_notice_message = (
                    "Opticutter is nog niet volledig ingevuld. "
                    "Werk de zaagplanning bij om brutemateriaalorders te tonen."
                )
            elif (
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

            return {
                "prods": prods,
                "finish_entries": finish_entries,
                "finish_label_lookup": finish_label_lookup,
                "selection_items": dict(selection_items),
                "opticutter_details": opticutter_details,
                "opticutter_notice_message": opticutter_notice_message,
            }

        def _show_supplier_selection_tab(
            self,
            *,
            select_tab: bool = True,
            prompt_opticutter: bool = True,
            pdf_dossier_context: bool = False,
        ) -> Optional["SupplierSelectionFrame"]:
            from tkinter import messagebox
            pdf_dossier_context = bool(pdf_dossier_context)

            payload = self._collect_supplier_selection_payload(
                prompt_opticutter=prompt_opticutter
            )
            if payload is None:
                return None

            prods = list(payload.get("prods") or [])
            finish_entries = list(payload.get("finish_entries") or [])
            finish_label_lookup = dict(payload.get("finish_label_lookup") or {})
            selection_items = dict(payload.get("selection_items") or {})
            opticutter_details = dict(payload.get("opticutter_details") or {})
            opticutter_notice_message = _to_str(
                payload.get("opticutter_notice_message")
            ).strip()
            sel_frame = None

            def on_sel(
                sel_map: Dict[str, str],
                doc_map: Dict[str, str],
                doc_num_map: Dict[str, str],
                delivery_map_raw: Dict[str, str],
                remarks_map_raw: Dict[str, str],
                group_map_raw: Dict[str, str],
                en1090_map_raw: Dict[str, bool],
                project_number: str,
                project_name: str,
                remember: bool,
                export_flags: Dict[str, bool] | None = None,
            ):
                if not self._ensure_bom_loaded():
                    return

                current_bom = self.bom_df
                attrs = getattr(current_bom, "attrs", {}) or {}
                missing_production = bool(attrs.get("production_column_missing"))
                if missing_production:
                    messagebox.showwarning(
                        "Let op",
                        "De geladen BOM mist de kolom 'Production'. "
                        "Vul de productie in de BOM in om bestelbonnen per productie te exporteren.",
                        parent=sel_frame or self,
                    )
                    return

                source_folder_snapshot = self.source_folder_var.get().strip()
                dest_folder_snapshot = self.dest_folder_var.get().strip()
                output_root_snapshot = dest_folder_snapshot or source_folder_snapshot
                current_exts = self._selected_exts() or []
                if pdf_dossier_context:
                    if not source_folder_snapshot or not output_root_snapshot:
                        messagebox.showwarning(
                            "Let op",
                            "Selecteer minstens een bronmap voor het PDF dossier.",
                            parent=sel_frame or self,
                        )
                        return
                    current_exts = []
                elif not current_exts or not source_folder_snapshot or not dest_folder_snapshot:
                    messagebox.showwarning(
                        "Let op",
                        "Selecteer bron, bestemming en extensies.",
                        parent=sel_frame or self,
                    )
                    return

                client = self._current_client()
                self._ensure_opticutter_analysis_current()
                opticutter_analysis_current = getattr(
                    self, "opticutter_last_analysis", None
                )

                export_state_snapshot: Optional[SupplierSelectionState] = None
                if sel_frame is not None:
                    try:
                        if sel_frame.winfo_exists():
                            export_state_snapshot = sel_frame.serialize_state()
                            self._last_supplier_selection_state = export_state_snapshot
                    except Exception:
                        pass
                if export_state_snapshot is None:
                    export_state_snapshot = getattr(self, "_last_supplier_selection_state", None)

                prod_override_map: Dict[str, str] = {}
                finish_override_map: Dict[str, str] = {}
                opticutter_override_map: Dict[str, str] = {}
                production_pricing_map: Dict[str, Dict[str, object]] = {}
                finish_pricing_map: Dict[str, Dict[str, object]] = {}
                opticutter_pricing_map: Dict[str, Dict[str, object]] = {}
                production_vat_rate_map: Dict[str, str] = {}
                finish_vat_rate_map: Dict[str, str] = {}
                opticutter_vat_rate_map: Dict[str, str] = {}
                export_flags = export_flags or {}
                prod_export_filter: Dict[str, bool] = {}
                finish_export_filter: Dict[str, bool] = {}
                opticutter_export_filter: Dict[str, bool] = {}
                for key, value in sel_map.items():
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_override_map[identifier] = value
                    elif kind == "opticutter":
                        opticutter_override_map[identifier] = value
                    else:
                        prod_override_map[identifier] = value
                for key, enabled in export_flags.items():
                    kind, identifier = parse_selection_key(key)
                    target: Dict[str, bool] | None
                    if kind == "finish":
                        target = finish_export_filter
                    elif kind == "opticutter":
                        target = opticutter_export_filter
                    else:
                        target = prod_export_filter
                    target[identifier] = bool(enabled)

                pricing_map_raw = (
                    getattr(export_state_snapshot, "pricing", {}) if export_state_snapshot else {}
                )
                for key, value in pricing_map_raw.items():
                    if not isinstance(value, Mapping):
                        continue
                    clean_value = _clean_supplier_pricing_value(value)
                    if not clean_value:
                        continue
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_pricing_map[identifier] = clean_value
                    elif kind == "opticutter":
                        opticutter_pricing_map[identifier] = clean_value
                    else:
                        production_pricing_map[identifier] = clean_value

                vat_rate_map_raw = (
                    getattr(export_state_snapshot, "vat_rates", {}) if export_state_snapshot else {}
                )
                for key, value in vat_rate_map_raw.items():
                    clean_rate = _clean_supplier_vat_rate(value)
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_vat_rate_map[identifier] = clean_rate
                    elif kind == "opticutter":
                        opticutter_vat_rate_map[identifier] = clean_rate
                    else:
                        production_vat_rate_map[identifier] = clean_rate

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
                    resolved = self._resolve_delivery_choice(name, client)
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

                en1090_override_map: Dict[str, bool] = {}
                for key, flag in en1090_map_raw.items():
                    kind, identifier = parse_selection_key(key)
                    if kind not in {"production", "opticutter"}:
                        continue
                    norm = normalize_en1090_key(identifier)
                    if norm:
                        en1090_override_map[norm] = bool(flag)

                custom_prefix_text = self.export_name_custom_prefix_text.get().strip()
                custom_prefix_enabled = bool(
                    self.export_name_custom_prefix_enabled_var.get()
                )
                custom_suffix_text = self.export_name_custom_suffix_text.get().strip()
                custom_suffix_enabled = bool(
                    self.export_name_custom_suffix_enabled_var.get()
                )
                client_name_snapshot = _to_str(getattr(client, "name", "") if client else "").strip()
                bom_source_path_snapshot = _to_str(getattr(self, "bom_source_path", "")).strip()
                pdf_order_document_base_snapshot = output_root_snapshot
                pdf_order_document_root_snapshot = ""
                if pdf_dossier_context:
                    pdf_order_document_root_snapshot = self._pdf_order_documents_root()

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
                    bundle = None
                    if pdf_dossier_context:
                        update_status("PDF dossier bonnenmap voorbereiden...")
                        try:
                            bundle_dest = pdf_order_document_root_snapshot
                            self._reset_pdf_order_documents_root(
                                bundle_dest,
                                base_folder=pdf_order_document_base_snapshot,
                            )
                        except Exception as exc:
                            def on_error():
                                messagebox.showerror(
                                    "Fout",
                                    f"Kon de PDF dossier documentenmap niet voorbereiden:\n{exc}",
                                    parent=self,
                                )
                                update_status("PDF dossier documentenmap maken mislukt.")
                                set_busy_state(False)

                            self.after(0, on_error)
                            return
                    else:
                        update_status("Bundelmap voorbereiden...")
                        try:
                            bundle = create_export_bundle(
                                dest_folder_snapshot,
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

                    update_status("KopiÃ«ren & bestelbonnen maken...")
                    if pdf_dossier_context:
                        update_status("Bon-PDF's voor PDF dossier maken...")
                    path_limit_messages: List[str] = []
                    document_status_lines: List[str] = []
                    generated_document_records: List[Dict[str, object]] = []
                    try:
                        if opticutter_choices_snapshot is None:
                            opticutter_choices_snapshot = dict(
                                self.opticutter_profile_selection_choice
                            )
                        cnt, chosen = copy_per_production_and_orders(
                            source_folder_snapshot,
                            bundle_dest,
                            current_bom,
                            current_exts,
                            self.db,
                            prod_override_map,
                            doc_type_map,
                            prod_doc_num_map,
                            remember,
                            client=client,
                            delivery_map=production_delivery_map,
                            footer_note=self.footer_note_var.get(),
                            zip_parts=(
                                False if pdf_dossier_context else bool(self.zip_var.get())
                            ),
                            date_prefix_exports=(
                                False
                                if pdf_dossier_context
                                else bool(self.export_date_prefix_var.get())
                            ),
                            date_suffix_exports=(
                                False
                                if pdf_dossier_context
                                else bool(self.export_date_suffix_var.get())
                            ),
                            project_number=project_number,
                            project_name=project_name,
                            copy_finish_exports=(
                                False
                                if pdf_dossier_context
                                else bool(self.finish_export_var.get())
                            ),
                            zip_finish_exports=(
                                False
                                if pdf_dossier_context
                                else bool(self.zip_finish_var.get())
                            ),
                            export_bom=(
                                False
                                if pdf_dossier_context
                                else bool(self.export_bom_var.get())
                            ),
                            export_related_files=(
                                False
                                if pdf_dossier_context
                                else bool(self.export_related_files_var.get())
                            ),
                            export_name_prefix_text=token_prefix_text,
                            export_name_prefix_enabled=token_prefix_enabled,
                            export_name_suffix_text=token_suffix_text,
                            export_name_suffix_enabled=token_suffix_enabled,
                            document_filename_profile=self.document_filename_profile_var.get(),
                            document_filename_show_doc_type=bool(
                                self.document_filename_show_doc_type_var.get()
                            ),
                            document_filename_show_doc_number=bool(
                                self.document_filename_show_doc_number_var.get()
                            ),
                            document_filename_show_context=bool(
                                self.document_filename_show_context_var.get()
                            ),
                            document_filename_show_date=bool(
                                self.document_filename_show_date_var.get()
                            ),
                            document_filename_compact_doc_number=bool(
                                self.document_filename_compact_doc_number_var.get()
                            ),
                            document_filename_separator=self.document_filename_separator_var.get(),
                            document_display_compact_doc_number=bool(
                                self.document_display_compact_doc_number_var.get()
                            ),
                            finish_override_map=finish_override_map,
                            finish_doc_type_map=finish_doc_type_map,
                            finish_doc_num_map=finish_doc_num_map,
                            finish_delivery_map=finish_delivery_map,
                            remarks_map=production_remarks_map,
                            finish_remarks_map=finish_remarks_map,
                            document_group_map=group_map_raw if group_map_raw else None,
                            bom_source_path=self.bom_source_path,
                            path_limit_warnings=path_limit_messages,
                            opticutter_analysis=opticutter_analysis_snapshot,
                            opticutter_choices=opticutter_choices_snapshot,
                            opticutter_override_map=opticutter_override_map,
                            opticutter_doc_type_map=opticutter_doc_type_map,
                            opticutter_doc_num_map=opticutter_doc_num_map,
                            opticutter_delivery_map=opticutter_delivery_map,
                            opticutter_remarks_map=opticutter_remarks_map,
                            pricing_map=production_pricing_map or None,
                            finish_pricing_map=finish_pricing_map or None,
                            opticutter_pricing_map=opticutter_pricing_map or None,
                            vat_rate_map=production_vat_rate_map or None,
                            finish_vat_rate_map=finish_vat_rate_map or None,
                            opticutter_vat_rate_map=opticutter_vat_rate_map or None,
                            production_export_filter=(
                                prod_export_filter if prod_export_filter else None
                            ),
                            finish_export_filter=(
                                finish_export_filter if finish_export_filter else None
                            ),
                            opticutter_export_filter=(
                                opticutter_export_filter
                                if opticutter_export_filter
                                else None
                            ),
                            en1090_overrides=en1090_override_map or None,
                            en1090_enabled=bool(self.en1090_enabled_var.get()),
                            en1090_note=self.en1090_note_var.get(),
                            document_status_messages=document_status_lines,
                            generated_documents=generated_document_records,
                        )
                        if export_state_snapshot is not None:
                            try:
                                log_payload = build_export_session_log(
                                    project_number=project_number,
                                    project_name=project_name,
                                    client_name=client_name_snapshot,
                                    bom_source_path=bom_source_path_snapshot,
                                    bom_df=current_bom,
                                    state=export_state_snapshot,
                                    app_version=APP_VERSION,
                                    generated_documents=generated_document_records,
                                    status_messages=document_status_lines,
                                    path_limit_warnings=path_limit_messages,
                                )
                                write_export_session_log(bundle_dest, log_payload)
                                document_status_lines.append(
                                    f"Exportlog opgeslagen: {EXPORT_SESSION_LOG_FILENAME}"
                                )
                            except Exception as log_exc:
                                document_status_lines.append(
                                    f"Exportlog niet opgeslagen: {log_exc}"
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
                        if pdf_dossier_context:
                            self._pdf_order_document_root = bundle_dest
                            self._pdf_generated_order_documents = list(
                                generated_document_records
                            )
                            self._pdf_order_context_signature = (
                                self._current_pdf_order_context_signature()
                            )
                            update_status(
                                f"Bon-PDF's voor PDF dossier aangemaakt in {bundle_dest}"
                            )
                            try:
                                current_sel = getattr(self, "sel_frame", None)
                                if current_sel is not None and current_sel.winfo_exists():
                                    self._last_supplier_selection_state = (
                                        current_sel.serialize_state()
                                    )
                            except Exception:
                                pass
                            try:
                                self.nb.select(self.pdf_workdossier_frame)
                            except tk.TclError:
                                pass
                            options_frame = getattr(
                                self, "pdf_workdossier_options_frame", None
                            )
                            if options_frame is not None:
                                try:
                                    options_frame.select_work_tab()
                                except tk.TclError:
                                    pass
                            self._refresh_pdf_workdossier_preview()
                            set_busy_state(False)
                            info_lines = [
                                "Bon-PDF's voor het PDF dossier zijn aangemaakt in:",
                                bundle_dest,
                            ]
                            if document_status_lines:
                                info_lines.append("")
                                info_lines.append("Details:")
                                info_lines.extend(
                                    f"- {line}" for line in document_status_lines
                                )
                            messagebox.showinfo(
                                "PDF dossier voorbereid",
                                "\n".join(info_lines),
                                parent=self,
                            )
                            return

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
                            f"Klaar. Gekopieerd: {cnt}. Leveranciers: {suppliers_text}. â†’ {bundle_dest}"
                        )
                        update_status(final_status)
                        try:
                            info_lines = ["Bestelbonnen aangemaakt in:", bundle_dest]
                            if bundle.latest_symlink:
                                info_lines.append(f"Symlink: {bundle.latest_symlink}")
                            if document_status_lines:
                                info_lines.append("")
                                info_lines.append("Details:")
                                info_lines.extend(
                                    f"- {line}" for line in document_status_lines
                                )
                            messagebox.showinfo("Klaar", "\n".join(info_lines), parent=self)
                            if path_limit_messages:
                                warning_lines = [
                                    "Sommige exportbestanden kregen een kortere naam omdat het pad te lang werd.",
                                    f"Windows laat maximaal {_WINDOWS_MAX_PATH} tekens per pad toe; Filehopper voegt dan automatisch een korte code toe.",
                                    "",
                                ]
                                warning_lines.extend(f"â€¢ {msg}" for msg in path_limit_messages)
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

                initial_busy_message = (
                    "PDF dossier bonnen voorbereiden..."
                    if pdf_dossier_context
                    else "Bundelmap voorbereiden..."
                )
                set_busy_state(True, initial_busy_message)

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
            previous_selected_tab = None
            try:
                previous_selected_tab = self.nb.select()
            except Exception:
                previous_selected_tab = None
            existing_frame = getattr(self, "sel_frame", None)
            existing_selected = False
            if existing_frame is not None:
                try:
                    if existing_frame.winfo_exists():
                        previous_state = existing_frame.serialize_state()
                        existing_selected = str(previous_selected_tab) == str(existing_frame)
                except Exception:
                    existing_selected = False
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
                    clients_db=self.client_db,
                    client_var=self.client_var,
                    presets_db=self.order_presets_db,
                    on_manage_presets=lambda: self.nb.select(self.preset_rules_frame),
                    opticutter_details=opticutter_details,
                    selection_items=selection_items,
                    initial_state=previous_state,
                    en1090_enabled=bool(self.en1090_enabled_var.get()),
                    en1090_getter=self._get_en1090_preference,
                    en1090_setter=self._set_en1090_preference,
                    pdf_dossier_context=pdf_dossier_context,
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
            settings_frame = getattr(self, "settings_frame", None)
            if settings_frame is not None:
                try:
                    settings_index = self.nb.index(settings_frame)
                except tk.TclError:
                    settings_index = None
            else:
                settings_index = None
            tab_text = "Bestelbonnen (PDF dossier)" if pdf_dossier_context else "Bestelbonnen"
            if settings_index is None:
                self.nb.add(sel_frame, text=tab_text)
            else:
                self.nb.insert(settings_index, sel_frame, text=tab_text)
            if select_tab or existing_selected:
                self.nb.select(sel_frame)

            if sup_search_restore and hasattr(sup_frame, "restore_search_filter"):
                def _restore_search(_event=None, frame=sup_frame, value=sup_search_restore):
                    try:
                        frame.restore_search_filter(value)
                    except Exception:
                        pass

                sel_frame.bind("<Destroy>", _restore_search, add="+")

            return sel_frame

        def _schedule_opticutter_refresh(self) -> None:
            after_id = getattr(self, "_opticutter_refresh_after_id", None)
            if after_id is not None:
                try:
                    self.after_cancel(after_id)
                except tk.TclError:
                    pass
            self._opticutter_refresh_after_id = self.after(200, self._refresh_opticutter_table)

        def _set_opticutter_update_status(
            self,
            message: str,
            *,
            color: str = "#6B7280",
        ) -> None:
            var = getattr(self, "opticutter_update_status_var", None)
            if var is not None:
                try:
                    var.set(message)
                except tk.TclError:
                    pass
            label = getattr(self, "opticutter_update_status_label", None)
            if label is not None:
                try:
                    label.configure(fg=color)
                except tk.TclError:
                    pass

        def _mark_opticutter_dirty(self) -> None:
            self._opticutter_dirty = True
            self._set_opticutter_update_status(
                "Wijzigingen actief",
                color="#B7791F",
            )

        def _format_opticutter_update_message(self) -> str:
            analysis = getattr(self, "opticutter_last_analysis", None)
            profiles = getattr(analysis, "profiles", None)
            try:
                profile_count = len(profiles or [])
            except TypeError:
                profile_count = 0
            if profile_count == 1:
                return "Opticutter bijgewerkt voor 1 profiel."
            if profile_count > 1:
                return f"Opticutter bijgewerkt voor {profile_count} profielen."
            return "Opticutter bijgewerkt."

        def _confirm_opticutter_update(self) -> bool:
            if not getattr(self, "_opticutter_dirty", False):
                return False

            if getattr(self, "_opticutter_refresh_after_id", None) is not None:
                try:
                    self._refresh_opticutter_table()
                except Exception:
                    pass

            self._opticutter_dirty = False
            message = self._format_opticutter_update_message()
            timestamp = datetime.datetime.now().strftime("%H:%M")
            self._set_opticutter_update_status(
                f"Laatst bijgewerkt {timestamp}",
                color="#2F855A",
            )
            status_var = getattr(self, "status_var", None)
            if status_var is not None:
                try:
                    status_var.set(message)
                except tk.TclError:
                    pass
            self._show_transient_toast(message)
            return True

        def _show_transient_toast(
            self,
            message: str,
            *,
            duration_ms: int = 3600,
        ) -> None:
            previous = getattr(self, "_opticutter_toast_window", None)
            if previous is not None:
                try:
                    previous.destroy()
                except tk.TclError:
                    pass
                self._opticutter_toast_window = None
            after_id = getattr(self, "_opticutter_toast_after_id", None)
            if after_id is not None:
                try:
                    self.after_cancel(after_id)
                except tk.TclError:
                    pass
                self._opticutter_toast_after_id = None

            try:
                toast = tk.Toplevel(self)
                toast.overrideredirect(True)
                toast.attributes("-topmost", True)
                background = "#174A2E"
                border = "#68D391"
                toast.configure(background=background)
                root_x = self.winfo_rootx()
                root_y = self.winfo_rooty()
                root_width = max(self.winfo_width(), 360)
                root_height = max(self.winfo_height(), 260)
                min_width = min(
                    max(int(root_width * 0.42), 360),
                    max(root_width - 72, 280),
                )

                container = tk.Frame(
                    toast,
                    background=background,
                    padx=24,
                    pady=16,
                    highlightthickness=2,
                    highlightbackground=border,
                )
                container.pack(fill="both", expand=True)
                title_font = tkfont.nametofont("TkDefaultFont").copy()
                try:
                    title_font.configure(
                        size=title_font.cget("size") + 2,
                        weight="bold",
                    )
                except tk.TclError:
                    pass
                tk.Label(
                    container,
                    text="Opticutter bijgewerkt",
                    background=background,
                    foreground="#F0FFF4",
                    font=title_font,
                    anchor="center",
                ).pack(fill="x")
                tk.Label(
                    container,
                    text=message,
                    background=background,
                    foreground="#DCFCE7",
                    font=tkfont.nametofont("TkDefaultFont"),
                    anchor="center",
                    wraplength=max(min_width - 48, 240),
                    justify="center",
                ).pack(fill="x", pady=(4, 0))
                toast.update_idletasks()
                width = max(toast.winfo_reqwidth(), min_width)
                height = toast.winfo_reqheight()
                x = root_x + ((root_width - width) // 2)
                y = root_y + root_height - height - 72
                toast.geometry(f"{width}x{height}+{max(x, 0)}+{max(y, 0)}")
                self._opticutter_toast_window = toast

                def _close_toast() -> None:
                    current = getattr(self, "_opticutter_toast_window", None)
                    self._opticutter_toast_window = None
                    self._opticutter_toast_after_id = None
                    if current is not None:
                        try:
                            current.destroy()
                        except tk.TclError:
                            pass

                self._opticutter_toast_after_id = self.after(duration_ms, _close_toast)
            except Exception:
                self._opticutter_toast_window = None
                self._opticutter_toast_after_id = None

        def _on_opticutter_kerf_change(self, *_args) -> None:
            self._mark_opticutter_dirty()
            self._schedule_opticutter_refresh()

        def _on_opticutter_custom_stock_change(self, *_args) -> None:
            self._mark_opticutter_dirty()
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
                self._mark_opticutter_dirty()
                self._refresh_opticutter_table()
                return
            previous_choice = self.opticutter_profile_selection_choice.get(key)
            self.opticutter_profile_selection_choice[key] = canonical
            if previous_choice != canonical:
                self._mark_opticutter_dirty()

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
            scrollbar = getattr(self, "opticutter_selection_scroll", None)
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
                if scrollbar is not None and scrollbar.winfo_ismapped():
                    scrollbar.pack_forget()
                custom_lengths.clear()
                return

            if empty_label is not None:
                empty_label.pack_forget()
            if scrollbar is not None and not scrollbar.winfo_ismapped():
                scrollbar.pack(side="left", fill="y", padx=(0, 8))

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

        def _refresh_opticutter_table(self, analysis_override=None) -> None:
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
                self._apply_opticutter_analysis_state(None)
                self._opticutter_analysis_stale = False
                self._opticutter_needs_refresh = False
                self.opticutter_profile_selection_scenarios = {}
                self._update_opticutter_selection_rows([])
                return
            if df.empty:
                if info_var is not None:
                    info_var.set("BOM is leeg. Geen profielen om te tonen.")
                self._apply_opticutter_analysis_state(None)
                self._opticutter_analysis_stale = False
                self._opticutter_needs_refresh = False
                self.opticutter_profile_selection_scenarios = {}
                self._update_opticutter_selection_rows([])
                return

            if analysis_override is None:
                self._opticutter_refresh_generation += 1
                self._opticutter_analysis_refresh_running = False
                kerf_mm = self._get_opticutter_kerf_mm()
                custom_stock_mm = self._get_opticutter_custom_stock_mm()
                analysis = analyse_profiles(
                    df,
                    kerf_mm=kerf_mm,
                    custom_stock_mm=custom_stock_mm,
                    manual_lengths=dict(self.opticutter_profile_custom_lengths),
                )
            else:
                analysis = analysis_override
                kerf_mm = getattr(analysis, "kerf_mm", self._get_opticutter_kerf_mm())
                custom_stock_mm = getattr(
                    analysis,
                    "custom_stock_mm",
                    self._get_opticutter_custom_stock_mm(),
                )
            self._apply_opticutter_analysis_state(analysis)
            self._opticutter_analysis_stale = False
            self._opticutter_needs_refresh = False
            if analysis is None:
                if info_var is not None:
                    info_var.set("Geen Opticutter-analyse beschikbaar.")
                self._update_opticutter_selection_rows([])
                return

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
            self._request_opticutter_refresh()
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

            self._request_opticutter_refresh()


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

        def _update_bom_file_status(self, exts):
            """Recalculate file status columns for the current BOM."""

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

        def _check_files(self):
            from tkinter import messagebox
            if not self._ensure_bom_loaded():
                return
            if not self.source_folder:
                messagebox.showwarning("Let op", "Selecteer een bronmap.")
                return
            exts = self._selected_exts()
            if not exts:
                messagebox.showwarning("Let op", "Selecteer minstens één bestandstype.")
                return
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
                messagebox.showwarning("Let op", "Selecteer bron, bestemming en extensies.")
                return
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

                if export_related_enabled and bom_source:
                    try:
                        for src_file in find_related_bom_exports(bom_source, idx):
                            if src_file in copied_paths:
                                continue
                            copied_paths.add(src_file)
                            transformed = _export_name(os.path.basename(src_file))
                            dst = os.path.join(bundle_dest, transformed)
                            shutil.copy2(src_file, dst)
                            related_copied += 1
                    except Exception:
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
            self._show_supplier_selection_tab(select_tab=True, prompt_opticutter=True)
            return

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

            self._ensure_opticutter_analysis_current()
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
                self._ensure_opticutter_analysis_current()
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
                group_map_raw: Dict[str, str],
                en1090_map_raw: Dict[str, bool],
                project_number: str,
                project_name: str,
                remember: bool,
                export_flags: Dict[str, bool] | None = None,
            ):
                if not self._ensure_bom_loaded():
                    return
                current_bom = self.bom_df
                client = self._current_client()
                self._ensure_opticutter_analysis_current()
                opticutter_analysis_current = getattr(
                    self, "opticutter_last_analysis", None
                )

                export_state_snapshot: Optional[SupplierSelectionState] = None
                if sel_frame is not None:
                    try:
                        if sel_frame.winfo_exists():
                            export_state_snapshot = sel_frame.serialize_state()
                            self._last_supplier_selection_state = export_state_snapshot
                    except Exception:
                        pass
                if export_state_snapshot is None:
                    export_state_snapshot = getattr(self, "_last_supplier_selection_state", None)

                prod_override_map: Dict[str, str] = {}
                finish_override_map: Dict[str, str] = {}
                opticutter_override_map: Dict[str, str] = {}
                production_pricing_map: Dict[str, Dict[str, object]] = {}
                finish_pricing_map: Dict[str, Dict[str, object]] = {}
                opticutter_pricing_map: Dict[str, Dict[str, object]] = {}
                production_vat_rate_map: Dict[str, str] = {}
                finish_vat_rate_map: Dict[str, str] = {}
                opticutter_vat_rate_map: Dict[str, str] = {}
                export_flags = export_flags or {}
                prod_export_filter: Dict[str, bool] = {}
                finish_export_filter: Dict[str, bool] = {}
                opticutter_export_filter: Dict[str, bool] = {}
                for key, value in sel_map.items():
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_override_map[identifier] = value
                    elif kind == "opticutter":
                        opticutter_override_map[identifier] = value
                    else:
                        prod_override_map[identifier] = value
                for key, enabled in export_flags.items():
                    kind, identifier = parse_selection_key(key)
                    target: Dict[str, bool] | None
                    if kind == "finish":
                        target = finish_export_filter
                    elif kind == "opticutter":
                        target = opticutter_export_filter
                    else:
                        target = prod_export_filter
                    target[identifier] = bool(enabled)

                pricing_map_raw = (
                    getattr(export_state_snapshot, "pricing", {}) if export_state_snapshot else {}
                )
                for key, value in pricing_map_raw.items():
                    if not isinstance(value, Mapping):
                        continue
                    clean_value = _clean_supplier_pricing_value(value)
                    if not clean_value:
                        continue
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_pricing_map[identifier] = clean_value
                    elif kind == "opticutter":
                        opticutter_pricing_map[identifier] = clean_value
                    else:
                        production_pricing_map[identifier] = clean_value

                vat_rate_map_raw = (
                    getattr(export_state_snapshot, "vat_rates", {}) if export_state_snapshot else {}
                )
                for key, value in vat_rate_map_raw.items():
                    clean_rate = _clean_supplier_vat_rate(value)
                    kind, identifier = parse_selection_key(key)
                    if kind == "finish":
                        finish_vat_rate_map[identifier] = clean_rate
                    elif kind == "opticutter":
                        opticutter_vat_rate_map[identifier] = clean_rate
                    else:
                        production_vat_rate_map[identifier] = clean_rate

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
                    resolved = self._resolve_delivery_choice(name, client)
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

                en1090_override_map: Dict[str, bool] = {}
                for key, flag in en1090_map_raw.items():
                    kind, identifier = parse_selection_key(key)
                    if kind not in {"production", "opticutter"}:
                        continue
                    norm = normalize_en1090_key(identifier)
                    if norm:
                        en1090_override_map[norm] = bool(flag)

                custom_prefix_text = self.export_name_custom_prefix_text.get().strip()
                custom_prefix_enabled = bool(
                    self.export_name_custom_prefix_enabled_var.get()
                )
                custom_suffix_text = self.export_name_custom_suffix_text.get().strip()
                custom_suffix_enabled = bool(
                    self.export_name_custom_suffix_enabled_var.get()
                )
                client_name_snapshot = _to_str(getattr(client, "name", "") if client else "").strip()
                bom_source_path_snapshot = _to_str(getattr(self, "bom_source_path", "")).strip()

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
                    except Exception:
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
                    path_limit_messages: List[str] = []
                    document_status_lines: List[str] = []
                    generated_document_records: List[Dict[str, object]] = []
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
                            document_filename_profile=self.document_filename_profile_var.get(),
                            document_filename_show_doc_type=bool(
                                self.document_filename_show_doc_type_var.get()
                            ),
                            document_filename_show_doc_number=bool(
                                self.document_filename_show_doc_number_var.get()
                            ),
                            document_filename_show_context=bool(
                                self.document_filename_show_context_var.get()
                            ),
                            document_filename_show_date=bool(
                                self.document_filename_show_date_var.get()
                            ),
                            document_filename_compact_doc_number=bool(
                                self.document_filename_compact_doc_number_var.get()
                            ),
                            document_filename_separator=self.document_filename_separator_var.get(),
                            document_display_compact_doc_number=bool(
                                self.document_display_compact_doc_number_var.get()
                            ),
                            finish_override_map=finish_override_map,
                            finish_doc_type_map=finish_doc_type_map,
                            finish_doc_num_map=finish_doc_num_map,
                            finish_delivery_map=finish_delivery_map,
                            remarks_map=production_remarks_map,
                            finish_remarks_map=finish_remarks_map,
                            document_group_map=group_map_raw if group_map_raw else None,
                            bom_source_path=self.bom_source_path,
                            path_limit_warnings=path_limit_messages,
                            opticutter_analysis=opticutter_analysis_snapshot,
                            opticutter_choices=opticutter_choices_snapshot,
                            opticutter_override_map=opticutter_override_map,
                            opticutter_doc_type_map=opticutter_doc_type_map,
                            opticutter_doc_num_map=opticutter_doc_num_map,
                            opticutter_delivery_map=opticutter_delivery_map,
                            opticutter_remarks_map=opticutter_remarks_map,
                            pricing_map=production_pricing_map or None,
                            finish_pricing_map=finish_pricing_map or None,
                            opticutter_pricing_map=opticutter_pricing_map or None,
                            vat_rate_map=production_vat_rate_map or None,
                            finish_vat_rate_map=finish_vat_rate_map or None,
                            opticutter_vat_rate_map=opticutter_vat_rate_map or None,
                            production_export_filter=(
                                prod_export_filter if prod_export_filter else None
                            ),
                            finish_export_filter=(
                                finish_export_filter if finish_export_filter else None
                            ),
                            opticutter_export_filter=(
                                opticutter_export_filter
                                if opticutter_export_filter
                                else None
                            ),
                            en1090_overrides=en1090_override_map or None,
                            en1090_enabled=bool(self.en1090_enabled_var.get()),
                            en1090_note=self.en1090_note_var.get(),
                            document_status_messages=document_status_lines,
                            generated_documents=generated_document_records,
                        )
                        if export_state_snapshot is not None:
                            try:
                                log_payload = build_export_session_log(
                                    project_number=project_number,
                                    project_name=project_name,
                                    client_name=client_name_snapshot,
                                    bom_source_path=bom_source_path_snapshot,
                                    bom_df=current_bom,
                                    state=export_state_snapshot,
                                    app_version=APP_VERSION,
                                    generated_documents=generated_document_records,
                                    status_messages=document_status_lines,
                                    path_limit_warnings=path_limit_messages,
                                )
                                write_export_session_log(bundle_dest, log_payload)
                                document_status_lines.append(
                                    f"Exportlog opgeslagen: {EXPORT_SESSION_LOG_FILENAME}"
                                )
                            except Exception as log_exc:
                                document_status_lines.append(
                                    f"Exportlog niet opgeslagen: {log_exc}"
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
                            if document_status_lines:
                                info_lines.append("")
                                info_lines.append("Details:")
                                info_lines.extend(
                                    f"- {line}" for line in document_status_lines
                                )
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
                    clients_db=self.client_db,
                    client_var=self.client_var,
                    presets_db=self.order_presets_db,
                    on_manage_presets=lambda: self.nb.select(self.preset_rules_frame),
                    opticutter_details=opticutter_details,
                    initial_state=previous_state,
                    en1090_enabled=bool(self.en1090_enabled_var.get()),
                    en1090_getter=self._get_en1090_preference,
                    en1090_setter=self._set_en1090_preference,
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
            settings_frame = getattr(self, "settings_frame", None)
            if settings_frame is not None:
                try:
                    settings_index = self.nb.index(settings_frame)
                except tk.TclError:
                    settings_index = None
            else:
                settings_index = None
            if settings_index is None:
                self.nb.add(sel_frame, text="Bestelbonnen")
            else:
                self.nb.insert(settings_index, sel_frame, text="Bestelbonnen")
            self.nb.select(sel_frame)

            if sup_search_restore and hasattr(sup_frame, "restore_search_filter"):
                def _restore_search(_event=None, frame=sup_frame, value=sup_search_restore):
                    try:
                        frame.restore_search_filter(value)
                    except Exception:
                        pass

                sel_frame.bind("<Destroy>", _restore_search, add="+")

        def _build_pdf_workdossier_tab(self) -> None:
            frame = self.pdf_workdossier_frame
            frame.columnconfigure(0, weight=1)
            frame.rowconfigure(0, weight=1)

            self.pdf_workdossier_options_frame = PdfWorkDossierOptionsFrame(
                frame,
                self.pdf_workdossier_presets_db,
                on_presets_changed=self._on_pdf_workdossier_presets_change,
                on_options_changed=self._schedule_pdf_workdossier_preview,
                on_prepare_orders=self._prepare_pdf_order_documents,
            )
            self.pdf_workdossier_options_frame.grid(row=0, column=0, sticky="nsew")
            self.pdf_workdossier_options_frame.subtabs.bind(
                "<<NotebookTabChanged>>",
                lambda _event: self._update_pdf_workdossier_actions_visibility(),
            )

            actions = tk.Frame(frame)
            self.pdf_workdossier_actions_frame = actions
            actions.grid(row=1, column=0, sticky="ew", pady=(10, 0))
            actions.columnconfigure(1, weight=1)
            self.pdf_workdossier_refresh_button = ttk.Button(
                actions,
                text="Tabel bijwerken",
                command=self._refresh_pdf_workdossier_preview,
            )
            self.pdf_workdossier_refresh_button.grid(row=0, column=0, sticky="w", padx=(0, 8))
            self.pdf_workdossier_progress_var = tk.DoubleVar(master=self, value=0)
            self.pdf_workdossier_progress = ttk.Progressbar(
                actions,
                variable=self.pdf_workdossier_progress_var,
                maximum=100,
                mode="determinate",
            )
            self.pdf_workdossier_progress.grid(row=0, column=1, sticky="ew", padx=(0, 8))
            self.pdf_workdossier_progress_label_var = tk.StringVar(master=self, value="0%")
            tk.Label(
                actions,
                textvariable=self.pdf_workdossier_progress_label_var,
                width=6,
                anchor="e",
            ).grid(row=0, column=2, sticky="e", padx=(0, 8))
            self.pdf_workdossier_export_button = tk.Button(
                actions,
                text="PDF dossier aanmaken",
                command=self._combine_pdf_from_tab,
                bg=MANUFACT_BRAND_COLOR,
                activebackground="#F4C46C",
                fg="black",
                activeforeground="black",
                highlightthickness=0,
                padx=14,
                pady=6,
            )
            self.pdf_workdossier_export_button.grid(row=0, column=3, sticky="e")
            self._update_pdf_workdossier_actions_visibility()
            self.after_idle(self._refresh_pdf_workdossier_preview)

        def _update_pdf_workdossier_actions_visibility(self) -> None:
            selected_tab = self.pdf_workdossier_options_frame.subtabs.select()
            work_tab_id = str(self.pdf_workdossier_options_frame.work_tab)
            if selected_tab == work_tab_id:
                self.pdf_workdossier_actions_frame.grid()
            else:
                self.pdf_workdossier_actions_frame.grid_remove()

        def _schedule_pdf_workdossier_preview(self, *_args) -> None:
            after_id = getattr(self, "_pdf_preview_after_id", None)
            if after_id is not None:
                try:
                    self.after_cancel(after_id)
                except tk.TclError:
                    pass
            self._pdf_preview_after_id = self.after(250, self._refresh_pdf_workdossier_preview)

        def _pdf_base_output_folder(self) -> str:
            return self.dest_folder_var.get().strip() or self.source_folder_var.get().strip()

        def _current_pdf_order_context_signature(self) -> Tuple[str, str, str, str, str]:
            def _abs(value: str) -> str:
                text = _to_str(value).strip()
                if not text:
                    return ""
                try:
                    return os.path.abspath(text)
                except (OSError, ValueError):
                    return text

            return (
                _abs(self.source_folder_var.get()),
                _abs(self._pdf_base_output_folder()),
                self.project_number_var.get().strip(),
                self.project_name_var.get().strip(),
                _abs(_to_str(self.bom_source_path)),
            )

        def _pdf_order_documents_root(self) -> str:
            base = self._pdf_base_output_folder()
            project_token = (
                _sanitize_component(self.project_number_var.get())
                or _sanitize_component(self.project_name_var.get())
                or "project"
            )
            return os.path.join(base, "PDF dossier documenten", project_token)

        def _reset_pdf_order_documents_root(
            self,
            folder: str,
            base_folder: str | None = None,
        ) -> None:
            root_base = _to_str(base_folder).strip() or self._pdf_base_output_folder()
            base = os.path.abspath(
                os.path.join(root_base, "PDF dossier documenten")
            )
            target = os.path.abspath(folder)
            try:
                common = os.path.commonpath([base, target])
            except ValueError as exc:
                raise RuntimeError("Ongeldige PDF dossier documentenmap.") from exc
            if common != base or target == base:
                raise RuntimeError("Ongeldige PDF dossier documentenmap.")
            if os.path.isdir(target):
                shutil.rmtree(target)
            os.makedirs(target, exist_ok=True)

        def _prepared_pdf_order_documents(self) -> Tuple[str, List[Dict[str, object]]]:
            root = _to_str(getattr(self, "_pdf_order_document_root", "")).strip()
            signature = getattr(self, "_pdf_order_context_signature", None)
            if not root or signature != self._current_pdf_order_context_signature():
                return "", []
            if not os.path.isdir(root):
                return "", []
            documents = getattr(self, "_pdf_generated_order_documents", []) or []
            return root, list(documents)

        def _pdf_export_name_preview(self, options: Mapping[str, object]) -> str:
            date_str = datetime.date.today().strftime("%Y-%m-%d")
            mode = _to_str(options.get("mode"))
            if mode == PdfWorkDossierOptionsFrame.MODE_PER_PRODUCTION:
                return f"<productie>_{date_str}_combined.pdf"
            return f"Werkdossier_{date_str}_combined.pdf"

        def _build_pdf_preview_items(
            self,
            options: Mapping[str, object],
        ) -> Tuple[List[PdfWorkDossierPlanItem], List[str], str]:
            source = self.source_folder_var.get().strip()
            bom_df = self.bom_df
            if not source:
                return [], [], "Selecteer een bronmap."
            if bom_df is None:
                return [], [], "Laad eerst een BOM."

            output_root = self._pdf_base_output_folder() or source
            prepared_order_root, generated_order_documents = (
                self._prepared_pdf_order_documents()
            )
            order_document_root = prepared_order_root or output_root
            mode = _to_str(options.get("mode"))
            if mode == PdfWorkDossierOptionsFrame.MODE_PER_PRODUCTION:
                idx = _build_file_index(source, [".pdf"])
                prod_to_files: Dict[str, List[str]] = defaultdict(list)
                for _, row in bom_df.iterrows():
                    production = _to_str(row.get("Production")).strip() or "_Onbekend"
                    part_number = _to_str(row.get("PartNumber")).strip()
                    if not part_number:
                        continue
                    prod_to_files[production].extend(idx.get(part_number, []))
                items: List[PdfWorkDossierPlanItem] = []
                for production in sorted(prod_to_files, key=lambda value: value.casefold()):
                    seen_paths: set[str] = set()
                    files: List[str] = []
                    for path in prod_to_files[production]:
                        key = os.path.abspath(path).casefold()
                        if key in seen_paths:
                            continue
                        seen_paths.add(key)
                        files.append(path)
                    for path in sorted(files, key=lambda value: os.path.basename(value).casefold()):
                        items.append(
                            PdfWorkDossierPlanItem(
                                path=path,
                                section_name=production,
                                production=production,
                                role="drawing",
                            )
                        )
                return items, [], ""

            include_order_documents = (
                mode == PdfWorkDossierOptionsFrame.MODE_WORKDOSSIER
                and bool(options.get("include_order_documents"))
            )
            items = build_pdf_workdossier_plan(
                source,
                bom_df,
                bom_source_path=self.bom_source_path,
                preset=options.get("preset")
                if mode == PdfWorkDossierOptionsFrame.MODE_WORKDOSSIER
                else None,
                include_order_documents=include_order_documents,
                order_document_root=order_document_root,
                include_offers=bool(options.get("include_offers")),
                generated_order_documents=generated_order_documents,
            )
            missing_orders: List[str] = []
            if include_order_documents:
                drawing_productions = {
                    _to_str(item.production).strip()
                    for item in items
                    if _to_str(item.role).strip().lower() == "drawing"
                    and _to_str(item.production).strip()
                }
                order_productions = {
                    _to_str(item.production).strip()
                    for item in items
                    if _to_str(item.role).strip().lower() == "order"
                    and _to_str(item.production).strip()
                }
                missing_orders = sorted(
                    drawing_productions - order_productions,
                    key=lambda value: value.casefold(),
                )
            return items, missing_orders, ""

        def _refresh_pdf_workdossier_preview(self) -> None:
            self._pdf_preview_after_id = None
            options_frame = getattr(self, "pdf_workdossier_options_frame", None)
            if options_frame is None:
                return
            options = options_frame.build_options()
            base_folder = self._pdf_base_output_folder()
            if base_folder:
                folder_text = f"Nieuwe Combined pdf-map in: {base_folder}"
            else:
                folder_text = "Selecteer bronmap of bestemmingsmap"
            options_frame.set_export_info(
                self._pdf_export_name_preview(options),
                folder_text,
            )
            try:
                items, missing_orders, error_message = self._build_pdf_preview_items(options)
            except Exception as exc:
                options_frame.clear_preview(f"Voorbeeld kon niet worden opgebouwd: {exc}")
                options_frame.set_order_flow_message("")
                return
            if error_message:
                options_frame.clear_preview(error_message)
                options_frame.set_order_flow_message("")
                return
            options_frame.set_preview_items(items)
            if missing_orders:
                shown = ", ".join(missing_orders[:8])
                if len(missing_orders) > 8:
                    shown += f", +{len(missing_orders) - 8}"
                options_frame.set_order_flow_message(
                    "Bestelbonnen/standaardbonnen ontbreken voor: "
                    f"{shown}. Gebruik 'Bestelbonnen klaarmaken' om de bon-PDF's "
                    "voor dit PDF dossier aan te maken.",
                    show_button=True,
                )
            else:
                options_frame.set_order_flow_message("")

        def _prepare_pdf_order_documents(self) -> None:
            if not self._ensure_bom_loaded():
                return
            self.status_var.set("Vul de bestelbonpagina aan voor het PDF dossier.")
            self._show_supplier_selection_tab(
                select_tab=True,
                prompt_opticutter=True,
                pdf_dossier_context=True,
            )

        def _handle_missing_pdf_order_documents(self, missing_orders: Sequence[str]) -> None:
            shown = ", ".join(list(missing_orders)[:10])
            if len(missing_orders) > 10:
                shown += f", +{len(missing_orders) - 10}"
            messagebox.showwarning(
                "Bestelbonnen ontbreken",
                (
                    "De optie om bestelbonnen/standaardbonnen in te voegen staat aan, "
                    "maar er ontbreken nog PDF-bonnen voor:\n\n"
                    f"{shown}\n\n"
                    "Maak eerst de bon-PDF's via de speciale PDF dossier-context; "
                    "daarna kom je automatisch terug naar deze pagina."
                ),
                parent=self,
            )
            self._prepare_pdf_order_documents()

        def _set_pdf_action_running(self, running: bool) -> None:
            self._pdf_action_running = bool(running)
            state = "disabled" if running else "normal"
            for attr in ("pdf_workdossier_export_button", "pdf_workdossier_refresh_button"):
                widget = getattr(self, attr, None)
                if widget is not None:
                    try:
                        widget.configure(state=state)
                    except tk.TclError:
                        pass
            options_frame = getattr(self, "pdf_workdossier_options_frame", None)
            if options_frame is not None:
                try:
                    options_frame.set_busy(running)
                except tk.TclError:
                    pass
            if running:
                self.pdf_workdossier_progress_var.set(0)
                self.pdf_workdossier_progress_label_var.set("0%")

        def _update_pdf_progress(self, done: int, total: int, path: str = "") -> None:
            percentage = 0
            if total > 0:
                percentage = max(0, min(100, int((done / total) * 100)))
            filename = os.path.basename(path) if path else ""

            def apply() -> None:
                self.pdf_workdossier_progress_var.set(percentage)
                self.pdf_workdossier_progress_label_var.set(f"{percentage}%")
                if filename:
                    self.status_var.set(f"PDF's combineren... {percentage}% - {filename}")
                else:
                    self.status_var.set(f"PDF's combineren... {percentage}%")

            self.after(0, apply)

        def _open_export_path(self, path: str) -> None:
            try:
                if sys.platform.startswith("win"):
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.run(["open", path], check=False)
                else:
                    subprocess.run(["xdg-open", path], check=False)
            except Exception as exc:
                messagebox.showwarning(
                    "Let op",
                    f"Kon pad niet openen:\n{exc}",
                    parent=self,
                )

        def _on_pdf_workdossier_presets_change(self) -> None:
            options_frame = getattr(self, "pdf_workdossier_options_frame", None)
            if options_frame is not None:
                try:
                    options_frame._reload_presets()
                except Exception:
                    pass
            self._schedule_pdf_workdossier_preview()

        def _combine_pdf_from_tab(self) -> None:
            options_frame = getattr(self, "pdf_workdossier_options_frame", None)
            if options_frame is None:
                return
            self._run_pdf_combine_with_options(options_frame.build_options())

        def _run_pdf_combine_with_options(self, options: Mapping[str, object]) -> None:
            if getattr(self, "_pdf_action_running", False):
                return
            if not self._ensure_bom_loaded():
                return
            bom_df = self.bom_df
            source_folder = self.source_folder_var.get().strip()
            out_dir = self.dest_folder_var.get().strip() or source_folder
            prepared_order_root, generated_order_documents = (
                self._prepared_pdf_order_documents()
            )
            order_document_root = prepared_order_root or out_dir
            if not source_folder or bom_df is None:
                messagebox.showwarning(
                    "Let op", "Selecteer bronmap en laad een BOM.", parent=self
                )
                return

            try:
                preview_items, missing_orders, preview_error = self._build_pdf_preview_items(options)
            except Exception as exc:
                messagebox.showerror(
                    "Fout",
                    f"Kon de PDF-volgorde niet voorbereiden:\n{exc}",
                    parent=self,
                )
                return
            self._refresh_pdf_workdossier_preview()
            if preview_error:
                messagebox.showwarning("Let op", preview_error, parent=self)
                return
            if missing_orders:
                self._handle_missing_pdf_order_documents(missing_orders)
                return
            if not preview_items:
                messagebox.showwarning(
                    "Let op",
                    "Er zijn geen PDF's gevonden om te combineren.",
                    parent=self,
                )
                return

            self._set_pdf_action_running(True)
            self.status_var.set("PDF's combineren... 0%")

            def work() -> None:
                try:
                    pn = self.project_number_var.get().strip() if self.project_number_var else ""
                    pname = self.project_name_var.get().strip() if self.project_name_var else ""
                    mode = _to_str(options.get("mode"))
                    if mode == PdfWorkDossierOptionsFrame.MODE_PER_PRODUCTION:
                        result = combine_pdfs_from_source(
                            source_folder,
                            bom_df,
                            out_dir,
                            project_number=pn or None,
                            project_name=pname or None,
                            combine_per_production=True,
                            bom_source_path=self.bom_source_path,
                            progress_callback=self._update_pdf_progress,
                        )
                    elif mode == PdfWorkDossierOptionsFrame.MODE_ALPHABETIC_SINGLE:
                        result = combine_workdossier_pdf_from_source(
                            source_folder,
                            bom_df,
                            out_dir,
                            project_number=pn or None,
                            project_name=pname or None,
                            bom_source_path=self.bom_source_path,
                            progress_callback=self._update_pdf_progress,
                        )
                    else:
                        result = combine_workdossier_pdf_from_source(
                            source_folder,
                            bom_df,
                            out_dir,
                            project_number=pn or None,
                            project_name=pname or None,
                            bom_source_path=self.bom_source_path,
                            preset=options.get("preset"),
                            include_order_documents=bool(
                                options.get("include_order_documents")
                            ),
                            order_document_root=order_document_root,
                            include_offers=bool(options.get("include_offers")),
                            generated_order_documents=generated_order_documents,
                            progress_callback=self._update_pdf_progress,
                        )
                except ModuleNotFoundError:
                    def on_missing_module() -> None:
                        self.status_var.set("PyPDF2 ontbreekt")
                        self._set_pdf_action_running(False)
                        messagebox.showwarning(
                            "PyPDF2 ontbreekt",
                            "Installeer PyPDF2 om PDF's te combineren.",
                            parent=self,
                        )

                    self.after(0, on_missing_module)
                    return
                except Exception as exc:
                    error_message = str(exc)

                    def on_error() -> None:
                        self.status_var.set("PDF combineren mislukt.")
                        self._set_pdf_action_running(False)
                        messagebox.showerror(
                            "Fout",
                            f"PDF combineren mislukt:\n{error_message}",
                            parent=self,
                        )

                    self.after(0, on_error)
                    return

                def on_done() -> None:
                    self.pdf_workdossier_progress_var.set(100)
                    self.pdf_workdossier_progress_label_var.set("100%")
                    self.status_var.set(
                        f"Gecombineerde pdf's: {result.count} -> {result.output_dir}"
                    )
                    output_files = list(getattr(result, "output_files", []) or [])
                    open_pdf = bool(options.get("open_pdf"))
                    target_to_open = (
                        output_files[0]
                        if open_pdf and len(output_files) == 1
                        else result.output_dir
                    )
                    message = "PDF's gecombineerd.\n\n" f"Map: {result.output_dir}"
                    if len(output_files) == 1:
                        message += f"\nPDF: {output_files[0]}"
                    self._set_pdf_action_running(False)
                    self._refresh_pdf_workdossier_preview()
                    messagebox.showinfo("Klaar", message, parent=self)
                    self._open_export_path(target_to_open)

                self.after(0, on_done)

            threading.Thread(target=work, daemon=True).start()

        def _select_pdf_workdossier_tab(self) -> None:
            try:
                self.nb.select(self.pdf_workdossier_frame)
            except tk.TclError:
                return
            options_frame = getattr(self, "pdf_workdossier_options_frame", None)
            if options_frame is not None:
                options_frame.select_work_tab()
            self._refresh_pdf_workdossier_preview()

        def _combine_pdf(self):
            self._select_pdf_workdossier_tab()
            return

    App().mainloop()

