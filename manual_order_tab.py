"""UI-tab voor het handmatig opstellen van bestel- of offertenbonnen."""

from __future__ import annotations

import math
import unicodedata
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple, TYPE_CHECKING

import tkinter as tk
from tkinter import font, messagebox, ttk

from orders import _prefix_for_doc_type

if TYPE_CHECKING:
    from clients_db import ClientsDB


def _normalize_numeric(value: str) -> object:
    """Try to convert ``value`` to ``int``/``float`` while respecting decimals."""

    text = value.strip().replace(",", ".")
    if not text:
        return ""
    try:
        number = float(text)
    except Exception:
        return value.strip()
    if math.isfinite(number):
        if number.is_integer():
            return int(number)
        return round(number, 4)
    return value.strip()


@dataclass
class _ManualRowWidgets:
    frame: tk.Frame
    vars: Dict[str, tk.StringVar]
    entries: Dict[str, tk.Entry]
    remove_btn: tk.Button


DEFAULT_MANUAL_CONTEXT = "Bestelbon-editor"


class SearchableCombobox(ttk.Combobox):
    """``ttk.Combobox`` variant met eenvoudige zoek/filter-functionaliteit."""

    def __init__(self, master: tk.Misc, *, values=(), **kwargs) -> None:
        self._all_values: List[str] = []
        self._normalized_values: List[Tuple[str, str]] = []
        self._last_query: str = ""
        super().__init__(master, values=values, **kwargs)
        self._store_all_values(list(values or ()))
        self.bind("<KeyRelease>", self._on_key_release, add="+")
        self.bind("<<ComboboxSelected>>", self._on_selection, add="+")
        self.bind("<FocusIn>", self._restore_values, add="+")
        self.bind("<FocusIn>", self._on_focus_in, add="+")
        self.bind("<FocusOut>", self._on_focus_out, add="+")

    # Public helpers -------------------------------------------------
    def set_choices(self, values: List[str]) -> None:
        """Update beschikbare opties en hergebruik de huidige filter."""

        self._store_all_values(values)
        # Houd de huidige invoer ongemoeid wanneer opties opnieuw geladen
        # worden, zodat reeds ingevulde waarden zichtbaar blijven.
        self._apply_filter(self._last_query, update_entry=False)

    # Internal -------------------------------------------------------
    def _store_all_values(self, values: List[str]) -> None:
        self._all_values = list(values)
        self._normalized_values = [
            (option, self._normalize_text(option)) for option in self._all_values
        ]

    def _apply_filter(self, query: str, *, update_entry: bool = True) -> None:
        display_text = query
        filtered = self._filter_values(query)
        self.configure(values=filtered)
        self._last_query = query
        if not update_entry:
            return
        if filtered:
            # Toon de dropdown zodat de gebruiker direct kan kiezen
            try:
                self.tk.call("ttk::combobox::Post", self._w)
            except Exception:
                try:
                    self.event_generate("<Down>")
                except Exception:
                    pass

            def _restore_entry() -> None:
                try:
                    self.delete(0, tk.END)
                    self.insert(0, display_text)
                    self.selection_clear()
                    self.icursor(tk.END)
                    self.focus_set()
                except Exception:
                    pass

            self.after_idle(_restore_entry)

    def _on_key_release(self, event: tk.Event) -> None:
        if event.keysym in {"Up", "Down", "Return", "Escape", "Tab"}:
            if event.keysym == "Escape":
                self._restore_values()
            return
        self.after_idle(lambda: self._apply_filter(self.get()))

    def _on_selection(self, _event: tk.Event) -> None:
        self._restore_values()
        self._clear_text_selection()
        self._unpost_dropdown()

    def _restore_values(self, _event: tk.Event | None = None) -> None:
        self.configure(values=self._all_values)
        self._last_query = ""

    def _on_focus_in(self, _event: tk.Event) -> None:
        def _select_all() -> None:
            try:
                self.selection_range(0, tk.END)
            except Exception:
                pass
            try:
                self.icursor(tk.END)
            except Exception:
                pass

        self.after_idle(_select_all)

    def _on_focus_out(self, _event: tk.Event) -> None:
        self._unpost_dropdown()

    def _clear_text_selection(self) -> None:
        try:
            self.selection_clear()
        except Exception:
            pass
        try:
            self.icursor(tk.END)
        except Exception:
            pass

    def _unpost_dropdown(self) -> None:
        try:
            self.tk.call("ttk::combobox::Unpost", self._w)
        except Exception:
            try:
                self.event_generate("<Escape>")
            except Exception:
                pass

    @staticmethod
    def _normalize_text(value: object) -> str:
        text = str(value or "")
        normalized = unicodedata.normalize("NFKD", text)
        normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
        normalized = normalized.casefold()
        return " ".join(normalized.split())

    def _filter_values(self, query: str) -> List[str]:
        tokens = self._normalize_text(query).split()
        if not tokens:
            return self._all_values
        matches: List[str] = []
        for option, normalized in self._normalized_values:
            if all(token in normalized for token in tokens):
                matches.append(option)
        return matches


class ManualOrderTab(tk.Frame):
    """Tab om handmatige orderregels in te geven."""

    DEFAULT_CONTEXT_LABEL = DEFAULT_MANUAL_CONTEXT
    DEFAULT_TEMPLATE = "Standaard"
    COLUMN_MIN_CHARS = 2  # Allow very small columns for compact layouts
    COLUMN_MAX_CHARS = 72
    COLUMN_SEPARATOR_COLOR = "#B9BEC7"
    COLUMN_SEPARATOR_ACTIVE_COLOR = "#6E7681"
    COLUMN_TEMPLATES: Dict[str, List[Dict[str, object]]] = {
        "Standaard": [
            {
                "key": "PartNumber",
                "label": "Artikel nr.",
                "width": 16,
                "justify": "left",
                "stretch": False,
                "wrap": False,
                "weight": 1.6,
            },
            {
                "key": "Description",
                "label": "Omschrijving",
                "width": 32,
                "justify": "left",
                "stretch": True,
                "wrap": True,
                "weight": 2.6,
            },
            {
                "key": "Materiaal",
                "label": "Materiaal",
                "width": 18,
                "justify": "left",
                "stretch": False,
                "wrap": False,
                "weight": 1.6,
            },
            {
                "key": "Aantal",
                "label": "Aantal",
                "width": 8,
                "justify": "right",
                "numeric": True,
                "stretch": False,
                "wrap": False,
                "weight": 0.8,
            },
            {
                "key": "Oppervlakte",
                "label": "Oppervlakte",
                "width": 10,
                "justify": "right",
                "numeric": True,
                "stretch": False,
                "wrap": False,
                "weight": 1.0,
            },
            {
                "key": "Gewicht",
                "label": "Gewicht (kg)",
                "width": 10,
                "justify": "right",
                "numeric": True,
                "stretch": False,
                "wrap": False,
                "weight": 1.0,
                "total_weight": True,
            },
        ],
        "Spare parts": [
            {
                "key": "Artikel",
                "label": "Artikel",
                "width": 20,
                "justify": "left",
                "stretch": True,
                "wrap": True,
                "weight": 2.0,
            },
            {
                "key": "Supplier",
                "label": "Supplier",
                "width": 18,
                "justify": "left",
                "stretch": False,
                "wrap": True,
                "weight": 1.6,
            },
            {
                "key": "SupplierCode",
                "label": "Supplier code",
                "width": 16,
                "justify": "left",
                "stretch": False,
                "wrap": False,
                "weight": 1.4,
            },
            {
                "key": "Manufacturer",
                "label": "Manufacturer",
                "width": 20,
                "justify": "left",
                "stretch": True,
                "wrap": True,
                "weight": 2.0,
            },
            {
                "key": "ManufacturerCode",
                "label": "Manufacturer code",
                "width": 16,
                "justify": "left",
                "stretch": False,
                "wrap": False,
                "weight": 1.4,
            },
            {
                "key": "Aantal",
                "label": "Aantal",
                "width": 8,
                "justify": "right",
                "numeric": True,
                "stretch": False,
                "wrap": False,
                "weight": 0.8,
            },
            {
                "key": "PrijsPerStuk",
                "label": "Prijs/st",
                "width": 10,
                "justify": "right",
                "numeric": True,
                "stretch": False,
                "wrap": False,
                "weight": 1.0,
            },
        ],
        "Profielen": [
            {
                "key": "ArtikelNummer",
                "label": "Artikel nr.",
                "width": 16,
                "justify": "left",
                "stretch": False,
                "wrap": False,
                "weight": 1.4,
            },
            {
                "key": "ProfielType",
                "label": "Profiel type",
                "width": 18,
                "justify": "left",
                "stretch": True,
                "wrap": True,
                "weight": 2.0,
            },
            {
                "key": "Materiaal",
                "label": "Materiaal",
                "width": 18,
                "justify": "left",
                "stretch": True,
                "wrap": False,
                "weight": 1.8,
            },
            {
                "key": "ProfielLengte",
                "label": "Profiel lengte",
                "width": 14,
                "justify": "right",
                "numeric": True,
                "stretch": False,
                "wrap": False,
                "weight": 1.2,
            },
            {
                "key": "Aantal",
                "label": "Aantal",
                "width": 8,
                "justify": "right",
                "numeric": True,
                "stretch": False,
                "wrap": False,
                "weight": 1.0,
            },
        ],
    }

    DOC_TYPE_OPTIONS: tuple[str, ...] = ("Bestelbon", "Standaard bon", "Offerteaanvraag")
    DELIVERY_PRESETS: tuple[str, ...] = (
        "Geen",
        "Bestelling wordt opgehaald",
        "Leveradres wordt nog meegedeeld",
    )

    def __init__(
        self,
        master: tk.Misc,
        *,
        suppliers_db,
        delivery_db,
        clients_db: Optional["ClientsDB"] = None,
        project_number_var: tk.StringVar,
        project_name_var: tk.StringVar,
        client_var: Optional[tk.StringVar] = None,
        dest_folder_var: Optional[tk.StringVar] = None,
        on_export: Callable[[Dict[str, object]], None],
        on_manage_clients: Optional[Callable[[], None]] = None,
        on_manage_suppliers: Optional[Callable[[], None]] = None,
        on_manage_deliveries: Optional[Callable[[], None]] = None,
    ) -> None:
        super().__init__(master)
        self.suppliers_db = suppliers_db
        self.delivery_db = delivery_db
        self.clients_db = clients_db
        self.project_number_var = project_number_var
        self.project_name_var = project_name_var
        self.client_var = client_var or tk.StringVar()
        self.dest_folder_var = dest_folder_var or tk.StringVar()
        self._on_export = on_export
        self._on_manage_suppliers = on_manage_suppliers
        self._on_manage_deliveries = on_manage_deliveries
        self._context_user_modified = False

        self.configure(padx=12, pady=12)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = tk.LabelFrame(self, text="Documentgegevens", labelanchor="n")
        header.configure(padx=12, pady=12)
        header.grid(row=0, column=0, sticky="nsew")

        field_width_px = int(self.winfo_fpixels("6c"))
        manage_spacing_px = int(self.winfo_fpixels("3m"))
        base_font = font.nametofont("TkDefaultFont")
        self._header_font = base_font.copy()
        try:
            self._header_font.configure(weight="bold")
        except Exception:
            pass
        char_width = max(1, base_font.measure("0"))
        field_char_width = max(1, round(field_width_px / char_width))
        self._entry_char_pixels = char_width
        self.profile_material_chars = field_char_width

        header.columnconfigure(1, weight=0, minsize=field_width_px)
        header.columnconfigure(2, weight=0)
        header.columnconfigure(3, weight=0)
        header.columnconfigure(4, weight=1)

        self.doc_type_var = tk.StringVar(value=self.DOC_TYPE_OPTIONS[0])
        tk.Label(header, text="Documenttype:").grid(row=0, column=0, sticky="w")
        self.doc_type_combo = ttk.Combobox(
            header,
            textvariable=self.doc_type_var,
            values=self.DOC_TYPE_OPTIONS,
            state="readonly",
            width=field_char_width,
            takefocus=False,
            exportselection=False,
        )
        self.doc_type_combo.grid(row=0, column=1, sticky="w", padx=(6, 0))

        info_spacing_px = int(self.winfo_fpixels("1c"))
        header.columnconfigure(2, minsize=info_spacing_px)
        tk.Label(header, text="Projectnummer:").grid(
            row=0, column=3, sticky="w"
        )
        tk.Entry(
            header,
            textvariable=self.project_number_var,
            width=field_char_width,
        ).grid(row=0, column=4, sticky="w", padx=(6, 0))

        tk.Label(header, text="Documentnummer:").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        self.doc_number_var = tk.StringVar()
        self.doc_number_entry = tk.Entry(
            header,
            textvariable=self.doc_number_var,
            width=field_char_width,
        )
        self.doc_number_entry.grid(row=1, column=1, sticky="w", padx=(6, 0), pady=(6, 0))

        tk.Label(header, text="Projectnaam:").grid(
            row=1, column=3, sticky="w", pady=(6, 0)
        )
        tk.Entry(
            header,
            textvariable=self.project_name_var,
            width=field_char_width,
        ).grid(row=1, column=4, sticky="w", padx=(6, 0), pady=(6, 0))

        self._doc_number_prefix = _prefix_for_doc_type(self.doc_type_var.get())
        if self._doc_number_prefix:
            self.doc_number_var.set(self._doc_number_prefix)

        def _handle_doc_type_change(*_args):
            old_prefix = getattr(self, "_doc_number_prefix", "")
            new_prefix = _prefix_for_doc_type(self.doc_type_var.get())
            current = self.doc_number_var.get().strip()
            if not current:
                if new_prefix:
                    self.doc_number_var.set(new_prefix)
            elif old_prefix and current.startswith(old_prefix):
                remainder = current[len(old_prefix) :].lstrip(" -_")
                if new_prefix:
                    self.doc_number_var.set(new_prefix + remainder)
                else:
                    self.doc_number_var.set(remainder)
            elif current == old_prefix and new_prefix:
                self.doc_number_var.set(new_prefix)
            self._doc_number_prefix = new_prefix

        self.doc_type_var.trace_add("write", _handle_doc_type_change)

        def _reset_doc_type_focus() -> None:
            try:
                self.doc_type_combo.selection_clear(0, tk.END)
            except Exception:
                pass
            try:
                self.doc_type_combo.icursor(tk.END)
            except Exception:
                pass

        self.after_idle(self.doc_number_entry.focus_set)
        self.after_idle(_reset_doc_type_focus)

        tk.Label(header, text="Opdrachtgever:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        client_field = tk.Frame(header)
        client_field.grid(
            row=2,
            column=1,
            columnspan=4,
            sticky="w",
            padx=(6, 0),
            pady=(8, 0),
        )
        self.client_combo = ttk.Combobox(
            client_field,
            textvariable=self.client_var,
            width=field_char_width,
            state="readonly",
        )
        self.client_combo.pack(side="left")
        if on_manage_clients:
            tk.Button(
                client_field,
                text="Beheer",
                command=on_manage_clients,
                width=10,
            ).pack(side="left", padx=(manage_spacing_px, 0))

        tk.Label(header, text="Bestemmingsmap:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        dest_field = tk.Frame(header)
        dest_field.grid(
            row=3,
            column=1,
            columnspan=4,
            sticky="w",
            padx=(6, 0),
            pady=(8, 0),
        )
        self.dest_entry = tk.Entry(
            dest_field,
            textvariable=self.dest_folder_var,
            width=field_char_width,
        )
        self.dest_entry.pack(side="left")
        tk.Button(
            dest_field,
            text="Bladeren",
            command=self._pick_dest_folder,
            width=10,
        ).pack(side="left", padx=(manage_spacing_px, 0))

        tk.Label(header, text="Leverancier:").grid(row=4, column=0, sticky="w", pady=(8, 0))
        self.supplier_var = tk.StringVar()
        supplier_field = tk.Frame(header)
        supplier_field.grid(
            row=4,
            column=1,
            columnspan=4,
            sticky="w",
            padx=(6, 0),
            pady=(8, 0),
        )
        self.supplier_combo = ttk.Combobox(
            supplier_field,
            textvariable=self.supplier_var,
            width=field_char_width,
            state="readonly",
        )
        self.supplier_combo.pack(side="left")
        if on_manage_suppliers:
            tk.Button(
                supplier_field,
                text="Beheer",
                command=on_manage_suppliers,
                width=10,
            ).pack(side="left", padx=(manage_spacing_px, 0))

        tk.Label(header, text="Leveradres:").grid(row=5, column=0, sticky="w", pady=(6, 0))
        self.delivery_var = tk.StringVar()
        delivery_field = tk.Frame(header)
        delivery_field.grid(
            row=5,
            column=1,
            columnspan=4,
            sticky="w",
            padx=(6, 0),
            pady=(6, 0),
        )
        self.delivery_combo = ttk.Combobox(
            delivery_field,
            textvariable=self.delivery_var,
            width=field_char_width,
            state="readonly",
        )
        self.delivery_combo.pack(side="left")
        if on_manage_deliveries:
            tk.Button(
                delivery_field,
                text="Beheer",
                command=on_manage_deliveries,
                width=10,
            ).pack(side="left", padx=(manage_spacing_px, 0))

        tk.Label(header, text="Documentnaam:").grid(row=6, column=0, sticky="w", pady=(6, 0))
        self.context_label_var = tk.StringVar(
            value=self.project_name_var.get().strip() or self.DEFAULT_CONTEXT_LABEL
        )
        context_entry = tk.Entry(
            header, textvariable=self.context_label_var, width=field_char_width
        )
        context_entry.grid(
            row=6,
            column=1,
            sticky="w",
            padx=(6, 0),
            pady=(6, 0),
        )

        def _mark_context_modified(*_args):
            self._context_user_modified = True

        self.context_label_var.trace_add("write", _mark_context_modified)

        def _sync_project_name(*_args):
            if not self._context_user_modified:
                self.context_label_var.set(
                    self.project_name_var.get().strip() or self.DEFAULT_CONTEXT_LABEL
                )

        self.project_name_var.trace_add("write", _sync_project_name)

        tk.Label(header, text="Opmerkingen:").grid(row=7, column=0, sticky="nw", pady=(8, 0))
        self.remark_text = tk.Text(header, height=4, wrap="word")
        self.remark_text.grid(
            row=7,
            column=1,
            columnspan=4,
            sticky="nsew",
            padx=(6, 0),
            pady=(8, 0),
        )
        header.rowconfigure(7, weight=1)

        table_container = tk.Frame(self)
        table_container.grid(row=1, column=0, sticky="nsew", padx=4, pady=(8, 0))
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(2, weight=1)

        template_row = tk.Frame(table_container)
        template_row.grid(row=0, column=0, columnspan=2, sticky="ew", padx=4, pady=(0, 6))
        tk.Label(template_row, text="Sjabloon:").pack(side="left")
        self.template_var = tk.StringVar(value=self.DEFAULT_TEMPLATE)

        template_style = "Manual.Template.TCombobox"
        style = ttk.Style(template_row)
        base_foreground = style.lookup("TCombobox", "foreground") or "black"
        base_background = self.cget("background")
        style.configure(
            template_style,
            foreground=base_foreground,
            fieldbackground=base_background,
            background=base_background,
        )
        style.map(
            template_style,
            fieldbackground=[("readonly", base_background)],
            background=[("readonly", base_background)],
            selectbackground=[("readonly", base_background)],
            selectforeground=[("readonly", base_foreground)],
        )

        self.template_combo = ttk.Combobox(
            template_row,
            textvariable=self.template_var,
            values=tuple(self.COLUMN_TEMPLATES.keys()),
            state="readonly",
            width=max(14, field_char_width // 2),
            takefocus=False,
            exportselection=False,
            style=template_style,
        )
        self.template_combo.pack(side="left", padx=(6, 0))

        def _reset_template_focus(_event: tk.Event | None = None) -> None:
            try:
                self.template_combo.selection_clear(0, tk.END)
            except Exception:
                pass
            try:
                self.template_combo.icursor(tk.END)
            except Exception:
                pass

        self.template_combo.bind("<FocusIn>", _reset_template_focus, add="+")
        self.template_combo.bind("<<ComboboxSelected>>", _reset_template_focus, add="+")
        self.template_combo.bind("<ButtonRelease-1>", _reset_template_focus, add="+")

        self.header_container = tk.Frame(table_container)
        self.header_container.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(6, 0), pady=(0, 0))
        self.header_container.columnconfigure(0, weight=1)

        # Verticale en horizontale scrollbars
        self.table_canvas = tk.Canvas(table_container, highlightthickness=0)
        self.table_canvas.grid(row=2, column=0, sticky="nsew", padx=(0, 0))
        
        v_scroll = ttk.Scrollbar(
            table_container, orient="vertical", command=self.table_canvas.yview
        )
        v_scroll.grid(row=2, column=1, sticky="ns")
        
        h_scroll = ttk.Scrollbar(
            table_container, orient="horizontal", command=self.table_canvas.xview
        )
        h_scroll.grid(row=3, column=0, sticky="ew")
        
        self.table_canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.rows_frame = tk.Frame(self.table_canvas)
        self.rows_window = self.table_canvas.create_window(
            (0, 0), window=self.rows_frame, anchor="nw"
        )
        self.rows_frame.columnconfigure(1, weight=1)
        
        # Initialiseer kolom-resize handles lijst
        self._column_resizer_handles = []

        def _update_scrollregion(_event=None):
            self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))

        self.rows_frame.bind("<Configure>", _update_scrollregion)

        def _resize_canvas(event: tk.Event) -> None:
            try:
                self.table_canvas.itemconfigure(self.rows_window, width=event.width)
            except Exception:
                pass

        self.table_canvas.bind("<Configure>", _resize_canvas)
        
        # Bind mouse wheel events for horizontal scrolling with Shift
        def _on_mousewheel(event):
            if event.state & 0x1:  # Shift key pressed
                delta = -1 if event.delta > 0 else 1
                self.table_canvas.xview_scroll(delta, "units")
                return "break"
        
        self.table_canvas.bind("<MouseWheel>", _on_mousewheel)
        self.table_canvas.bind("<Shift-MouseWheel>", _on_mousewheel)
        
        # Also bind to arrow keys for horizontal scroll when canvas has focus
        def _on_arrow_key(event):
            if event.keysym == "Left":
                self.table_canvas.xview_scroll(-3, "units")
                return "break"
            elif event.keysym == "Right":
                self.table_canvas.xview_scroll(3, "units")
                return "break"
        
        self.table_canvas.bind("<Left>", _on_arrow_key)
        self.table_canvas.bind("<Right>", _on_arrow_key)

        controls = tk.Frame(table_container)
        controls.grid(row=4, column=0, columnspan=2, sticky="ew", padx=4, pady=(8, 0))
        controls.columnconfigure(3, weight=1)

        tk.Label(controls, text="Nieuwe regels:").grid(row=0, column=0, sticky="w")
        self.add_count_var = tk.StringVar(value="1")
        add_count_entry = tk.Entry(
            controls, textvariable=self.add_count_var, width=6, justify="right"
        )
        add_count_entry.grid(row=0, column=1, sticky="w", padx=(4, 10))
        add_count_entry.bind("<Return>", lambda _e: self.add_rows_from_input())

        tk.Button(
            controls,
            text="Regels toevoegen",
            command=self.add_rows_from_input,
        ).grid(row=0, column=2, sticky="w")

        self.total_weight_var = tk.StringVar(value="Totaal gewicht: —")
        tk.Label(controls, textvariable=self.total_weight_var, anchor="e").grid(
            row=0, column=3, sticky="e"
        )

        footer = tk.Frame(self)
        footer.grid(row=2, column=0, sticky="ew", padx=4, pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        tk.Button(footer, text="Bestelbon opslaan", command=self._handle_export).grid(
            row=0, column=1, sticky="e"
        )

        self.rows: List[_ManualRowWidgets] = []
        self.current_template_name: str = ""
        self.current_columns: List[Dict[str, object]] = []
        self._template_rows_cache: Dict[str, List[Dict[str, str]]] = {}
        self._template_layout_cache: Dict[str, List[Dict[str, object]]] = {}
        self._column_resizer_handles: List[tk.Widget] = []
        self._column_resize_state: Optional[Dict[str, object]] = None
        self._resizer_update_job: Optional[str] = None
        self._row_grid_indices: Dict[int, int] = {}  # Maps rows-list index to grid row number

        # Maak EERST de header-rij IN de canvas (voor alle andere inits!)
        self._create_header_row_in_canvas()

        def _handle_template_change(*_args) -> None:
            self._apply_template(self.template_var.get())

        self.template_var.trace_add("write", _handle_template_change)

        self._apply_template(self.template_var.get(), store_previous=False)

        try:
            self.after_idle(self.refresh_data)
        except Exception:
            self.refresh_data()

    # Public helpers -------------------------------------------------
    def refresh_data(self) -> None:
        """Reload client/supplier/delivery options from databases."""

        if self.clients_db is not None:
            client_opts = [
                self.clients_db.display_name(c)
                for c in self.clients_db.clients_sorted()
            ]
        else:
            client_opts = []
        current_client = self.client_var.get()
        self.client_combo.configure(values=client_opts)
        if current_client not in client_opts:
            if client_opts:
                self.client_var.set(client_opts[0])
            else:
                self.client_var.set("")

        supplier_opts = ["Geen"]
        if self.suppliers_db is not None:
            supplier_opts.extend(
                self.suppliers_db.display_name(s)
                for s in self.suppliers_db.suppliers_sorted()
            )
        current_supplier = self.supplier_var.get()
        self.supplier_combo.configure(values=supplier_opts)
        if current_supplier not in supplier_opts:
            self.supplier_var.set("Geen")

        delivery_opts = list(self.DELIVERY_PRESETS)
        if self.delivery_db is not None:
            delivery_opts.extend(
                self.delivery_db.display_name(a)
                for a in self.delivery_db.addresses_sorted()
            )
        current_delivery = self.delivery_var.get()
        self.delivery_combo.configure(values=delivery_opts)
        if current_delivery not in delivery_opts:
            self.delivery_var.set(self.DELIVERY_PRESETS[0])

    def set_doc_number(self, value: str) -> None:
        self.doc_number_var.set(value)
        self._doc_number_prefix = _prefix_for_doc_type(self.doc_type_var.get())

    # Row management -------------------------------------------------
    def _create_header_row_in_canvas(self) -> None:
        """Creëer de header-rij als rij 0 in rows_frame centrale grid."""
        # rows_frame gebruikt ONE centrale grid waar ALLES in zit
        # Col 0: Delete-knop/spacer (width=20)
        # Col 1,3,5,...: Data kolommen (weight varies)
        # Col 2,4,6,...: Separators (width=2)
        
        # Header delete-spacer in kolom 0
        header_delete = tk.Label(self.rows_frame, text="", width=2)
        header_delete.grid(row=0, column=0, sticky="w")
        self.rows_frame.columnconfigure(0, weight=0, minsize=20)
        
        # Header-widgets worden direct in rows_frame geplaatst (niet in nested frame)
        self.header_row = None  # We don't use a separate container anymore
        self._header_labels = {}  # dict {column_index: tk.Label}
        self._header_separators = []  # list van separator frames
        self._next_data_row = 1
    
    def add_row(self, values: Optional[Dict[str, object]] = None) -> None:
        # Maak een button-frame voor delete/copy/add knoppen
        buttons_frame = tk.Frame(self.rows_frame)
        row_idx = self._next_data_row
        buttons_frame.grid(row=row_idx, column=0, sticky="w")
        
        # Bepaal de rij-index in self.rows (dit is de lengte voordat we toevoegen)
        row_list_idx = len(self.rows)
        
        # Delete-knop met correct index
        remove_btn = tk.Button(
            buttons_frame,
            text="✕",
            width=2,
            bg="#ff6b6b",
            fg="white",
            command=lambda idx=row_list_idx: self._safe_delete_row(idx),
        )
        remove_btn.pack(side="left", padx=(0, 2))
        
        # Copy-knop met correct index
        copy_btn = tk.Button(
            buttons_frame,
            text="⧉",
            width=2,
            bg="#4ecdc4",
            fg="white",
            command=lambda idx=row_list_idx: self._copy_row(idx),
        )
        copy_btn.pack(side="left", padx=(0, 2))
        
        # Add-knop (nieuwe rij toevoegen)
        add_btn = tk.Button(
            buttons_frame,
            text="+",
            width=2,
            bg="#51cf66",
            fg="white",
            command=lambda: self.add_row(),
        )
        add_btn.pack(side="left", padx=(0, 0))
        
        # Data entries en separators direkt in rows_frame (GEEN nested frame!)
        widgets = _ManualRowWidgets(
            frame=buttons_frame,  # Store the button frame
            vars={},
            entries={},
            remove_btn=remove_btn,
        )

        for idx, column in enumerate(self.current_columns):
            grid_col = 1 + idx * 2  # Kolom 1, 3, 5, 7, ...
            
            # Create entry widget
            var = tk.StringVar()
            if values is not None and column["key"] in values:
                value = values[column["key"]]
                var.set("" if value is None else str(value))
            
            display_chars, min_width_px = self._column_display_metrics(column)
            entry = tk.Entry(
                self.rows_frame,
                textvariable=var,
                width=display_chars,
                justify=column.get("justify", "left"),
            )
            entry.grid(row=row_idx, column=grid_col, sticky="ew", padx=(6, 6))
            
            # Configure column width
            if column.get("stretch"):
                weight = 1
            else:
                weight = 0
            self.rows_frame.columnconfigure(grid_col, weight=weight, minsize=min_width_px)
            
            # Add separator BETWEEN columns (not after last)
            if idx < len(self.current_columns) - 1:
                sep_col = grid_col + 1  # Kolom 2, 4, 6, 8, ...
                separator = tk.Frame(
                    self.rows_frame,
                    width=2,
                    background=self.COLUMN_SEPARATOR_COLOR,
                )
                separator.grid(row=row_idx, column=sep_col, sticky="ns", padx=0)
                self.rows_frame.columnconfigure(sep_col, weight=0, minsize=2)
            
            var.trace_add("write", lambda *_args: self._update_totals())
            widgets.vars[column["key"]] = var
            widgets.entries[column["key"]] = entry
        
        self.rows.append(widgets)
        self._row_grid_indices[row_list_idx] = row_idx  # Track grid row voor deze data row
        self._next_data_row += 1
        
        if self.current_columns:
            first_key = self.current_columns[0]["key"]
            self.after_idle(lambda: widgets.entries[first_key].focus_set())
        self._update_totals()

    def remove_row(self, row_idx: int) -> None:
        """Remove a data row by its index in self.rows."""
        if not (0 <= row_idx < len(self.rows)):
            return
        
        row = self.rows[row_idx]
        self.rows.pop(row_idx)
        
        # Destroy button frame
        try:
            if row.frame is not None:
                row.frame.destroy()
        except Exception:
            pass
        
        # Destroy all entry widgets in this row
        for entry in row.entries.values():
            try:
                entry.destroy()
            except Exception:
                pass
        
        # Remove grid row tracking
        grid_row = self._row_grid_indices.pop(row_idx, None)
        
        # Ensure at least one empty row exists
        if len(self.rows) == 0:
            self.add_row()
        
        self._update_totals()
    
    def _safe_delete_row(self, row_idx: int) -> None:
        """Delete row with confirmation."""
        if not (0 <= row_idx < len(self.rows)):
            return
        
        if messagebox.askyesno("Rij verwijderen", f"Weet je zeker dat je rij {row_idx + 1} wilt verwijderen?"):
            self.remove_row(row_idx)
    
    def _copy_row(self, row_idx: int) -> None:
        """Duplicate a row."""
        if not (0 <= row_idx < len(self.rows)):
            return
        
        # Get values from source row
        source_row = self.rows[row_idx]
        source_values = {}
        for key, var in source_row.vars.items():
            source_values[key] = var.get()
        
        # Add new row with same values
        self.add_row(values=source_values)
        self._update_totals()

    def add_rows_from_input(self) -> None:
        text = self.add_count_var.get().strip()
        try:
            desired = int(text)
        except Exception:
            desired = 1
        desired = max(1, min(desired, 500))
        for _ in range(desired):
            self.add_row()

    # Data collection ------------------------------------------------
    def _collect_items(self) -> Dict[str, object]:
        items: List[Dict[str, object]] = []
        total_weight = 0.0
        weight_found = False
        numeric_keys = {col["key"] for col in self.current_columns if col.get("numeric")}
        weight_columns = [col["key"] for col in self.current_columns if col.get("total_weight")]

        for widgets in self.rows:
            raw = {key: var.get().strip() for key, var in widgets.vars.items()}
            if not any(raw.values()):
                continue
            record: Dict[str, object] = {}
            for column in self.current_columns:
                key = column["key"]
                value = raw.get(key, "")
                if key in numeric_keys:
                    normalized = _normalize_numeric(value)
                else:
                    normalized = value
                record[key] = normalized
            if weight_columns:
                weight_key = weight_columns[0]
                weight_raw = raw.get(weight_key, "")
            else:
                weight_raw = ""
            if weight_raw:
                try:
                    weight_total = float(weight_raw.replace(",", "."))
                except Exception:
                    weight_total = None
                if weight_total is not None and math.isfinite(weight_total):
                    total_weight += weight_total
                    weight_found = True
            items.append(record)

        return {
            "items": items,
            "total_weight": total_weight if weight_found else None,
        }

    # Export ---------------------------------------------------------
    def _pick_dest_folder(self) -> None:
        """Open a folder picker to select the destination folder."""
        from tkinter import filedialog
        p = filedialog.askdirectory(
            title="Kies bestemmingsmap voor handmatige bestelbon"
        )
        if p:
            self.dest_folder_var.set(p)

    def _handle_export(self) -> None:
        payload = self._collect_items()
        items: List[Dict[str, object]] = payload["items"]
        if not items:
            messagebox.showwarning(
                "Geen gegevens",
                "Voeg minstens één regel met gegevens toe voordat je exporteert.",
                parent=self,
            )
            return
        remark = self.remark_text.get("1.0", "end").strip()
        export_payload = {
            "doc_type": self.doc_type_var.get().strip() or self.DOC_TYPE_OPTIONS[0],
            "doc_number": self.doc_number_var.get().strip(),
            "supplier": self.supplier_var.get().strip(),
            "delivery": self.delivery_var.get().strip(),
            "context_label": self.context_label_var.get().strip()
            or self.DEFAULT_CONTEXT_LABEL,
            "context_kind": "document",
            "remark": remark,
            "items": items,
            "total_weight": payload["total_weight"],
            "template": self.current_template_name,
            "column_layout": [dict(col) for col in self.current_columns],
        }
        self._on_export(export_payload)

    # Internal -------------------------------------------------------
    def _update_totals(self) -> None:
        payload = self._collect_items()
        total_weight = payload["total_weight"]
        if total_weight is None:
            text = "Totaal gewicht: —"
        else:
            text = f"Totaal gewicht: {total_weight:.2f} kg"
        self.total_weight_var.set(text)

    def _clone_columns(self, template: str) -> List[Dict[str, object]]:
        columns = self.COLUMN_TEMPLATES.get(template, [])
        cloned = [dict(col) for col in columns]
        if template == "Profielen":
            material_width = getattr(self, "profile_material_chars", None)
            if material_width:
                for column in cloned:
                    if column.get("key") == "Materiaal":
                        column["width"] = material_width
                        break
        for column in cloned:
            self._ensure_column_metrics(column)
        return cloned

    def _ensure_column_metrics(self, column: Dict[str, object]) -> None:
        """Ensure helper sizing metadata is present on a column definition."""

        width_value = column.get("width", 12)
        try:
            base_width = int(width_value)
        except Exception:
            try:
                base_width = int(float(width_value))
            except Exception:
                base_width = 12
        
        # display_chars is ONLY based on the width setting
        # This allows columns to be resized independently of header text length
        display_chars = base_width
        column["_display_chars"] = display_chars
        
        entry_char_px = getattr(self, "_entry_char_pixels", 1)
        min_width_px = max(1, int(round(display_chars * entry_char_px)))
        
        # Don't enforce header width as minimum - let columns be smaller than their headers
        # The header text will just wrap or be cut off if needed
        column["_min_width_px"] = min_width_px

    def _column_display_metrics(self, column: Dict[str, object]) -> tuple[int, int]:
        """Return the preferred width in characters and pixels for a column."""

        if "_display_chars" not in column or "_min_width_px" not in column:
            self._ensure_column_metrics(column)
        return column["_display_chars"], column["_min_width_px"]

    def _capture_rows(self) -> List[Dict[str, str]]:
        captured: List[Dict[str, str]] = []
        for widgets in self.rows:
            captured.append({key: var.get() for key, var in widgets.vars.items()})
        return captured

    def _clear_rows(self) -> None:
        # Destroy all data row widgets
        for widgets in self.rows:
            # Destroy button frame (which contains all buttons)
            try:
                if widgets.frame is not None:
                    widgets.frame.destroy()
            except Exception:
                pass
            
            # Also destroy all entry widgets directly
            for entry in widgets.entries.values():
                try:
                    entry.destroy()
                except Exception:
                    pass
        
        self.rows.clear()
        self._row_grid_indices.clear()
        self._next_data_row = 1  # Reset naar rij 1 (header is rij 0)

    def _render_header(self) -> None:
        """Render header-labels en separators direkt in rows_frame rij 0."""
        # Clear old header widgets
        for lbl in self._header_labels.values():
            try:
                lbl.destroy()
            except Exception:
                pass
        self._header_labels.clear()
        
        for sep in self._header_separators:
            try:
                sep.destroy()
            except Exception:
                pass
        self._header_separators.clear()
        
        # Verwijder oude resize-handles
        for handle in self._column_resizer_handles:
            try:
                handle.destroy()
            except Exception:
                pass
        self._column_resizer_handles.clear()
        
        # IMPORTANT: Clear ALL grid column configurations from rows_frame
        # This prevents old columns from showing up when switching templates
        for col in list(self.rows_frame.grid_slaves()):
            try:
                col.grid_remove()
            except Exception:
                pass
        
        # Reset all grid column weights and sizes
        for col_idx in range(100):  # Clear up to column 100
            try:
                self.rows_frame.columnconfigure(col_idx, weight=0, minsize=0)
            except Exception:
                pass
        
        # Render header-labels EN separators direkt in rows_frame grid
        for idx, column in enumerate(self.current_columns):
            grid_col = 1 + idx * 2  # Grid kolom 1, 3, 5, 7, ...
            display_chars, min_width_px = self._column_display_metrics(column)
            
            # Header label
            lbl = tk.Label(
                self.rows_frame,
                text=column.get("label", column.get("key", "")),
                anchor="w",
                font=getattr(self, "_header_font", None) or ("TkDefaultFont", 10, "bold"),
            )
            lbl.grid(row=0, column=grid_col, sticky="ew", padx=(6, 6))
            self.rows_frame.columnconfigure(grid_col, weight=1 if column.get("stretch") else 0, minsize=min_width_px)
            self._header_labels[idx] = lbl
            
            # Separator TUSSEN kolommen
            if idx < len(self.current_columns) - 1:
                sep_col = grid_col + 1  # Grid kolom 2, 4, 6, 8, ...
                separator = tk.Frame(
                    self.rows_frame,
                    width=2,
                    background=self.COLUMN_SEPARATOR_COLOR,
                    cursor="sb_h_double_arrow",
                )
                separator.grid(row=0, column=sep_col, sticky="ns", padx=0)
                self.rows_frame.columnconfigure(sep_col, weight=0, minsize=2)
                
                # Bind resize events with correct column_index
                # Use a helper function to create proper closures
                self._bind_separator_events(separator, idx)
                
                self._column_resizer_handles.append(separator)
                self._header_separators.append(separator)
        
        self._schedule_resizer_position_update()
    
    def _bind_separator_events(self, separator: tk.Widget, column_index: int) -> None:
        """Bind mouse events to a separator for resizing column_index."""
        def on_press(event):
            self._start_column_resize(column_index, event)
        
        def on_drag(event):
            self._drag_column_resize(column_index, event)
        
        def on_release(event):
            self._end_column_resize()
        
        separator.bind("<ButtonPress-1>", on_press, add="+")
        separator.bind("<B1-Motion>", on_drag, add="+")
        separator.bind("<ButtonRelease-1>", on_release, add="+")

    def _bind_resizer_events(self, widget: tk.Widget, column_index: int) -> None:
        widget.bind(
            "<ButtonPress-1>",
            lambda event, idx=column_index: self._start_column_resize(idx, event),
            add="+",
        )
        widget.bind(
            "<B1-Motion>",
            lambda event, idx=column_index: self._drag_column_resize(idx, event),
            add="+",
        )
        widget.bind("<ButtonRelease-1>", lambda _e: self._end_column_resize(), add="+")

    def _start_column_resize(self, column_index: int, event: tk.Event) -> None:
        """Start resizing: record which column and direction will determine resize."""
        if not (0 <= column_index < len(self.current_columns)):
            return
        
        left_column = self.current_columns[column_index]
        left_chars, _ = self._column_display_metrics(left_column)
        
        # Als er een rechter kolom is, we kunnen ook die resizen
        right_column = None
        right_chars = 0
        if column_index + 1 < len(self.current_columns):
            right_column = self.current_columns[column_index + 1]
            right_chars, _ = self._column_display_metrics(right_column)
        
        self._column_resize_state = {
            "left_index": column_index,
            "right_index": column_index + 1 if right_column else None,
            "start_x": event.x_root,
            "left_chars": left_chars,
            "right_chars": right_chars,
            "handle": event.widget,
        }
        
        handle = self._get_resizer_handle(column_index)
        if handle is not None:
            try:
                handle.configure(background=self.COLUMN_SEPARATOR_ACTIVE_COLOR)
            except Exception:
                pass

    def _drag_column_resize(self, column_index: int, event: tk.Event) -> None:
        """Drag resize: determine which column to resize based on drag direction."""
        state = self._column_resize_state
        if not state or state.get("left_index") != column_index:
            return
        
        try:
            start_x = state["start_x"]
            left_chars = state["left_chars"]
            right_chars = state.get("right_chars", 0)
            right_index = state.get("right_index")
        except KeyError:
            return
        
        delta_px = event.x_root - start_x
        char_width = max(1, int(getattr(self, "_entry_char_pixels", 1)))
        delta_chars = delta_px / char_width
        
        if delta_px > 0:
            # Sleep naar RECHTS = verbreed LINKER kolom
            desired = int(round(left_chars + delta_chars))
            self._set_column_width(column_index, desired)
        elif delta_px < 0 and right_index is not None:
            # Sleep naar LINKS = verbreed RECHTER kolom
            # Delta is negatief, dus we willen right_chars groter maken
            desired = int(round(right_chars - delta_chars))  # -delta_chars want delta is negatief
            self._set_column_width(right_index, desired)
        
        self._schedule_resizer_position_update()

    def _end_column_resize(self) -> None:
        state = self._column_resize_state
        if not state:
            return
        # Use left_index (de nieuwe key) niet index
        left_index = state.get("left_index")
        handle = self._get_resizer_handle(left_index)
        if handle is not None:
            try:
                handle.configure(background=self.COLUMN_SEPARATOR_COLOR)
            except Exception:
                pass
        self._column_resize_state = None

    def _column_padx(self, idx: int) -> tuple[int, int]:
        """Return consistent horizontal padding for column cells."""

        return (6 if idx else 0, 6)

    def _get_resizer_handle(self, column_index: Optional[int]) -> Optional[tk.Widget]:
        if column_index is None:
            return None
        if not (0 <= column_index < len(self._column_resizer_handles)):
            return None
        handle = self._column_resizer_handles[column_index]
        if handle is None or not handle.winfo_exists():
            return None
        return handle

    def _schedule_resizer_position_update(self) -> None:
        if self._resizer_update_job is not None:
            try:
                self.after_cancel(self._resizer_update_job)
            except Exception:
                pass
        self._resizer_update_job = self.after_idle(self._update_resizer_positions)

    def _update_resizer_positions(self) -> None:
        self._resizer_update_job = None
        if not self._column_resizer_handles:
            return
        try:
            self.header_row.update_idletasks()
            self.header_container.update_idletasks()
        except Exception:
            pass
        container_height = max(1, self.header_container.winfo_height())
        for idx, handle in enumerate(self._column_resizer_handles):
            if handle is None or not handle.winfo_exists():
                continue
            try:
                bbox = self.header_row.grid_bbox(idx, 0)
            except Exception:
                bbox = None
            if not bbox:
                try:
                    handle.place_forget()
                except Exception:
                    pass
                continue
            x = bbox[0] + bbox[2]
            try:
                handle.place(
                    in_=self.header_container,
                    x=x - 1,
                    y=0,
                    width=2,
                    height=container_height,
                )
                handle.lift()
            except Exception:
                pass

    def _apply_template(self, template: str, *, store_previous: bool = True) -> None:
        self.current_template_name = template
        if template in self._template_layout_cache:
            cached_layout = [dict(col) for col in self._template_layout_cache[template]]
            for column in cached_layout:
                self._ensure_column_metrics(column)
            self.current_columns = cached_layout
        else:
            self.current_columns = self._clone_columns(template)
        if not self.current_columns:
            self.current_columns = self._clone_columns(self.DEFAULT_TEMPLATE)
            self.current_template_name = self.DEFAULT_TEMPLATE

        # Clear rows BEFORE rendering header (so grid columns are reset)
        self._clear_rows()
        
        # Now render header and add one empty row
        self._render_header()
        self.add_row()
        self._update_totals()

    def _set_column_width(self, column_index: int, desired_chars: int) -> None:
        if not (0 <= column_index < len(self.current_columns)):
            return
        
        column = self.current_columns[column_index]
        
        # Clamp desired width between global min/max
        # Allow columns to be smaller than their header text
        desired = max(self.COLUMN_MIN_CHARS, min(self.COLUMN_MAX_CHARS, desired_chars))
        
        column["width"] = desired
        column.pop("_display_chars", None)
        column.pop("_min_width_px", None)
        self._ensure_column_metrics(column)
        self._apply_column_width(column_index)

    def _apply_column_width(self, column_index: int) -> None:
        """Update column width in rows_frame grid."""
        column = self.current_columns[column_index]
        display_chars, min_width_px = self._column_display_metrics(column)
        weight = 1 if column.get("stretch") else 0
        
        # Grid column is 1 + column_index * 2
        grid_col = 1 + column_index * 2
        self.rows_frame.columnconfigure(grid_col, weight=weight, minsize=min_width_px)
        
        # Update all data rows
        key = column.get("key")
        for widgets in self.rows:
            entry = widgets.entries.get(key)
            if entry is None:
                continue
            try:
                entry.configure(width=display_chars)
            except Exception:
                pass
        
        self._schedule_resizer_position_update()

