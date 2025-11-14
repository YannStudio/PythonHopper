"""UI-tab voor het handmatig opstellen van bestel- of offertenbonnen."""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, TYPE_CHECKING

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
        self._all_values: List[str] = list(values or ())
        self._last_query: str = ""
        super().__init__(master, values=values, **kwargs)
        self.bind("<KeyRelease>", self._on_key_release, add="+")
        self.bind("<<ComboboxSelected>>", self._on_selection, add="+")
        self.bind("<FocusIn>", self._restore_values, add="+")
        self.bind("<FocusIn>", self._on_focus_in, add="+")
        self.bind("<FocusOut>", self._on_focus_out, add="+")

    # Public helpers -------------------------------------------------
    def set_choices(self, values: List[str]) -> None:
        """Update beschikbare opties en hergebruik de huidige filter."""

        self._all_values = list(values)
        # Houd de huidige invoer ongemoeid wanneer opties opnieuw geladen
        # worden, zodat reeds ingevulde waarden zichtbaar blijven.
        self._apply_filter(self._last_query, update_entry=False)

    # Internal -------------------------------------------------------
    def _apply_filter(self, query: str, *, update_entry: bool = True) -> None:
        display_text = query
        normalized = display_text.casefold().strip()
        if not normalized:
            filtered = self._all_values
        else:
            filtered = [
                option
                for option in self._all_values
                if normalized in option.casefold()
            ]
        self.configure(values=filtered)
        self._last_query = normalized
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


class ManualOrderTab(tk.Frame):
    """Tab om handmatige orderregels in te geven."""

    DEFAULT_CONTEXT_LABEL = DEFAULT_MANUAL_CONTEXT
    DEFAULT_TEMPLATE = "Standaard"
    COLUMN_MIN_CHARS = 6
    COLUMN_MAX_CHARS = 72
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
        tk.Label(header, textvariable=self.project_number_var, anchor="w").grid(
            row=0, column=4, sticky="w", padx=(6, 0)
        )

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
        tk.Label(header, textvariable=self.project_name_var, anchor="w").grid(
            row=1, column=4, sticky="w", padx=(6, 0), pady=(6, 0)
        )

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
        self.client_combo = SearchableCombobox(
            client_field,
            textvariable=self.client_var,
            width=field_char_width,
        )
        self.client_combo.pack(side="left")
        if on_manage_clients:
            tk.Button(
                client_field,
                text="Beheer",
                command=on_manage_clients,
                width=10,
            ).pack(side="left", padx=(manage_spacing_px, 0))

        tk.Label(header, text="Leverancier:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        self.supplier_var = tk.StringVar()
        supplier_field = tk.Frame(header)
        supplier_field.grid(
            row=3,
            column=1,
            columnspan=4,
            sticky="w",
            padx=(6, 0),
            pady=(8, 0),
        )
        self.supplier_combo = SearchableCombobox(
            supplier_field, textvariable=self.supplier_var, width=field_char_width
        )
        self.supplier_combo.pack(side="left")
        if on_manage_suppliers:
            tk.Button(
                supplier_field,
                text="Beheer",
                command=on_manage_suppliers,
                width=10,
            ).pack(side="left", padx=(manage_spacing_px, 0))

        tk.Label(header, text="Leveradres:").grid(row=4, column=0, sticky="w", pady=(6, 0))
        self.delivery_var = tk.StringVar()
        delivery_field = tk.Frame(header)
        delivery_field.grid(
            row=4,
            column=1,
            columnspan=4,
            sticky="w",
            padx=(6, 0),
            pady=(6, 0),
        )
        self.delivery_combo = SearchableCombobox(
            delivery_field, textvariable=self.delivery_var, width=field_char_width
        )
        self.delivery_combo.pack(side="left")
        if on_manage_deliveries:
            tk.Button(
                delivery_field,
                text="Beheer",
                command=on_manage_deliveries,
                width=10,
            ).pack(side="left", padx=(manage_spacing_px, 0))

        tk.Label(header, text="Documentnaam:").grid(row=5, column=0, sticky="w", pady=(6, 0))
        self.context_label_var = tk.StringVar(
            value=self.project_name_var.get().strip() or self.DEFAULT_CONTEXT_LABEL
        )
        context_entry = tk.Entry(
            header, textvariable=self.context_label_var, width=field_char_width
        )
        context_entry.grid(
            row=5,
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

        tk.Label(header, text="Opmerkingen:").grid(row=6, column=0, sticky="nw", pady=(8, 0))
        self.remark_text = tk.Text(header, height=4, wrap="word")
        self.remark_text.grid(
            row=6,
            column=1,
            columnspan=4,
            sticky="nsew",
            padx=(6, 0),
            pady=(8, 0),
        )
        header.rowconfigure(6, weight=1)

        ttk.Separator(self, orient="horizontal").grid(row=1, column=0, sticky="ew", pady=12)

        table_container = tk.Frame(self)
        table_container.grid(row=2, column=0, sticky="nsew", padx=4, pady=(12, 0))
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
        self.header_container.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 6))
        self.header_container.columnconfigure(0, weight=1)
        self.header_row = tk.Frame(self.header_container)
        self.header_row.grid(row=0, column=0, sticky="ew")
        self.column_slider_row = tk.Frame(self.header_container)
        self.column_slider_row.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        self.table_canvas = tk.Canvas(table_container, highlightthickness=0)
        self.table_canvas.grid(row=2, column=0, sticky="nsew", padx=(0, 0))
        table_scroll = ttk.Scrollbar(
            table_container, orient="vertical", command=self.table_canvas.yview
        )
        table_scroll.grid(row=2, column=1, sticky="ns")
        self.table_canvas.configure(yscrollcommand=table_scroll.set)

        self.rows_frame = tk.Frame(self.table_canvas)
        self.rows_window = self.table_canvas.create_window(
            (0, 0), window=self.rows_frame, anchor="nw"
        )
        self.rows_frame.columnconfigure(0, weight=1)

        def _update_scrollregion(_event=None):
            self.table_canvas.configure(scrollregion=self.table_canvas.bbox("all"))

        self.rows_frame.bind("<Configure>", _update_scrollregion)

        def _resize_canvas(event: tk.Event) -> None:
            try:
                self.table_canvas.itemconfigure(self.rows_window, width=event.width)
            except Exception:
                pass

        self.table_canvas.bind("<Configure>", _resize_canvas)

        controls = tk.Frame(table_container)
        controls.grid(row=3, column=0, columnspan=2, sticky="ew", padx=4, pady=(8, 0))
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
        footer.grid(row=3, column=0, sticky="ew", padx=4, pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        tk.Button(footer, text="Bestelbon opslaan", command=self._handle_export).grid(
            row=0, column=1, sticky="e"
        )

        self.rows: List[_ManualRowWidgets] = []
        self.current_template_name: str = ""
        self.current_columns: List[Dict[str, object]] = []
        self._template_rows_cache: Dict[str, List[Dict[str, str]]] = {}
        self._template_layout_cache: Dict[str, List[Dict[str, object]]] = {}
        self._column_slider_vars: List[tk.StringVar] = []
        self._column_slider_value_vars: List[tk.StringVar] = []
        self._column_width_updating = False

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
        try:
            self.client_combo.set_choices(client_opts)
        except Exception:
            self.client_combo.configure(values=client_opts)
        if current_client not in client_opts:
            if client_opts:
                self.client_var.set(client_opts[0])
            else:
                self.client_var.set("")

        supplier_opts = ["Geen"] + [
            self.suppliers_db.display_name(s)
            for s in self.suppliers_db.suppliers_sorted()
        ]
        current_supplier = self.supplier_var.get()
        try:
            self.supplier_combo.set_choices(supplier_opts)
        except Exception:
            self.supplier_combo.configure(values=supplier_opts)
        if current_supplier not in supplier_opts:
            self.supplier_var.set("Geen")

        delivery_opts = list(self.DELIVERY_PRESETS) + [
            self.delivery_db.display_name(a)
            for a in self.delivery_db.addresses_sorted()
        ]
        current_delivery = self.delivery_var.get()
        try:
            self.delivery_combo.set_choices(delivery_opts)
        except Exception:
            self.delivery_combo.configure(values=delivery_opts)
        if current_delivery not in delivery_opts:
            self.delivery_var.set(self.DELIVERY_PRESETS[0])

    def set_doc_number(self, value: str) -> None:
        self.doc_number_var.set(value)
        self._doc_number_prefix = _prefix_for_doc_type(self.doc_type_var.get())

    # Row management -------------------------------------------------
    def add_row(self, values: Optional[Dict[str, object]] = None) -> None:
        widgets = _ManualRowWidgets(
            frame=tk.Frame(self.rows_frame),
            vars={},
            entries={},
            remove_btn=None,
        )
        widgets.frame.pack(fill="x", padx=6, pady=4)
        widgets.frame.columnconfigure(len(self.current_columns), weight=0)

        for idx, column in enumerate(self.current_columns):
            var = tk.StringVar()
            if values is not None and column["key"] in values:
                value = values[column["key"]]
                var.set("" if value is None else str(value))
            display_chars, min_width_px = self._column_display_metrics(column)
            entry = tk.Entry(
                widgets.frame,
                textvariable=var,
                width=display_chars,
                justify=column.get("justify", "left"),
            )
            entry.grid(row=0, column=idx, sticky="ew", padx=self._column_padx(idx))
            if column.get("stretch"):
                widgets.frame.columnconfigure(idx, weight=1, minsize=min_width_px)
            else:
                widgets.frame.columnconfigure(idx, weight=0, minsize=min_width_px)
            var.trace_add("write", lambda *_args: self._update_totals())
            widgets.vars[column["key"]] = var
            widgets.entries[column["key"]] = entry

        remove_btn = tk.Button(
            widgets.frame,
            text="✕",
            width=3,
            command=lambda row=widgets: self.remove_row(row),
        )
        remove_btn.grid(row=0, column=len(self.current_columns), padx=(0, 4))
        widgets.remove_btn = remove_btn
        self.rows.append(widgets)
        if self.current_columns:
            first_key = self.current_columns[0]["key"]
            self.after_idle(lambda: widgets.entries[first_key].focus_set())
        self._update_totals()

    def remove_row(self, row: _ManualRowWidgets) -> None:
        if row not in self.rows:
            return
        self.rows.remove(row)
        try:
            row.frame.destroy()
        except Exception:
            pass
        if not self.rows:
            self.add_row()
        else:
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
        label_text = str(column.get("label") or column.get("key") or "")
        display_chars = max(base_width, len(label_text))
        column["_display_chars"] = display_chars
        entry_char_px = getattr(self, "_entry_char_pixels", 1)
        min_width_px = max(1, int(round(display_chars * entry_char_px)))
        header_font = getattr(self, "_header_font", None)
        if header_font is not None:
            try:
                min_width_px = max(min_width_px, header_font.measure(label_text))
            except Exception:
                pass
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
        for widgets in self.rows:
            try:
                widgets.frame.destroy()
            except Exception:
                pass
        self.rows.clear()

    def _render_header(self) -> None:
        for child in self.header_row.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        for child in self.column_slider_row.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        self._column_slider_vars.clear()
        self._column_slider_value_vars.clear()
        for idx, column in enumerate(self.current_columns):
            display_chars, min_width_px = self._column_display_metrics(column)
            lbl = tk.Label(
                self.header_row,
                text=column.get("label", column.get("key", "")),
                anchor="w",
                font=getattr(self, "_header_font", None) or ("TkDefaultFont", 10, "bold"),
            )
            lbl.grid(row=0, column=idx, sticky="w", padx=self._column_padx(idx))
            if column.get("stretch"):
                weight = 1
            else:
                weight = 0
            self.header_row.columnconfigure(idx, weight=weight, minsize=min_width_px)
            self.column_slider_row.columnconfigure(idx, weight=weight, minsize=min_width_px)

            slider_frame = tk.Frame(self.column_slider_row)
            slider_frame.grid(row=0, column=idx, sticky="ew", padx=self._column_padx(idx))
            slider_frame.columnconfigure(0, weight=1)
            slider_var = tk.StringVar(value=str(display_chars))
            self._column_slider_vars.append(slider_var)
            spinbox = tk.Spinbox(
                slider_frame,
                from_=self.COLUMN_MIN_CHARS,
                to=self.COLUMN_MAX_CHARS,
                increment=1,
                textvariable=slider_var,
                width=6,
                justify="right",
            )
            spinbox.pack(fill="x")

            def _commit_spinbox_change(_event=None, column_index=idx, var=slider_var):
                try:
                    value = var.get()
                except tk.TclError:
                    value = ""
                self._handle_column_slider(column_index, value)

            spinbox.configure(
                command=lambda column_index=idx, var=slider_var: self._handle_column_slider(
                    column_index, var.get()
                )
            )
            spinbox.bind("<FocusOut>", _commit_spinbox_change, add="+")
            spinbox.bind("<Return>", _commit_spinbox_change, add="+")
            value_var = tk.StringVar(value=self._format_slider_value(display_chars))
            self._column_slider_value_vars.append(value_var)
            tk.Label(
                slider_frame,
                textvariable=value_var,
                anchor="e",
                font=("TkDefaultFont", 8),
            ).pack(fill="x", pady=(2, 0))

    def _column_padx(self, idx: int) -> tuple[int, int]:
        """Return consistent horizontal padding for column cells."""

        return (6 if idx else 0, 6)

    def _format_slider_value(self, chars: int) -> str:
        return f"{int(chars)} ch"

    def _apply_template(self, template: str, *, store_previous: bool = True) -> None:
        if store_previous and self.current_template_name:
            self._template_rows_cache[self.current_template_name] = self._capture_rows()
            self._template_layout_cache[self.current_template_name] = [
                dict(col) for col in self.current_columns
            ]

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

        self._render_header()
        self._clear_rows()

        cached_rows = self._template_rows_cache.get(self.current_template_name, [])
        if cached_rows:
            for row_values in cached_rows:
                self.add_row(row_values)
        else:
            self.add_row()
        self._update_totals()

    def _handle_column_slider(self, column_index: int, raw_value: object) -> None:
        if self._column_width_updating:
            return
        if not (0 <= column_index < len(self.current_columns)):
            return
        try:
            value = float(raw_value)
        except Exception:
            column = self.current_columns[column_index]
            slider_var = self._column_slider_vars[column_index]
            try:
                slider_var.set(str(int(column.get("width", self.COLUMN_MIN_CHARS))))
            except Exception:
                slider_var.set(str(self.COLUMN_MIN_CHARS))
            return
        desired = int(round(value))
        desired = max(self.COLUMN_MIN_CHARS, min(self.COLUMN_MAX_CHARS, desired))
        column = self.current_columns[column_index]
        column["width"] = desired
        column.pop("_display_chars", None)
        column.pop("_min_width_px", None)
        self._ensure_column_metrics(column)
        slider_var = self._column_slider_vars[column_index]
        new_text = str(desired)
        if slider_var.get() != new_text:
            self._column_width_updating = True
            try:
                slider_var.set(new_text)
            finally:
                self._column_width_updating = False
        value_var = self._column_slider_value_vars[column_index]
        value_var.set(self._format_slider_value(desired))
        self._apply_column_width(column_index)

    def _apply_column_width(self, column_index: int) -> None:
        column = self.current_columns[column_index]
        display_chars, min_width_px = self._column_display_metrics(column)
        weight = 1 if column.get("stretch") else 0
        self.header_row.columnconfigure(column_index, weight=weight, minsize=min_width_px)
        self.column_slider_row.columnconfigure(column_index, weight=weight, minsize=min_width_px)
        key = column.get("key")
        for widgets in self.rows:
            entry = widgets.entries.get(key)
            if entry is None:
                continue
            try:
                entry.configure(width=display_chars)
            except Exception:
                pass
            try:
                widgets.frame.columnconfigure(column_index, weight=weight, minsize=min_width_px)
            except Exception:
                pass
        if 0 <= column_index < len(self._column_slider_value_vars):
            self._column_slider_value_vars[column_index].set(
                self._format_slider_value(display_chars)
            )
        if 0 <= column_index < len(self._column_slider_vars):
            slider_var = self._column_slider_vars[column_index]
            try:
                current_value = float(slider_var.get())
            except Exception:
                current_value = display_chars
            if int(round(current_value)) != display_chars:
                self._column_width_updating = True
                try:
                    slider_var.set(str(display_chars))
                finally:
                    self._column_width_updating = False

