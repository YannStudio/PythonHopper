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

    # Public helpers -------------------------------------------------
    def set_choices(self, values: List[str]) -> None:
        """Update beschikbare opties en hergebruik de huidige filter."""

        self._all_values = list(values)
        self._apply_filter(self._last_query)

    # Internal -------------------------------------------------------
    def _apply_filter(self, query: str) -> None:
        normalized = query.casefold().strip()
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
        if filtered:
            # Toon de dropdown zodat de gebruiker direct kan kiezen
            try:
                self.event_generate("<Down>")
            except Exception:
                pass

    def _on_key_release(self, event: tk.Event) -> None:
        if event.keysym in {"Up", "Down", "Return", "Escape", "Tab"}:
            if event.keysym == "Escape":
                self._restore_values()
            return
        self.after_idle(lambda: self._apply_filter(self.get()))

    def _on_selection(self, _event: tk.Event) -> None:
        self._restore_values()

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


class ManualOrderTab(tk.Frame):
    """Tab om handmatige orderregels in te geven."""

    DEFAULT_CONTEXT_LABEL = DEFAULT_MANUAL_CONTEXT
    DEFAULT_TEMPLATE = "Standaard"
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
                "key": "Supplier",
                "label": "Supplier",
                "width": 20,
                "justify": "left",
                "stretch": True,
                "wrap": True,
                "weight": 2.0,
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
        char_width = max(1, base_font.measure("0"))
        field_char_width = max(1, round(field_width_px / char_width))

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
        table_container.grid(row=2, column=0, sticky="nsew", padx=4)
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(2, weight=1)

        template_row = tk.Frame(table_container)
        template_row.grid(row=0, column=0, columnspan=2, sticky="ew", padx=4, pady=(0, 6))
        tk.Label(template_row, text="Sjabloon:").pack(side="left")
        self.template_var = tk.StringVar(value=self.DEFAULT_TEMPLATE)
        self.template_combo = ttk.Combobox(
            template_row,
            textvariable=self.template_var,
            values=tuple(self.COLUMN_TEMPLATES.keys()),
            state="readonly",
            width=max(14, field_char_width // 2),
            takefocus=False,
        )
        self.template_combo.pack(side="left", padx=(6, 0))

        self.header_row = tk.Frame(table_container)
        self.header_row.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 6))
        self.header_row.columnconfigure(0, weight=1)

        self.table_canvas = tk.Canvas(table_container, highlightthickness=0)
        self.table_canvas.grid(row=2, column=0, sticky="nsew", padx=(4, 0))
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

        def _handle_template_change(*_args) -> None:
            self._apply_template(self.template_var.get())

        self.template_var.trace_add("write", _handle_template_change)

        self._apply_template(self.template_var.get(), store_previous=False)
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
            entry = tk.Entry(
                widgets.frame,
                textvariable=var,
                width=column.get("width", 12),
                justify=column.get("justify", "left"),
            )
            entry.grid(row=0, column=idx, sticky="ew", padx=(6 if idx else 0, 6))
            if column.get("stretch"):
                widgets.frame.columnconfigure(idx, weight=1)
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
        current = len(self.rows)
        to_add = max(0, desired - current)
        if to_add == 0:
            return
        for _ in range(to_add):
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
        return [dict(col) for col in columns]

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
        for idx, column in enumerate(self.current_columns):
            anchor = "w" if column.get("justify") != "right" else "e"
            lbl = tk.Label(
                self.header_row,
                text=column.get("label", column.get("key", "")),
                anchor=anchor,
                font=("TkDefaultFont", 10, "bold"),
            )
            lbl.grid(row=0, column=idx, sticky="ew", padx=(6 if idx else 0, 6))
            if column.get("stretch"):
                self.header_row.columnconfigure(idx, weight=1)
            else:
                self.header_row.columnconfigure(idx, weight=0)

    def _apply_template(self, template: str, *, store_previous: bool = True) -> None:
        if store_previous and self.current_template_name:
            self._template_rows_cache[self.current_template_name] = self._capture_rows()

        self.current_template_name = template
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

