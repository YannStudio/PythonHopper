"""UI-tab voor het handmatig opstellen van bestel- of offertenbonnen."""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional

import tkinter as tk
from tkinter import ttk, messagebox


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


class ManualOrderTab(tk.Frame):
    """Tab om handmatige orderregels in te geven."""

    COLUMNS: List[Dict[str, object]] = [
        {"key": "PartNumber", "label": "Artikel nr.", "width": 16, "justify": "left"},
        {"key": "Description", "label": "Omschrijving", "width": 32, "justify": "left"},
        {"key": "Materiaal", "label": "Materiaal", "width": 18, "justify": "left"},
        {"key": "Aantal", "label": "Aantal", "width": 8, "justify": "right"},
        {"key": "Oppervlakte", "label": "Oppervlakte", "width": 10, "justify": "right"},
        {"key": "Gewicht", "label": "Gewicht (kg)", "width": 10, "justify": "right"},
    ]

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
        project_number_var: tk.StringVar,
        project_name_var: tk.StringVar,
        on_export: Callable[[Dict[str, object]], None],
        on_manage_suppliers: Optional[Callable[[], None]] = None,
        on_manage_deliveries: Optional[Callable[[], None]] = None,
    ) -> None:
        super().__init__(master)
        self.suppliers_db = suppliers_db
        self.delivery_db = delivery_db
        self.project_number_var = project_number_var
        self.project_name_var = project_name_var
        self._on_export = on_export
        self._on_manage_suppliers = on_manage_suppliers
        self._on_manage_deliveries = on_manage_deliveries
        self._context_user_modified = False

        self.configure(padx=12, pady=12)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = tk.LabelFrame(self, text="Documentgegevens", labelanchor="n")
        header.grid(row=0, column=0, sticky="nsew")
        header.columnconfigure(1, weight=1)
        header.columnconfigure(3, weight=1)

        self.doc_type_var = tk.StringVar(value=self.DOC_TYPE_OPTIONS[0])
        tk.Label(header, text="Documenttype:").grid(row=0, column=0, sticky="w")
        self.doc_type_combo = ttk.Combobox(
            header,
            textvariable=self.doc_type_var,
            values=self.DOC_TYPE_OPTIONS,
            state="readonly",
            width=20,
        )
        self.doc_type_combo.grid(row=0, column=1, sticky="w", padx=(6, 0))

        tk.Label(header, text="Documentnummer:").grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.doc_number_var = tk.StringVar()
        self.doc_number_entry = tk.Entry(header, textvariable=self.doc_number_var, width=18)
        self.doc_number_entry.grid(row=0, column=3, sticky="w", padx=(6, 0))

        tk.Label(header, text="Leverancier:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.supplier_var = tk.StringVar()
        self.supplier_combo = ttk.Combobox(header, textvariable=self.supplier_var, width=40)
        self.supplier_combo.grid(row=1, column=1, sticky="ew", pady=(8, 0))
        if on_manage_suppliers:
            tk.Button(
                header,
                text="Beheer",
                command=on_manage_suppliers,
                width=10,
            ).grid(row=1, column=2, columnspan=2, sticky="e", pady=(8, 0))

        tk.Label(header, text="Leveradres:").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.delivery_var = tk.StringVar()
        self.delivery_combo = ttk.Combobox(header, textvariable=self.delivery_var, width=40)
        self.delivery_combo.grid(row=2, column=1, sticky="ew", pady=(6, 0))
        if on_manage_deliveries:
            tk.Button(
                header,
                text="Beheer",
                command=on_manage_deliveries,
                width=10,
            ).grid(row=2, column=2, columnspan=2, sticky="e", pady=(6, 0))

        tk.Label(header, text="Projectnummer:").grid(row=3, column=0, sticky="w", pady=(6, 0))
        tk.Label(header, textvariable=self.project_number_var, anchor="w").grid(
            row=3, column=1, sticky="w", pady=(6, 0)
        )
        tk.Label(header, text="Projectnaam:").grid(row=3, column=2, sticky="w", padx=(12, 0), pady=(6, 0))
        tk.Label(header, textvariable=self.project_name_var, anchor="w").grid(
            row=3, column=3, sticky="w", pady=(6, 0)
        )

        tk.Label(header, text="Documentnaam:").grid(row=4, column=0, sticky="w", pady=(6, 0))
        self.context_label_var = tk.StringVar(value=self.project_name_var.get().strip() or "Handmatige bon")
        context_entry = tk.Entry(header, textvariable=self.context_label_var)
        context_entry.grid(row=4, column=1, columnspan=3, sticky="ew", pady=(6, 0))

        def _mark_context_modified(*_args):
            self._context_user_modified = True

        self.context_label_var.trace_add("write", _mark_context_modified)

        def _sync_project_name(*_args):
            if not self._context_user_modified:
                self.context_label_var.set(
                    self.project_name_var.get().strip() or "Handmatige bon"
                )

        self.project_name_var.trace_add("write", _sync_project_name)

        tk.Label(header, text="Opmerkingen:").grid(row=5, column=0, sticky="nw", pady=(8, 0))
        self.remark_text = tk.Text(header, height=4, wrap="word")
        self.remark_text.grid(row=5, column=1, columnspan=3, sticky="nsew", pady=(8, 0))
        header.rowconfigure(5, weight=1)

        ttk.Separator(self, orient="horizontal").grid(row=1, column=0, sticky="ew", pady=12)

        table_container = tk.Frame(self)
        table_container.grid(row=2, column=0, sticky="nsew")
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(1, weight=1)

        header_row = tk.Frame(table_container)
        header_row.grid(row=0, column=0, sticky="ew")
        header_row.columnconfigure(len(self.COLUMNS), weight=0)
        for idx, column in enumerate(self.COLUMNS):
            tk.Label(
                header_row,
                text=column["label"],
                anchor="w" if column["justify"] != "right" else "e",
                font=("TkDefaultFont", 10, "bold"),
            ).grid(row=0, column=idx, sticky="ew", padx=(4 if idx else 0, 4))
            header_row.columnconfigure(idx, weight=1 if column["key"] == "Description" else 0)

        self.table_canvas = tk.Canvas(table_container, highlightthickness=0)
        self.table_canvas.grid(row=1, column=0, sticky="nsew")
        table_scroll = ttk.Scrollbar(
            table_container, orient="vertical", command=self.table_canvas.yview
        )
        table_scroll.grid(row=1, column=1, sticky="ns")
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
        controls.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        controls.columnconfigure(0, weight=1)

        tk.Button(
            controls,
            text="Regel toevoegen",
            command=self.add_row,
        ).grid(row=0, column=0, sticky="w")

        self.total_weight_var = tk.StringVar(value="Totaal gewicht: 0.00 kg")
        tk.Label(controls, textvariable=self.total_weight_var, anchor="e").grid(
            row=0, column=1, sticky="e"
        )

        footer = tk.Frame(self)
        footer.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        tk.Button(footer, text="Bestelbon opslaan", command=self._handle_export).grid(
            row=0, column=1, sticky="e"
        )

        self.rows: List[_ManualRowWidgets] = []
        self.add_row()
        self.refresh_data()

    # Public helpers -------------------------------------------------
    def refresh_data(self) -> None:
        """Reload supplier/delivery options from databases."""

        supplier_opts = ["Geen"] + [
            self.suppliers_db.display_name(s)
            for s in self.suppliers_db.suppliers_sorted()
        ]
        current_supplier = self.supplier_var.get()
        self.supplier_combo["values"] = supplier_opts
        if current_supplier not in supplier_opts:
            self.supplier_var.set("Geen")

        delivery_opts = list(self.DELIVERY_PRESETS) + [
            self.delivery_db.display_name(a)
            for a in self.delivery_db.addresses_sorted()
        ]
        current_delivery = self.delivery_var.get()
        self.delivery_combo["values"] = delivery_opts
        if current_delivery not in delivery_opts:
            self.delivery_var.set(self.DELIVERY_PRESETS[0])

    def set_doc_number(self, value: str) -> None:
        self.doc_number_var.set(value)

    # Row management -------------------------------------------------
    def add_row(self) -> None:
        widgets = _ManualRowWidgets(frame=tk.Frame(self.rows_frame), vars={}, entries={}, remove_btn=None)  # type: ignore[arg-type]
        widgets.frame.pack(fill="x", pady=2)
        widgets.frame.columnconfigure(len(self.COLUMNS), weight=0)

        for idx, column in enumerate(self.COLUMNS):
            var = tk.StringVar()
            entry = tk.Entry(
                widgets.frame,
                textvariable=var,
                width=column.get("width", 12),
                justify=column.get("justify", "left"),
            )
            entry.grid(row=0, column=idx, sticky="ew", padx=(4 if idx else 0, 4))
            if column["key"] == "Description":
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
        remove_btn.grid(row=0, column=len(self.COLUMNS), padx=(0, 4))
        widgets.remove_btn = remove_btn
        self.rows.append(widgets)
        self.after_idle(lambda: widgets.entries[self.COLUMNS[0]["key"]].focus_set())
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

    # Data collection ------------------------------------------------
    def _collect_items(self) -> Dict[str, object]:
        items: List[Dict[str, object]] = []
        total_weight = 0.0
        weight_found = False

        for widgets in self.rows:
            raw = {key: var.get().strip() for key, var in widgets.vars.items()}
            if not any(raw.values()):
                continue
            record: Dict[str, object] = {}
            for column in self.COLUMNS:
                key = column["key"]
                value = raw.get(key, "")
                if key in {"Aantal", "Oppervlakte", "Gewicht"}:
                    normalized = _normalize_numeric(value)
                else:
                    normalized = value
                record[key] = normalized
            weight_raw = raw.get("Gewicht", "")
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
            "context_label": self.context_label_var.get().strip() or "Handmatige bon",
            "context_kind": "document",
            "remark": remark,
            "items": items,
            "total_weight": payload["total_weight"],
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

