import os
import shutil
import subprocess
import sys
import threading
import unicodedata
from typing import Dict, List, Optional

import pandas as pd

from helpers import _to_str, _build_file_index
from models import Supplier, Client, DeliveryAddress
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from clients_db import ClientsDB, CLIENTS_DB_FILE
from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE
from bom import read_csv_flex, load_bom
from orders import (
    copy_per_production_and_orders,
    DEFAULT_FOOTER_NOTE,
    combine_pdfs_per_production,
    combine_pdfs_from_source,
    _prefix_for_doc_type,
)


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
    from tkinter import ttk, filedialog, messagebox, simpledialog

    TREE_ODD_BG = "#FFFFFF"
    TREE_EVEN_BG = "#F5F5F5"

    class ClientsManagerFrame(tk.Frame):
        def __init__(self, master, db: ClientsDB, on_change=None):
            super().__init__(master)
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
            fields = [
                ("Naam", "name"),
                ("Adres", "address"),
                ("BTW", "vat"),
                ("E-mail", "email"),
            ]
            entries = {}
            for i, (lbl, key) in enumerate(fields):
                tk.Label(win, text=lbl + ":").grid(row=i, column=0, sticky="e", padx=4, pady=2)
                ent = tk.Entry(win, width=40)
                ent.grid(row=i, column=1, padx=4, pady=2)
                if client:
                    ent.insert(0, _to_str(getattr(client, key)))
                entries[key] = ent
            fav_var = tk.BooleanVar(value=client.favorite if client else False)
            tk.Checkbutton(win, text="Favoriet", variable=fav_var).grid(row=len(fields), column=1, sticky="w", padx=4, pady=2)

            def _save():
                rec = {k: e.get().strip() for k, e in entries.items()}
                rec["favorite"] = fav_var.get()
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
            btnf.grid(row=len(fields)+1, column=0, columnspan=2, pady=6)
            tk.Button(btnf, text="Opslaan", command=_save).pack(side="left", padx=4)
            tk.Button(btnf, text="Annuleer", command=win.destroy).pack(side="left", padx=4)
            win.transient(self)
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

    class SupplierSelectionFrame(tk.Frame):
        """Per productie: type-to-filter of dropdown; rechts detailkaart (klik = selecteer).
           Knoppen altijd zichtbaar onderaan.
        """
        def __init__(
            self,
            master,
            productions: List[str],
            db: SuppliersDB,
            delivery_db: DeliveryAddressesDB,
            callback,
        ):
            super().__init__(master)
            self.db = db
            self.delivery_db = delivery_db
            self.callback = callback
            self._preview_supplier: Optional[Supplier] = None
            self._active_prod: Optional[str] = None  # laatst gefocuste rij
            self.sel_vars: Dict[str, tk.StringVar] = {}
            self.doc_vars: Dict[str, tk.StringVar] = {}
            self.doc_num_vars: Dict[str, tk.StringVar] = {}
            self.delivery_vars: Dict[str, tk.StringVar] = {}
            self.delivery_combos: Dict[str, ttk.Combobox] = {}

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

            # Project info entries above production rows
            proj_frame = tk.LabelFrame(left, text="Projectgegevens", labelanchor="n")
            proj_frame.pack(fill="x", pady=(0, 6))
            tk.Label(
                proj_frame,
                text="Deze gegevens gelden voor het volledige project",
                anchor="w",
            ).pack(fill="x", padx=6, pady=(0, 3))

            pn_row = tk.Frame(proj_frame)
            pn_row.pack(fill="x", pady=3)
            tk.Label(pn_row, text="Projectnr.", width=18, anchor="w").pack(side="left")
            self.project_number_var = tk.StringVar()
            tk.Entry(pn_row, textvariable=self.project_number_var, width=50).pack(side="left", padx=6)

            name_row = tk.Frame(proj_frame)
            name_row.pack(fill="x", pady=3)
            tk.Label(name_row, text="Projectnaam", width=18, anchor="w").pack(side="left")
            self.project_name_var = tk.StringVar()
            tk.Entry(name_row, textvariable=self.project_name_var, width=50).pack(side="left", padx=6)

            tk.Label(left, text="Kies leveranciers per productie", anchor="w").pack(
                fill="x", pady=(0, 6)
            )

            delivery_opts = [
                "Geen",
                "Bestelling wordt opgehaald",
                "Leveradres wordt nog meegedeeld",
            ] + [
                self.delivery_db.display_name(a)
                for a in self.delivery_db.addresses_sorted()
            ]

            doc_type_opts = ["Bestelbon", "Offerteaanvraag"]
            self._doc_type_prefixes = {
                _prefix_for_doc_type(t) for t in doc_type_opts
            }

            self.rows = []
            for prod in productions:
                row = tk.Frame(left)
                row.pack(fill="x", pady=3)
                tk.Label(row, text=prod, width=18, anchor="w").pack(side="left")
                var = tk.StringVar()
                self.sel_vars[prod] = var
                combo = ttk.Combobox(row, textvariable=var, state="normal", width=50)
                combo.pack(side="left", padx=6)
                combo.bind("<<ComboboxSelected>>", self._on_combo_change)
                combo.bind("<FocusIn>", lambda _e, p=prod: self._on_focus_prod(p))
                combo.bind("<KeyRelease>", lambda ev, p=prod, c=combo: self._on_combo_type(ev, p, c))

                doc_var = tk.StringVar(value="Bestelbon")
                self.doc_vars[prod] = doc_var
                doc_combo = ttk.Combobox(
                    row,
                    textvariable=doc_var,
                    values=doc_type_opts,
                    state="readonly",
                    width=18,
                )
                doc_combo.pack(side="left", padx=6)
                doc_combo.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, p=prod: self._on_doc_type_change(p),
                )

                doc_num_var = tk.StringVar()
                self.doc_num_vars[prod] = doc_num_var
                tk.Entry(row, textvariable=doc_num_var, width=8).pack(side="left", padx=6)

                dvar = tk.StringVar(value="Geen")
                self.delivery_vars[prod] = dvar
                dcombo = ttk.Combobox(
                    row,
                    textvariable=dvar,
                    values=delivery_opts,
                    state="readonly",
                    width=50,
                )
                dcombo.pack(side="left", padx=6)
                self.delivery_combos[prod] = dcombo

                self.rows.append((prod, combo))

            # Container voor kaarten
            preview_frame = tk.LabelFrame(
                content,
                text="Leverancier details\n(klik om te selecteren)",
                labelanchor="n",
            )
            preview_frame.grid(row=1, column=0, sticky="nsew")
            preview_frame.grid_rowconfigure(0, weight=1)
            preview_frame.grid_columnconfigure(0, weight=1)

            self.cards_frame = tk.Frame(preview_frame)
            self.cards_frame.grid(row=0, column=0, sticky="nsew")

            # Mapping voor combobox per productie
            self.combo_by_prod = {prod: combo for prod, combo in self.rows}

            # Buttons bar (altijd zichtbaar)
            btns = tk.Frame(self)
            btns.grid(row=1, column=0, sticky="ew", padx=10, pady=(6,10))
            btns.grid_columnconfigure(0, weight=1)
            self.remember_var = tk.BooleanVar(value=True)
            tk.Checkbutton(btns, text="Onthoud keuze per productie", variable=self.remember_var).grid(row=0, column=0, sticky="w")
            tk.Button(btns, text="Annuleer", command=self._cancel).grid(row=0, column=1, sticky="e", padx=(4,0))
            tk.Button(btns, text="Bevestig", command=self._confirm).grid(row=0, column=2, sticky="e")

            # Init
            self._refresh_options(initial=True)
            self._update_preview_from_any_combo()

        def _on_focus_prod(self, prod: str):
            self._active_prod = prod

        def _display_list(self) -> List[str]:
            sups = self.db.suppliers_sorted()
            opts = [self.db.display_name(s) for s in sups]
            opts.insert(0, "(geen)")
            return opts

        def _refresh_options(self, initial=False):
            self._base_options = self._display_list()
            self._disp_to_name = {}
            src = self.db.suppliers_sorted()
            for s in src:
                self._disp_to_name[self.db.display_name(s)] = s.supplier

            for prod, combo in self.rows:
                typed = combo.get()
                combo["values"] = self._base_options
                lower_prod = prod.strip().lower()
                if lower_prod in ("dummy part", "nan", "spare part"):
                    combo.set(self._base_options[0])
                    continue
                if typed:
                    combo.set(typed)
                    continue
                name = self.db.get_default(prod)
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
                    combo.set(disp)
                elif self._base_options:
                    combo.set(
                        self._base_options[1]
                        if len(self._base_options) > 1
                        else self._base_options[0]
                    )

            delivery_opts = [
                "Geen",
                "Bestelling wordt opgehaald",
                "Leveradres wordt nog meegedeeld",
            ] + [
                self.delivery_db.display_name(a)
                for a in self.delivery_db.addresses_sorted()
            ]
            for prod, dcombo in self.delivery_combos.items():
                cur = dcombo.get()
                dcombo["values"] = delivery_opts
                if cur:
                    dcombo.set(cur)

        def _on_combo_change(self, _evt=None):
            self._update_preview_from_any_combo()

        def _on_doc_type_change(self, prod: str):
            doc_var = self.doc_vars.get(prod)
            doc_num_var = self.doc_num_vars.get(prod)
            if not doc_var or not doc_num_var:
                return
            cur = doc_num_var.get()
            prefix = _prefix_for_doc_type(doc_var.get())
            prefixes = getattr(self, "_doc_type_prefixes", {prefix})
            if not cur or cur in prefixes:
                doc_num_var.set(prefix)

        def _on_combo_type(self, evt, production: str, combo):
            self._active_prod = production
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
            self._populate_cards(filtered, production)
            if evt.keysym == "Return" and len(filtered) == 1:
                combo.set(filtered[0])
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
            for prod, combo in self.rows:
                t = combo.get()
                if t:
                    self._active_prod = prod
                    self._update_preview_for_text(t)
                    self._populate_cards([t], prod)
                    return
            self._preview_supplier = None
            self._populate_cards([], self._active_prod if self._active_prod else None)

        def _on_card_click(self, option: str, production: str):
            combo = self.combo_by_prod.get(production)
            if combo:
                combo.set(option)
            self._active_prod = production
            self._update_preview_for_text(option)
            self._populate_cards([option], production)

        def _populate_cards(self, options, production):
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
                handler = lambda _e, o=opt, p=production: self._on_card_click(o, p)
                card.bind("<Button-1>", handler)
                for w in widgets:
                    w.bind("<Button-1>", handler)

        def _cancel(self):
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
            sel_map: Dict[str, str] = {}
            doc_map: Dict[str, str] = {}
            for prod, combo in self.rows:
                typed = combo.get().strip()
                if not typed or typed.lower() in ("(geen)", "geen"):
                    sel_map[prod] = ""
                else:
                    s = self._resolve_text_to_supplier(typed)
                    if s:
                        sel_map[prod] = s.supplier
                doc_map[prod] = self.doc_vars.get(prod, tk.StringVar(value="Bestelbon")).get()

            doc_num_map: Dict[str, str] = {}
            delivery_map: Dict[str, str] = {}
            for prod, _combo in self.rows:
                doc_num_map[prod] = self.doc_num_vars[prod].get().strip()
                delivery_map[prod] = self.delivery_vars.get(prod, tk.StringVar(value="Geen")).get()

            project_number = self.project_number_var.get().strip()
            project_name = self.project_name_var.get().strip()

            self.callback(
                sel_map,
                doc_map,
                doc_num_map,
                delivery_map,
                project_number,
                project_name,
                bool(self.remember_var.get()),
            )

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            import sys
            style = ttk.Style(self)
            if sys.platform == "darwin":
                style.theme_use("aqua")
            else:
                style.theme_use("clam")
            self.title("File Hopper – Miami Vice Edition (Dual-mode)")
            self.minsize(1024, 720)

            self.db = SuppliersDB.load(SUPPLIERS_DB_FILE)
            self.client_db = ClientsDB.load(CLIENTS_DB_FILE)
            self.delivery_db = DeliveryAddressesDB.load(DELIVERY_DB_FILE)

            self.source_folder = ""
            self.dest_folder = ""
            self.bom_df: Optional[pd.DataFrame] = None

            self.nb = ttk.Notebook(self)
            self.nb.pack(fill="both", expand=True)
            main = tk.Frame(self.nb)
            self.nb.add(main, text="Main")
            self.main_frame = main
            self.clients_frame = ClientsManagerFrame(
                self.nb, self.client_db, on_change=self._on_db_change
            )
            self.nb.add(self.clients_frame, text="Klant beheer")
            self.delivery_frame = DeliveryAddressesManagerFrame(
                self.nb, self.delivery_db, on_change=self._on_db_change
            )
            self.nb.add(self.delivery_frame, text="Leveradres beheer")
            self.suppliers_frame = SuppliersManagerFrame(
                self.nb, self.db, on_change=self._on_db_change
            )
            self.nb.add(self.suppliers_frame, text="Leverancier beheer")

            # Top folders
            top = tk.Frame(main); top.pack(fill="x", padx=8, pady=6)
            tk.Label(top, text="Bronmap:").grid(row=0, column=0, sticky="w")
            self.src_entry = tk.Entry(top, width=60); self.src_entry.grid(row=0, column=1, padx=4)
            tk.Button(top, text="Bladeren", command=self._pick_src).grid(row=0, column=2, padx=4)

            tk.Label(top, text="Bestemmingsmap:").grid(row=1, column=0, sticky="w")
            self.dst_entry = tk.Entry(top, width=60); self.dst_entry.grid(row=1, column=1, padx=4)
            tk.Button(top, text="Bladeren", command=self._pick_dst).grid(row=1, column=2, padx=4)

            tk.Label(top, text="Opdrachtgever:").grid(row=2, column=0, sticky="w")
            self.client_var = tk.StringVar()
            self.client_combo = ttk.Combobox(top, textvariable=self.client_var, state="readonly", width=40)
            self.client_combo.grid(row=2, column=1, padx=4)
            tk.Button(top, text="Beheer", command=lambda: self.nb.select(self.clients_frame)).grid(row=2, column=2, padx=4)
            self._refresh_clients_combo()



            # Filters
            filt = tk.LabelFrame(main, text="Selecteer bestandstypen om te kopiëren", labelanchor="n"); filt.pack(fill="x", padx=8, pady=6)
            self.pdf_var = tk.IntVar(); self.step_var = tk.IntVar(); self.dxf_var = tk.IntVar(); self.dwg_var = tk.IntVar()
            self.zip_var = tk.IntVar()
            tk.Checkbutton(filt, text="PDF (.pdf)", variable=self.pdf_var).pack(anchor="w", padx=8)
            tk.Checkbutton(filt, text="STEP (.step, .stp)", variable=self.step_var).pack(anchor="w", padx=8)
            tk.Checkbutton(filt, text="DXF (.dxf)", variable=self.dxf_var).pack(anchor="w", padx=8)
            tk.Checkbutton(filt, text="DWG (.dwg)", variable=self.dwg_var).pack(anchor="w", padx=8)

            # BOM controls
            bf = tk.Frame(main); bf.pack(fill="x", padx=8, pady=6)
            tk.Button(bf, text="Laad BOM (CSV/Excel)", command=self._load_bom).pack(side="left", padx=6)
            tk.Button(bf, text="Controleer Bestanden", command=self._check_files).pack(side="left", padx=6)

            pnf = tk.Frame(main); pnf.pack(fill="x", padx=8, pady=(0,6))
            tk.Label(pnf, text="PartNumbers (één per lijn):").pack(anchor="w")
            txtf = tk.Frame(pnf); txtf.pack(fill="x")
            self.pn_text = tk.Text(txtf, height=4)
            pn_scroll = ttk.Scrollbar(txtf, orient="vertical", command=self.pn_text.yview)
            self.pn_text.configure(yscrollcommand=pn_scroll.set)
            self.pn_text.pack(side="left", fill="both", expand=True)
            pn_scroll.pack(side="left", fill="y")
            tk.Button(pnf, text="Gebruik PartNumbers", command=self._load_manual_pns).pack(anchor="w", pady=4)

            # Tree
            style.configure("Treeview", rowheight=24)
            treef = tk.Frame(main)
            treef.pack(fill="both", expand=True, padx=8, pady=6)
            self.tree = ttk.Treeview(treef, columns=("PartNumber","Description","Production","Bestanden gevonden","Status"), show="headings")
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
            self.item_links: Dict[str, str] = {}

            # Actions
            act = tk.Frame(main); act.pack(fill="x", padx=8, pady=8)
            tk.Button(act, text="Kopieer zonder submappen", command=self._copy_flat).pack(side="left", padx=6)
            tk.Button(act, text="Kopieer per productie + bestelbonnen", command=self._copy_per_prod).pack(side="left", padx=6)
            tk.Checkbutton(act, text="Zip per productie", variable=self.zip_var).pack(side="left", padx=6)
            tk.Button(act, text="Combine pdf", command=self._combine_pdf).pack(side="left", padx=6)

            # Status
            self.status_var = tk.StringVar(value="Klaar")
            tk.Label(main, textvariable=self.status_var, anchor="w").pack(fill="x", padx=8, pady=(0,8))

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

        def _pick_src(self):
            from tkinter import filedialog
            p = filedialog.askdirectory()
            if p: self.source_folder = p; self.src_entry.delete(0, "end"); self.src_entry.insert(0, p)

        def _pick_dst(self):
            from tkinter import filedialog
            p = filedialog.askdirectory()
            if p: self.dest_folder = p; self.dst_entry.delete(0, "end"); self.dst_entry.insert(0, p)

        def _selected_exts(self) -> Optional[List[str]]:
            exts = []
            if self.pdf_var.get(): exts.append(".pdf")
            if self.step_var.get(): exts += [".step",".stp"]
            if self.dxf_var.get(): exts.append(".dxf")
            if self.dwg_var.get(): exts.append(".dwg")
            return exts or None

        def _load_bom(self):
            from tkinter import filedialog, messagebox
            start_dir = self.source_folder if self.source_folder else os.getcwd()
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("Excel","*.xlsx;*.xls")], initialdir=start_dir)
            if not path: return
            try:
                self.bom_df = load_bom(path)
                if "Bestanden gevonden" not in self.bom_df.columns: self.bom_df["Bestanden gevonden"]=""
                if "Status" not in self.bom_df.columns: self.bom_df["Status"]=""
                self._refresh_tree()
                self.status_var.set(f"BOM geladen: {len(self.bom_df)} rijen")
            except Exception as e:
                messagebox.showerror("Fout", str(e))

        def _load_manual_pns(self):
            text = self.pn_text.get("1.0", "end").strip()
            if not text:
                return
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if not lines:
                return
            n = len(lines)
            self.bom_df = pd.DataFrame(
                {
                    "PartNumber": lines,
                    "Description": ["" for _ in range(n)],
                    "Production": ["" for _ in range(n)],
                    "Bestanden gevonden": ["" for _ in range(n)],
                    "Status": ["" for _ in range(n)],
                    "Materiaal": ["" for _ in range(n)],
                    "Aantal": [1 for _ in range(n)],
                    "Oppervlakte": ["" for _ in range(n)],
                    "Gewicht": ["" for _ in range(n)],
                }
            )
            self._refresh_tree()
            self.status_var.set(f"Partnummers geladen: {n} rijen")

        def _refresh_tree(self):
            self.item_links.clear()
            for it in self.tree.get_children():
                self.tree.delete(it)
            if self.bom_df is None:
                return
            for _, row in self.bom_df.iterrows():
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
            if self.bom_df is None:
                messagebox.showwarning("Let op", "Laad eerst een BOM."); return
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
            exts = self._selected_exts()
            if not exts or not self.source_folder or not self.dest_folder:
                messagebox.showwarning("Let op", "Selecteer bron, bestemming en extensies."); return
            def work():
                self.status_var.set("Kopiëren...")
                idx = _build_file_index(self.source_folder, exts)
                cnt = 0
                for _, paths in idx.items():
                    for p in paths:
                        dst = os.path.join(self.dest_folder, os.path.basename(p))
                        shutil.copy2(p, dst)
                        cnt += 1
                self.status_var.set(f"Gekopieerd: {cnt}")
            threading.Thread(target=work, daemon=True).start()

        def _copy_per_prod(self):
            from tkinter import messagebox
            if self.bom_df is None:
                messagebox.showwarning("Let op", "Laad eerst een BOM."); return
            exts = self._selected_exts()
            if not exts or not self.source_folder or not self.dest_folder:
                messagebox.showwarning("Let op", "Selecteer bron, bestemming en extensies."); return

            prods = sorted(
                set(
                    (str(r.get("Production") or "").strip() or "_Onbekend")
                    for _, r in self.bom_df.iterrows()
                )
            )
            sel_frame = None

            def on_sel(
                sel_map: Dict[str, str],
                doc_map: Dict[str, str],
                doc_num_map: Dict[str, str],
                delivery_map_raw: Dict[str, str],
                project_number: str,
                project_name: str,
                remember: bool,
            ):
                def work():
                    self.status_var.set("Kopiëren & bestelbonnen maken...")
                    client = self.client_db.get(
                        self.client_var.get().replace("★ ", "", 1)
                    )
                    resolved_delivery_map: Dict[str, DeliveryAddress | None] = {}
                    for prod, name in delivery_map_raw.items():
                        clean = name.replace("★ ", "", 1)
                        if clean == "Geen":
                            resolved_delivery_map[prod] = None
                        elif clean in (
                            "Bestelling wordt opgehaald",
                            "Leveradres wordt nog meegedeeld",
                        ):
                            resolved_delivery_map[prod] = DeliveryAddress(name=clean)
                        else:
                            addr = self.delivery_db.get(clean)
                            resolved_delivery_map[prod] = addr
                    cnt, chosen = copy_per_production_and_orders(
                        self.source_folder,
                        self.dest_folder,
                        self.bom_df,
                        exts,
                        self.db,
                        sel_map,
                        doc_map,
                        doc_num_map,
                        remember,
                        client=client,
                        delivery_map=resolved_delivery_map,
                        footer_note=DEFAULT_FOOTER_NOTE,
                        zip_parts=bool(self.zip_var.get()),
                        project_number=project_number,
                        project_name=project_name,
                    )

                    def on_done():
                        self.status_var.set(
                            f"Klaar. Gekopieerd: {cnt}. Leveranciers: {chosen}"
                        )
                        messagebox.showinfo("Klaar", "Bestelbonnen aangemaakt.")
                        if sys.platform.startswith("win"):
                            try:
                                os.startfile(self.dest_folder)
                            except Exception:
                                pass
                        if getattr(self, "sel_frame", None):
                            try:
                                self.nb.forget(self.sel_frame)
                                self.sel_frame.destroy()
                            except Exception:
                                pass
                            self.sel_frame = None
                        self.nb.select(self.main_frame)

                    self.after(0, on_done)

                threading.Thread(target=work, daemon=True).start()

            sel_frame = SupplierSelectionFrame(
                self.nb, prods, self.db, self.delivery_db, on_sel
            )
            self.sel_frame = sel_frame
            self.nb.add(sel_frame, state="hidden")
            self.nb.select(sel_frame)

        def _combine_pdf(self):
            from tkinter import messagebox
            if self.source_folder and self.bom_df is not None:
                def work():
                    self.status_var.set("PDF's combineren...")
                    try:
                        out_dir = self.dest_folder or self.source_folder
                        cnt = combine_pdfs_from_source(
                            self.source_folder, self.bom_df, out_dir
                        )
                    except ModuleNotFoundError:
                        self.status_var.set("PyPDF2 ontbreekt")
                        messagebox.showwarning(
                            "PyPDF2 ontbreekt",
                            "Installeer PyPDF2 om PDF's te combineren.",
                        )
                        return
                    self.status_var.set(f"Gecombineerde pdf's: {cnt}")
                    messagebox.showinfo("Klaar", "PDF's gecombineerd.")
                threading.Thread(target=work, daemon=True).start()
            elif self.dest_folder:
                def work():
                    self.status_var.set("PDF's combineren...")
                    try:
                        cnt = combine_pdfs_per_production(self.dest_folder)
                    except ModuleNotFoundError:
                        self.status_var.set("PyPDF2 ontbreekt")
                        messagebox.showwarning(
                            "PyPDF2 ontbreekt",
                            "Installeer PyPDF2 om PDF's te combineren.",
                        )
                        return
                    self.status_var.set(f"Gecombineerde pdf's: {cnt}")
                    messagebox.showinfo("Klaar", "PDF's gecombineerd.")
                threading.Thread(target=work, daemon=True).start()
            else:
                messagebox.showwarning(
                    "Let op", "Selecteer bron + BOM of bestemmingsmap."
                )

    App().mainloop()

