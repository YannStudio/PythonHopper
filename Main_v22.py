#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
File Hopper – dual-mode (GUI if Tkinter available, else CLI)

- Suppliers JSON (suppliers_db.json)
  Fields:
    supplier (naam, required, uniek), description (beschrijving),
    supplier_id, adres_1, adres_2, postcode, gemeente, land,
    btw, contact_sales, sales_email, phone, favorite (bool)

- Defaults per production: defaults_by_production: {production: supplier_name}

- BOM min: PartNumber, Description, Production
  optioneel: Aantal (default 1, max 999), Oppervlakte (m²/m2/Area), Gewicht (kg/Weight), Materiaal (Material)
"""

import os, sys, json, shutil, argparse, tempfile, datetime, threading
from typing import List, Dict, Optional, Tuple, Any

import pandas as pd


# --------------------------- Helpers ---------------------------
from helpers import (
    _to_str,
    _build_file_index,
    _unique_path,
)
from models import Supplier
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE

# --------------------------- BOM & files ---------------------------
from bom import read_csv_flex, load_bom

# --------------------------- Orders ---------------------------
from orders import copy_per_production_and_orders, DEFAULT_FOOTER_NOTE

# --------------------------- CLI ---------------------------

def parse_exts(s: str) -> List[str]:
    parts = [p.strip().lower() for p in s.split(",") if p.strip()]
    m = []
    for p in parts:
        if p in ("pdf","dxf","dwg","step","stp"):
            if p=="step": m += [".step",".stp"]
            elif p=="stp": m += [".stp",".step"]
            else: m.append("." + p)
    if not m:
        raise ValueError("Geen geldige extensies opgegeven (pdf, dxf, dwg, step, stp).")
    return sorted(set(m))

def cli_suppliers(args):
    db = SuppliersDB.load(SUPPLIERS_DB_FILE)
    if args.action == "list":
        rows = db.suppliers_sorted()
        if not rows:
            print("(geen leveranciers)")
            return 0
        for s in rows:
            star = "★" if s.favorite else " "
            if s.adres_1 or s.adres_2:
                addr = ", ".join([x for x in [s.adres_1, s.adres_2] if x])
            else:
                addr = " | ".join([x for x in [s.adres_1, " ".join([_to_str(s.postcode), _to_str(s.gemeente)]).strip(), s.land] if x])
            print(f"{star} {s.supplier}  | Desc: {s.description or '-'} | BTW: {s.btw or '-'} | {addr or '-'} | Mail: {s.sales_email or '-'} | Tel: {s.phone or '-'}")
        return 0
    if args.action == "add":
        rec = {"supplier": args.name}
        if args.description: rec["description"]=args.description
        if args.btw: rec["btw"]=args.btw
        if args.adres_1: rec["adres_1"]=args.adres_1
        if args.adres_2: rec["adres_2"]=args.adres_2
        if args.email: rec["sales_email"]=args.email
        if args.phone: rec["phone"]=args.phone
        s = Supplier.from_any(rec)
        db.upsert(s)
        db.save(SUPPLIERS_DB_FILE)
        print("Toegevoegd/bijgewerkt")
        return 0
    if args.action == "remove":
        ok = db.remove(args.name); db.save(SUPPLIERS_DB_FILE)
        print("Verwijderd" if ok else "Niet gevonden")
        return 0
    if args.action == "fav":
        ok = db.toggle_fav(args.name); db.save(SUPPLIERS_DB_FILE)
        print("Favoriet gewisseld" if ok else "Niet gevonden")
        return 0
    if args.action == "set-default":
        db.set_default(args.production, args.name); db.save(SUPPLIERS_DB_FILE)
        print(f"Default voor {args.production}: {args.name}")
        return 0
    if args.action == "get-default":
        print(db.get_default(args.production) or "(geen)")
        return 0
    if args.action == "import-csv":
        path = args.csv
        if not os.path.exists(path):
            print("CSV niet gevonden."); return 2
        try:
            df = pd.read_csv(path, encoding="latin1", sep=";")
        except Exception:
            df = read_csv_flex(path)
        changed = 0
        for _, row in df.iterrows():
            raw_name = _to_str(row.get("Supplier")).strip()
            if not raw_name or raw_name == "-":
                continue
            try:
                s = Supplier.from_any({k:row[k] for k in df.columns if k in row})
                db.upsert(s); changed += 1
            except Exception:
                pass
        db.save(SUPPLIERS_DB_FILE)
        print(f"Verwerkt (upsert): {changed}")
        return 0
    if args.action == "clear":
        db.clear_all(); db.save(SUPPLIERS_DB_FILE)
        print("Alle leveranciers verwijderd.")
        return 0
    print("Onbekende actie"); return 2

def cli_bom_check(args):
    exts = parse_exts(args.exts)
    df = load_bom(args.bom)
    if not os.path.isdir(args.source):
        print("Bronmap ongeldig"); return 2
    file_index = _build_file_index(args.source, exts)
    found, status = [], []
    for _, row in df.iterrows():
        pn = row["PartNumber"]
        hits = file_index.get(pn, [])
        found.append(", ".join(sorted({os.path.splitext(h)[1].lstrip('.') for h in hits})))
        status.append("Gevonden" if hits else "Ontbrekend")
    df["Bestanden gevonden"] = found
    df["Status"] = status
    if args.out:
        if args.out.lower().endswith(".xlsx"):
            df.to_excel(args.out, index=False, engine="openpyxl")
        else:
            df.to_csv(args.out, index=False)
        print("Weergegeven naar", args.out)
    else:
        print(df.head(20).to_string(index=False))
    return 0

def cli_copy(args):
    exts = parse_exts(args.exts)
    if not os.path.isdir(args.source) or not os.path.isdir(args.dest):
        print("Bron of bestemming ongeldig"); return 2
    idx = _build_file_index(args.source, exts)
    cnt = 0
    for _, paths in idx.items():
        for p in paths:
            dst = _unique_path(os.path.join(args.dest, os.path.basename(p)))
            shutil.copy2(p, dst)
            cnt += 1
    print("Gekopieerd:", cnt)
    return 0

def cli_copy_per_prod(args):
    exts = parse_exts(args.exts)
    df = load_bom(args.bom)
    db = SuppliersDB.load(SUPPLIERS_DB_FILE)
    override_map = dict(kv.split("=",1) for kv in (args.supplier or []))
    cnt, chosen = copy_per_production_and_orders(
        args.source, args.dest, df, exts, db, override_map, args.remember_defaults, footer_note=args.note or DEFAULT_FOOTER_NOTE
    )
    print("Gekopieerd:", cnt)
    for k,v in chosen.items():
        print(f"  {k} → {v}")
    return 0

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="File Hopper dual-mode")
    p.add_argument("--run-tests", action="store_true", help="Run basic self-tests and exit")
    sub = p.add_subparsers(dest="cmd")

    sp = sub.add_parser("suppliers", help="Beheer leveranciers")
    ssp = sp.add_subparsers(dest="action", required=True)
    ssp.add_parser("list")
    ap = ssp.add_parser("add"); ap.add_argument("name"); ap.add_argument("--description"); ap.add_argument("--btw"); ap.add_argument("--adres-1", dest="adres_1"); ap.add_argument("--adres-2", dest="adres_2"); ap.add_argument("--email"); ap.add_argument("--phone")
    rp = ssp.add_parser("remove"); rp.add_argument("name")
    fp = ssp.add_parser("fav"); fp.add_argument("name")
    sdp = ssp.add_parser("set-default"); sdp.add_argument("production"); sdp.add_argument("name")
    gdp = ssp.add_parser("get-default"); gdp.add_argument("production")
    ip = ssp.add_parser("import-csv"); ip.add_argument("csv")
    ssp.add_parser("clear")

    bp = sub.add_parser("bom", help="BOM acties"); bsp = bp.add_subparsers(dest="bact", required=True)
    bcp = bsp.add_parser("check"); bcp.add_argument("--source", required=True); bcp.add_argument("--bom", required=True); bcp.add_argument("--exts", required=True); bcp.add_argument("--out")

    cp = sub.add_parser("copy", help="Kopieer vlak"); cp.add_argument("--source", required=True); cp.add_argument("--dest", required=True); cp.add_argument("--exts", required=True)

    cpp = sub.add_parser("copy-per-prod", help="Kopieer per productie + bestelbonnen")
    cpp.add_argument("--source", required=True); cpp.add_argument("--dest", required=True); cpp.add_argument("--bom", required=True); cpp.add_argument("--exts", required=True)
    cpp.add_argument("--supplier", action="append", help="Override: Production=Supplier (meerdere keren mogelijk)")
    cpp.add_argument("--remember-defaults", action="store_true")
    cpp.add_argument("--note", help="Optioneel voetnootje op de bestelbon", default="")

    return p


# --------------------------- Tests ---------------------------

def run_tests() -> int:
    print("Running self-tests...")
    db = SuppliersDB()
    db.upsert(Supplier.from_any({
        "supplier":"ACME","description":"Snijwerk","btw":"BE123",
        "adress_1":"Teststraat 1 bus 2","address_2":"BE-2000 Antwerpen","land":"BE",
        "e-mail sales":"x@y.z","tel. sales":"+32 123"
    }))
    assert db.suppliers and db.suppliers[0].adres_2 == "BE-2000 Antwerpen"
    db.toggle_fav("ACME"); assert db.suppliers[0].favorite
    db.set_default("Laser","ACME"); assert db.get_default("Laser") == "ACME"

    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td,"src"); dst = os.path.join(td,"dst"); os.makedirs(src); os.makedirs(dst)
        open(os.path.join(src,"PN1.pdf"),"wb").write(b"%PDF-1.4")
        open(os.path.join(src,"PN1.stp"),"wb").write(b"step")
        bom = os.path.join(td,"bom.xlsx")
        df = pd.DataFrame([
            {"PartNumber":"PN1-THIS-IS-A-VERY-LONG-CODE-OVER-25CHARS","Description":"Lange omschrijving die netjes moet wrappen.","Production":"Laser","Aantal":2,"Materiaal":"S235JR","Oppervlakte (m²)":"1,23","Gewicht (kg)":"4,56"},
            {"PartNumber":"PN2","Description":"Geen files","Production":"Laser","Aantal":1000,"Material":"Alu 5754","Area":"0.50","Weight":"1.00"},
        ])
        df.to_excel(bom, index=False, engine="openpyxl")
        ldf = load_bom(bom)
        assert ldf["Aantal"].max() <= 999  # capped
        cnt, chosen = copy_per_production_and_orders(src, dst, ldf, [".pdf",".stp"], db, {}, True, footer_note=DEFAULT_FOOTER_NOTE)
        assert cnt == 2
        assert chosen.get("Laser") == "ACME"
        prod_folder = os.path.join(dst,"Laser")
        assert os.path.exists(os.path.join(prod_folder, "PN1.pdf"))
        assert os.path.exists(os.path.join(prod_folder, "PN1.stp"))
        xlsx = [f for f in os.listdir(prod_folder) if f.lower().endswith(".xlsx")]
        assert xlsx, "Excel bestelbon niet aangemaakt"
    print("All tests passed.")
    return 0


# --------------------------- GUI (Tkinter) ---------------------------

def start_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog

    BG_COLOR = "#2B1B2B"
    BTN_BG = "#FF77FF"
    BTN_FG = "#1A1A1A"
    TXT_FG = "#F8C4FF"
    TREE_ODD_BG = "#3A213A"
    TREE_EVEN_BG = "#4D2C4D"

    class SuppliersManagerWin(tk.Toplevel):
        def __init__(self, master, db: SuppliersDB, on_change=None):
            super().__init__(master)
            self.title("Leveranciers Beheer")
            self.configure(bg=BG_COLOR)
            self.db = db
            self.on_change = on_change
            self.minsize(960, 480)

            # Bovenbalk: Zoek links, knoppen rechts
            topbar = tk.Frame(self, bg=BG_COLOR); topbar.pack(fill="x", padx=8, pady=(8,4))
            left = tk.Frame(topbar, bg=BG_COLOR); left.pack(side="left", fill="x", expand=True)
            tk.Label(left, text="Zoek:", bg=BG_COLOR, fg=TXT_FG).pack(side="left")
            self.search_var = tk.StringVar()
            se = tk.Entry(left, textvariable=self.search_var, bg="#3A213A", fg=TXT_FG, insertbackground=TXT_FG, width=32)
            se.pack(side="left", padx=6)
            se.bind("<KeyRelease>", lambda e: self.refresh())

            btns = tk.Frame(topbar, bg=BG_COLOR); btns.pack(side="right")
            tk.Button(btns, text="Toevoegen", bg=BTN_BG, fg=BTN_FG, command=self.add_supplier).pack(side="left", padx=4)
            tk.Button(btns, text="Verwijderen", bg=BTN_BG, fg=BTN_FG, command=self.remove_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Favoriet ★", bg=BTN_BG, fg=BTN_FG, command=self.toggle_fav_sel).pack(side="left", padx=4)
            tk.Button(btns, text="Update uit CSV (merge)", bg=BTN_BG, fg=BTN_FG, command=self.update_from_csv).pack(side="left", padx=4)
            tk.Button(btns, text="Alles verwijderen", bg="#ffb3b3", fg="#1A1A1A", command=self.clear_all).pack(side="left", padx=4)
            tk.Button(btns, text="Sluiten", bg=BTN_BG, fg=BTN_FG, command=self.destroy).pack(side="left", padx=4)

            # Tabel: verberg postcode, gemeente, land
            cols = ("★","Supplier","Description","BTW","E-mail","Tel","Adres_1","Adres_2")
            self.tree = ttk.Treeview(self, columns=cols, show="headings")
            widths = {"★":40,"Supplier":190,"Description":240,"BTW":120,"E-mail":200,"Tel":130,"Adres_1":220,"Adres_2":220}
            for c in cols:
                self.tree.heading(c, text=c)
                self.tree.column(c, width=widths.get(c,120), anchor="w")
            self.tree.pack(fill="both", expand=True, padx=8, pady=(4,8))

            style = ttk.Style(self)
            style.theme_use("clam")
            style.configure("Treeview", background=BG_COLOR, fieldbackground=BG_COLOR, foreground=TXT_FG, rowheight=22)
            self.tree.tag_configure("oddrow", background=TREE_ODD_BG)
            self.tree.tag_configure("evenrow", background=TREE_EVEN_BG)

            self.refresh()

        def refresh(self):
            for i in self.tree.get_children():
                self.tree.delete(i)
            q = self.search_var.get()
            rows = self.db.find(q) if q else self.db.suppliers_sorted()
            for idx, s in enumerate(rows):
                star = "★" if s.favorite else ""
                vals = (star, s.supplier, s.description or "", s.btw or "", s.sales_email or "", s.phone or "",
                        s.adres_1 or "", s.adres_2 or "")
                self.tree.insert("", "end", values=vals, tags=("evenrow" if idx%2==0 else "oddrow",))

        def _sel_name(self) -> Optional[str]:
            it = self.tree.selection()
            if not it: return None
            vals = self.tree.item(it[0], "values")
            return vals[1]

        def add_supplier(self):
            name = simpledialog.askstring("Nieuwe leverancier","Naam (Supplier):", parent=self)
            if not name: return
            s = Supplier.from_any({"supplier":name})
            self.db.upsert(s)
            self.db.save(SUPPLIERS_DB_FILE)
            self.refresh()
            if self.on_change: self.on_change()

        def remove_sel(self):
            n = self._sel_name()
            if not n: return
            from tkinter import messagebox
            if messagebox.askyesno("Bevestigen", f"Verwijder '{n}'?"):
                if self.db.remove(n):
                    self.db.save(SUPPLIERS_DB_FILE)
                    self.refresh()
                    if self.on_change: self.on_change()

        def toggle_fav_sel(self):
            n = self._sel_name()
            if not n: return
            if self.db.toggle_fav(n):
                self.db.save(SUPPLIERS_DB_FILE)
                self.refresh()
                if self.on_change: self.on_change()

        def update_from_csv(self):
            from tkinter import filedialog, messagebox
            path = filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("All","*.*")], initialdir=os.getcwd())
            if not path: return
            try:
                try:
                    df = pd.read_csv(path, encoding="latin1", sep=";")
                except Exception:
                    df = read_csv_flex(path)
                changed = 0
                for _, row in df.iterrows():
                    raw_name = _to_str(row.get("Supplier")).strip()
                    if not raw_name or raw_name == "-":
                        continue
                    try:
                        s = Supplier.from_any({k:row[k] for k in df.columns if k in row})
                        self.db.upsert(s)
                        changed += 1
                    except Exception:
                        pass
                self.db.save(SUPPLIERS_DB_FILE)
                messagebox.showinfo("CSV update", f"{changed} records verwerkt (merge/upsert).")
                self.refresh()
                if self.on_change: self.on_change()
            except Exception as e:
                messagebox.showerror("Fout", f"Update mislukt:\n{e}")

        def clear_all(self):
            from tkinter import messagebox
            if messagebox.askyesno("Bevestigen", "Wil je echt ALLE leveranciers verwijderen?"):
                self.db.clear_all()
                self.db.save(SUPPLIERS_DB_FILE)
                self.refresh()
                if self.on_change: self.on_change()

    class SupplierSelectionPopup(tk.Toplevel):
        """Per productie: type-to-filter of dropdown; rechts detailkaart (klik = selecteer).
           Knoppen altijd zichtbaar onderaan.
        """
        def __init__(self, master, productions: List[str], db: SuppliersDB, callback):
            super().__init__(master)
            self.title("Selecteer leveranciers per productie")
            self.configure(bg=BG_COLOR)
            self.db = db
            self.callback = callback
            self._preview_supplier: Optional[Supplier] = None
            self._active_prod: Optional[str] = None  # laatst gefocuste rij
            self.sel_vars: Dict[str, tk.StringVar] = {}

            # Grid layout: content (row=0, weight=1), buttons (row=1)
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)

            content = tk.Frame(self, bg=BG_COLOR)
            content.grid(row=0, column=0, sticky="nsew", padx=10, pady=6)
            content.grid_columnconfigure(0, weight=1)  # left
            content.grid_columnconfigure(1, weight=0)  # right

            # Left: per productie comboboxen
            left = tk.Frame(content, bg=BG_COLOR)
            left.grid(row=0, column=0, sticky="nw", padx=(0,8))
            self.rows = []
            for prod in productions:
                row = tk.Frame(left, bg=BG_COLOR)
                row.pack(fill="x", pady=3)
                tk.Label(row, text=prod, bg=BG_COLOR, fg=TXT_FG, width=18, anchor="w").pack(side="left")
                var = tk.StringVar()
                self.sel_vars[prod] = var
                combo = ttk.Combobox(row, textvariable=var, state="normal", width=50)
                combo.pack(side="left", padx=6)
                combo.bind("<<ComboboxSelected>>", self._on_combo_change)
                combo.bind("<FocusIn>", lambda _e, p=prod: self._on_focus_prod(p))
                combo.bind("<KeyRelease>", lambda ev, p=prod, c=combo: self._on_combo_type(ev, p, c))
                self.rows.append((prod, combo))

            # Right: preview details (klikbaar) in LabelFrame met ondertitel
            right = tk.LabelFrame(content,
                                  text="Leverancier details\n(klik om te selecteren)",
                                  fg=TXT_FG, bg=BG_COLOR, labelanchor="n")
            right.grid(row=0, column=1, sticky="ne", padx=(8,0))
            self.preview = tk.Label(right, text="", justify="left", bg=BG_COLOR, fg=TXT_FG, anchor="nw", cursor="hand2")
            self.preview.pack(fill="both", expand=True, padx=8, pady=8)
            self.preview.configure(wraplength=360)
            self.preview.bind("<Button-1>", self._on_preview_click)

            # Buttons bar (altijd zichtbaar)
            btns = tk.Frame(self, bg=BG_COLOR)
            btns.grid(row=1, column=0, sticky="ew", padx=10, pady=(6,10))
            btns.grid_columnconfigure(0, weight=1)
            self.remember_var = tk.BooleanVar(value=True)
            tk.Checkbutton(btns, text="Onthoud keuze per productie", variable=self.remember_var,
                           bg=BG_COLOR, fg=TXT_FG, selectcolor=BG_COLOR,
                           activebackground=BG_COLOR, activeforeground=TXT_FG).grid(row=0, column=0, sticky="w")
            tk.Button(btns, text="Annuleer", bg=BTN_BG, fg=BTN_FG, command=self.destroy).grid(row=0, column=1, sticky="e", padx=(4,0))
            tk.Button(btns, text="Bevestig", bg=BTN_BG, fg=BTN_FG, command=self._confirm).grid(row=0, column=2, sticky="e")

            # Init
            self._refresh_options(initial=True)
            self._update_preview_from_any_combo()

            # Compact formaat
            self.update_idletasks()
            w = min(920, self.winfo_reqwidth()+40)
            h = min(600, self.winfo_reqheight()+20)
            self.geometry(f"{w}x{h}")

        def _on_focus_prod(self, prod: str):
            self._active_prod = prod

        def _display_list(self) -> List[str]:
            sups = self.db.suppliers_sorted()
            return [self.db.display_name(s) for s in sups]

        def _refresh_options(self, initial=False):
            self._base_options = self._display_list()
            self._disp_to_name = {}
            src = self.db.suppliers_sorted()
            for s in src:
                self._disp_to_name[self.db.display_name(s)] = s.supplier

            for prod, combo in self.rows:
                typed = combo.get()
                combo["values"] = self._base_options
                name = self.db.get_default(prod)
                if not typed:
                    if not name and initial:
                        favs = [x for x in src if x.favorite]
                        name = (favs[0].supplier if favs else (src[0].supplier if src else ""))
                    disp = None
                    for k, v in self._disp_to_name.items():
                        if v and name and v.lower() == name.lower():
                            disp = k; break
                    if disp:
                        combo.set(disp)
                    elif self._base_options:
                        combo.set(self._base_options[0])

        def _on_combo_change(self, _evt=None):
            self._update_preview_from_any_combo()

        def _on_combo_type(self, evt, production: str, combo):
            self._active_prod = production
            text = combo.get().strip().lower()
            if not hasattr(self, "_base_options"):
                return
            if evt.keysym in ("Up","Down","Return","Escape"):
                return
            if not text:
                combo["values"] = self._base_options
            else:
                filtered = [opt for opt in self._base_options if text in opt.lower()]
                combo["values"] = filtered or self._base_options
            self._update_preview_for_text(combo.get())

        def _resolve_text_to_supplier(self, text: str) -> Optional[Supplier]:
            if not text:
                return None
            if hasattr(self, "_disp_to_name") and text in self._disp_to_name:
                target = self._disp_to_name[text]
                for s in self.db.suppliers:
                    if s.supplier.lower() == target.lower():
                        return s
            for s in self.db.suppliers:
                if s.supplier.lower() == text.lower():
                    return s
            cand = [s for s in self.db.suppliers if s.supplier.lower().startswith(text.lower())]
            if cand:
                return sorted(cand, key=lambda x: (not x.favorite, x.supplier.lower()))[0]
            cand = [s for s in self.db.suppliers if text.lower() in s.supplier.lower()]
            if cand:
                return sorted(cand, key=lambda x: (not x.favorite, x.supplier.lower()))[0]
            return None

        def _update_preview_for_text(self, text: str):
            s = self._resolve_text_to_supplier(text)
            self._preview_supplier = s
            if not s:
                self.preview.config(text="")
                return
            addr_line = None
            if s.adres_1 or s.adres_2:
                addr_line = f"{s.adres_1}, {s.adres_2}" if (s.adres_1 and s.adres_2) else (s.adres_1 or s.adres_2)
            lines = [f"{s.supplier}"]
            if s.description: lines.append(f"({s.description})")
            if addr_line: lines.append(addr_line)
            if not addr_line:
                pc_gem = " ".join(x for x in [s.postcode, s.gemeente] if x)
                if pc_gem: lines.append(pc_gem)
                if s.land: lines.append(s.land)
            if s.btw: lines.append(f"BTW: {s.btw}")
            if s.contact_sales: lines.append(f"Contact sales: {s.contact_sales}")
            if s.sales_email: lines.append(f"E-mail: {s.sales_email}")
            if s.phone: lines.append(f"Tel: {s.phone}")
            self.preview.config(text="\n".join(lines))

        def _update_preview_from_any_combo(self):
            for prod, combo in self.rows:
                t = combo.get()
                if t:
                    self._active_prod = prod
                    self._update_preview_for_text(t)
                    return
            self.preview.config(text="")
            self._preview_supplier = None

        def _on_preview_click(self, _evt=None):
            if not self._preview_supplier:
                return
            if not self._active_prod and self.rows:
                self._active_prod = self.rows[0][0]
            for prod, combo in self.rows:
                if prod == self._active_prod:
                    disp = None
                    if not hasattr(self, "_disp_to_name"): self._refresh_options()
                    for k, v in self._disp_to_name.items():
                        if v.lower() == self._preview_supplier.supplier.lower():
                            disp = k; break
                    combo.set(disp or self._preview_supplier.supplier)
                    break

        def _confirm(self):
            result = {}
            for prod, combo in self.rows:
                typed = combo.get()
                s = self._resolve_text_to_supplier(typed)
                if s:
                    result[prod] = s.supplier
            self.callback(result, True if self.remember_var.get() else False)
            self.destroy()

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("File Hopper – Miami Vice Edition (Dual-mode)")
            self.configure(bg=BG_COLOR)
            self.minsize(1024, 720)

            self.db = SuppliersDB.load(SUPPLIERS_DB_FILE)

            self.source_folder = ""
            self.dest_folder = ""
            self.bom_df: Optional[pd.DataFrame] = None

            # Top folders
            top = tk.Frame(self, bg=BG_COLOR); top.pack(fill="x", padx=8, pady=6)
            tk.Label(top, text="Bronmap:", fg=TXT_FG, bg=BG_COLOR).grid(row=0, column=0, sticky="w")
            self.src_entry = tk.Entry(top, width=60, bg="#3A213A", fg=TXT_FG, insertbackground=TXT_FG); self.src_entry.grid(row=0, column=1, padx=4)
            tk.Button(top, text="Bladeren", command=self._pick_src, bg=BTN_BG, fg=BTN_FG).grid(row=0, column=2, padx=4)

            tk.Label(top, text="Bestemmingsmap:", fg=TXT_FG, bg=BG_COLOR).grid(row=1, column=0, sticky="w")
            self.dst_entry = tk.Entry(top, width=60, bg="#3A213A", fg=TXT_FG, insertbackground=TXT_FG); self.dst_entry.grid(row=1, column=1, padx=4)
            tk.Button(top, text="Bladeren", command=self._pick_dst, bg=BTN_BG, fg=BTN_FG).grid(row=1, column=2, padx=4)

            # Filters
            filt = tk.LabelFrame(self, text="Selecteer bestandstypen om te kopiëren", fg=TXT_FG, bg=BG_COLOR, labelanchor="n"); filt.pack(fill="x", padx=8, pady=6)
            self.pdf_var = tk.IntVar(); self.step_var = tk.IntVar(); self.dxf_var = tk.IntVar(); self.dwg_var = tk.IntVar()
            tk.Checkbutton(filt, text="PDF (.pdf)", variable=self.pdf_var, fg=TXT_FG, bg=BG_COLOR, selectcolor=BG_COLOR).pack(anchor="w", padx=8)
            tk.Checkbutton(filt, text="STEP (.step, .stp)", variable=self.step_var, fg=TXT_FG, bg=BG_COLOR, selectcolor=BG_COLOR).pack(anchor="w", padx=8)
            tk.Checkbutton(filt, text="DXF (.dxf)", variable=self.dxf_var, fg=TXT_FG, bg=BG_COLOR, selectcolor=BG_COLOR).pack(anchor="w", padx=8)
            tk.Checkbutton(filt, text="DWG (.dwg)", variable=self.dwg_var, fg=TXT_FG, bg=BG_COLOR, selectcolor=BG_COLOR).pack(anchor="w", padx=8)

            # BOM controls
            bf = tk.Frame(self, bg=BG_COLOR); bf.pack(fill="x", padx=8, pady=6)
            tk.Button(bf, text="Laad BOM (CSV/Excel)", command=self._load_bom, bg=BTN_BG, fg=BTN_FG).pack(side="left", padx=6)
            tk.Button(bf, text="Leveranciers Beheer", command=self._open_suppliers, bg=BTN_BG, fg=BTN_FG).pack(side="left", padx=6)
            tk.Button(bf, text="Controleer Bestanden", command=self._check_files, bg=BTN_BG, fg=BTN_FG).pack(side="left", padx=6)

            # Tree
            style = ttk.Style(self); style.theme_use("clam")
            style.configure("Treeview", background=BG_COLOR, fieldbackground=BG_COLOR, foreground=TXT_FG, rowheight=24)
            self.tree = ttk.Treeview(self, columns=("PartNumber","Description","Production","Bestanden gevonden","Status"), show="headings")
            for col in ("PartNumber","Description","Production","Bestanden gevonden","Status"):
                self.tree.heading(col, text=col)
                w = 140
                if col=="Description": w=320
                if col=="Bestanden gevonden": w=180
                if col=="Status": w=120
                self.tree.column(col, width=w, anchor="w")
            self.tree.pack(fill="both", expand=True, padx=8, pady=6)

            # Actions
            act = tk.Frame(self, bg=BG_COLOR); act.pack(fill="x", padx=8, pady=8)
            tk.Button(act, text="Kopieer zonder submappen", command=self._copy_flat, bg="#c8f7c5", fg="#1A1A1A").pack(side="left", padx=6)
            tk.Button(act, text="Kopieer per productie + bestelbonnen", command=self._copy_per_prod, bg="#bfe0ff", fg="#1A1A1A").pack(side="left", padx=6)

            # Status
            self.status_var = tk.StringVar(value="Klaar")
            tk.Label(self, textvariable=self.status_var, anchor="w", bg=BG_COLOR, fg=TXT_FG).pack(fill="x", padx=8, pady=(0,8))

        def _open_suppliers(self):
            SuppliersManagerWin(self, self.db, on_change=lambda: None)

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

        def _refresh_tree(self):
            for it in self.tree.get_children():
                self.tree.delete(it)
            if self.bom_df is None: return
            for _, row in self.bom_df.iterrows():
                vals = (
                    row.get("PartNumber",""),
                    row.get("Description",""),
                    row.get("Production",""),
                    row.get("Bestanden gevonden",""),
                    row.get("Status","")
                )
                self.tree.insert("", "end", values=vals)

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
            found, status = [], []
            for _, row in self.bom_df.iterrows():
                pn = row["PartNumber"]
                hits = idx.get(pn, [])
                found.append(", ".join(sorted({os.path.splitext(h)[1].lstrip('.') for h in hits})))
                status.append("Gevonden" if hits else "Ontbrekend")
            self.bom_df["Bestanden gevonden"] = found
            self.bom_df["Status"] = status
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
                        dst = _unique_path(os.path.join(self.dest_folder, os.path.basename(p)))
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

            prods = sorted(set((str(r.get("Production") or "").strip() or "_Onbekend") for _, r in self.bom_df.iterrows()))
            def on_sel(sel_map: Dict[str,str], remember: bool):
                def work():
                    self.status_var.set("Kopiëren & bestelbonnen maken...")
                    cnt, chosen = copy_per_production_and_orders(
                        self.source_folder, self.dest_folder, self.bom_df, exts, self.db, sel_map, remember,
                        footer_note=DEFAULT_FOOTER_NOTE
                    )
                    self.status_var.set(f"Klaar. Gekopieerd: {cnt}. Leveranciers: {chosen}")
                    messagebox.showinfo("Klaar", "Bestelbonnen aangemaakt.")
                threading.Thread(target=work, daemon=True).start()
            SupplierSelectionPopup(self, prods, self.db, on_sel)

    App().mainloop()


# --------------------------- Main ---------------------------

def main(argv: Optional[List[str]] = None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.run_tests:
        return run_tests()

    if not args.cmd:
        try:
            import tkinter  # noqa
            start_gui()
            return 0
        except Exception:
            parser.print_help()
            return 0

    if args.cmd == "suppliers": return cli_suppliers(args)
    if args.cmd == "bom": return cli_bom_check(args)
    if args.cmd == "copy": return cli_copy(args)
    if args.cmd == "copy-per-prod": return cli_copy_per_prod(args)
    parser.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
