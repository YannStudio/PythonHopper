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
from gui import start_gui

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
