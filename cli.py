#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Command-line interface helpers for File Hopper."""

import os
import shutil
import argparse
from typing import Iterable, List, Optional, Union

import pandas as pd
from dataclasses import asdict

from helpers import _to_str, _build_file_index, _unique_path, validate_vat
from models import Supplier, Client, DeliveryAddress
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from clients_db import ClientsDB, CLIENTS_DB_FILE
from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_DB_FILE
from bom import read_csv_flex, load_bom
from orders import copy_per_production_and_orders, DEFAULT_FOOTER_NOTE

DEFAULT_ALLOWED_EXTS = "pdf,dxf,dwg,step,stp"


def _normalize_ext(ext: str) -> str:
    ext = ext.strip().lower()
    if ext.startswith("*"):
        ext = ext[1:]
    if not ext.startswith("."):
        ext = "." + ext
    return ext


def _parse_ext_list(src: Union[Iterable[str], str]) -> List[str]:
    if isinstance(src, str):
        parts = src.split(",")
    else:
        parts = list(src)
    return [_normalize_ext(p) for p in parts if p and p.strip()]


def parse_exts(s: str, allowed_exts: Optional[Union[Iterable[str], str]] = None) -> List[str]:
    allowed = _parse_ext_list(allowed_exts or DEFAULT_ALLOWED_EXTS)
    allowed_set = set(allowed)
    if ".step" in allowed_set or ".stp" in allowed_set:
        allowed_set.update({".step", ".stp"})
    parts = _parse_ext_list(s)
    invalid = [p for p in parts if p not in allowed_set]
    if invalid:
        raise ValueError(
            "Ongeldige extensies: {}. Toegestane extensies: {}.".format(
                ", ".join(sorted(i.lstrip(".") for i in invalid)),
                ", ".join(sorted(e.lstrip(".") for e in allowed_set)),
            )
        )
    result: List[str] = []
    for p in parts:
        if p in {".step", ".stp"}:
            result.extend([".step", ".stp"])
        else:
            result.append(p)
    if not result:
        raise ValueError(
            "Geen geldige extensies opgegeven ({}).".format(
                ", ".join(sorted(e.lstrip(".")) for e in allowed_set)
            )
        )
    return sorted(set(result))


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
                addr = " | ".join(
                    [
                        x
                        for x in [
                            s.adres_1,
                            " ".join([
                                _to_str(s.postcode),
                                _to_str(s.gemeente),
                            ]).strip(),
                            s.land,
                        ]
                        if x
                    ]
                )
            print(
                f"{star} {s.supplier}  | Desc: {s.description or '-'} | BTW: {s.btw or '-'} | {addr or '-'} | Mail: {s.sales_email or '-'} | Tel: {s.phone or '-'}"
            )
        return 0
    if args.action == "add":
        rec = {"supplier": args.name}
        if args.description:
            rec["description"] = args.description
        if args.btw:
            vat = validate_vat(args.btw)
            if not vat:
                print("Ongeldig BTW-nummer")
                return 2
            rec["btw"] = vat
        if args.adres_1:
            rec["adres_1"] = args.adres_1
        if args.adres_2:
            rec["adres_2"] = args.adres_2
        if args.email:
            rec["sales_email"] = args.email
        if args.phone:
            rec["phone"] = args.phone
        s = Supplier.from_any(rec)
        db.upsert(s)
        db.save(SUPPLIERS_DB_FILE)
        print("Toegevoegd/bijgewerkt")
        return 0
    if args.action == "remove":
        ok = db.remove(args.name)
        db.save(SUPPLIERS_DB_FILE)
        print("Verwijderd" if ok else "Niet gevonden")
        return 0
    if args.action == "fav":
        ok = db.toggle_fav(args.name)
        db.save(SUPPLIERS_DB_FILE)
        print("Favoriet gewisseld" if ok else "Niet gevonden")
        return 0
    if args.action == "set-default":
        db.set_default(args.production, args.name)
        db.save(SUPPLIERS_DB_FILE)
        print(f"Default voor {args.production}: {args.name}")
        return 0
    if args.action == "get-default":
        print(db.get_default(args.production) or "(geen)")
        return 0
    if args.action == "import-csv":
        path = args.csv
        if not os.path.exists(path):
            print("CSV niet gevonden.")
            return 2
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
                s = Supplier.from_any({k: row[k] for k in df.columns if k in row})
                if args.btw and not s.btw:
                    s.btw = args.btw
                if args.adres_1 and not s.adres_1:
                    s.adres_1 = args.adres_1
                if args.adres_2 and not s.adres_2:
                    s.adres_2 = args.adres_2
                if args.email and not s.sales_email:
                    s.sales_email = args.email
                if args.phone and not s.phone:
                    s.phone = args.phone
                db.upsert(s)
                changed += 1
            except Exception:
                pass
        db.save(SUPPLIERS_DB_FILE)
        print(f"Verwerkt (upsert): {changed}")
        return 0
    if args.action == "clear":
        db.clear_all()
        db.save(SUPPLIERS_DB_FILE)
        print("Alle leveranciers verwijderd.")
        return 0
    print("Onbekende actie")
    return 2


def cli_clients(args):
    db = ClientsDB.load(CLIENTS_DB_FILE)
    if args.action == "list":
        rows = db.clients_sorted()
        if not rows:
            print("(geen opdrachtgevers)")
            return 0
        for c in rows:
            star = "★" if c.favorite else " "
            print(
                f"{star} {c.name} | {c.address or '-'} | BTW: {c.vat or '-'} | Mail: {c.email or '-'}"
            )
        return 0
    if args.action == "add":
        rec = {"name": args.name}
        if args.address:
            rec["address"] = args.address
        if args.vat:
            vat = validate_vat(args.vat)
            if not vat:
                print("Ongeldig BTW-nummer")
                return 2
            rec["vat"] = vat
        if args.email:
            rec["email"] = args.email
        c = Client.from_any(rec)
        db.upsert(c)
        db.save(CLIENTS_DB_FILE)
        print("Toegevoegd/bijgewerkt")
        return 0
    if args.action == "remove":
        ok = db.remove(args.name)
        db.save(CLIENTS_DB_FILE)
        print("Verwijderd" if ok else "Niet gevonden")
        return 0
    if args.action == "fav":
        ok = db.toggle_fav(args.name)
        db.save(CLIENTS_DB_FILE)
        print("Favoriet gewisseld" if ok else "Niet gevonden")
        return 0
    if args.action == "import-csv":
        path = args.csv
        if not os.path.exists(path):
            print("CSV niet gevonden.")
            return 2
        try:
            df = pd.read_csv(path, encoding="latin1", sep=";")
        except Exception:
            df = read_csv_flex(path)
        changed = 0
        for _, row in df.iterrows():
            raw_name = _to_str(row.get("Name")).strip()
            if not raw_name or raw_name == "-":
                continue
            try:
                c = Client.from_any({k: row[k] for k in df.columns if k in row})
                if args.address and not c.address:
                    c.address = args.address
                if args.vat and not c.vat:
                    c.vat = args.vat
                if args.email and not c.email:
                    c.email = args.email
                db.upsert(c)
                changed += 1
            except Exception:
                pass
        db.save(CLIENTS_DB_FILE)
        print(f"Verwerkt (upsert): {changed}")
        return 0
    if args.action == "export-csv":
        rows = [asdict(c) for c in db.clients]
        pd.DataFrame(rows).to_csv(args.csv, index=False, sep=";", encoding="utf-8")
        print(f"Geëxporteerd: {len(rows)}")
        return 0
    print("Onbekende actie")
    return 2


def cli_delivery_addresses(args):
    db = DeliveryAddressesDB.load(DELIVERY_DB_FILE)
    if args.action == "rename":
        addr = db.get(args.old_name)
        if not addr:
            print("Niet gevonden")
            return 2
        new_addr = DeliveryAddress(
            name=args.new_name,
            address=addr.address,
            remarks=addr.remarks,
            favorite=addr.favorite,
        )
        db.upsert(new_addr, old_name=args.old_name)
        db.save(DELIVERY_DB_FILE)
        print("Hernoemd")
        return 0
    print("Onbekende actie")
    return 2


def cli_bom_check(args):
    exts = parse_exts(args.exts, args.allowed_exts)
    df = load_bom(args.bom)
    if not os.path.isdir(args.source):
        print("Bronmap ongeldig")
        return 2
    file_index = _build_file_index(args.source, exts)
    found, status = [], []
    groups = []
    exts_set = set(e.lower() for e in exts)
    if ".step" in exts_set or ".stp" in exts_set:
        groups.append({".step", ".stp"})
        exts_set -= {".step", ".stp"}
    for e in exts_set:
        groups.append({e})
    for _, row in df.iterrows():
        pn = row["PartNumber"]
        hits = file_index.get(pn, [])
        hit_exts = {os.path.splitext(h)[1].lower() for h in hits}
        all_present = all(any(ext in hit_exts for ext in g) for g in groups)
        found.append(", ".join(sorted(e.lstrip('.') for e in hit_exts)))
        status.append("✅ Gevonden" if all_present else "❌ Ontbrekend")
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
    exts = parse_exts(args.exts, args.allowed_exts)
    if not os.path.isdir(args.source) or not os.path.isdir(args.dest):
        print("Bron of bestemming ongeldig")
        return 2
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
    exts = parse_exts(args.exts, args.allowed_exts)
    df = load_bom(args.bom)
    db = SuppliersDB.load(SUPPLIERS_DB_FILE)
    override_map = dict(kv.split("=", 1) for kv in (args.supplier or []))
    cdb = ClientsDB.load(CLIENTS_DB_FILE)
    ddb = DeliveryAddressesDB.load(DELIVERY_DB_FILE)
    client = None
    if args.client:
        client = cdb.get(args.client)
        if not client:
            print("Client niet gevonden")
            return 2
    else:
        cl = cdb.clients_sorted()
        client = cl[0] if cl else None
    delivery_map = {}
    if args.delivery:
        for kv in args.delivery:
            if "=" not in kv:
                print("Leveradres optie moet PROD=NAAM zijn")
                return 2
            prod, name = kv.split("=", 1)
            prod = prod.strip()
            name = name.strip()
            lname = name.lower()
            if lname == "none":
                delivery_map[prod] = None
            elif lname == "pickup":
                delivery_map[prod] = DeliveryAddress(
                    name="Bestelling wordt opgehaald"
                )
            elif lname == "tbd":
                delivery_map[prod] = DeliveryAddress(
                    name="Leveradres wordt nog meegedeeld"
                )
            else:
                addr = ddb.get(name)
                if not addr:
                    print("Leveradres niet gevonden")
                    return 2
                delivery_map[prod] = addr
    doc_type_map = {}
    if args.doc_type:
        for kv in args.doc_type:
            if "=" not in kv:
                print("Documenttype optie moet PROD=TYPE zijn")
                return 2
            prod, dtyp = kv.split("=", 1)
            doc_type_map[prod.strip()] = dtyp.strip()
    doc_num_map = {}
    if args.doc_number:
        for kv in args.doc_number:
            if "=" not in kv:
                print("Documentnummer optie moet PROD=NUM zijn")
                return 2
            prod, num = kv.split("=", 1)
            doc_num_map[prod.strip()] = num.strip()
    cnt, chosen = copy_per_production_and_orders(
        args.source,
        args.dest,
        df,
        exts,
        db,
        override_map,
        doc_type_map=doc_type_map,
        doc_num_map=doc_num_map,
        remember_defaults=args.remember_defaults,
        client=client,
        delivery_map=delivery_map,
        footer_note=args.note or DEFAULT_FOOTER_NOTE,
        project_number=args.project_number,
        project_name=args.project_name,
    )
    print("Gekopieerd:", cnt)
    for k, v in chosen.items():
        print(f"  {k} → {v}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="File Hopper dual-mode")
    p.add_argument(
        "--run-tests", action="store_true", help="Run basic self-tests and exit"
    )
    sub = p.add_subparsers(dest="cmd")

    sp = sub.add_parser("suppliers", help="Beheer leveranciers")
    ssp = sp.add_subparsers(dest="action", required=True)
    ssp.add_parser("list")
    ap = ssp.add_parser("add")
    ap.add_argument("name")
    ap.add_argument("--description", help="Beschrijving van de leverancier")
    ap.add_argument("--btw", help="BTW-nummer")
    ap.add_argument("--adres-1", dest="adres_1", help="Adresregel 1")
    ap.add_argument("--adres-2", dest="adres_2", help="Adresregel 2")
    ap.add_argument("--email", help="E-mail adres voor verkoop")
    ap.add_argument("--tel", dest="phone", help="Telefoonnummer")
    rp = ssp.add_parser("remove")
    rp.add_argument("name")
    fp = ssp.add_parser("fav")
    fp.add_argument("name")
    sdp = ssp.add_parser("set-default")
    sdp.add_argument("production")
    sdp.add_argument("name")
    gdp = ssp.add_parser("get-default")
    gdp.add_argument("production")
    ip = ssp.add_parser("import-csv")
    ip.add_argument("csv", help="CSV-bestand met leveranciers")
    ip.add_argument("--btw", help="Fallback BTW-nummer")
    ip.add_argument("--adres-1", dest="adres_1", help="Fallback adresregel 1")
    ip.add_argument("--adres-2", dest="adres_2", help="Fallback adresregel 2")
    ip.add_argument("--email", help="Fallback e-mail adres voor verkoop")
    ip.add_argument("--tel", dest="phone", help="Fallback telefoonnummer")
    ssp.add_parser("clear")

    cp = sub.add_parser("clients", help="Beheer opdrachtgevers")
    csp = cp.add_subparsers(dest="action", required=True)
    csp.add_parser("list")
    cap = csp.add_parser("add")
    cap.add_argument("name")
    cap.add_argument("--address", help="Adres van de opdrachtgever")
    cap.add_argument("--vat", help="BTW-nummer van de opdrachtgever")
    cap.add_argument("--email", help="E-mail adres van de opdrachtgever")
    crp = csp.add_parser("remove")
    crp.add_argument("name")
    cfp = csp.add_parser("fav")
    cfp.add_argument("name")
    cip = csp.add_parser("import-csv")
    cip.add_argument("csv", help="CSV-bestand met opdrachtgevers")
    cip.add_argument("--address", help="Fallback adres")
    cip.add_argument("--vat", help="Fallback BTW-nummer")
    cip.add_argument("--email", help="Fallback e-mail adres")
    cep = csp.add_parser("export-csv")
    cep.add_argument("csv", help="Bestand om naar te exporteren")

    dp = sub.add_parser("delivery-addresses", help="Beheer leveradressen")
    ddsp = dp.add_subparsers(dest="action", required=True)
    rnp = ddsp.add_parser("rename")
    rnp.add_argument("old_name")
    rnp.add_argument("new_name")

    bp = sub.add_parser("bom", help="BOM acties")
    bsp = bp.add_subparsers(dest="bact", required=True)
    bcp = bsp.add_parser("check")
    bcp.add_argument("--source", required=True)
    bcp.add_argument("--bom", required=True)
    bcp.add_argument("--exts", required=True)
    bcp.add_argument(
        "--allowed-exts",
        default=DEFAULT_ALLOWED_EXTS,
        help="Toegestane extensies (komma gescheiden, wildcards toegestaan)",
    )
    bcp.add_argument("--out")

    cp = sub.add_parser("copy", help="Kopieer vlak")
    cp.add_argument("--source", required=True)
    cp.add_argument("--dest", required=True)
    cp.add_argument("--exts", required=True)
    cp.add_argument(
        "--allowed-exts",
        default=DEFAULT_ALLOWED_EXTS,
        help="Toegestane extensies (komma gescheiden, wildcards toegestaan)",
    )

    cpp = sub.add_parser(
        "copy-per-prod", help="Kopieer per productie + bestelbonnen"
    )
    cpp.add_argument("--source", required=True)
    cpp.add_argument("--dest", required=True)
    cpp.add_argument("--bom", required=True)
    cpp.add_argument("--exts", required=True)
    cpp.add_argument(
        "--allowed-exts",
        default=DEFAULT_ALLOWED_EXTS,
        help="Toegestane extensies (komma gescheiden, wildcards toegestaan)",
    )
    cpp.add_argument(
        "--supplier",
        action="append",
        help="Override: Production=Supplier (meerdere keren mogelijk)",
    )
    cpp.add_argument("--remember-defaults", action="store_true")
    cpp.add_argument(
        "--note", help="Optioneel voetnootje op de bestelbon", default=""
    )
    cpp.add_argument("--client", help="Gebruik opdrachtgever met deze naam")
    cpp.add_argument(
        "--delivery",
        action="append",
        metavar="PROD=NAME",
        help=(
            "Leveradres voor productie: PROD=NAAM. Speciale waarden: "
            "none, pickup, tbd (meerdere keren mogelijk)"
        ),
    )
    cpp.add_argument(
        "--doc-type",
        action="append",
        metavar="PROD=TYPE",
        help="Documenttype per productie (meerdere keren mogelijk)",
    )
    cpp.add_argument(
        "--doc-number",
        action="append",
        metavar="PROD=NUM",
        help="Documentnummer per productie (meerdere keren mogelijk)",
    )
    cpp.add_argument(
        "--project-number",
        dest="project_number",
        help="Projectnummer voor documentkoppen",
    )
    cpp.add_argument(
        "--project-name",
        dest="project_name",
        help="Projectnaam voor documentkoppen",
    )


    return p

