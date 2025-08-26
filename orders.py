"""Order-related utilities for creating purchase documents and copying files.

This module groups functions previously in Main_v22.py and depends on helpers,
models, suppliers_db, and bom modules.
"""

import os
import sys
import shutil
import datetime
from collections import defaultdict
from typing import Dict, List, Tuple

import pandas as pd

# ReportLab (PDF). Script works without it (PDF generation is skipped).
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        LongTable,
        TableStyle,
        Spacer,
    )
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False


MIAMI_PINK = "#FF77FF"
DEFAULT_FOOTER_NOTE = (
    "Gelieve afwijkingen schriftelijk te bevestigen. "
    "Levertermijn in overleg. Betalingsvoorwaarden: 30 dagen netto. "
    "Vermeld onze productiereferentie bij levering."
)

from helpers import (
    _to_str,
    _num_to_2dec,
    _pn_wrap_25,
    _material_nowrap,
    _build_file_index,
    _unique_path,
)
from models import Supplier, Client
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from bom import load_bom  # noqa: F401 - imported for module dependency


def generate_pdf_order_platypus(
    path: str,
    company_info: Dict[str, str],
    supplier: Supplier,
    production: str,
    items: List[Dict[str, str]],
    footer_note: str = "",
) -> None:
    """Generate a PDF order using ReportLab if available."""
    if not REPORTLAB_OK:
        return

    margin = 18 * mm
    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
    )
    width, _ = A4
    styles = getSampleStyleSheet()
    title_style = styles["Heading1"]
    title_style.textColor = colors.HexColor(MIAMI_PINK)
    title_style.fontName = "Helvetica-Bold"
    title_style.fontSize = 18
    text_style = styles["Normal"]
    text_style.leading = 13
    small_style = ParagraphStyle("small", parent=text_style, fontSize=8.5, leading=10.5)

    # Address: Adres_1, Adres_2 as one line
    addr_line = None
    if supplier.adres_1 or supplier.adres_2:
        if supplier.adres_1 and supplier.adres_2:
            addr_line = f"{supplier.adres_1}, {supplier.adres_2}"
        else:
            addr_line = supplier.adres_1 or supplier.adres_2

    company_lines = [
        f"<b>{company_info.get('name','')}</b>",
        f"{company_info.get('address','')}",
        f"BTW: {company_info.get('vat','')}",
        f"E-mail: {company_info.get('email','')}",
    ]

    supp_lines = [f"<b>Besteld bij:</b> {supplier.supplier}"]
    if addr_line:
        supp_lines.append(addr_line)
    if not addr_line:
        pc_gem = " ".join(x for x in [supplier.postcode, supplier.gemeente] if x)
        if pc_gem:
            supp_lines.append(pc_gem)
        if supplier.land:
            supp_lines.append(supplier.land)

    supp_lines.append(f"BTW: {supplier.btw or ''}")
    if supplier.contact_sales:
        supp_lines.append(f"Contact sales: {supplier.contact_sales}")
    if supplier.sales_email:
        supp_lines.append(f"Sales e-mail: {supplier.sales_email}")
    if supplier.phone:
        supp_lines.append(f"Tel: {supplier.phone}")

    story = []
    story.append(Paragraph(f"Bestelbon productie: {production}", title_style))
    story.append(Spacer(0, 6))
    story.append(Paragraph("<br/>".join(company_lines), text_style))
    story.append(Spacer(0, 6))
    story.append(Paragraph("<br/>".join(supp_lines), text_style))
    story.append(Spacer(0, 10))

    # Headers and data
    head = ["PartNumber", "Omschrijving", "Materiaal", "St.", "m²", "kg"]

    def wrap_cell_html(val: str, small=False, align=None):
        style = ParagraphStyle(
            "cellsmall" if small else "cell",
            fontName="Helvetica",
            fontSize=8.5 if small else 9,
            leading=10.5 if small else 11,
            wordWrap="CJK",
        )
        if align:
            style.alignment = {"LEFT": 0, "CENTER": 1, "RIGHT": 2}.get(align.upper(), 0)
        return Paragraph(str(val if (val is not None) else ""), style)

    data = [head]
    for it in items:
        pn = _pn_wrap_25(it.get("PartNumber", ""))
        desc = _to_str(it.get("Description", ""))
        mat = _material_nowrap(it.get("Materiaal", ""))
        qty = int(
            pd.to_numeric(_to_str(it.get("Aantal", "")).strip() or 1, errors="coerce") or 1
        )
        qty = max(1, min(999, qty))
        opp = _num_to_2dec(it.get("Oppervlakte", ""))
        gew = _num_to_2dec(it.get("Gewicht", ""))
        data.append(
            [
                wrap_cell_html(pn, small=False, align="LEFT"),
                wrap_cell_html(desc, small=False, align="LEFT"),
                wrap_cell_html(mat, small=True, align="LEFT"),
                wrap_cell_html(qty, small=True, align="RIGHT"),
                wrap_cell_html(opp, small=True, align="RIGHT"),
                wrap_cell_html(gew, small=True, align="RIGHT"),
            ]
        )

    usable_w = width - 2 * margin
    col_fracs = [0.22, 0.40, 0.14, 0.06, 0.09, 0.09]
    col_widths = [usable_w * f for f in col_fracs]

    tbl = LongTable(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(MIAMI_PINK)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (3, 0), (5, 0), "RIGHT"),
                ("ALIGN", (3, 1), (5, -1), "RIGHT"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ]
        )
    )
    story.append(tbl)

    note = footer_note or DEFAULT_FOOTER_NOTE
    if note:
        story.append(Spacer(0, 8))
        story.append(Paragraph(note, small_style))

    doc.build(story)


def write_order_excel(path: str, items: List[Dict[str, str]]) -> None:
    """Write order information to an Excel file."""
    df = pd.DataFrame(
        items,
        columns=["PartNumber", "Description", "Materiaal", "Aantal", "Oppervlakte", "Gewicht"],
    )
    df.to_excel(path, index=False, engine="openpyxl")


def pick_supplier_for_production(
    prod: str, db: SuppliersDB, override_map: Dict[str, str]
) -> Supplier:
    """Select a supplier for a given production."""
    name = override_map.get(prod) or db.get_default(prod)
    sups = db.suppliers_sorted()
    if name:
        for s in sups:
            if s.supplier.lower() == name.lower():
                return s
    return sups[0] if sups else Supplier(supplier="Onbekend")


def copy_per_production_and_orders(
    source: str,
    dest: str,
    bom_df: pd.DataFrame,
    selected_exts: List[str],
    db: SuppliersDB,
    override_map: Dict[str, str],
    remember_defaults: bool,
    client: Client | None = None,
    footer_note: str = "",
) -> Tuple[int, Dict[str, str]]:
    """Copy files per production and create accompanying order documents."""
    os.makedirs(dest, exist_ok=True)
    file_index = _build_file_index(source, selected_exts)
    count_copied = 0
    chosen: Dict[str, str] = {}

    prod_to_rows: Dict[str, List[dict]] = defaultdict(list)
    for _, row in bom_df.iterrows():
        prod = (row.get("Production") or "").strip() or "_Onbekend"
        prod_to_rows[prod].append(row)

    today = datetime.date.today().strftime("%Y-%m-%d")
    for prod, rows in prod_to_rows.items():
        prod_folder = os.path.join(dest, prod)
        os.makedirs(prod_folder, exist_ok=True)

        for row in rows:
            pn = str(row["PartNumber"])
            for src_file in file_index.get(pn, []):
                dst = _unique_path(os.path.join(prod_folder, os.path.basename(src_file)))
                shutil.copy2(src_file, dst)
                count_copied += 1

        supplier = pick_supplier_for_production(prod, db, override_map)
        chosen[prod] = supplier.supplier
        if remember_defaults and supplier.supplier != "Onbekend":
            db.set_default(prod, supplier.supplier)

        items = []
        for row in rows:
            items.append(
                {
                    "PartNumber": row.get("PartNumber", ""),
                    "Description": row.get("Description", ""),
                    "Materiaal": row.get("Materiaal", ""),
                    "Aantal": int(
                        pd.to_numeric(_to_str(row.get("Aantal", "") or 1), errors="coerce") or 1
                    ),
                    "Oppervlakte": row.get("Oppervlakte", ""),
                    "Gewicht": row.get("Gewicht", ""),
                }
            )

        excel_path = os.path.join(prod_folder, f"Bestelbon_{prod}_{today}.xlsx")
        write_order_excel(excel_path, items)

        pdf_path = os.path.join(prod_folder, f"Bestelbon_{prod}_{today}.pdf")
        company = {
            "name": client.name if client else "",
            "address": client.address if client else "",
            "vat": client.vat if client else "",
            "email": client.email if client else "",
        }
        try:
            generate_pdf_order_platypus(
                pdf_path,
                company,
                supplier,
                prod,
                items,
                footer_note=footer_note or DEFAULT_FOOTER_NOTE,
            )
        except Exception as e:
            print(f"[WAARSCHUWING] PDF mislukt voor {prod}: {e}", file=sys.stderr)

    # Persist any (possibly unchanged) supplier defaults so that callers can rely on
    # the database reflecting the latest state on disk.
    db.save(SUPPLIERS_DB_FILE)

    return count_copied, chosen
