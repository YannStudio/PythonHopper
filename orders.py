"""Order-related utilities for creating purchase documents and copying files.

This module groups functions previously in Main_v22.py and depends on helpers,
models, suppliers_db, and bom modules.
"""

import os
import sys
import shutil
import datetime
import re
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
    from reportlab.pdfbase.pdfmetrics import stringWidth
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


def _parse_qty(val: object) -> int:
    """Parse quantity values to int within [1, 999]."""
    s = _to_str(val).strip()
    if not s:
        return 1
    s = s.replace(",", ".")
    s = re.sub(r"[^0-9.]+", "", s)
    try:
        q = int(float(s))
    except Exception:
        q = 1
    return max(1, min(999, q))


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

    company_lines = [
        f"<b>{company_info.get('name','')}</b>",
        f"{company_info.get('address','')}",
        f"BTW: {company_info.get('vat','')}",
        f"E-mail: {company_info.get('email','')}",
    ]

    # Supplier info with full address and contact details
    addr_parts = []
    if supplier.adres_1:
        addr_parts.append(supplier.adres_1)
    if supplier.adres_2:
        addr_parts.append(supplier.adres_2)
    pc_gem = " ".join(x for x in [supplier.postcode, supplier.gemeente] if x)
    if pc_gem:
        addr_parts.append(pc_gem)
    if supplier.land:
        addr_parts.append(supplier.land)
    full_addr = ", ".join(addr_parts)

    supp_lines = [f"<b>Besteld bij:</b> {supplier.supplier}"]
    if full_addr:
        supp_lines.append(full_addr)
    supp_lines.append(f"BTW: {supplier.btw or ''}")
    if supplier.sales_email:
        supp_lines.append(f"E-mail: {supplier.sales_email}")
    if supplier.phone:
        supp_lines.append(f"Tel: {supplier.phone}")
    if supplier.contact_sales:
        supp_lines.append(f"Contact sales: {supplier.contact_sales}")

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
        qty = _parse_qty(it.get("Aantal", ""))
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
    desc_w = usable_w * col_fracs[1]
    mat_w = usable_w * col_fracs[2]
    try:
        max_desc = (
            max(stringWidth(_to_str(it.get("Description", "")), "Helvetica", 9) for it in items)
            + 6
        )
        if max_desc < desc_w:
            mat_w += desc_w - max_desc
            desc_w = max_desc
    except Exception:
        pass
    col_widths = [
        usable_w * col_fracs[0],
        desc_w,
        mat_w,
        usable_w * col_fracs[3],
        usable_w * col_fracs[4],
        usable_w * col_fracs[5],
    ]

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


def write_order_excel(
    path: str,
    items: List[Dict[str, str]],
    company_info: Dict[str, str] | None = None,
    supplier: Supplier | None = None,
) -> None:
    """Write order information to an Excel file with header info."""
    df = pd.DataFrame(
        items,
        columns=["PartNumber", "Description", "Materiaal", "Aantal", "Oppervlakte", "Gewicht"],
    )

    header_lines: List[Tuple[str, str]] = []
    if company_info:
        header_lines.extend(
            [
                ("Bedrijf", company_info.get("name", "")),
                ("Adres", company_info.get("address", "")),
                ("BTW", company_info.get("vat", "")),
                ("E-mail", company_info.get("email", "")),
                ("", ""),
            ]
        )
    if supplier:
        addr_parts = []
        if supplier.adres_1:
            addr_parts.append(supplier.adres_1)
        if supplier.adres_2:
            addr_parts.append(supplier.adres_2)
        pc_gem = " ".join(x for x in [supplier.postcode, supplier.gemeente] if x)
        if pc_gem:
            addr_parts.append(pc_gem)
        if supplier.land:
            addr_parts.append(supplier.land)
        full_addr = ", ".join(addr_parts)
        header_lines.extend(
            [
                ("Leverancier", supplier.supplier),
                ("Adres", full_addr),
                ("BTW", supplier.btw or ""),
                ("E-mail", supplier.sales_email or ""),
                ("Tel", supplier.phone or ""),
                ("", ""),
            ]
        )

    startrow = len(header_lines)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=startrow)
        ws = writer.sheets[list(writer.sheets.keys())[0]]
        for r, (label, value) in enumerate(header_lines, start=1):
            ws.cell(row=r, column=1, value=label)
            ws.cell(row=r, column=2, value=value)


def pick_supplier_for_production(
    prod: str, db: SuppliersDB, override_map: Dict[str, str]
) -> Supplier:
    """Select a supplier for a given production."""
    name = override_map.get(prod)
    sups = db.suppliers_sorted()
    if name is not None:
        if not name.strip():
            return Supplier(supplier="")
        for s in sups:
            if s.supplier.lower() == name.lower():
                return s
        return Supplier(supplier=name)
    default = db.get_default(prod)
    if default:
        for s in sups:
            if s.supplier.lower() == default.lower():
                return s
    return sups[0] if sups else Supplier(supplier="")


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
        if remember_defaults and supplier.supplier not in ("", "Onbekend"):
            db.set_default(prod, supplier.supplier)

        items = []
        for row in rows:
            items.append(
                {
                    "PartNumber": row.get("PartNumber", ""),
                    "Description": row.get("Description", ""),
                    "Materiaal": row.get("Materiaal", ""),
                    "Aantal": _parse_qty(row.get("Aantal", "")),
                    "Oppervlakte": row.get("Oppervlakte", ""),
                    "Gewicht": row.get("Gewicht", ""),
                }
            )

        company = {
            "name": client.name if client else "",
            "address": client.address if client else "",
            "vat": client.vat if client else "",
            "email": client.email if client else "",
        }
        if supplier.supplier:
            excel_path = os.path.join(prod_folder, f"Bestelbon_{prod}_{today}.xlsx")
            write_order_excel(excel_path, items, company, supplier)

            pdf_path = os.path.join(prod_folder, f"Bestelbon_{prod}_{today}.pdf")
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
