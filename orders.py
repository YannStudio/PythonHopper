"""Order-related utilities for creating purchase documents and copying files.

This module groups functions previously in Main_v22.py and depends on helpers,
models, suppliers_db, and bom modules.
"""

import os
import sys
import shutil
import datetime
import re
import zipfile
import io
from collections import defaultdict
from typing import Dict, List, Tuple

import pandas as pd
from openpyxl.styles import Alignment

try:
    from PyPDF2 import PdfMerger
except Exception:  # pragma: no cover - PyPDF2 might be absent
    PdfMerger = None

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
)
from models import Supplier, Client
from suppliers_db import SuppliersDB, SUPPLIERS_DB_FILE
from bom import load_bom  # noqa: F401 - imported for module dependency
from delivery_addresses_db import DeliveryAddressesDB, DELIVERY_ADDRESSES_DB_FILE


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

    doc_title = "Offerte" if doc_type == "offerte" else "Bestelbon"
    supp_label = "Offerte bij:" if doc_type == "offerte" else "Besteld bij:"
    supp_lines = [f"<b>{supp_label}</b> {supplier.supplier}"]
    if full_addr:
        supp_lines.append(full_addr)
    supp_lines.append(f"BTW: {supplier.btw or ''}")
    if supplier.contact_sales:
        supp_lines.append(f"Contact sales: {supplier.contact_sales}")
    if supplier.sales_email:
        supp_lines.append(f"E-mail: {supplier.sales_email}")
    if supplier.phone:
        supp_lines.append(f"Tel: {supplier.phone}")
    if delivery_address:
        supp_lines.append(f"Leveradres: {delivery_address}")

    story = []
    story.append(Paragraph(f"{doc_title} productie: {production}", title_style))
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
        qty = it.get("Aantal", "")
        opp = _num_to_2dec(it.get("Oppervlakte", ""))
        gew = _num_to_2dec(it.get("Gewicht", ""))
        data.append(
            [
                wrap_cell_html(pn, small=False, align="LEFT"),
                wrap_cell_html(desc, small=False, align="LEFT"),
                wrap_cell_html(mat, small=True, align="RIGHT"),
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
        header_width = stringWidth("Materiaal", "Helvetica-Bold", 10) + 6
        value_width = (
            max(
                stringWidth(
                    _material_nowrap(it.get("Materiaal", "")), "Helvetica", 9
                )
                for it in items
            )
            + 6
        )
        max_mat = max(header_width, value_width)
        if max_mat < mat_w:
            desc_w += mat_w - max_mat
            mat_w = max_mat
        elif max_mat > mat_w:
            desc_w -= max_mat - mat_w
            mat_w = max_mat
        min_desc_w = 40 * mm
        if desc_w < min_desc_w:
            diff = min_desc_w - desc_w
            desc_w = min_desc_w
            mat_w = max(0, mat_w - diff)
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
                ("ALIGN", (2, 0), (5, 0), "RIGHT"),
                ("ALIGN", (2, 1), (5, -1), "RIGHT"),
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
    delivery_address: str = "",
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
            ]
        )
    if delivery_address:
        header_lines.append(("Leveradres", delivery_address))
    if supplier or delivery_address:
        header_lines.append(("", ""))

    startrow = len(header_lines)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=startrow)
        ws = writer.sheets[list(writer.sheets.keys())[0]]
        for r, (label, value) in enumerate(header_lines, start=1):
            ws.cell(row=r, column=1, value=label)
            ws.cell(row=r, column=2, value=value)

        left_cols = {"PartNumber", "Description"}
        for col_idx, col_name in enumerate(df.columns, start=1):
            align = Alignment(horizontal="left" if col_name in left_cols else "right")
            for row in range(startrow + 1, startrow + len(df) + 2):
                ws.cell(row=row, column=col_idx).alignment = align


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
    if prod.strip().lower() in {"dummy part", "nan", "spare part"}:
        return Supplier(supplier="")
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
    delivery_map: Dict[str, str] | None = None,
    client: Client | None = None,
    footer_note: str = "",
    zip_parts: bool = False,
    doc_type: str = "bestelbon",
) -> Tuple[int, Dict[str, str]]:
    """Copy files per production and create accompanying order documents.

    If ``zip_parts`` is ``True``, all export files for a production are
    collected into a single ``<production>.zip`` archive instead of individual
    ``PartNumber`` files. Only the generated order Excel/PDF remain unzipped in
    the production folder.
    """
    os.makedirs(dest, exist_ok=True)
    file_index = _build_file_index(source, selected_exts)
    count_copied = 0
    chosen: Dict[str, str] = {}
    addr_db = DeliveryAddressesDB.load(DELIVERY_ADDRESSES_DB_FILE)

    prod_to_rows: Dict[str, List[dict]] = defaultdict(list)
    for _, row in bom_df.iterrows():
        prod = (row.get("Production") or "").strip() or "_Onbekend"
        prod_to_rows[prod].append(row)

    today = datetime.date.today().strftime("%Y-%m-%d")
    for prod, rows in prod_to_rows.items():
        prod_folder = os.path.join(dest, prod)
        os.makedirs(prod_folder, exist_ok=True)
        zf = None
        if zip_parts:
            zip_path = os.path.join(prod_folder, f"{prod}.zip")
            zf = zipfile.ZipFile(zip_path, "w")

        for row in rows:
            pn = str(row["PartNumber"])
            files = file_index.get(pn, [])
            if zip_parts:
                for src_file in files:
                    if zf is not None:
                        zf.write(src_file, arcname=os.path.basename(src_file))
                        count_copied += 1
            else:
                for src_file in files:
                    dst = os.path.join(prod_folder, os.path.basename(src_file))
                    shutil.copy2(src_file, dst)
                    count_copied += 1

        if zf is not None:
            zf.close()

        supplier = pick_supplier_for_production(prod, db, override_map)
        chosen[prod] = supplier.supplier
        if remember_defaults and supplier.supplier not in ("", "Onbekend"):
            db.set_default(prod, supplier.supplier)

        choice = (delivery_map or {}).get(prod, "").strip()
        lc = choice.lower()
        if lc in ("", "geen", "(geen)"):
            delivery_address = ""
        elif lc == "zelfde als klantadres":
            delivery_address = client.address if client else ""
        elif lc == "klant haalt zelf op":
            delivery_address = "Klant haalt zelf op"
        else:
            rec = addr_db.get(choice)
            delivery_address = rec.address if rec else choice

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


            pdf_path = os.path.join(prod_folder, f"{doc_prefix}_{prod}_{today}.pdf")
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


def combine_pdfs_from_source(
    source: str,
    bom_df: pd.DataFrame,
    dest: str,
    date_str: str | None = None,
) -> int:
    """Combine PDF drawing files per production directly from ``source``.

    The BOM dataframe provides ``PartNumber`` to ``Production`` mappings.
    PDFs matching the part numbers are searched in ``source`` using
    :func:`_build_file_index` and merged per production. The resulting files
    are written to ``dest/Combined pdf``. Output filenames contain the
    production name and current date. Returns the number of combined PDF
    files created.
    """
    if PdfMerger is None:
        raise ModuleNotFoundError(
            "PyPDF2 must be installed to combine PDF files"
        )

    date_str = date_str or datetime.date.today().strftime("%Y-%m-%d")
    idx = _build_file_index(source, [".pdf"])

    prod_to_files: Dict[str, List[str]] = defaultdict(list)
    for _, row in bom_df.iterrows():
        prod = (row.get("Production") or "").strip() or "_Onbekend"
        pn = str(row.get("PartNumber", ""))
        prod_to_files[prod].extend(idx.get(pn, []))

    out_dir = os.path.join(dest, "Combined pdf")
    os.makedirs(out_dir, exist_ok=True)
    count = 0
    for prod, files in prod_to_files.items():
        if not files:
            continue
        merger = PdfMerger()
        for path in sorted(files, key=lambda x: os.path.basename(x).lower()):
            merger.append(path)
        out_name = f"{prod}_{date_str}_combined.pdf"
        merger.write(os.path.join(out_dir, out_name))
        merger.close()
        count += 1
    return count


def combine_pdfs_per_production(dest: str, date_str: str | None = None) -> int:
    """Combine PDF drawing files per production folder into single PDFs.

    The resulting files are written to a subdirectory ``Combined pdf`` inside
    ``dest``. Output filenames contain the production name and current date.
    Returns the number of combined PDF files created.
    """
    if PdfMerger is None:
        raise ModuleNotFoundError(
            "PyPDF2 must be installed to combine PDF files"
        )

    date_str = date_str or datetime.date.today().strftime("%Y-%m-%d")
    out_dir = os.path.join(dest, "Combined pdf")
    os.makedirs(out_dir, exist_ok=True)
    count = 0
    for prod in sorted(os.listdir(dest)):
        prod_path = os.path.join(dest, prod)
        if not os.path.isdir(prod_path):
            continue
        pdfs = [
            f
            for f in os.listdir(prod_path)
            if f.lower().endswith(".pdf") and not f.startswith(("Bestelbon_", "Offerte_"))
        ]
        if pdfs:
            merger = PdfMerger()
            pdfs.sort(key=lambda x: x.lower())
            for fname in pdfs:
                merger.append(os.path.join(prod_path, fname))
        else:
            zip_path = os.path.join(prod_path, f"{prod}.zip")
            if not os.path.isfile(zip_path):
                continue
            with zipfile.ZipFile(zip_path) as zf:
                zip_pdfs = [
                    name
                    for name in zf.namelist()
                    if name.lower().endswith(".pdf")
                    and not os.path.basename(name).startswith(("Bestelbon_", "Offerte_"))
                ]
                if not zip_pdfs:
                    continue
                merger = PdfMerger()
                for name in sorted(
                    zip_pdfs, key=lambda x: os.path.basename(x).lower()
                ):
                    with zf.open(name) as fh:
                        merger.append(io.BytesIO(fh.read()))
        out_name = f"{prod}_{date_str}_combined.pdf"
        merger.write(os.path.join(out_dir, out_name))
        merger.close()
        count += 1
    return count
