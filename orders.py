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
try:
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover - optional dependency
    Alignment = None
    get_column_letter = None

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
        Table,
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
from models import Supplier, Client, DeliveryAddress
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


def _prefix_for_doc_type(doc_type: str) -> str:
    """Return standard document number prefix for a ``doc_type``.

    ``"Bestelbon"`` uses ``"BB-"`` while ``"Offerteaanvraag"`` uses ``"OFF-"``.
    Unknown types return an empty prefix.
    """
    t = (doc_type or "").strip().lower()
    if t.startswith("bestel"):
        return "BB-"
    if t.startswith("offerte"):
        return "OFF-"
    return ""


def generate_pdf_order_platypus(
    path: str,
    company_info: Dict[str, str],
    supplier: Supplier,
    production: str,
    items: List[Dict[str, str]],
    doc_type: str = "Bestelbon",
    doc_number: str | None = None,
    footer_note: str = "",
    delivery: DeliveryAddress | None = None,
    project_number: str | None = None,
    project_name: str | None = None,
) -> None:
    """Generate a PDF order using ReportLab if available.

    ``doc_type`` determines the document title, e.g. ``"Bestelbon"`` or
    ``"Offerteaanvraag"``.
    """
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

    doc_lines: List[str] = []
    if doc_number:
        doc_lines.append(f"Nummer: {doc_number}")
    today = datetime.date.today().strftime("%Y-%m-%d")
    doc_lines.append(f"Datum: {today}")
    if project_number:
        doc_lines.append(f"Projectnummer: {project_number}")
    if project_name:
        doc_lines.append(f"Projectnaam: {project_name}")

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
    if supplier.contact_sales:
        supp_lines.append(f"Contact sales: {supplier.contact_sales}")
    if supplier.sales_email:
        supp_lines.append(f"E-mail: {supplier.sales_email}")
    if supplier.phone:
        supp_lines.append(f"Tel: {supplier.phone}")

    left_lines = company_lines + [""] + supp_lines

    right_lines: List[str] = []
    if delivery:
        # Delivery address block with each piece of information on its own line
        right_lines.append("<b>Leveradres:</b>")
        right_lines.append(delivery.name)
        if delivery.address:
            right_lines.extend(delivery.address.splitlines())
        if delivery.remarks:
            right_lines.append(delivery.remarks)

    story = []
    title = f"{doc_type} productie: {production}"
    story.append(Paragraph(title, title_style))
    if doc_lines:
        story.append(Paragraph("<br/>".join(doc_lines), text_style))
    story.append(Spacer(0, 6))
    header_tbl = LongTable(
        [
            [
                Paragraph("<br/>".join(left_lines), text_style),
                Paragraph("<br/>".join(right_lines), text_style),
            ]
        ],
        colWidths=[(width - 2 * margin) / 2, (width - 2 * margin) / 2],
    )
    header_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header_tbl)
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
    delivery: DeliveryAddress | None = None,
    doc_type: str = "Bestelbon",
    doc_number: str | None = None,
    project_number: str | None = None,
    project_name: str | None = None,
) -> None:
    """Write order information to an Excel file with header info."""
    df = pd.DataFrame(
        items,
        columns=["PartNumber", "Description", "Materiaal", "Aantal", "Oppervlakte", "Gewicht"],
    )

    header_lines: List[Tuple[str, str]] = []
    today = datetime.date.today().strftime("%Y-%m-%d")
    if doc_number:
        header_lines.append(("Nummer", str(doc_number)))
    header_lines.append(("Datum", today))
    if project_number:
        header_lines.append(("Projectnummer", project_number))
    if project_name:
        header_lines.append(("Projectnaam", project_name))
    header_lines.append(("", ""))
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
    if delivery:
        header_lines.extend(
            [
                ("Leveradres", ""),
                ("", delivery.name),
                ("Adres", delivery.address or ""),
                ("Opmerking", delivery.remarks or ""),
                ("", ""),
            ]
        )

    startrow = len(header_lines)
    if Alignment is not None and hasattr(pd, "ExcelWriter"):
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, startrow=startrow)
            ws = writer.sheets[list(writer.sheets.keys())[0]]
            for r, (label, value) in enumerate(header_lines, start=1):
                ws.cell(row=r, column=1, value=label)
                ws.cell(row=r, column=2, value=value)

            left_cols = {"PartNumber", "Description"}
            wrap_cols = {"PartNumber", "Description"}
            for col_idx, col_name in enumerate(df.columns, start=1):
                align = Alignment(
                    horizontal="left" if col_name in left_cols else "right",
                    wrap_text=col_name in wrap_cols,
                )
                if col_name == "PartNumber" and get_column_letter is not None:
                    column_letter = get_column_letter(col_idx)
                    ws.column_dimensions[column_letter].width = 25
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
    doc_type_map: Dict[str, str] | None,
    doc_num_map: Dict[str, str] | None,
    remember_defaults: bool,
    client: Client | None = None,
    delivery_map: Dict[str, DeliveryAddress] | None = None,
    footer_note: str = "",
    zip_parts: bool = False,
    date_prefix_exports: bool = False,
    date_suffix_exports: bool = False,
    project_number: str | None = None,
    project_name: str | None = None,
    export_name_prefix_text: str = "",
    export_name_prefix_enabled: bool | None = None,
    export_name_suffix_text: str = "",
    export_name_suffix_enabled: bool | None = None,
) -> Tuple[int, Dict[str, str]]:
    """Copy files per production and create accompanying order documents.

    ``doc_type_map`` may specify per production whether a *Bestelbon* or an
    *Offerteaanvraag* should be generated. Missing entries default to
    ``"Bestelbon"``.

    ``doc_num_map`` provides document numbers per production which are used in
    filenames and document headers.

    ``delivery_map`` can provide a :class:`DeliveryAddress` per production.
    
    If ``zip_parts`` is ``True``, all export files for a production are
    collected into a single ``<production>.zip`` archive instead of individual
    ``PartNumber`` files. Only the generated order Excel/PDF remain unzipped in
    the production folder.

    When ``date_prefix_exports`` is ``True`` the copied export files will start
    with ``YYYYMMDD-``. When ``date_suffix_exports`` is ``True`` they will end
    with ``-YYYYMMDD`` before the extension. Both transformations are applied
    consistently to copied files and ZIP archive members.

    When custom export prefix/suffix tokens are provided and enabled they are
    added before and/or after the filename. Both transformations are applied
    consistently to copied files and ZIP archive members. The enable flags
    default to active when the corresponding text is non-empty unless
    explicitly set to ``False``.
    """
    os.makedirs(dest, exist_ok=True)
    file_index = _build_file_index(source, selected_exts)
    count_copied = 0
    chosen: Dict[str, str] = {}
    doc_type_map = doc_type_map or {}
    doc_num_map = doc_num_map or {}

    prod_to_rows: Dict[str, List[dict]] = defaultdict(list)
    for _, row in bom_df.iterrows():
        prod = (row.get("Production") or "").strip() or "_Onbekend"
        prod_to_rows[prod].append(row)

    today_date = datetime.date.today()
    today = today_date.strftime("%Y-%m-%d")
    date_token = today_date.strftime("%Y%m%d")
    delivery_map = delivery_map or {}
    export_name_prefix_text = (export_name_prefix_text or "").strip()
    export_name_suffix_text = (export_name_suffix_text or "").strip()
    prefix_has_text = bool(export_name_prefix_text)
    suffix_has_text = bool(export_name_suffix_text)
    if export_name_prefix_enabled is None:
        token_prefix_active = prefix_has_text
    else:
        token_prefix_active = bool(export_name_prefix_enabled) and prefix_has_text
    if export_name_suffix_enabled is None:
        token_suffix_active = suffix_has_text
    else:
        token_suffix_active = bool(export_name_suffix_enabled) and suffix_has_text

    def _transform_export_name(filename: str) -> str:
        """Apply date/custom tokens to ``filename`` when requested."""

        if not (
            date_prefix_exports
            or date_suffix_exports
            or token_prefix_active
            or token_suffix_active
        ):
            return filename
        stem, ext = os.path.splitext(filename)
        prefix_parts: List[str] = []
        if date_prefix_exports:
            prefix_parts.append(date_token)
        if token_prefix_active:
            prefix_parts.append(export_name_prefix_text)
        suffix_parts: List[str] = []
        if date_suffix_exports:
            suffix_parts.append(date_token)
        if token_suffix_active:
            suffix_parts.append(export_name_suffix_text)
        new_stem = "-".join(prefix_parts + [stem] + suffix_parts)
        return f"{new_stem}{ext}"
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
            for src_file in files:
                transformed = _transform_export_name(os.path.basename(src_file))
                if zip_parts:
                    if zf is not None:
                        zf.write(src_file, arcname=transformed)
                        count_copied += 1
                else:
                    dst = os.path.join(prod_folder, transformed)
                    shutil.copy2(src_file, dst)
                    count_copied += 1

        if zf is not None:
            zf.close()

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
            doc_type = doc_type_map.get(prod, "Bestelbon")
            doc_num = _to_str(doc_num_map.get(prod, "")).strip()
            prefix = _prefix_for_doc_type(doc_type)
            if doc_num:
                if prefix and not doc_num.upper().startswith(prefix.upper()):
                    doc_num = f"{prefix}{doc_num}"
            num_part = f"_{doc_num}" if doc_num else ""
            excel_path = os.path.join(
                prod_folder, f"{doc_type}{num_part}_{prod}_{today}.xlsx"
            )
            delivery = delivery_map.get(prod)
            write_order_excel(
                excel_path,
                items,
                company,
                supplier,
                delivery,
                doc_type,
                doc_num or None,
                project_number=project_number,
                project_name=project_name,
            )

            pdf_path = os.path.join(
                prod_folder, f"{doc_type}{num_part}_{prod}_{today}.pdf"
            )
            try:
                generate_pdf_order_platypus(
                    pdf_path,
                    company,
                    supplier,
                    prod,
                    items,
                    doc_type=doc_type,
                    doc_number=doc_num or None,
                    footer_note=footer_note or DEFAULT_FOOTER_NOTE,
                    delivery=delivery,
                    project_number=project_number,
                    project_name=project_name,
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
            if f.lower().endswith(".pdf") and not f.startswith(("Bestelbon_", "Offerteaanvraag_"))
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
                    and not os.path.basename(name).startswith(("Bestelbon_", "Offerteaanvraag_"))
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
