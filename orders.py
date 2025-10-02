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
from typing import Dict, List, Optional, Tuple

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
        Image as RLImage,
        KeepTogether,
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

STEP_EXTS = {".step", ".stp"}


def _normalize_crop_box(
    crop: object, width: int, height: int
) -> Optional[tuple[int, int, int, int]]:
    """Validate and clamp crop data against an image size."""

    if not crop or width <= 0 or height <= 0:
        return None

    left = top = 0
    right, bottom = width, height

    try:
        if isinstance(crop, dict):
            left = int(float(crop.get("left", 0)))
            top = int(float(crop.get("top", 0)))
            right = int(float(crop.get("right", width)))
            bottom = int(float(crop.get("bottom", height)))
        elif isinstance(crop, (list, tuple)) and len(crop) == 4:
            left, top, right, bottom = [int(float(v)) for v in crop]
        else:
            return None
    except Exception:
        return None

    left = max(0, min(width, left))
    top = max(0, min(height, top))
    right = max(left + 1, min(width, right))
    bottom = max(top + 1, min(height, bottom))

    if right <= left or bottom <= top:
        return None
    return left, top, right, bottom


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


def _normalize_finish_folder(value: object) -> str:
    """Return a filesystem-friendly folder component for finish/RAL names."""

    text = _to_str(value).strip()
    if not text:
        text = "_Onbekend"
    text = re.sub(r"[\\/:]+", "-", text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^0-9A-Za-z._ \-]+", "_", text)
    text = text.strip(" .-_")
    if not text:
        text = "_Onbekend"
    return text


def generate_pdf_order_platypus(
    path: str,
    company_info: Dict[str, object],
    supplier: Supplier,
    production: str,
    items: List[Dict[str, str]],
    doc_type: str = "Bestelbon",
    doc_number: str | None = None,
    footer_note: Optional[str] = None,
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

    logo_flowable = None
    logo_path_info = company_info.get("logo_path") if company_info else None
    if logo_path_info:
        logo_path = str(logo_path_info)
        if not os.path.isabs(logo_path):
            logo_path = os.path.join(os.getcwd(), logo_path)
        if os.path.exists(logo_path):
            try:
                from PIL import Image as PILImage  # type: ignore
            except Exception:  # pragma: no cover - Pillow missing during runtime
                PILImage = None  # type: ignore
            if PILImage is not None:
                try:
                    with PILImage.open(logo_path) as src_logo:  # type: ignore[union-attr]
                        logo_img = src_logo.convert("RGBA")
                        crop_box = _normalize_crop_box(
                            company_info.get("logo_crop"),
                            logo_img.width,
                            logo_img.height,
                        )
                        if crop_box:
                            logo_img = logo_img.crop(crop_box)
                        if logo_img.width > 0 and logo_img.height > 0:
                            buffer = io.BytesIO()
                            logo_img.save(buffer, format="PNG")
                            buffer.seek(0)
                            aspect = (
                                logo_img.width / logo_img.height
                                if logo_img.height
                                else 1.0
                            )
                            max_width = 50 * mm
                            max_height = 25 * mm
                            width_pt = max_width
                            height_pt = width_pt / aspect if aspect else max_height
                            if height_pt > max_height:
                                height_pt = max_height
                                width_pt = height_pt * aspect
                            logo_flowable = RLImage(
                                buffer, width=width_pt, height=height_pt
                            )
                            logo_flowable.hAlign = "LEFT"
                except Exception:
                    logo_flowable = None

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
    left_paragraph = Paragraph("<br/>".join(left_lines), text_style)
    left_elements: List[object] = []
    if logo_flowable is not None:
        left_elements.extend([logo_flowable, Spacer(0, 4)])
    left_elements.append(left_paragraph)
    if len(left_elements) == 1:
        left_cell = left_elements[0]
    else:
        left_cell = KeepTogether(left_elements)

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
                left_cell,
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

    if footer_note is None:
        note = DEFAULT_FOOTER_NOTE
    else:
        note = _to_str(footer_note)
    if note:
        story.append(Spacer(0, 8))
        story.append(Paragraph(note, small_style))

    doc.build(story)


def generate_packlist_pdf(
    path: str,
    production: str,
    previews: List[dict],
    doc_date: str | None = None,
    columns: int = 2,
) -> bool:
    """Generate a packing list PDF containing thumbnails."""

    if not REPORTLAB_OK or not previews:
        return False
    columns = max(1, int(columns))
    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm,
        title=f"Paklijst {production}",
    )
    styles = getSampleStyleSheet()
    title_style = styles["Heading1"]
    subtitle_style = styles["Normal"]
    subtitle_style.leading = 12
    label_style = ParagraphStyle(
        "PacklistLabel",
        parent=styles["Normal"],
        alignment=1,
        leading=11,
    )
    story: List[object] = []
    story.append(Paragraph(f"Paklijst – {production}", title_style))
    if doc_date:
        story.append(Paragraph(f"Datum: {doc_date}", subtitle_style))
    story.append(Spacer(0, 8 * mm))

    data: List[List[object]] = []
    row: List[object] = []
    usable_width = doc.width
    col_width = usable_width / columns
    image_width = col_width
    image_height = col_width
    for entry in previews:
        thumb = entry.get("thumbnail")
        label = entry.get("label") or os.path.basename(entry.get("source", ""))
        try:
            img = RLImage(thumb, width=image_width, height=image_height)
        except Exception:
            continue
        cell = KeepTogether([img, Spacer(0, 4), Paragraph(label, label_style)])
        row.append(cell)
        if len(row) == columns:
            data.append(row)
            row = []
    if row:
        while len(row) < columns:
            row.append(Spacer(0, 0))
        data.append(row)

    if not data:
        return False

    table = Table(data, colWidths=[col_width] * columns, hAlign="CENTER")
    table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(table)
    doc.build(story)
    return True


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
    prod: str,
    db: SuppliersDB,
    override_map: Dict[str, str],
    suppliers_sorted: List[Supplier] | None = None,
) -> Supplier:
    """Select a supplier for a given production.

    ``suppliers_sorted`` allows callers to provide a pre-sorted supplier list in
    order to avoid repeated :meth:`SuppliersDB.suppliers_sorted` lookups.
    """
    name = override_map.get(prod)
    sups = suppliers_sorted if suppliers_sorted is not None else db.suppliers_sorted()
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
    footer_note: Optional[str] = None,
    zip_parts: bool = False,
    date_prefix_exports: bool = False,
    date_suffix_exports: bool = False,
    project_number: str | None = None,
    project_name: str | None = None,
    export_name_prefix_text: str = "",
    export_name_prefix_enabled: bool | None = None,
    export_name_suffix_text: str = "",
    export_name_suffix_enabled: bool | None = None,
    copy_finish_exports: bool = False,
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

    When ``copy_finish_exports`` is ``True`` each referenced export file is
    additionally copied to folders named ``Finish-<finish>`` with an optional
    ``-<ral>`` suffix when the BOM "RAL color" property is filled in. The
    folder components are normalized versions of the BOM "Finish" and
    "RAL color" values.
    """
    os.makedirs(dest, exist_ok=True)
    file_index = _build_file_index(source, selected_exts)
    count_copied = 0
    chosen: Dict[str, str] = {}
    doc_type_map = doc_type_map or {}
    doc_num_map = doc_num_map or {}

    prod_to_rows: Dict[str, List[dict]] = defaultdict(list)
    step_entries: Dict[str, List[tuple[str, str]]] = defaultdict(list)
    step_seen: Dict[str, set[str]] = defaultdict(set)
    finish_combo_parts: Dict[tuple[str, str], set[str]] = defaultdict(set)
    for _, row in bom_df.iterrows():
        prod = (row.get("Production") or "").strip() or "_Onbekend"
        prod_to_rows[prod].append(row)
        if copy_finish_exports:
            pn = _to_str(row.get("PartNumber")).strip()
            finish_value = _to_str(row.get("Finish")).strip()
            if pn and finish_value:
                finish_name = _normalize_finish_folder(finish_value)
                ral_value = _to_str(row.get("RAL color")).strip()
                ral_name = (
                    _normalize_finish_folder(ral_value)
                    if ral_value
                    else ""
                )
                finish_combo_parts[(finish_name, ral_name)].add(pn)

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
    suppliers_sorted = db.suppliers_sorted()

    footer_note_text = (
        DEFAULT_FOOTER_NOTE
        if footer_note is None
        else _to_str(footer_note).replace("\r\n", "\n")
    )

    for prod, rows in prod_to_rows.items():
        prod_folder = os.path.join(dest, prod)
        os.makedirs(prod_folder, exist_ok=True)

        raw_doc_type = doc_type_map.get(prod, "Bestelbon")
        doc_type = _to_str(raw_doc_type).strip() or "Bestelbon"
        doc_num = _to_str(doc_num_map.get(prod, "")).strip()
        prefix = _prefix_for_doc_type(doc_type)
        if doc_num and prefix and not doc_num.upper().startswith(prefix.upper()):
            doc_num = f"{prefix}{doc_num}"
        num_part = f"_{doc_num}" if doc_num else ""

        zf = None
        if zip_parts:
            zip_name = f"{prod}{num_part}.zip"
            zip_path = os.path.join(prod_folder, zip_name)
            try:
                zf = zipfile.ZipFile(
                    zip_path,
                    "w",
                    compression=zipfile.ZIP_DEFLATED,
                    compresslevel=6,
                )
            except TypeError:
                # ``compresslevel`` not supported (older Python). Retry without it.
                zf = zipfile.ZipFile(
                    zip_path,
                    "w",
                    compression=zipfile.ZIP_DEFLATED,
                )
            except (RuntimeError, NotImplementedError):
                print(
                    "[WAARSCHUWING] ZIP_DEFLATED niet beschikbaar, val terug op ZIP_STORED",
                    file=sys.stderr,
                )
                zf = zipfile.ZipFile(
                    zip_path,
                    "w",
                    compression=zipfile.ZIP_STORED,
                )

        processed_pairs: set[tuple[str, str]] = set()
        for row in rows:
            pn = str(row["PartNumber"])
            files = file_index.get(pn, [])
            for src_file in files:
                transformed = _transform_export_name(os.path.basename(src_file))
                combo = (src_file, transformed)
                if combo in processed_pairs:
                    continue
                processed_pairs.add(combo)
                ext = os.path.splitext(src_file)[1].lower()
                if ext in STEP_EXTS:
                    seen_paths = step_seen[prod]
                    if src_file not in seen_paths:
                        seen_paths.add(src_file)
                        label = f"{pn} — {os.path.basename(src_file)}"
                        step_entries[prod].append((label, src_file))
                if zip_parts:
                    if zf is not None:
                        zf.write(src_file, arcname=transformed)
                else:
                    dst = os.path.join(prod_folder, transformed)
                    shutil.copy2(src_file, dst)
                count_copied += 1

        if zf is not None:
            zf.close()

        supplier = pick_supplier_for_production(
            prod, db, override_map, suppliers_sorted=suppliers_sorted
        )
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
            "logo_path": client.logo_path if client else "",
            "logo_crop": client.logo_crop if client else None,
        }
        if supplier.supplier:
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
                    footer_note=footer_note_text,
                    delivery=delivery,
                    project_number=project_number,
                    project_name=project_name,
                )
            except Exception as e:
                print(f"[WAARSCHUWING] PDF mislukt voor {prod}: {e}", file=sys.stderr)

        packlist_items = step_entries.get(prod, [])
        if packlist_items and REPORTLAB_OK:
            try:
                with tempfile.TemporaryDirectory(prefix="previews_", dir=prod_folder) as preview_dir:
                    rendered_previews = step_previews.render_step_files(
                        packlist_items, preview_dir
                    )
                    if rendered_previews:
                        packlist_path = os.path.join(
                            prod_folder, f"Paklijst_{prod}_{today}.pdf"
                        )
                        try:
                            if not generate_packlist_pdf(
                                packlist_path,
                                production=prod,
                                previews=rendered_previews,
                                doc_date=today,
                            ) and os.path.exists(packlist_path):
                                os.unlink(packlist_path)
                        except Exception as exc:
                            print(
                                f"[WAARSCHUWING] Paklijst mislukt voor {prod}: {exc}",
                                file=sys.stderr,
                            )
            except Exception as exc:
                print(
                    f"[WAARSCHUWING] Previews genereren mislukt voor {prod}: {exc}",
                    file=sys.stderr,
                )

    if copy_finish_exports and finish_combo_parts:
        finish_seen: Dict[tuple[str, str], set[tuple[str, str]]] = defaultdict(set)
        for (finish_name, ral_name), part_numbers in finish_combo_parts.items():
            folder_name = f"Finish-{finish_name}"
            if ral_name:
                folder_name = f"{folder_name}-{ral_name}"
            target_dir = os.path.join(dest, folder_name)
            os.makedirs(target_dir, exist_ok=True)
            seen_pairs = finish_seen[(finish_name, ral_name)]
            for pn in sorted(part_numbers):
                files = file_index.get(pn, [])
                for src_file in files:
                    transformed = _transform_export_name(os.path.basename(src_file))
                    combo = (src_file, transformed)
                    if combo in seen_pairs:
                        continue
                    seen_pairs.add(combo)
                    shutil.copy2(src_file, os.path.join(target_dir, transformed))

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
            zip_path = None
            prod_prefix = f"{prod}_"
            for fname in sorted(os.listdir(prod_path)):
                if not fname.lower().endswith(".zip"):
                    continue
                stem, _ = os.path.splitext(fname)
                if stem.startswith(prod_prefix):
                    zip_path = os.path.join(prod_path, fname)
                    break
            if zip_path is None:
                fallback = os.path.join(prod_path, f"{prod}.zip")
                if not os.path.isfile(fallback):
                    continue
                zip_path = fallback
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
