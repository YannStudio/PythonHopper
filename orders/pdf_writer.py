"""PDF generation functions for orders and packing lists."""

import os
import datetime
import io
from html import escape
from typing import Dict, List, Mapping, Optional, Sequence, Tuple


from app_paths import resolve_runtime_path
from helpers import _to_str, _num_to_2dec, _material_nowrap
from models import Supplier, DeliveryAddress
from en1090 import EN1090_NOTE_TEXT

from . import core

_PRICE_UNIT_KEY = "Eenheidsprijs"
_PRICE_TOTAL_KEY = "Totaalprijs"

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


def _pdf_order_column_label(column: Mapping[str, object]) -> str:
    """Return a compact header label for PDF tables."""

    key = _to_str(column.get("key")).strip().lower()
    label = _to_str(column.get("label") or column.get("key") or "").strip()
    label_lower = label.lower()
    label_compact = (
        label_lower.replace("€", "")
        .replace("(euro)", "")
        .replace("\u20ac", "")
        .replace("(", "")
        .replace(")", "")
        .replace(" ", "")
    )

    if key == "oppervlakte" or label_lower in {"oppervlakte", "oppervlakte/st"}:
        return "m\u00b2/st"
    if key == "gewicht" or label_lower in {"gewicht", "gewicht (kg)", "gewicht/st"}:
        return "kg/st"
    if key == "eenheidsprijs" or label_compact in {"eenheidsprijs", "unitprice"}:
        return "Prijs/st. (\u20ac)"
    if key == "totaalprijs" or label_compact in {"totaalprijs", "totalprice"}:
        return "Totaal (\u20ac)"
    return label or _to_str(column.get("key")).strip()


def _order_table_header_min_width(label: object, font_size: float) -> float:
    """Return the minimum width needed to keep header words intact."""

    text = " ".join(_to_str(label).replace("\n", " ").split())
    tokens = [token for token in text.split(" ") if token]
    if not tokens:
        return 0.0
    try:
        token_width = max(
            stringWidth(token, "Helvetica-Bold", font_size) for token in tokens
        )
    except Exception:
        token_width = max(len(token) for token in tokens) * font_size * 0.55
    return token_width + 14.0


def _weighted_widths_with_minimums(
    total_width: float,
    weights: Sequence[float],
    minimums: Sequence[float],
) -> List[float]:
    """Distribute width by weight while honoring column minimums when possible."""

    count = len(weights)
    if count == 0:
        return []
    safe_weights = [weight if weight > 0 else 1.0 for weight in weights]
    safe_minimums = [max(0.0, minimum) for minimum in minimums]
    while len(safe_minimums) < count:
        safe_minimums.append(0.0)
    safe_minimums = safe_minimums[:count]
    minimum_total = sum(safe_minimums)
    if minimum_total >= total_width and minimum_total > 0:
        return [total_width * (minimum / minimum_total) for minimum in safe_minimums]

    fixed: set[int] = set()
    widths = [0.0] * count
    while True:
        fixed_width = sum(safe_minimums[index] for index in fixed)
        flexible = [index for index in range(count) if index not in fixed]
        if not flexible:
            break
        flexible_width = max(0.0, total_width - fixed_width)
        flexible_weight = sum(safe_weights[index] for index in flexible) or len(flexible)
        for index in fixed:
            widths[index] = safe_minimums[index]
        for index in flexible:
            widths[index] = flexible_width * (safe_weights[index] / flexible_weight)
        new_fixed = [
            index
            for index in flexible
            if widths[index] < safe_minimums[index]
        ]
        if not new_fixed:
            break
        fixed.update(new_fixed)
    return widths


def _order_metadata_table(
    rows: Sequence[Tuple[str, object]],
    *,
    table_width: float,
    base_style,
    value_max_width: float | None = None,
    right_gap: float | None = None,
):
    """Return a compact label/value table for order PDF metadata."""

    visible_rows = [
        (_to_str(label).strip(), _to_str(value).strip())
        for label, value in rows
        if _to_str(label).strip() or _to_str(value).strip()
    ]
    if not visible_rows:
        return Paragraph("", base_style)

    font_size = float(getattr(base_style, "fontSize", 10) or 10)
    leading = float(getattr(base_style, "leading", font_size + 2.2) or font_size + 2.2)
    text_color = getattr(base_style, "textColor", colors.HexColor(core.ORDER_TEXT_COLOR))
    try:
        widest_label = max(
            stringWidth(f"{label}: ", "Helvetica-Bold", font_size)
            for label, _value in visible_rows
            if label
        )
    except Exception:
        widest_label = max((len(label) for label, _value in visible_rows), default=10) * font_size * 0.55

    gap = right_gap if right_gap is not None else 11 * mm
    value_cap = value_max_width if value_max_width is not None else 60 * mm
    available_width = max(80.0, table_width - gap)
    metadata_width = min(available_width, widest_label + value_cap)
    metadata_width = max(min(available_width, widest_label + 56.0), metadata_width)

    def value_html(value: str) -> str:
        lines = value.splitlines() or [value]
        return "<br/>".join(escape(line.strip()) for line in lines)

    data = []
    for index, (label, value) in enumerate(visible_rows):
        try:
            indent = stringWidth(f"{label}: ", "Helvetica-Bold", font_size) + 2.0
        except Exception:
            indent = (len(label) + 2) * font_size * 0.55
        row_style = ParagraphStyle(
            f"OrderMetaRow{index}",
            parent=base_style,
            fontSize=font_size,
            leading=leading,
            textColor=text_color,
            leftIndent=indent,
            firstLineIndent=-indent,
            wordWrap="CJK",
        )
        data.append(
            [
                Paragraph(
                    f"<b>{escape(label)}:</b> {value_html(value)}" if label else value_html(value),
                    row_style,
                )
            ]
        )
    table = Table(data, colWidths=[metadata_width], hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0.8),
            ]
        )
    )
    return table


def _normalized_price_summary_label(value: object) -> str:
    return " ".join(_to_str(value).strip().lower().split())


def _price_summary_label_keys(
    column_layout: List[Dict[str, object]],
    *,
    is_raw: bool,
) -> List[str]:
    candidates = ["Profiel" if is_raw else "Description"]
    for column in column_layout:
        key = _to_str(column.get("key")).strip()
        if key and key not in candidates and key not in {_PRICE_UNIT_KEY, _PRICE_TOTAL_KEY}:
            candidates.append(key)
    return candidates


def _split_order_price_summary_rows(
    items: List[Dict[str, object]],
    column_layout: List[Dict[str, object]],
    *,
    is_raw: bool,
) -> tuple[List[Dict[str, object]], object, List[tuple[str, object]]]:
    label_keys = _price_summary_label_keys(column_layout, is_raw=is_raw)
    data_items: List[Dict[str, object]] = []
    subtotal: object = ""
    summary_rows: List[tuple[str, object]] = []

    for item in items:
        label_text = ""
        for key in label_keys:
            text = _to_str(item.get(key)).strip()
            if text:
                label_text = text
                break
        normalized = _normalized_price_summary_label(label_text)
        total = item.get(_PRICE_TOTAL_KEY, "")
        if normalized == "subtotaal excl. btw":
            subtotal = total
        elif normalized in {"btw", "totaal incl. btw"} or normalized.startswith("btw "):
            summary_rows.append((label_text, total))
        else:
            data_items.append(item)

    return data_items, subtotal, summary_rows


def _append_order_price_summary_story(
    story: List[object],
    summary_rows: List[tuple[str, object]],
    *,
    usable_w: float,
    palette: Mapping[str, str],
) -> None:
    if not summary_rows:
        return
    label_style = ParagraphStyle(
        "OrderPriceSummaryLabel",
        fontName="Helvetica",
        fontSize=8.7,
        leading=10.8,
        alignment=0,
    )
    value_style = ParagraphStyle(
        "OrderPriceSummaryValue",
        fontName="Helvetica",
        fontSize=8.7,
        leading=10.8,
        alignment=2,
    )
    data = [
        [
            Paragraph(escape(_to_str(label)), label_style),
            Paragraph(escape(_num_to_2dec(total)), value_style),
        ]
        for label, total in summary_rows
    ]
    table_width = min(usable_w * 0.42, 82 * mm)
    value_width = min(32 * mm, table_width * 0.42)
    label_width = table_width - value_width
    summary_tbl = Table(data, colWidths=[label_width, value_width], hAlign="RIGHT")
    style_cmds = [
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor(core.ORDER_TEXT_COLOR)),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
    ]
    last_row = len(data) - 1
    style_cmds.extend(
        [
            ("FONTNAME", (0, last_row), (-1, last_row), "Helvetica-Bold"),
            (
                "LINEABOVE",
                (0, last_row),
                (-1, last_row),
                0.45,
                colors.HexColor(core.ORDER_TABLE_OUTLINE_COLOR),
            ),
            (
                "BACKGROUND",
                (0, last_row),
                (-1, last_row),
                colors.HexColor(palette["total_fill"]),
            ),
        ]
    )
    summary_tbl.setStyle(TableStyle(style_cmds))
    story.append(Spacer(0, 4))
    story.append(summary_tbl)



def generate_pdf_order_platypus(
    path: str,
    company_info: Dict[str, object],
    supplier: Supplier | None,
    production: str,
    items: List[Dict[str, object]],
    doc_type: str = "Bestelbon",
    doc_number: str | None = None,
    footer_note: Optional[str] = None,
    quote_footer_note: Optional[str] = None,
    delivery: DeliveryAddress | None = None,
    project_number: str | None = None,
    project_name: str | None = None,
    label_kind: str = "productie",
    order_remark: str | None = None,
    total_surface_m2: float | None = None,
    total_weight_kg: float | None = None,
    en1090_required: bool = False,
    en1090_note: Optional[str] = None,
    column_layout: Optional[List[Dict[str, object]]] = None,
) -> None:
    """Generate a PDF order using ReportLab if available.

    ``doc_type`` determines the document title, e.g. ``"Bestelbon"`` or
    ``"Offerteaanvraag"``.
    """
    if not REPORTLAB_OK:
        return

    column_layout = [dict(col) for col in column_layout] if column_layout else []
    custom_layout = bool(column_layout)

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
    usable_w = width - 2 * margin
    left_col_width = usable_w * 0.58
    right_col_width = usable_w - left_col_width
    palette = core._order_palette(company_info)
    styles = getSampleStyleSheet()
    title_style = styles["Heading1"]
    title_style.textColor = colors.HexColor(palette["accent"])
    title_style.fontName = "Helvetica-Bold"
    title_style.fontSize = 20
    title_style.leading = 22
    title_style.spaceAfter = 1
    text_style = styles["Normal"]
    text_style.fontSize = 10
    text_style.leading = 12.2
    text_style.textColor = colors.HexColor(core.ORDER_TEXT_COLOR)
    meta_style = ParagraphStyle("meta", parent=text_style, leading=12.4)
    delivery_style = ParagraphStyle(
        "delivery",
        parent=text_style,
        fontSize=9.2,
        leading=11.2,
        textColor=colors.HexColor(core.ORDER_TEXT_COLOR),
    )
    small_style = ParagraphStyle(
        "small",
        parent=text_style,
        fontSize=8.4,
        leading=10.3,
        textColor=colors.HexColor(core.ORDER_MUTED_TEXT_COLOR),
    )

    doc_type_text = (_to_str(doc_type).strip() or "Bestelbon")
    doc_type_text_lower = doc_type_text.lower()
    doc_type_text_slug = __import__("re").sub(r"[^0-9a-z]+", "", doc_type_text_lower)
    is_standaard_doc = doc_type_text_lower.startswith("standaard")
    order_remark_text = _to_str(order_remark) if order_remark is not None else ""
    order_remark_has_content = bool(order_remark_text.strip())
    place_remark_in_delivery_block = core._should_place_remark_in_delivery_block(
        order_remark_has_content=order_remark_has_content,
        doc_type_text_slug=doc_type_text_slug,
        is_standaard_doc=is_standaard_doc,
        delivery=delivery,
    )

    doc_meta_rows: List[Tuple[str, object]] = []
    if doc_number:
        doc_meta_rows.append(("Nummer", doc_number))
    today = datetime.date.today().strftime("%Y-%m-%d")
    doc_meta_rows.append(("Datum", today))
    label_kind_clean = (_to_str(label_kind) or "productie").strip() or "productie"
    label_title = label_kind_clean[0].upper() + label_kind_clean[1:]
    is_raw_material_order = label_kind_clean.lower().startswith("brutemateriaal")
    if production:
        doc_meta_rows.append((label_title, production))
    if project_number:
        doc_meta_rows.append(("Projectnummer", project_number))
    if project_name:
        doc_meta_rows.append(("Projectnaam", project_name))
    if order_remark_has_content and not place_remark_in_delivery_block:
        doc_meta_rows.append(("Opmerking", order_remark_text))

    company_lines = [
        f"<b>{company_info.get('name','')}</b>",
        f"{company_info.get('address','')}",
        f"BTW: {company_info.get('vat','')}",
        f"E-mail: {company_info.get('email','')}",
    ]

    logo_flowable = None
    logo_path_info = company_info.get("logo_path") if company_info else None
    if logo_path_info:
        logo_path = resolve_runtime_path(str(logo_path_info))
        if logo_path and logo_path.exists():
            try:
                from PIL import Image as PILImage
            except Exception:
                PILImage = None
            if PILImage is not None:
                try:
                    with PILImage.open(logo_path) as src_logo:
                        logo_img = src_logo.convert("RGBA")
                        crop_box = core._normalize_crop_box(
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
                            max_width = 38 * mm
                            max_height = 18 * mm
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

    supp_lines: List[str] = []
    if supplier is not None and not is_standaard_doc:
        full_addr = core.format_supplier_address(supplier)
        supplier_label = (
            "Offerte aangevraagd bij:"
            if doc_type_text_slug.startswith("offerte")
            else "Besteld bij:"
        )

        supp_lines = [f"<b>{supplier_label}</b> {supplier.supplier}"]
        if full_addr:
            supp_lines.append(full_addr)
        supp_lines.append(f"BTW: {supplier.btw or ''}")
        if supplier.contact_sales:
            supp_lines.append(f"Contact sales: {supplier.contact_sales}")
        if supplier.sales_email:
            supp_lines.append(f"E-mail: {supplier.sales_email}")
        if supplier.phone:
            supp_lines.append(f"Tel: {supplier.phone}")

    client_block = Paragraph("<br/>".join(company_lines), text_style)
    doc_block = _order_metadata_table(
        doc_meta_rows,
        table_width=left_col_width,
        base_style=meta_style,
        value_max_width=58 * mm,
    )

    supplier_block_parts: List[object] = []
    if supp_lines:
        supplier_block_parts.append(Paragraph("<br/>".join(supp_lines), text_style))

    delivery_block: object | None = None
    include_delivery_block = not is_standaard_doc and (
        delivery is not None or place_remark_in_delivery_block
    )
    if include_delivery_block:
        delivery_text_parts: List[str] = []
        if delivery:
            if _to_str(delivery.name).strip():
                delivery_text_parts.append(_to_str(delivery.name).strip())
            address_text = ", ".join(
                line.strip()
                for line in _to_str(delivery.address).splitlines()
                if line.strip()
            )
            if address_text:
                delivery_text_parts.append(address_text)
            if _to_str(delivery.remarks).strip():
                delivery_text_parts.append(_to_str(delivery.remarks).strip())
        delivery_rows: List[Tuple[str, object]] = []
        if delivery_text_parts:
            delivery_rows.append(("Leveradres", " | ".join(delivery_text_parts)))
        if place_remark_in_delivery_block:
            remark_lines = order_remark_text.splitlines()
            if not remark_lines:
                remark_lines = [order_remark_text]
            delivery_rows.append(
                ("Opmerking", "\n".join(line for line in remark_lines if line.strip()))
            )
        if delivery_rows:
            delivery_block = _order_metadata_table(
                delivery_rows,
                table_width=min(usable_w, 106 * mm),
                base_style=delivery_style,
                value_max_width=70 * mm,
                right_gap=0,
            )

    left_block_parts: List[object] = [doc_block]
    if supplier_block_parts:
        left_block_parts.append(Spacer(0, 8))
        left_block_parts.extend(supplier_block_parts)
    left_block: object = left_block_parts

    story = []
    title = (
        f"{doc_type_text} {label_kind_clean}: {production}"
        if production
        else f"{doc_type_text}"
    )
    story.append(Paragraph(title, title_style))
    title_rule = Table([[""]], colWidths=[usable_w], rowHeights=[2])
    title_rule.setStyle(
        TableStyle(
            [
                ("LINEBELOW", (0, 0), (-1, -1), 0.55, colors.HexColor(core.ORDER_RULE_COLOR)),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(title_rule)
    story.append(Spacer(0, 12))

    right_block_parts: List[object] = []
    if logo_flowable is not None:
        logo_flowable.hAlign = "LEFT"
        right_block_parts.append(logo_flowable)
        right_block_parts.append(Spacer(0, 6))
    right_block_parts.append(client_block)
    right_block: object = right_block_parts

    header_tbl = LongTable(
        [[left_block, right_block]],
        colWidths=[left_col_width, right_col_width],
    )
    header_tbl.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(header_tbl)
    if delivery_block:
        story.append(Spacer(0, 6))
        story.append(delivery_block)
        story.append(Spacer(0, 12))
    else:
        story.append(Spacer(0, 10))

    # Headers and data
    if custom_layout:
        head = []
        for column in column_layout:
            header = _pdf_order_column_label(column)
            if not header:
                header = column.get("key", "")
            column["label"] = header
            head.append(header)
    elif is_raw_material_order:
        head = ["Profiel", "Materiaal", "Lengte", "St.", "kg"]
    else:
        head = ["PartNumber", "Omschrijving", "Materiaal", "St.", "m\u00b2/st", "kg/st"]

    display_items, price_subtotal, price_summary_rows = _split_order_price_summary_rows(
        items,
        column_layout,
        is_raw=is_raw_material_order,
    )

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

    standard_col_widths: List[float] | None = None
    if not custom_layout and not is_raw_material_order:
        col_fracs = [0.22, 0.40, 0.14, 0.06, 0.09, 0.09]
        non_empty_desc_count = sum(
            1 for it in display_items if core._clean_order_cell_text(it.get("Description", ""))
        )
        if display_items:
            empty_desc_ratio = 1.0 - (non_empty_desc_count / len(display_items))
        else:
            empty_desc_ratio = 0.0
        extra_pn_frac = 0.12 * max(0.0, min(1.0, empty_desc_ratio))
        col_fracs[0] += extra_pn_frac
        col_fracs[1] -= extra_pn_frac

        desc_w = usable_w * col_fracs[1]
        mat_w = usable_w * col_fracs[2]
        try:
            header_width = stringWidth("Materiaal", "Helvetica-Bold", 10) + 6
            material_values = [
                stringWidth(
                    _material_nowrap(core._clean_order_cell_text(it.get("Materiaal", ""))),
                    "Helvetica",
                    9,
                )
                for it in display_items
                if core._clean_order_cell_text(it.get("Materiaal", ""))
            ]
            value_width = (max(material_values) if material_values else 0) + 6
            max_mat = max(header_width, value_width)
            if max_mat < mat_w:
                desc_w += mat_w - max_mat
                mat_w = max_mat
            elif max_mat > mat_w:
                desc_w -= max_mat - mat_w
                mat_w = max_mat
            min_desc_w = 42 * mm
            if desc_w < min_desc_w:
                diff = min_desc_w - desc_w
                desc_w = min_desc_w
                mat_w = max(0, mat_w - diff)
        except Exception:
            pass
        standard_col_widths = [
            usable_w * col_fracs[0],
            desc_w,
            mat_w,
            usable_w * col_fracs[3],
            usable_w * col_fracs[4],
            usable_w * col_fracs[5],
        ]

    def description_cell_html(val: object, width: float) -> str:
        lines = core._wrap_words_to_lines(
            core._clean_order_cell_text(val),
            max(24.0, width),
            "Helvetica",
            9,
            max_lines=2,
        )
        return "<br/>".join(escape(line) for line in lines)

    data = [head]
    total_row_index: int | None = None
    if custom_layout:
        surface_idx: int | None = None
        weight_idx: int | None = None
        total_price_idx: int | None = None
        for idx, column in enumerate(column_layout):
            if surface_idx is None and core.is_order_surface_column(column):
                surface_idx = idx
            if weight_idx is None and core.is_order_weight_column(column):
                weight_idx = idx
            if total_price_idx is None and _to_str(column.get("key")).strip().lower() == _PRICE_TOTAL_KEY.lower():
                total_price_idx = idx
        for it in display_items:
            row_cells: List[object] = []
            for idx, column in enumerate(column_layout):
                key = column.get("key")
                value = it.get(key, "") if key else ""
                if column.get("numeric"):
                    if column.get("integer"):
                        value = core._format_integer_like(value)
                    else:
                        value = _num_to_2dec(value)
                    small = True
                else:
                    value = _to_str(value)
                    small = False
                align = _to_str(column.get("justify") or "left").strip().upper() or "LEFT"
                if align not in {"LEFT", "RIGHT", "CENTER"}:
                    align = "LEFT"
                row_cells.append(wrap_cell_html(value, small=small, align=align))
            data.append(row_cells)

        if (
            (surface_idx is not None and total_surface_m2 is not None)
            or (weight_idx is not None and total_weight_kg is not None)
            or (total_price_idx is not None and _to_str(price_subtotal).strip())
        ):
            total_row: List[object] = []
            for idx, column in enumerate(column_layout):
                align = _to_str(column.get("justify") or "left").strip().upper() or "LEFT"
                if align not in {"LEFT", "RIGHT", "CENTER"}:
                    align = "LEFT"
                if idx == surface_idx and total_surface_m2 is not None:
                    surface_text = _num_to_2dec(total_surface_m2)
                    total_row.append(wrap_cell_html(surface_text, small=True, align=align))
                elif idx == weight_idx and total_weight_kg is not None:
                    weight_text = _num_to_2dec(total_weight_kg)
                    total_row.append(wrap_cell_html(weight_text, small=True, align=align))
                elif idx == total_price_idx and _to_str(price_subtotal).strip():
                    subtotal_text = _num_to_2dec(price_subtotal)
                    total_row.append(wrap_cell_html(subtotal_text, small=True, align=align))
                elif idx == 0:
                    total_row.append(wrap_cell_html("Totaal", small=False, align="LEFT"))
                else:
                    total_row.append(
                        wrap_cell_html("", small=bool(column.get("numeric")), align=align)
                    )
            data.append(total_row)
            total_row_index = len(data) - 1
    elif is_raw_material_order:
        for it in display_items:
            prof = _to_str(it.get("Profiel", ""))
            mat = _to_str(it.get("Materiaal", ""))
            length_val = it.get("Lengte", "")
            length = _to_str("" if length_val in (None, "") else length_val)
            qty_val = it.get("St.", "")
            qty = _to_str("" if qty_val in (None, "") else qty_val)
            weight_val = it.get("kg", "")
            weight = _num_to_2dec(weight_val)
            data.append(
                [
                    wrap_cell_html(prof, small=False, align="LEFT"),
                    wrap_cell_html(mat, small=False, align="LEFT"),
                    wrap_cell_html(length, small=True, align="RIGHT"),
                    wrap_cell_html(qty, small=True, align="RIGHT"),
                    wrap_cell_html(weight, small=True, align="RIGHT"),
                ]
            )
        if total_weight_kg is not None:
            total_row_index = len(data)
            total_row = [
                wrap_cell_html("Totaal", small=False, align="LEFT"),
                wrap_cell_html("", small=False, align="LEFT"),
                wrap_cell_html("", small=True, align="RIGHT"),
                wrap_cell_html("", small=True, align="RIGHT"),
                wrap_cell_html(_num_to_2dec(total_weight_kg), small=True, align="RIGHT"),
            ]
            data.append(total_row)
            total_row_index = len(data) - 1
    else:
        for it in display_items:
            pn = escape(core._clean_order_cell_text(it.get("PartNumber", "")))
            desc_width = (
                (standard_col_widths[1] - 10) if standard_col_widths else (usable_w * 0.40)
            )
            desc = description_cell_html(it.get("Description", ""), desc_width)
            mat = _material_nowrap(core._clean_order_cell_text(it.get("Materiaal", "")))
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
        if total_surface_m2 is not None or total_weight_kg is not None:
            total_row = [
                wrap_cell_html("Totaal", small=False, align="LEFT"),
                wrap_cell_html("", small=False, align="LEFT"),
                wrap_cell_html("", small=True, align="RIGHT"),
                wrap_cell_html("", small=True, align="RIGHT"),
                wrap_cell_html(
                    _num_to_2dec(total_surface_m2)
                    if total_surface_m2 is not None
                    else "",
                    small=True,
                    align="RIGHT",
                ),
                wrap_cell_html(
                    _num_to_2dec(total_weight_kg)
                    if total_weight_kg is not None
                    else "",
                    small=True,
                    align="RIGHT",
                ),
            ]
            data.append(total_row)
            total_row_index = len(data) - 1

    if custom_layout and column_layout:
        weights: List[float] = []
        for column in column_layout:
            try:
                weight_val = float(column.get("weight", 0))
            except Exception:
                weight_val = 0.0
            weights.append(weight_val if weight_val > 0 else 1.0)
        minimums = [
            _order_table_header_min_width(column.get("label"), 10)
            for column in column_layout
        ]
        col_widths = _weighted_widths_with_minimums(usable_w, weights, minimums)
    elif is_raw_material_order:
        col_fracs = [0.32, 0.24, 0.16, 0.12, 0.16]
        col_widths = [usable_w * frac for frac in col_fracs]
    else:
        col_widths = standard_col_widths or [
            usable_w * 0.22,
            usable_w * 0.40,
            usable_w * 0.14,
            usable_w * 0.06,
            usable_w * 0.09,
            usable_w * 0.09,
        ]

    tbl = LongTable(data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(palette["accent"])),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(palette["accent_text"])),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor(core.ORDER_TEXT_COLOR)),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOX", (0, 0), (-1, -1), 0.45, colors.HexColor(core.ORDER_TABLE_OUTLINE_COLOR)),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor(core.ORDER_TABLE_GRID_COLOR)),
        (
            "ROWBACKGROUNDS",
            (0, 1),
            (-1, -1),
            [colors.white, colors.HexColor(core.ORDER_TABLE_ALT_ROW_COLOR)],
        ),
        ("LEFTPADDING", (0, 0), (-1, 0), 6),
        ("RIGHTPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 5),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("LEFTPADDING", (0, 1), (-1, -1), 5),
        ("RIGHTPADDING", (0, 1), (-1, -1), 5),
        ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
    ]
    if custom_layout and column_layout:
        for idx, column in enumerate(column_layout):
            align = _to_str(column.get("justify") or "left").strip().upper() or "LEFT"
            if align not in {"LEFT", "RIGHT", "CENTER"}:
                align = "LEFT"
            style_cmds.append(("ALIGN", (idx, 0), (idx, -1), align))
    elif is_raw_material_order:
        style_cmds.append(("ALIGN", (2, 0), (4, -1), "RIGHT"))
    else:
        style_cmds.extend(
            [
                ("ALIGN", (2, 0), (5, 0), "RIGHT"),
                ("ALIGN", (2, 1), (5, -1), "RIGHT"),
            ]
        )
    if total_row_index is not None:
        style_cmds.extend(
            [
                ("FONTNAME", (0, total_row_index), (-1, total_row_index), "Helvetica-Bold"),
                (
                    "BACKGROUND",
                    (0, total_row_index),
                    (-1, total_row_index),
                    colors.HexColor(palette["total_fill"]),
                ),
                (
                    "LINEABOVE",
                    (0, total_row_index),
                    (-1, total_row_index),
                    0.45,
                    colors.HexColor(core.ORDER_TABLE_OUTLINE_COLOR),
                ),
            ]
        )
    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)
    _append_order_price_summary_story(
        story,
        price_summary_rows,
        usable_w=usable_w,
        palette=palette,
    )

    if en1090_required:
        note_text = EN1090_NOTE_TEXT if en1090_note is None else _to_str(en1090_note)
        if note_text:
            story.append(Spacer(0, 12))
            en1090_note_html = note_text.replace("\n", "<br/>")
            if note_text == EN1090_NOTE_TEXT:
                en1090_note_html = f"<b>{en1090_note_html}</b>"
            story.append(Paragraph(en1090_note_html, small_style))

    # Add a brief explanatory note for raw material (brutemateriaal) orders
    # Only include when a supplier is present and the production/supplier indicates
    # that the supplier will perform cutting/laser work (heuristic).
    if is_raw_material_order and supplier is not None:
        prod_lower = _to_str(production).strip().lower() if production else ""
        supp_prod_type = (_to_str(supplier.product_type).strip().lower() if getattr(supplier, 'product_type', None) else "")

        keywords = ("tube", "laser", "cut", "snij", "snijden", "metal processing")
        def _has_kw(text: str) -> bool:
            return any(k in text for k in keywords if k)

        show_bruto_note = _has_kw(prod_lower) or _has_kw(supp_prod_type)

        if show_bruto_note:
            story.append(Spacer(0, 12))
            prod_part = f" (productie: {escape(_to_str(production))})" if production else ""
            bruto_note = (
                f"Deze bruto materiaalbon is aanvullende productie-informatie voor snijwerk{prod_part}. "
                "De bon geeft aan hoeveel bruto profielmateriaal nodig is om de snedes uit te voeren. "
                "Alleen ter ondersteuning van productie; vervangt geen order of factuur."
            )
            story.append(Paragraph(bruto_note, small_style))

    note = core.footer_note_for_doc_type(doc_type_text, footer_note, quote_footer_note)
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
        leftMargin=15 * __import__("reportlab").lib.units.mm,
        rightMargin=15 * __import__("reportlab").lib.units.mm,
        topMargin=20 * __import__("reportlab").lib.units.mm,
        bottomMargin=20 * __import__("reportlab").lib.units.mm,
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
