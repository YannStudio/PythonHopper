"""Core utilities, constants, and data structures for order generation.

This module provides color palettes, text manipulation, path utilities,
document number formatting, selection keys, opticutter utilities, and more.
"""

import os
import re
import unicodedata
import hashlib
import math
import datetime
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd

from helpers import _to_str
from models import color_to_rgb, normalize_rgb_color, Supplier, DeliveryAddress
from suppliers_db import SuppliersDB
from opticutter import parse_length_to_mm

# Color Constants
MIAMI_PINK = "#FF77FF"
ORDER_RULE_COLOR = "#B9C1CA"
ORDER_TEXT_COLOR = "#1F2329"
ORDER_MUTED_TEXT_COLOR = "#5D6670"
ORDER_TABLE_OUTLINE_COLOR = "#B7BEC8"
ORDER_TABLE_GRID_COLOR = "#D5DAE1"
ORDER_TABLE_ALT_ROW_COLOR = "#FBFCFD"
ORDER_TOTAL_FILL_COLOR = "#FFF4FF"
ORDER_DELIVERY_FILL_COLOR = "#FFF8FD"

DEFAULT_FOOTER_NOTE = (
    "Gelieve afwijkingen schriftelijk te bevestigen. "
    "Levertermijn in overleg. Betalingsvoorwaarden: 30 dagen netto. "
    "Vermeld onze productiereferentie bij levering."
)

STEP_EXTS = {".step", ".stp"}
NO_SUPPLIER_PLACEHOLDER = "(geen)"


def _clean_address_part(value: object) -> str:
    text = _to_str(value).strip()
    return "" if text.lower() == "nan" else text


def _normalize_address_part(value: str) -> str:
    text = unicodedata.normalize("NFKD", _to_str(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def _address_part_already_present(parts: Sequence[str], candidate: str) -> bool:
    normalized_candidate = _normalize_address_part(candidate)
    if not normalized_candidate:
        return True
    for part in parts:
        normalized_part = _normalize_address_part(part)
        if (
            normalized_candidate == normalized_part
            or normalized_candidate in normalized_part
        ):
            return True
    return False


def format_supplier_address(supplier: Supplier | None) -> str:
    """Return a compact supplier address without repeated city/postcode data."""

    if supplier is None:
        return ""
    parts: List[str] = []

    def add_part(value: object) -> None:
        text = _clean_address_part(value)
        if text and not _address_part_already_present(parts, text):
            parts.append(text)

    add_part(supplier.adres_1)
    add_part(supplier.adres_2)
    add_part(" ".join(
        part
        for part in (
            _clean_address_part(supplier.postcode),
            _clean_address_part(supplier.gemeente),
        )
        if part
    ))
    add_part(supplier.land)
    return ", ".join(parts)


# BOM columns
_BOM_STATUS_COLUMNS: Tuple[str, ...] = ("Bestanden gevonden", "Status", "Link")
_BOM_EXPORT_BASE_COLUMNS: Tuple[str, ...] = (
    "PartNumber",
    "Description",
    "QTY.",
    "Profile",
    "Length profile",
    "Production",
    "Materiaal",
    "Supplier",
    "Supplier code",
    "Manufacturer",
    "Manufacturer code",
    "Finish",
    "RAL color",
    "Oppervlakte",
    "Gewicht",
)

# Document filename profiles
DOCUMENT_FILENAME_PROFILE_STANDARD = "standard"
DOCUMENT_FILENAME_PROFILE_SHORT = "short"
DOCUMENT_FILENAME_PROFILE_COMPACT = "compact"
DOCUMENT_FILENAME_PROFILE_CUSTOM = "custom"
DOCUMENT_FILENAME_PROFILES = {
    DOCUMENT_FILENAME_PROFILE_STANDARD,
    DOCUMENT_FILENAME_PROFILE_SHORT,
    DOCUMENT_FILENAME_PROFILE_COMPACT,
    DOCUMENT_FILENAME_PROFILE_CUSTOM,
}
DOCUMENT_FILENAME_SEPARATOR_MAP = {
    "underscore": "_",
    "dash": "-",
    "none": "",
}

# Selection key prefixes
FINISH_KEY_PREFIX = "finish::"
PRODUCTION_KEY_PREFIX = "production::"
OPTICUTTER_KEY_PREFIX = "opticutter::"
OPTICUTTER_DEFAULT_SUFFIX = "::Opticutter"

# Path limits
_INVALID_PATH_CHARS = set('<>:"/\\|?*')
_WINDOWS_MAX_PATH = 240


@dataclass(slots=True)
class CombinedPdfResult:
    """Metadata for combined PDF export operations."""

    count: int
    output_dir: str


@dataclass(slots=True)
class OpticutterProfileStats:
    """Aggregated length/weight data for a single Opticutter profile."""

    total_length_mm: float = 0.0
    total_weight_kg: float = 0.0

    @property
    def weight_per_mm(self) -> float | None:
        if self.total_length_mm <= 0 or self.total_weight_kg <= 0:
            return None
        return self.total_weight_kg / self.total_length_mm


@dataclass(slots=True)
class OpticutterOrderComputation:
    """Computed data for Opticutter raw material exports per production."""

    scenario_rows: List[Dict[str, object]]
    piece_rows: List[Dict[str, object]]
    order_rows: List[Dict[str, object]]
    raw_items: List[Dict[str, object]]
    has_valid_bars: bool
    total_bars: int
    total_weight_kg: float | None
    selection_count: int


# ========== Color Functions ==========

def _mix_color_with_white(color: str, whiteness: float) -> str:
    rgb = color_to_rgb(color)
    if rgb is None:
        return color
    ratio = max(0.0, min(1.0, float(whiteness)))
    mixed = tuple(int(round(channel + (255 - channel) * ratio)) for channel in rgb)
    return "#{:02X}{:02X}{:02X}".format(*mixed)


def _accent_text_color(fill_color: str) -> str:
    rgb = color_to_rgb(fill_color)
    if rgb is None:
        return ORDER_TEXT_COLOR
    luminance = ((0.299 * rgb[0]) + (0.587 * rgb[1]) + (0.114 * rgb[2])) / 255.0
    return ORDER_TEXT_COLOR if luminance >= 0.68 else "#FFFFFF"


def _order_palette(company_info: Mapping[str, object] | None) -> Dict[str, str]:
    accent_color = normalize_rgb_color(
        company_info.get("accent_color") if company_info else None
    ) or MIAMI_PINK
    return {
        "accent": accent_color,
        "accent_text": _accent_text_color(accent_color),
        "total_fill": _mix_color_with_white(accent_color, 0.88),
    }


# ========== Text Cleaning & Wrapping ==========

def _clean_order_cell_text(value: object) -> str:
    text = _to_str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    return " ".join(text.split())


def _fit_text_to_width(
    text: str,
    width: float,
    font_name: str,
    font_size: float,
) -> Tuple[str, str]:
    try:
        from reportlab.pdfbase.pdfmetrics import stringWidth
    except Exception:
        return text, ""

    if not text:
        return "", ""
    if stringWidth(text, font_name, font_size) <= width:
        return text, ""
    lo = 1
    hi = len(text)
    best = 1
    while lo <= hi:
        mid = (lo + hi) // 2
        probe = text[:mid]
        if stringWidth(probe, font_name, font_size) <= width:
            best = mid
            lo = mid + 1
        else:
            hi = mid - 1
    head = text[:best].rstrip()
    tail = text[best:].lstrip()
    if not head:
        head = text[:1]
        tail = text[1:].lstrip()
    return head, tail


def _truncate_text_to_width(
    text: str,
    width: float,
    font_name: str,
    font_size: float,
    suffix: str = "...",
    force_suffix: bool = False,
) -> str:
    try:
        from reportlab.pdfbase.pdfmetrics import stringWidth
    except Exception:
        return text

    text = text.rstrip()
    if not text:
        return ""
    if stringWidth(text, font_name, font_size) <= width and not force_suffix:
        return text
    while text and stringWidth(text + suffix, font_name, font_size) > width:
        text = text[:-1].rstrip()
    return (text + suffix) if text else suffix


def _wrap_words_to_lines(
    text: str,
    width: float,
    font_name: str,
    font_size: float,
    max_lines: int,
) -> List[str]:
    try:
        from reportlab.pdfbase.pdfmetrics import stringWidth
    except Exception:
        clean = _clean_order_cell_text(text)
        return [clean] if clean else []

    clean = _clean_order_cell_text(text)
    if not clean:
        return []
    pending = clean.split()
    lines: List[str] = []
    while pending and len(lines) < max_lines:
        current = pending.pop(0)
        if stringWidth(current, font_name, font_size) > width:
            current, remainder = _fit_text_to_width(
                current,
                width,
                font_name,
                font_size,
            )
            if remainder:
                pending.insert(0, remainder)
        while pending:
            probe = f"{current} {pending[0]}"
            if stringWidth(probe, font_name, font_size) > width:
                break
            current = probe
            pending.pop(0)
        lines.append(current)
    if pending and lines:
        lines[-1] = _truncate_text_to_width(
            lines[-1],
            width,
            font_name,
            font_size,
            force_suffix=True,
        )
    return lines


# ========== Path & File Utilities ==========

def _sanitize_component(value: object) -> str:
    """Return a filesystem-friendly representation of ``value``."""
    text = "" if value is None else str(value).strip()
    if not text:
        return ""
    text = " ".join(text.split())
    cleaned: List[str] = []
    for ch in text:
        if ch in _INVALID_PATH_CHARS or ord(ch) < 32:
            cleaned.append("_")
        elif ch == os.sep or (os.altsep and ch == os.altsep):
            cleaned.append("-")
        else:
            cleaned.append(ch)
    return "".join(cleaned).strip(" .-_")


def _slugify_name(value: object, fallback: str) -> str:
    """Slugify ``value`` similar to export bundle directories."""
    normalized = unicodedata.normalize("NFKD", "" if value is None else str(value))
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = ascii_text.lower()
    ascii_text = ascii_text.replace(" ", "-")
    ascii_text = re.sub(r"[^a-z0-9-]", "", ascii_text)
    ascii_text = re.sub(r"-+", "-", ascii_text).strip("-")
    if len(ascii_text) > 40:
        ascii_text = ascii_text[:40].rstrip("-")
    if ascii_text:
        return ascii_text
    fallback_norm = unicodedata.normalize("NFKD", fallback)
    ascii_fallback = fallback_norm.encode("ascii", "ignore").decode("ascii") or fallback
    ascii_fallback = re.sub(r"[^a-zA-Z0-9-]", "", ascii_fallback)
    ascii_fallback = ascii_fallback.lower()[:40].rstrip("-")
    return ascii_fallback or "export"


def _fit_filename_within_path(
    directory: str,
    filename: str,
    *,
    max_path: int = _WINDOWS_MAX_PATH,
) -> str:
    """Return ``filename`` possibly shortened so ``directory/filename`` fits ``max_path``."""
    if max_path is None:
        return filename

    directory_abs = os.path.abspath(directory)
    full_abs = os.path.join(directory_abs, filename)
    if len(full_abs) <= max_path:
        return filename

    stem, ext = os.path.splitext(filename)
    if not stem:
        stem = "_"

    available = max_path - len(directory_abs) - len(os.sep) - len(ext)
    if available <= 0:
        raise OSError(
            "Basispad is te lang voor het genereren van exportbestanden. Verkort het exportpad."
        )

    digest = hashlib.sha1(full_abs.encode("utf-8")).hexdigest()[:8]
    if available <= len(digest):
        safe_stem = digest[:available]
    else:
        keep = max(1, available - len(digest) - 1)
        safe_stem = f"{stem[:keep].rstrip(' _-.')}_{digest}" if keep < len(stem) else f"{stem}_{digest}"

    safe_filename = f"{safe_stem}{ext}"
    safe_abs = os.path.join(directory_abs, safe_filename)
    if len(safe_abs) > max_path and available > len(digest):
        safe_stem = digest[: available]
        safe_filename = f"{safe_stem}{ext}"

    return safe_filename


def _create_combined_output_dir(
    base_dir: str,
    project_number: Optional[str],
    project_name: Optional[str],
    *,
    timestamp: Optional[datetime.datetime] = None,
) -> str:
    """Create a combined PDF export directory inside ``base_dir``."""
    os.makedirs(base_dir, exist_ok=True)
    ts = timestamp or datetime.datetime.now()
    ts_token = ts.strftime("%Y-%m-%dT%H%M%S")
    pn_clean = _sanitize_component(project_number) or "project"
    slug = _slugify_name(project_name, pn_clean)
    name_parts = ["Combined pdf", ts_token, pn_clean]
    if slug and slug != pn_clean.lower():
        name_parts.append(slug)
    folder_name = "_".join(part for part in name_parts if part)
    base_path = os.path.join(base_dir, folder_name)
    candidate = base_path
    index = 1
    while os.path.exists(candidate):
        candidate = f"{base_path}_{index}"
        index += 1
    os.makedirs(candidate, exist_ok=True)
    return candidate


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


# ========== Quantity & Number Parsing ==========

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


def _coerce_integer_like(value: object) -> object:
    """Return ``value`` as an int when it represents a whole number."""
    if value in ("", None):
        return ""
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(round(value)) if math.isfinite(value) else value
    text = _to_str(value).strip()
    if not text:
        return ""
    text = text.replace(",", ".")
    try:
        number = float(text)
    except Exception:
        return value
    if not math.isfinite(number):
        return value
    return int(round(number))


def _format_integer_like(value: object) -> str:
    """Format integer-like values without decimal places."""
    coerced = _coerce_integer_like(value)
    if coerced in ("", None):
        return ""
    if isinstance(coerced, int):
        return str(coerced)
    return _to_str(coerced)


# ========== Document Number Utilities ==========

def _prefix_for_doc_type(doc_type: str) -> str:
    """Return standard document number prefix for a ``doc_type``."""
    t = (doc_type or "").strip().lower()
    if t.startswith("standaard"):
        return "BOM-"
    if t.startswith("bestel"):
        return "BB-"
    if t.startswith("offerte"):
        return "OFF-"
    return ""


def _normalize_doc_number(value: object, doc_type: object) -> str:
    """Return a cleaned document number for a given ``doc_type``."""
    doc_num = _to_str(value).strip()
    if not doc_num:
        return ""

    prefix = _prefix_for_doc_type(_to_str(doc_type))
    if not prefix:
        return doc_num

    prefix_upper = prefix.upper()
    doc_upper = doc_num.upper()
    prefix_compact = re.sub(r"[^A-Z0-9]", "", prefix_upper)

    if doc_upper.startswith(prefix_upper):
        remainder = doc_num[len(prefix) :]
        remainder_stripped = remainder.lstrip(" -_")
        remainder_upper = remainder_stripped.upper()
        if remainder_upper.startswith(prefix_upper):
            remainder = remainder_stripped[len(prefix) :]
            doc_num = prefix + remainder.lstrip(" -_")
        elif prefix_compact and remainder_upper.startswith(prefix_compact):
            remainder = remainder_stripped[len(prefix_compact) :]
            doc_num = prefix + remainder.lstrip(" -_")
    elif prefix_compact and doc_upper.startswith(prefix_compact):
        remainder = doc_num[len(prefix_compact) :]
        doc_num = prefix + remainder.lstrip(" -_")

    return doc_num


def normalize_document_filename_profile(value: object) -> str:
    """Return a supported document filename profile key."""
    text = _to_str(value).strip().lower()
    if text in DOCUMENT_FILENAME_PROFILES:
        return text
    return DOCUMENT_FILENAME_PROFILE_STANDARD


def normalize_document_filename_separator(value: object) -> str:
    """Return a supported document filename separator key."""
    text = _to_str(value).strip().lower()
    if text in DOCUMENT_FILENAME_SEPARATOR_MAP:
        return text
    if text == "_":
        return "underscore"
    if text == "-":
        return "dash"
    if text in {"", "geen"}:
        return "none"
    return "underscore"


def _format_doc_number_for_filename(
    doc_number: object,
    doc_type: object,
    *,
    compact: bool = False,
) -> str:
    """Return the filename-safe document number component."""
    normalized = _normalize_doc_number(doc_number, doc_type)
    if not normalized:
        return ""
    if compact:
        normalized = re.sub(r"[\s\-_]+", "", normalized)
    return _sanitize_component(normalized)


def _join_filename_parts(parts: Sequence[str], separator: str) -> str:
    cleaned = [part for part in (_sanitize_component(part) for part in parts) if part]
    if not cleaned:
        return ""
    if separator:
        return separator.join(cleaned)
    return "".join(cleaned)


def build_document_export_basename(
    doc_type: object,
    doc_number: object = "",
    context_label: object = "",
    export_date: object = None,
    *,
    profile: object = DOCUMENT_FILENAME_PROFILE_STANDARD,
    show_doc_type: bool = True,
    show_doc_number: bool = True,
    show_context: bool = True,
    show_date: bool = True,
    compact_doc_number: bool = False,
    separator: object = "underscore",
    extra_context_label: object = "",
) -> str:
    """Return the export basename for order-related PDF/XLSX files."""
    doc_type_text = _to_str(doc_type).strip() or "Bestelbon"
    context_text = _to_str(context_label).strip()
    extra_context_text = _to_str(extra_context_label).strip()
    date_text = _to_str(export_date).strip() or datetime.date.today().strftime("%Y-%m-%d")
    profile_key = normalize_document_filename_profile(profile)
    separator_key = normalize_document_filename_separator(separator)
    separator_text = DOCUMENT_FILENAME_SEPARATOR_MAP[separator_key]

    doc_num_default = _format_doc_number_for_filename(doc_number, doc_type_text, compact=False)
    doc_num_compact = _format_doc_number_for_filename(doc_number, doc_type_text, compact=True)

    def _standard_name() -> str:
        return _join_filename_parts(
            [
                doc_type_text,
                doc_num_default,
                context_text,
                extra_context_text,
                date_text,
            ],
            "_",
        )

    if profile_key == DOCUMENT_FILENAME_PROFILE_SHORT:
        if doc_num_default:
            return doc_num_default
        profile_key = DOCUMENT_FILENAME_PROFILE_STANDARD
    elif profile_key == DOCUMENT_FILENAME_PROFILE_COMPACT:
        if doc_num_compact:
            return doc_num_compact
        profile_key = DOCUMENT_FILENAME_PROFILE_STANDARD

    if profile_key == DOCUMENT_FILENAME_PROFILE_STANDARD:
        basename = _standard_name()
        return basename or "document"

    doc_num_custom = doc_num_compact if compact_doc_number else doc_num_default
    parts: List[str] = []
    if show_doc_type:
        parts.append(doc_type_text)
    if show_doc_number and doc_num_custom:
        parts.append(doc_num_custom)
    if show_context:
        if context_text:
            parts.append(context_text)
        if extra_context_text:
            parts.append(extra_context_text)
    if show_date and date_text:
        parts.append(date_text)

    basename = _join_filename_parts(parts, separator_text)
    if basename:
        return basename
    fallback = _standard_name()
    return fallback or "document"


def _should_place_remark_in_delivery_block(
    *,
    order_remark_has_content: bool,
    doc_type_text_slug: str,
    is_standaard_doc: bool,
    delivery: DeliveryAddress | None,
) -> bool:
    """Return :data:`True` when remarks belong in the delivery block."""
    doc_type_is_export = "export" in doc_type_text_slug
    return (
        order_remark_has_content
        and doc_type_is_export
        and not is_standaard_doc
    )


# ========== Finish Utilities ==========

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


def describe_finish_combo(
    finish_value: object,
    ral_value: object,
) -> Dict[str, str]:
    """Return normalized metadata for a finish/RAL combination."""
    finish_text = _to_str(finish_value).strip()
    ral_text = _to_str(ral_value).strip()
    finish_norm = _normalize_finish_folder(finish_text)
    if not finish_norm:
        finish_norm = "_Onbekend"
    finish_display = finish_text if finish_text else "Onbekend"
    ral_norm = _normalize_finish_folder(ral_text) if ral_text else ""
    ral_display = ral_text
    folder_name = f"Finish-{finish_norm}"
    if ral_norm:
        folder_name = f"{folder_name}-{ral_norm}"
    label = finish_display
    if ral_display:
        label = f"{finish_display} – {ral_display}"
    filename_component = _normalize_finish_folder(label)
    if not filename_component:
        suffix = f"-{ral_norm}" if ral_norm else ""
        filename_component = f"{finish_norm}{suffix}"
    return {
        "finish_display": finish_display,
        "finish_norm": finish_norm,
        "ral_display": ral_display,
        "ral_norm": ral_norm,
        "folder_name": folder_name,
        "label": label,
        "filename_component": filename_component,
        "key": folder_name,
    }


# ========== Selection Key Functions ==========

def _selection_key(kind: str, identifier: str) -> str:
    """Return a disambiguated key for supplier/doc selections."""
    identifier = _to_str(identifier)
    if kind == "finish":
        return f"{FINISH_KEY_PREFIX}{identifier}"
    if kind == "opticutter":
        return f"{OPTICUTTER_KEY_PREFIX}{identifier}"
    return f"{PRODUCTION_KEY_PREFIX}{identifier}"


def make_production_selection_key(name: str) -> str:
    """Return a stable selection key for a production name."""
    return _selection_key("production", name)


def make_finish_selection_key(finish_key: str) -> str:
    """Return a stable selection key for a finish combination."""
    return _selection_key("finish", finish_key)


def make_opticutter_selection_key(name: str) -> str:
    """Return a stable selection key for Opticutter raw material orders."""
    return _selection_key("opticutter", name)


def make_opticutter_default_key(name: str) -> str:
    """Return the SuppliersDB default key for Opticutter raw material orders."""
    base = _to_str(name)
    return f"{base}{OPTICUTTER_DEFAULT_SUFFIX}" if base else OPTICUTTER_DEFAULT_SUFFIX


def parse_selection_key(key: str) -> Tuple[str, str]:
    """Return the kind (``"production"``/``"finish"``) and identifier."""
    if key.startswith(FINISH_KEY_PREFIX):
        return "finish", key[len(FINISH_KEY_PREFIX) :]
    if key.startswith(PRODUCTION_KEY_PREFIX):
        return "production", key[len(PRODUCTION_KEY_PREFIX) :]
    if key.startswith(OPTICUTTER_KEY_PREFIX):
        return "opticutter", key[len(OPTICUTTER_KEY_PREFIX) :]
    return "production", key


# ========== Opticutter Utilities ==========

def _parse_weight_kg(value: object) -> float | None:
    """Parse a textual kilogram value to float."""
    return _parse_order_measure_value(value)


def _parse_order_measure_value(value: object) -> float | None:
    """Parse a textual surface/weight value to float."""

    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None
    text = _to_str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "nat"}:
        return None
    text = text.replace("\u00a0", "")
    text = text.replace(" ", "")
    text = text.replace(",", ".")
    cleaned = []
    for ch in text:
        if ch.isdigit():
            cleaned.append(ch)
        elif ch in "+-":
            if not cleaned:
                cleaned.append(ch)
        elif ch == ".":
            cleaned.append(".")
    if not cleaned:
        return None
    candidate = "".join(cleaned)
    if candidate.count(".") > 1:
        first = candidate.find(".")
        candidate = candidate[: first + 1] + candidate[first + 1 :].replace(".", "")
    if candidate in {"", "+", "-", ".", "+.", "-."}:
        return None
    try:
        number = float(candidate)
    except Exception:
        return None
    return number if math.isfinite(number) else None


_ORDER_QUANTITY_KEY_HINTS = {"aantal", "qty", "quantity", "st", "stukken", "pieces"}
_ORDER_SURFACE_KEY_HINTS = {
    "oppervlakte",
    "oppervlaktem2",
    "m2",
    "surface",
    "surfacearea",
    "surfaceaream2",
}
_ORDER_WEIGHT_KEY_HINTS = {
    "gewicht",
    "gewichtkg",
    "kg",
    "weight",
    "weightkg",
}


def _normalize_order_measure_key(value: object) -> str:
    text = _to_str(value).strip().lower().replace("\u00b2", "2")
    return re.sub(r"[^a-z0-9]+", "", text)


def is_order_surface_column(column: Mapping[str, object]) -> bool:
    """Return whether an order column should receive surface totals."""

    if bool(column.get("total_surface")):
        return True
    key = _normalize_order_measure_key(column.get("key"))
    label = _normalize_order_measure_key(column.get("label"))
    return key in _ORDER_SURFACE_KEY_HINTS or label in _ORDER_SURFACE_KEY_HINTS


def is_order_weight_column(column: Mapping[str, object]) -> bool:
    """Return whether an order column should receive weight totals."""

    if bool(column.get("total_weight")):
        return True
    key = _normalize_order_measure_key(column.get("key"))
    label = _normalize_order_measure_key(column.get("label"))
    return key in _ORDER_WEIGHT_KEY_HINTS or label in _ORDER_WEIGHT_KEY_HINTS


def _first_order_measure_value(
    item: Mapping[str, object],
    key_hints: set[str],
) -> object:
    for key, value in item.items():
        if _normalize_order_measure_key(key) in key_hints:
            return value
    return None


def _order_item_quantity(item: Mapping[str, object]) -> float:
    qty_value = _first_order_measure_value(item, _ORDER_QUANTITY_KEY_HINTS)
    qty = _parse_order_measure_value(qty_value)
    if qty is None or qty <= 0:
        return 1.0
    return qty


def calculate_order_measure_totals(
    items: Sequence[Mapping[str, object]],
) -> tuple[float | None, float | None]:
    """Return total surface m2 and weight kg for order items.

    Surface and weight values are treated as per-piece values and multiplied by
    the row quantity when a quantity column is present.
    """

    surface_total = 0.0
    weight_total = 0.0
    surface_found = False
    weight_found = False

    for item in items:
        qty = _order_item_quantity(item)
        surface_each = _parse_order_measure_value(
            _first_order_measure_value(item, _ORDER_SURFACE_KEY_HINTS)
        )
        if surface_each is not None:
            surface_total += surface_each * qty
            surface_found = True
        weight_each = _parse_order_measure_value(
            _first_order_measure_value(item, _ORDER_WEIGHT_KEY_HINTS)
        )
        if weight_each is not None:
            weight_total += weight_each * qty
            weight_found = True

    return (
        surface_total if surface_found else None,
        weight_total if weight_found else None,
    )


def _collect_opticutter_profile_stats(
    bom_df: pd.DataFrame,
) -> Dict[tuple[str, str, str], OpticutterProfileStats]:
    """Aggregate length and weight per (profile, material, production)."""
    stats: Dict[tuple[str, str, str], OpticutterProfileStats] = defaultdict(
        OpticutterProfileStats
    )
    for _, row in bom_df.iterrows():
        profile_name = _to_str(row.get("Profile")).strip()
        if not profile_name:
            continue
        material_name = _to_str(row.get("Materiaal") or row.get("Material")).strip()
        production_name = _to_str(row.get("Production")).strip() or "_Onbekend"
        weight_each = _parse_weight_kg(row.get("Gewicht"))
        if weight_each is None or weight_each <= 0:
            continue
        qty_raw = row.get("Aantal", 0)
        try:
            qty = int(float(qty_raw))
        except Exception:
            qty = 0
        if qty <= 0:
            continue
        length_mm: int | None = None
        length_value = row.get("Length profile mm")
        if length_value is not None and not pd.isna(length_value):
            try:
                length_mm = int(round(float(length_value)))
            except Exception:
                length_mm = None
        if length_mm is None or length_mm <= 0:
            length_mm = parse_length_to_mm(row.get("Length profile"))
        if length_mm is None or length_mm <= 0:
            continue
        key = (profile_name, material_name, production_name)
        entry = stats[key]
        entry.total_length_mm += float(length_mm) * qty
        entry.total_weight_kg += float(weight_each) * qty
    return stats


def _format_weight_kg(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.2f}"


def _compute_opticutter_order_exports(
    opticutter_prod,
    stats_map: Mapping[tuple[str, str, str], OpticutterProfileStats],
) -> OpticutterOrderComputation:
    """Compute Opticutter order exports from production and stats."""
    scenario_rows: List[Dict[str, object]] = []
    piece_rows: List[Dict[str, object]] = []
    order_rows: List[Dict[str, object]] = []
    raw_items: List[Dict[str, object]] = []
    has_valid_bars = False
    total_bars = 0
    total_weight_known = False
    total_weight_kg = 0.0
    selection_count = 0

    for selection in opticutter_prod.selections:
        selection_count += 1
        profile = selection.profile
        result = selection.result
        stock_length = selection.stock_length_mm

        remark_lines: List[str] = []
        if selection.choice == "input":
            remark_lines.append("Per stuk zagen (inputlengte).")
        elif result is None:
            remark_lines.append("Geen scenario berekend.")
        elif result.dropped_pieces:
            remark_lines.append(
                f"Niet mogelijk: {result.dropped_pieces} stuk(ken) zijn te lang."
            )
        if selection.blockers:
            remark_lines.append("Blokkerende stukken:")
            remark_lines.extend(f"- {text}" for text in selection.blockers)
        elif result is not None and not result.dropped_pieces:
            remark_lines.append(f"Totale restlengte: {result.waste_mm:.0f} mm.")
        scenario_remark = "\n".join(remark_lines).strip()

        bars_value = (
            result.bars if result is not None and not result.dropped_pieces else None
        )
        waste_pct_value = (
            round(result.waste_pct, 1)
            if result is not None and not result.dropped_pieces
            else None
        )
        waste_mm_value = (
            round(result.waste_mm, 0)
            if result is not None and not result.dropped_pieces
            else None
        )
        cuts_value = (
            result.cuts if result is not None and not result.dropped_pieces else None
        )

        scenario_rows.append(
            {
                "Profiel": profile.profile,
                "Materiaal": profile.material,
                "Productie": profile.production,
                "Keuze": selection.choice_label,
                "Staaflengte (mm)": stock_length,
                "Aantal staven": bars_value,
                "Afval %": waste_pct_value,
                "Afval (mm)": waste_mm_value,
                "Zaagsneden": cuts_value,
                "Opmerking": scenario_remark,
            }
        )

        for piece in profile.pieces:
            piece_rows.append(
                {
                    "Profiel": profile.profile,
                    "Materiaal": profile.material,
                    "Productie": profile.production,
                    "Onderdeel": piece.label,
                    "Lengte (mm)": piece.length_mm,
                }
            )

        if selection.choice == "input":
            order_remark_text = "Handmatig zagen per stuk."
        elif result is None or result.dropped_pieces:
            blockers_text = "; ".join(selection.blockers)
            if blockers_text:
                order_remark_text = f"Handmatig controleren – {blockers_text}"
            else:
                order_remark_text = "Handmatig controleren – scenario niet mogelijk."
        else:
            order_remark_text = (
                f"Afval {result.waste_pct:.1f}% ({result.waste_mm:.0f} mm)."
            )

        total_length_m = None
        if result is not None and not result.dropped_pieces and stock_length is not None:
            total_length_m = round(result.bars * stock_length / 1000, 3)

        weight_total = None
        stats_key = (profile.profile, profile.material, profile.production)
        stats_entry = stats_map.get(stats_key)
        if (
            stats_entry is not None
            and stats_entry.weight_per_mm is not None
            and stock_length is not None
            and bars_value
        ):
            weight_total = (
                float(stats_entry.weight_per_mm)
                * float(stock_length)
                * float(bars_value)
                / 1.0
            )
            total_weight_known = True
            total_weight_kg += weight_total

        order_rows.append(
            {
                "Profiel": profile.profile,
                "Materiaal": profile.material,
                "Productie": profile.production,
                "Keuze": selection.choice_label,
                "Staaflengte (mm)": stock_length,
                "Aantal staven": bars_value,
                "Totale lengte (m)": total_length_m,
                "Totaal gewicht (kg)": round(weight_total, 3)
                if weight_total is not None
                else None,
                "Opmerking": order_remark_text,
            }
        )

        if bars_value and bars_value > 0:
            has_valid_bars = True
            total_bars += int(bars_value)
            raw_items.append(
                {
                    "Profiel": profile.profile or "Brutemateriaal",
                    "Materiaal": profile.material,
                    "Lengte": stock_length or "",
                    "St.": int(bars_value),
                    "kg": _format_weight_kg(weight_total),
                }
            )

    summary_weight = total_weight_kg if total_weight_known else None
    return OpticutterOrderComputation(
        scenario_rows=scenario_rows,
        piece_rows=piece_rows,
        order_rows=order_rows,
        raw_items=raw_items,
        has_valid_bars=has_valid_bars,
        total_bars=total_bars,
        total_weight_kg=summary_weight,
        selection_count=selection_count,
    )


def compute_opticutter_order_details(
    bom_df: pd.DataFrame,
    context,
) -> Dict[str, OpticutterOrderComputation]:
    """Return computed Opticutter export data per production."""
    if context is None:
        return {}
    stats_map = _collect_opticutter_profile_stats(bom_df)
    details: Dict[str, OpticutterOrderComputation] = {}
    for prod_key, export in context.productions.items():
        details[prod_key] = _compute_opticutter_order_exports(export, stats_map)
    return details


# ========== Supplier Selection ==========

def pick_supplier_for_production(
    prod: str,
    db: SuppliersDB,
    override_map: Dict[str, str],
    suppliers_sorted: List[Supplier] | None = None,
) -> Supplier:
    """Select a supplier for a given production."""
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
    return Supplier(supplier=NO_SUPPLIER_PLACEHOLDER)


def pick_supplier_for_opticutter(
    prod: str,
    db: SuppliersDB,
    override_map: Dict[str, str],
    suppliers_sorted: List[Supplier] | None = None,
) -> Supplier:
    """Select a supplier for Opticutter raw material orders."""
    key = make_opticutter_default_key(prod)
    name = override_map.get(prod)
    sups = suppliers_sorted if suppliers_sorted is not None else db.suppliers_sorted()
    if name is not None:
        if not name.strip():
            return Supplier(supplier="")
        for s in sups:
            if s.supplier.lower() == name.lower():
                return s
        return Supplier(supplier=name)
    default = db.get_default(key)
    if default:
        for s in sups:
            if s.supplier.lower() == default.lower():
                return s
    return Supplier(supplier=NO_SUPPLIER_PLACEHOLDER)


def pick_supplier_for_finish(
    finish_key: str,
    db: SuppliersDB,
    override_map: Dict[str, str],
    suppliers_sorted: List[Supplier] | None = None,
) -> Supplier:
    """Select a supplier for a finish combination."""
    name = override_map.get(finish_key)
    sups = suppliers_sorted if suppliers_sorted is not None else db.suppliers_sorted()
    if name is not None:
        if not name.strip():
            return Supplier(supplier="")
        for s in sups:
            if s.supplier.lower() == name.lower():
                return s
        return Supplier(supplier=name)
    default = db.get_default_finish(finish_key)
    if default:
        for s in sups:
            if s.supplier.lower() == default.lower():
                return s
    return Supplier(supplier=NO_SUPPLIER_PLACEHOLDER)
