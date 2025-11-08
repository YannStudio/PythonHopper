"""Order-related utilities for creating purchase documents and copying files.

This module groups functions previously in Main_v22.py and depends on helpers,
models, suppliers_db, and bom modules.
"""

import os
import sys
import shutil
import datetime
import re
import unicodedata
import zipfile
import io
import tempfile
import hashlib
from collections import defaultdict
from pathlib import Path
from typing import Callable, Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd
from dataclasses import dataclass
try:
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover - optional dependency
    Alignment = None
    Font = None
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

from en1090 import EN1090_NOTE_TEXT, should_require_en1090

from opticutter import (
    OpticutterAnalysis,
    OpticutterExportContext,
    OpticutterProductionExport,
    OpticutterSelection,
    parse_length_to_mm,
    prepare_opticutter_export,
)

MIAMI_PINK = "#FF77FF"
DEFAULT_FOOTER_NOTE = (
    "Gelieve afwijkingen schriftelijk te bevestigen. "
    "Levertermijn in overleg. Betalingsvoorwaarden: 30 dagen netto. "
    "Vermeld onze productiereferentie bij levering."
)

STEP_EXTS = {".step", ".stp"}

NO_SUPPLIER_PLACEHOLDER = "(geen)"


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


_INVALID_PATH_CHARS = set('<>:"/\\|?*')
_WINDOWS_MAX_PATH = 240


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
    """Return ``filename`` possibly shortened so ``directory/filename`` fits ``max_path``.

    On Windows, opening files with paths longer than 260 characters raises ``FileNotFoundError``.
    ``max_path`` defaults to a slightly smaller value (240) to provide some safety margin for
    runtime conversions that may add characters (e.g. via extended paths). When the composed path
    exceeds the limit, the base filename is truncated and suffixed with an 8-character hash so the
    resulting name remains unique and deterministic while staying within the limit. ``max_path`` can
    be overridden (primarily for tests). When the directory path already exceeds the limit the
    function raises :class:`OSError` as there is no filename that could satisfy the constraint.
    """

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


def _export_bom_workbook(bom_df: pd.DataFrame, dest: str, filename: str) -> str:
    """Write the processed BOM dataframe to an Excel workbook."""

    if not filename.lower().endswith(".xlsx"):
        filename = f"{filename}.xlsx"
    target_path = os.path.join(dest, filename)
    export_df = bom_df.reset_index(drop=True).copy()
    # Drop status-related columns that are only useful inside the app.
    to_drop = [col for col in _BOM_STATUS_COLUMNS if col in export_df.columns]
    if to_drop:
        export_df = export_df.drop(columns=to_drop)

    # Normalise quantity column naming to ``QTY.`` and drop aliases.
    qty_aliases: Tuple[str, ...] = ("QTY.", "Qty.", "Qty", "Quantity", "Aantal")
    qty_columns = [col for col in qty_aliases if col in export_df.columns]
    if "QTY." not in export_df.columns:
        if qty_columns:
            export_df = export_df.rename(columns={qty_columns[0]: "QTY."})
        else:
            export_df["QTY."] = ""
    for alias in qty_columns:
        if alias == "QTY." or alias not in export_df.columns:
            continue
        source = export_df[alias]
        destination = export_df["QTY."]
        dest_str = destination.astype(str)
        missing_mask = dest_str.str.strip().isin(("", "nan"))
        export_df.loc[missing_mask, "QTY."] = source[missing_mask]
        export_df = export_df.drop(columns=alias)

    # Ensure all primary BOM columns are present and appear first.
    for column in _BOM_EXPORT_BASE_COLUMNS:
        if column not in export_df.columns:
            export_df[column] = ""
    ordered_columns = [c for c in _BOM_EXPORT_BASE_COLUMNS if c in export_df.columns]
    remaining_columns = [c for c in export_df.columns if c not in ordered_columns]
    export_df = export_df[ordered_columns + remaining_columns]

    with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="BOM")
        if Alignment is not None and get_column_letter is not None:
            ws = writer.sheets["BOM"]
            alignment = Alignment(wrap_text=False, vertical="top")
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = alignment
            for col_idx in range(1, ws.max_column + 1):
                column_letter = get_column_letter(col_idx)
                max_length = 0
                for cell in ws[column_letter]:
                    value = cell.value
                    cell_length = len(str(value)) if value is not None else 0
                    if cell_length > max_length:
                        max_length = cell_length
                ws.column_dimensions[column_letter].width = min(max(12, max_length + 2), 80)

    return target_path


def make_bom_export_filename(
    bom_source_path: Optional[str],
    date_iso: str,
    transform: Callable[[str], str],
) -> str:
    """Return a normalized filename for exporting the processed BOM workbook."""

    source_stem = ""
    if bom_source_path:
        source_stem = Path(bom_source_path).stem
        match = re.search(r"(.*?\bBOM\b)", source_stem, flags=re.IGNORECASE)
        if match:
            source_stem = match.group(1)
        source_stem = source_stem.rstrip(" -_.")
    stem = source_stem or "BOM-FileHopper-Export"
    stem_with_date = f"{stem}-{date_iso}"
    filename = f"{stem_with_date}.xlsx"
    return transform(filename)


def find_related_bom_exports(
    bom_source_path: Optional[str],
    file_index: Mapping[str, Sequence[str]],
) -> List[str]:
    """Return export files whose stem appears in the BOM filename."""

    if not bom_source_path:
        return []
    stem = Path(bom_source_path).stem.lower()
    if not stem:
        return []
    matches: List[str] = []
    seen: set[str] = set()
    for key in sorted(file_index.keys(), key=len, reverse=True):
        if not key:
            continue
        key_lower = key.lower()
        if len(key_lower) < 4 and not any(ch.isdigit() for ch in key_lower):
            continue
        idx = stem.find(key_lower)
        if idx == -1:
            continue
        if idx > 0 and stem[idx - 1].isalnum():
            continue
        end = idx + len(key_lower)
        if end < len(stem) and stem[end].isalnum():
            continue
        for src_file in file_index.get(key, []):
            if src_file not in seen:
                matches.append(src_file)
                seen.add(src_file)
    return matches


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
    if t.startswith("standaard"):
        return "BOM-"
    if t.startswith("bestel"):
        return "BB-"
    if t.startswith("offerte"):
        return "OFF-"
    return ""


def _normalize_doc_number(value: object, doc_type: object) -> str:
    """Return a cleaned document number for a given ``doc_type``.

    The GUI provides placeholder prefixes such as ``"BB-"``. When the user
    pastes a value that already contains the prefix the placeholder should be
    replaced instead of duplicated (``"BB-BB123"`` → ``"BB-123"``).
    """

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


FINISH_KEY_PREFIX = "finish::"
PRODUCTION_KEY_PREFIX = "production::"
OPTICUTTER_KEY_PREFIX = "opticutter::"
OPTICUTTER_DEFAULT_SUFFIX = "::Opticutter"


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
    # Fallback for legacy keys without explicit prefix.
    return "production", key


def _parse_weight_kg(value: object) -> float | None:
    """Parse a textual kilogram value to float."""

    text = _to_str(value).strip()
    if not text:
        return None
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
        return float(candidate)
    except Exception:
        return None


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
    opticutter_prod: OpticutterProductionExport,
    stats_map: Mapping[tuple[str, str, str], OpticutterProfileStats],
) -> OpticutterOrderComputation:
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
    context: OpticutterExportContext | None,
) -> Dict[str, OpticutterOrderComputation]:
    """Return computed Opticutter export data per production."""

    if context is None:
        return {}
    stats_map = _collect_opticutter_profile_stats(bom_df)
    details: Dict[str, OpticutterOrderComputation] = {}
    for prod_key, export in context.productions.items():
        details[prod_key] = _compute_opticutter_order_exports(export, stats_map)
    return details


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


def generate_pdf_order_platypus(
    path: str,
    company_info: Dict[str, object],
    supplier: Supplier | None,
    production: str,
    items: List[Dict[str, object]],
    doc_type: str = "Bestelbon",
    doc_number: str | None = None,
    footer_note: Optional[str] = None,
    delivery: DeliveryAddress | None = None,
    project_number: str | None = None,
    project_name: str | None = None,
    label_kind: str = "productie",
    order_remark: str | None = None,
    total_weight_kg: float | None = None,
    en1090_required: bool = False,
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

    doc_type_text = (_to_str(doc_type).strip() or "Bestelbon")
    doc_type_text_lower = doc_type_text.lower()
    doc_type_text_slug = re.sub(r"[^0-9a-z]+", "", doc_type_text_lower)
    is_standaard_doc = doc_type_text_lower.startswith("standaard")
    order_remark_text = _to_str(order_remark) if order_remark is not None else ""
    order_remark_has_content = bool(order_remark_text.strip())
    place_remark_in_delivery_block = _should_place_remark_in_delivery_block(
        order_remark_has_content=order_remark_has_content,
        doc_type_text_slug=doc_type_text_slug,
        is_standaard_doc=is_standaard_doc,
        delivery=delivery,
    )

    doc_lines: List[str] = []
    if doc_number:
        doc_lines.append(f"Nummer: {doc_number}")
    today = datetime.date.today().strftime("%Y-%m-%d")
    doc_lines.append(f"Datum: {today}")
    label_kind_clean = (_to_str(label_kind) or "productie").strip() or "productie"
    label_title = label_kind_clean[0].upper() + label_kind_clean[1:]
    is_raw_material_order = label_kind_clean.lower().startswith("brutemateriaal")
    if production:
        doc_lines.append(f"{label_title}: {production}")
    if project_number:
        doc_lines.append(f"Projectnummer: {project_number}")
    if project_name:
        doc_lines.append(f"Projectnaam: {project_name}")
    if order_remark_has_content and not place_remark_in_delivery_block:
        doc_lines.append(f"Opmerking: {order_remark_text}")

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

    supp_lines: List[str] = []
    if supplier is not None and not is_standaard_doc:
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

    left_lines = list(company_lines)
    if supp_lines:
        left_lines.append("")
        left_lines.extend(supp_lines)
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
    include_right_block = not is_standaard_doc and (
        delivery is not None or place_remark_in_delivery_block
    )
    if include_right_block:
        if delivery:
            # Delivery address block with each piece of information on its own line
            right_lines.append("<b>Leveradres:</b>")
            right_lines.append(delivery.name)
            if delivery.address:
                right_lines.extend(delivery.address.splitlines())
            if delivery.remarks:
                right_lines.append(delivery.remarks)
        if place_remark_in_delivery_block:
            right_lines.append("<b>Opmerking:</b>")
            remark_lines = order_remark_text.splitlines()
            if not remark_lines:
                remark_lines = [order_remark_text]
            right_lines.extend(remark_lines)

    story = []
    title = (
        f"{doc_type_text} {label_kind_clean}: {production}"
        if production
        else f"{doc_type_text}"
    )
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
    if is_raw_material_order:
        head = ["Profiel", "Materiaal", "Lengte", "St.", "kg"]
    else:
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
    total_row_index: int | None = None
    if is_raw_material_order:
        for it in items:
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
    if is_raw_material_order:
        col_fracs = [0.32, 0.24, 0.16, 0.12, 0.16]
        col_widths = [usable_w * frac for frac in col_fracs]
    else:
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
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(MIAMI_PINK)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]
    if is_raw_material_order:
        style_cmds.append(("ALIGN", (2, 0), (4, -1), "RIGHT"))
    else:
        style_cmds.extend(
            [
                ("ALIGN", (2, 0), (5, 0), "RIGHT"),
                ("ALIGN", (2, 1), (5, -1), "RIGHT"),
            ]
        )
    if total_row_index is not None:
        style_cmds.append(("FONTNAME", (0, total_row_index), (-1, total_row_index), "Helvetica-Bold"))
    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)

    if en1090_required:
        story.append(Spacer(0, 6))
        story.append(Paragraph(f"<b>{EN1090_NOTE_TEXT}</b>", text_style))

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
    items: List[Dict[str, object]],
    company_info: Dict[str, str] | None = None,
    supplier: Supplier | None = None,
    delivery: DeliveryAddress | None = None,
    doc_type: str = "Bestelbon",
    doc_number: str | None = None,
    project_number: str | None = None,
    project_name: str | None = None,
    context_label: str | None = None,
    context_kind: str = "productie",
    order_remark: str | None = None,
    total_weight_kg: float | None = None,
    en1090_required: bool = False,
) -> None:
    """Write order information to an Excel file with header info."""
    context_kind_clean = (_to_str(context_kind) or "productie").strip() or "productie"
    is_raw_material_order = context_kind_clean.lower().startswith("brutemateriaal")
    if is_raw_material_order:
        df_columns = ["Profiel", "Materiaal", "Lengte", "St.", "kg"]
    else:
        df_columns = ["PartNumber", "Description", "Materiaal", "Aantal", "Oppervlakte", "Gewicht"]
    df = pd.DataFrame(items, columns=df_columns)
    if is_raw_material_order and total_weight_kg is not None:
        total_row = {
            "Profiel": "Totaal",
            "Materiaal": "",
            "Lengte": "",
            "St.": "",
            "kg": _format_weight_kg(total_weight_kg),
        }
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    append_note_to_df = en1090_required and (
        Alignment is None or not hasattr(pd, "ExcelWriter")
    )
    if append_note_to_df:
        blank_row = {col: "" for col in df_columns}
        note_row = {col: "" for col in df_columns}
        if df_columns:
            note_row[df_columns[0]] = EN1090_NOTE_TEXT
        df = pd.concat(
            [df, pd.DataFrame([blank_row, note_row])], ignore_index=True
        )

    doc_type_text = (_to_str(doc_type).strip() or "Bestelbon")
    doc_type_text_lower = doc_type_text.lower()
    doc_type_text_slug = re.sub(r"[^0-9a-z]+", "", doc_type_text_lower)
    is_standaard_doc = doc_type_text_lower.startswith("standaard")
    order_remark_text = _to_str(order_remark) if order_remark is not None else ""
    order_remark_has_content = bool(order_remark_text.strip())
    place_remark_in_delivery_block = _should_place_remark_in_delivery_block(
        order_remark_has_content=order_remark_has_content,
        doc_type_text_slug=doc_type_text_slug,
        is_standaard_doc=is_standaard_doc,
        delivery=delivery,
    )

    header_lines: List[Tuple[str, str]] = []
    today = datetime.date.today().strftime("%Y-%m-%d")
    if doc_number:
        header_lines.append(("Nummer", str(doc_number)))
    header_lines.append(("Datum", today))
    if context_label:
        header_lines.append((context_kind_clean.capitalize(), context_label))
    if project_number:
        header_lines.append(("Projectnummer", project_number))
    if project_name:
        header_lines.append(("Projectnaam", project_name))
    if order_remark_has_content and not place_remark_in_delivery_block:
        header_lines.append(("Opmerking", order_remark_text))
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
    supplier_name = _to_str(supplier.supplier).strip() if supplier else ""
    include_supplier_block = supplier is not None and (
        not is_standaard_doc or bool(supplier_name)
    )
    if include_supplier_block:
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
    include_delivery_block = False
    if delivery is not None:
        delivery_has_content = any(
            _to_str(value).strip()
            for value in (delivery.name, delivery.address, delivery.remarks)
        )
        include_delivery_block = (
            not is_standaard_doc
            or delivery_has_content
            or place_remark_in_delivery_block
        )
    elif place_remark_in_delivery_block and not is_standaard_doc:
        include_delivery_block = True
    if include_delivery_block:
        if delivery:
            header_lines.extend(
                [
                    ("Leveradres", ""),
                    ("", delivery.name),
                    ("Adres", delivery.address or ""),
                    ("Opmerking", delivery.remarks or ""),
                ]
            )
        if place_remark_in_delivery_block and order_remark_has_content:
            header_lines.append(("Opmerking", order_remark_text))
        header_lines.append(("", ""))

    startrow = len(header_lines)
    if Alignment is not None and hasattr(pd, "ExcelWriter"):
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, startrow=startrow)
            ws = writer.sheets[list(writer.sheets.keys())[0]]
            for r, (label, value) in enumerate(header_lines, start=1):
                ws.cell(row=r, column=1, value=label)
                ws.cell(row=r, column=2, value=value)

            if is_raw_material_order:
                left_cols = {"Profiel", "Materiaal"}
                wrap_cols = {"Profiel", "Materiaal"}
            else:
                left_cols = {"PartNumber", "Description"}
                wrap_cols = {"PartNumber", "Description"}
            for col_idx, col_name in enumerate(df.columns, start=1):
                align = Alignment(
                    horizontal="left" if col_name in left_cols else "right",
                    wrap_text=col_name in wrap_cols,
                )
                if (
                    col_name in {"PartNumber", "Profiel"}
                    and get_column_letter is not None
                ):
                    column_letter = get_column_letter(col_idx)
                    ws.column_dimensions[column_letter].width = 25
                for row in range(startrow + 1, startrow + len(df) + 2):
                    ws.cell(row=row, column=col_idx).alignment = align

            if en1090_required and not append_note_to_df:
                note_row = ws.max_row + 2
                cell = ws.cell(row=note_row, column=1, value=EN1090_NOTE_TEXT)
                if Font is not None:
                    cell.font = Font(bold=True)
                if Alignment is not None:
                    cell.alignment = Alignment(horizontal="left", wrap_text=True)


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
    # Geen eerder geselecteerde leverancier onthouden: vul de placeholder in.
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
    zip_finish_exports: bool = True,
    export_bom: bool = True,
    export_related_files: bool = True,
    finish_override_map: Dict[str, str] | None = None,
    finish_doc_type_map: Dict[str, str] | None = None,
    finish_doc_num_map: Dict[str, str] | None = None,
    finish_delivery_map: Dict[str, DeliveryAddress | None] | None = None,
    remarks_map: Dict[str, str] | None = None,
    finish_remarks_map: Dict[str, str] | None = None,
    bom_source_path: str | None = None,
    path_limit_warnings: List[str] | None = None,
    opticutter_analysis: OpticutterAnalysis | None = None,
    opticutter_choices: Mapping[tuple[str, str, str], str] | None = None,
    opticutter_override_map: Dict[str, str] | None = None,
    opticutter_doc_type_map: Dict[str, str] | None = None,
    opticutter_doc_num_map: Dict[str, str] | None = None,
    opticutter_delivery_map: Dict[str, DeliveryAddress | None] | None = None,
    opticutter_remarks_map: Dict[str, str] | None = None,
    en1090_overrides: Mapping[str, bool] | None = None,
) -> Tuple[int, Dict[str, str]]:
    """Copy files per production and create accompanying order documents.

    ``doc_type_map`` may specify per production whether a *Bestelbon* or an
    *Offerteaanvraag* should be generated. Missing entries default to
    ``"Bestelbon"``.

    ``doc_num_map`` provides document numbers per production which are used in
    filenames and document headers.

    ``delivery_map`` can provide a :class:`DeliveryAddress` per production.

    ``remarks_map`` and ``finish_remarks_map`` allow passing additional notes for
    productions and finishes respectively. When provided, the remarks are added
    to the generated Excel- en PDF-bestanden.

    ``opticutter_analysis`` en ``opticutter_choices`` laten toe om per
    productie extra Opticutter-overzichten en bestelbonnen voor brutemateriaal
    aan te maken. Wanneer beide waarden aanwezig zijn wordt in iedere
    productiemap een Opticutter-werkboek met scenario-informatie en een
    besteloverzicht voor volle lengten geschreven. Extra kaarten voor
    Opticutter-bestellingen kunnen voorzien worden via
    ``opticutter_override_map``, ``opticutter_doc_type_map``,
    ``opticutter_doc_num_map``, ``opticutter_delivery_map`` en
    ``opticutter_remarks_map``.

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
    ``RAL color`` values. When ``zip_finish_exports`` is ``True`` (default)
    the finish folders receive a single ZIP archive containing the export
    files instead of loose copies.

    When ``export_bom`` is ``True`` (default) the processed BOM, including any
    changes made inside Filehopper, is written as an Excel workbook in the root
    of ``dest``. The filename contains the export date in ISO-formaat.

    When ``bom_source_path`` refers to the loaded BOM file, Filehopper tries to
    copy export files from ``source`` whose filename stem appears inside the BOM
    name (for example ``123-BOM-PartsOnly`` → ``123.pdf``). Matching files are
    placed next to the exported BOM workbook and use the same optional
    name-transformations (date/prefix/suffix). When no related files are found
    nothing is copied. Set ``export_related_files`` to ``False`` to skip copying
    these auxiliary files even when a BOM export is created.

    Finish-specific overrides, document types/numbers and deliveries can be
    provided via the ``finish_*`` mappings. Keys correspond to the normalized
    ``Finish-...`` folder names produced by :func:`describe_finish_combo`.

    The returned ``chosen`` mapping uses selection keys produced by
    :func:`make_production_selection_key` for productions and
    :func:`make_finish_selection_key` for finish selections.
    """
    os.makedirs(dest, exist_ok=True)
    file_index = _build_file_index(source, selected_exts)
    selected_exts_set = {ext.lower() for ext in selected_exts}
    count_copied = 0
    chosen: Dict[str, str] = {}
    doc_type_map = doc_type_map or {}
    doc_num_map = doc_num_map or {}
    finish_override_map = finish_override_map or {}
    finish_doc_type_map = finish_doc_type_map or {}
    finish_doc_num_map = finish_doc_num_map or {}
    finish_delivery_map = finish_delivery_map or {}
    opticutter_override_map = opticutter_override_map or {}
    opticutter_doc_type_map = opticutter_doc_type_map or {}
    opticutter_doc_num_map = opticutter_doc_num_map or {}
    opticutter_delivery_map = opticutter_delivery_map or {}
    opticutter_remarks_map = {
        key: _to_str(value).strip()
        for key, value in (opticutter_remarks_map or {}).items()
        if _to_str(value).strip()
    }
    remarks_clean: Dict[str, str] = {}
    for key, value in (remarks_map or {}).items():
        text = _to_str(value).strip()
        if text:
            remarks_clean[key] = text
    remarks_map = remarks_clean

    finish_remarks_clean: Dict[str, str] = {}
    for key, value in (finish_remarks_map or {}).items():
        text = _to_str(value).strip()
        if text:
            finish_remarks_clean[key] = text
    finish_remarks_map = finish_remarks_clean

    opticutter_context: OpticutterExportContext | None = None
    if opticutter_analysis is not None and opticutter_analysis.profiles:
        try:
            opticutter_context = prepare_opticutter_export(
                opticutter_analysis, opticutter_choices or {}
            )
        except Exception:
            opticutter_context = None

    opticutter_details_map: Dict[str, OpticutterOrderComputation] = {}
    opticutter_stats_map: Dict[
        tuple[str, str, str], OpticutterProfileStats
    ] = {}
    if opticutter_context is not None:
        try:
            opticutter_stats_map = _collect_opticutter_profile_stats(bom_df)
            for prod_key, export in opticutter_context.productions.items():
                opticutter_details_map[prod_key] = _compute_opticutter_order_exports(
                    export, opticutter_stats_map
                )
        except Exception:
            opticutter_details_map = {}
            opticutter_stats_map = {}

    prod_to_rows: Dict[str, List[dict]] = defaultdict(list)
    step_entries: Dict[str, List[tuple[str, str]]] = defaultdict(list)
    step_seen: Dict[str, set[str]] = defaultdict(set)
    finish_groups: Dict[str, Dict[str, object]] = {}
    for _, row in bom_df.iterrows():
        prod = (row.get("Production") or "").strip() or "_Onbekend"
        prod_to_rows[prod].append(row)
        pn = _to_str(row.get("PartNumber")).strip()
        finish_text = _to_str(row.get("Finish")).strip()
        if finish_text:
            finish_meta = describe_finish_combo(
                row.get("Finish"), row.get("RAL color")
            )
            finish_key = finish_meta["key"]
            group = finish_groups.get(finish_key)
            if group is None:
                group = {
                    **finish_meta,
                    "rows": [],
                    "part_numbers": set(),
                }
                finish_groups[finish_key] = group
            group["rows"].append(row)
            if pn:
                group["part_numbers"].add(pn)

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

    def _record_path_warning(
        directory: str,
        requested: str,
        final: str,
        *,
        context: str,
    ) -> None:
        if path_limit_warnings is None or requested == final:
            return
        directory_abs = os.path.abspath(directory)
        original_abs = os.path.join(directory_abs, requested)
        detail = (
            f"{context}: '{requested}' → '{final}' "
            f"(padlengte {len(original_abs)} tekens, limiet {_WINDOWS_MAX_PATH})"
        )
        if detail not in path_limit_warnings:
            path_limit_warnings.append(detail)

    footer_note_text = (
        DEFAULT_FOOTER_NOTE
        if footer_note is None
        else _to_str(footer_note).replace("\r\n", "\n")
    )

    for prod, rows in prod_to_rows.items():
        prod_folder = os.path.join(dest, prod)
        os.makedirs(prod_folder, exist_ok=True)

        opticutter_prod = None
        opticutter_comp: OpticutterOrderComputation | None = None
        if opticutter_context is not None:
            opticutter_prod = opticutter_context.productions.get(prod)
            opticutter_comp = opticutter_details_map.get(prod)

        raw_doc_type = doc_type_map.get(prod, "Bestelbon")
        doc_type = _to_str(raw_doc_type).strip() or "Bestelbon"
        doc_num = _normalize_doc_number(doc_num_map.get(prod, ""), doc_type)
        prefix = _prefix_for_doc_type(doc_type)
        if doc_num and prefix and doc_num.upper() == prefix.upper():
            doc_num = ""
        doc_num_token = _sanitize_component(doc_num) if doc_num else ""
        num_part = f"_{doc_num_token}" if doc_num_token else ""
        doc_type_lower = doc_type.lower()
        is_standaard_doc = doc_type_lower.startswith("standaard")

        zf = None
        if zip_parts:
            zip_name = _fit_filename_within_path(
                prod_folder, f"{prod}{num_part}.zip"
            )
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
                if selected_exts_set and ext not in selected_exts_set:
                    continue
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
        chosen[make_production_selection_key(prod)] = supplier.supplier
        if remember_defaults and supplier.supplier not in ("", "Onbekend", NO_SUPPLIER_PLACEHOLDER):
            db.set_default(prod, supplier.supplier)

        en1090_required = should_require_en1090(prod, en1090_overrides)

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
        supplier_name_clean = _to_str(supplier.supplier).strip()
        delivery = delivery_map.get(prod)
        order_remark = (remarks_map.get(prod, "") if remarks_map else "").strip()
        supplier_for_docs: Supplier | None = supplier
        delivery_for_docs = delivery
        if is_standaard_doc and not supplier_name_clean:
            supplier_for_docs = None
            delivery_for_docs = None

        if supplier_name_clean or is_standaard_doc:
            excel_requested = f"{doc_type}{num_part}_{prod}_{today}.xlsx"
            excel_filename = _fit_filename_within_path(prod_folder, excel_requested)
            _record_path_warning(
                prod_folder,
                excel_requested,
                excel_filename,
                context=f"Productie '{prod}' – {doc_type}",
            )
            excel_path = os.path.join(prod_folder, excel_filename)
            write_order_excel(
                excel_path,
                items,
                company,
                supplier_for_docs,
                delivery_for_docs,
                doc_type,
                doc_num or None,
                project_number=project_number,
                project_name=project_name,
                context_label=prod,
                context_kind="Productie",
                order_remark=order_remark or None,
                en1090_required=en1090_required,
            )

            pdf_requested = f"{doc_type}{num_part}_{prod}_{today}.pdf"
            pdf_filename = _fit_filename_within_path(prod_folder, pdf_requested)
            _record_path_warning(
                prod_folder,
                pdf_requested,
                pdf_filename,
                context=f"Productie '{prod}' – {doc_type}",
            )
            pdf_path = os.path.join(prod_folder, pdf_filename)
            try:
                generate_pdf_order_platypus(
                    pdf_path,
                    company,
                    supplier_for_docs,
                    prod,
                    items,
                    doc_type=doc_type,
                    doc_number=doc_num or None,
                    footer_note=footer_note_text,
                    delivery=delivery_for_docs,
                    project_number=project_number,
                    project_name=project_name,
                    label_kind="productie",
                    order_remark=order_remark or None,
                    en1090_required=en1090_required,
                )
            except Exception as e:
                print(f"[WAARSCHUWING] PDF mislukt voor {prod}: {e}", file=sys.stderr)

        opticutter_order_items: List[Dict[str, object]] = []
        opticutter_total_weight: float | None = None
        if opticutter_prod is not None and opticutter_prod.selections:
            comp = opticutter_comp
            if comp is None:
                stats_map = opticutter_stats_map or {}
                comp = _compute_opticutter_order_exports(opticutter_prod, stats_map)

            settings_rows = [
                {"Parameter": "Exportdatum", "Waarde": today},
                {
                    "Parameter": "Zaagbreedte (kerf) [mm]",
                    "Waarde": opticutter_context.kerf_mm if opticutter_context else None,
                },
            ]
            if opticutter_context and opticutter_context.custom_stock_mm is not None:
                settings_rows.append(
                    {
                        "Parameter": "Aangepaste staaflengte (mm)",
                        "Waarde": opticutter_context.custom_stock_mm,
                    }
                )
            if comp.total_weight_kg is not None:
                settings_rows.append(
                    {
                        "Parameter": "Totaal brutogewicht (kg)",
                        "Waarde": round(comp.total_weight_kg, 2),
                    }
                )
            settings_rows.append({"Parameter": "Productie", "Waarde": prod})

            scenario_df = pd.DataFrame(comp.scenario_rows)
            pieces_df = pd.DataFrame(comp.piece_rows)
            settings_df = pd.DataFrame(settings_rows)
            order_df = pd.DataFrame(comp.order_rows)

            opticutter_requested = f"Opticutter_{prod}_{today}.xlsx"
            opticutter_filename = _fit_filename_within_path(
                prod_folder, opticutter_requested
            )
            _record_path_warning(
                prod_folder,
                opticutter_requested,
                opticutter_filename,
                context=f"Productie '{prod}' – Opticutter",
            )
            opticutter_path = os.path.join(prod_folder, opticutter_filename)
            with pd.ExcelWriter(opticutter_path) as writer:
                scenario_df.to_excel(writer, sheet_name="Scenario", index=False)
                pieces_df.to_excel(writer, sheet_name="Stukken", index=False)
                settings_df.to_excel(writer, sheet_name="Instellingen", index=False)
                if not order_df.empty:
                    order_df.to_excel(writer, sheet_name="Bestelling", index=False)

            order_overview_path: str | None = None
            should_write_order_overview = not order_df.empty
            if should_write_order_overview:
                order_requested = f"Bestelbon_brutemateriaal_{prod}_{today}.xlsx"
                order_filename = _fit_filename_within_path(
                    prod_folder, order_requested
                )
                _record_path_warning(
                    prod_folder,
                    order_requested,
                    order_filename,
                    context=f"Productie '{prod}' – Brutebestelling",
                )
                order_overview_path = os.path.join(prod_folder, order_filename)

            opticutter_order_items = list(comp.raw_items)
            opticutter_total_weight = comp.total_weight_kg

        if opticutter_prod is not None and opticutter_prod.selections:
            opticutter_sel_key = make_opticutter_selection_key(prod)
            opticutter_supplier = pick_supplier_for_opticutter(
                prod, db, opticutter_override_map, suppliers_sorted=suppliers_sorted
            )
            chosen[opticutter_sel_key] = opticutter_supplier.supplier
            if remember_defaults and opticutter_supplier.supplier not in (
                "",
                "Onbekend",
                NO_SUPPLIER_PLACEHOLDER,
            ):
                db.set_default(
                    make_opticutter_default_key(prod), opticutter_supplier.supplier
                )

            opticutter_doc_type_raw = opticutter_doc_type_map.get(prod, "Bestelbon")
            opticutter_doc_type = (
                _to_str(opticutter_doc_type_raw).strip() or "Bestelbon"
            )
            opticutter_doc_num = _normalize_doc_number(
                opticutter_doc_num_map.get(prod, ""), opticutter_doc_type
            )
            opticutter_prefix = _prefix_for_doc_type(opticutter_doc_type)
            if (
                opticutter_doc_num
                and opticutter_prefix
                and opticutter_doc_num.upper() == opticutter_prefix.upper()
            ):
                opticutter_doc_num = ""
            opticutter_doc_token = (
                _sanitize_component(opticutter_doc_num) if opticutter_doc_num else ""
            )
            opticutter_num_part = (
                f"_{opticutter_doc_token}" if opticutter_doc_token else ""
            )
            opticutter_doc_lower = opticutter_doc_type.lower()
            opticutter_is_standaard = opticutter_doc_lower.startswith("standaard")

            opticutter_delivery = opticutter_delivery_map.get(prod)
            opticutter_remark_text = opticutter_remarks_map.get(prod, "")
            if opticutter_total_weight is not None:
                weight_line = f"Totaal brutogewicht: {opticutter_total_weight:.2f} kg"
                if opticutter_remark_text:
                    if weight_line not in opticutter_remark_text:
                        opticutter_remark_text = (
                            f"{opticutter_remark_text}\n{weight_line}"
                        )
                else:
                    opticutter_remark_text = weight_line

            opticutter_supplier_name = _to_str(opticutter_supplier.supplier).strip()
            supplier_for_opticutter_docs: Supplier | None = opticutter_supplier
            delivery_for_opticutter_docs = opticutter_delivery
            if opticutter_is_standaard and not opticutter_supplier_name:
                supplier_for_opticutter_docs = None
                delivery_for_opticutter_docs = None

            should_generate_opticutter_order = (
                bool(opticutter_order_items)
                and (opticutter_supplier_name or opticutter_is_standaard)
            )
            if should_generate_opticutter_order:
                should_write_order_overview = False

                opticutter_excel_requested = (
                    f"{opticutter_doc_type}{opticutter_num_part}_"
                    f"{prod}_Brutemateriaal_{today}.xlsx"
                )
                opticutter_excel_filename = _fit_filename_within_path(
                    prod_folder, opticutter_excel_requested
                )
                _record_path_warning(
                    prod_folder,
                    opticutter_excel_requested,
                    opticutter_excel_filename,
                    context=f"Productie '{prod}' – Brutemateriaal {opticutter_doc_type}",
                )
                opticutter_excel_path = os.path.join(
                    prod_folder, opticutter_excel_filename
                )
                opticutter_en1090 = should_require_en1090(prod, en1090_overrides)

                write_order_excel(
                    opticutter_excel_path,
                    opticutter_order_items,
                    company,
                    supplier_for_opticutter_docs,
                    delivery_for_opticutter_docs,
                    opticutter_doc_type,
                    opticutter_doc_num or None,
                    project_number=project_number,
                    project_name=project_name,
                    context_label=prod,
                    context_kind="Brutemateriaal",
                    order_remark=opticutter_remark_text or None,
                    total_weight_kg=opticutter_total_weight,
                    en1090_required=opticutter_en1090,
                )

                opticutter_pdf_requested = (
                    f"{opticutter_doc_type}{opticutter_num_part}_"
                    f"{prod}_Brutemateriaal_{today}.pdf"
                )
                opticutter_pdf_filename = _fit_filename_within_path(
                    prod_folder, opticutter_pdf_requested
                )
                _record_path_warning(
                    prod_folder,
                    opticutter_pdf_requested,
                    opticutter_pdf_filename,
                    context=f"Productie '{prod}' – Brutemateriaal {opticutter_doc_type}",
                )
                opticutter_pdf_path = os.path.join(prod_folder, opticutter_pdf_filename)
                try:
                    generate_pdf_order_platypus(
                        opticutter_pdf_path,
                        company,
                        supplier_for_opticutter_docs,
                        prod,
                        opticutter_order_items,
                        doc_type=opticutter_doc_type,
                        doc_number=opticutter_doc_num or None,
                        footer_note=footer_note_text,
                        delivery=delivery_for_opticutter_docs,
                        project_number=project_number,
                        project_name=project_name,
                        label_kind="brutemateriaal",
                        order_remark=opticutter_remark_text or None,
                        total_weight_kg=opticutter_total_weight,
                        en1090_required=opticutter_en1090,
                    )
                except Exception as exc:
                    print(
                        f"[WAARSCHUWING] PDF brutemateriaal mislukt voor {prod}: {exc}",
                        file=sys.stderr,
                    )

            if should_write_order_overview and order_overview_path:
                order_df.to_excel(order_overview_path, index=False)

        packlist_items = step_entries.get(prod, [])
        if packlist_items and REPORTLAB_OK:
            try:
                with tempfile.TemporaryDirectory(prefix="previews_", dir=prod_folder) as preview_dir:
                    rendered_previews = step_previews.render_step_files(
                        packlist_items, preview_dir
                    )
                    if rendered_previews:
                        packlist_requested = f"Paklijst_{prod}_{today}.pdf"
                        packlist_filename = _fit_filename_within_path(
                            prod_folder, packlist_requested
                        )
                        _record_path_warning(
                            prod_folder,
                            packlist_requested,
                            packlist_filename,
                            context=f"Productie '{prod}' – Paklijst",
                        )
                        packlist_path = os.path.join(prod_folder, packlist_filename)
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

    if copy_finish_exports and finish_groups:
        finish_seen: Dict[str, set[tuple[str, str]]] = defaultdict(set)
        for finish_key, info in finish_groups.items():
            part_numbers = info.get("part_numbers") or set()
            if not part_numbers:
                continue
            folder_name = info.get("folder_name", finish_key)
            target_dir = os.path.join(dest, folder_name)
            os.makedirs(target_dir, exist_ok=True)
            seen_pairs = finish_seen[finish_key]
            zf = None
            if zip_finish_exports:
                zip_name = _fit_filename_within_path(
                    target_dir, f"{folder_name}.zip"
                )
                zip_path = os.path.join(target_dir, zip_name)
                try:
                    zf = zipfile.ZipFile(
                        zip_path,
                        "w",
                        compression=zipfile.ZIP_DEFLATED,
                        compresslevel=6,
                    )
                except TypeError:
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
            for pn in sorted(part_numbers):
                files = file_index.get(pn, [])
                for src_file in files:
                    transformed = _transform_export_name(os.path.basename(src_file))
                    combo = (src_file, transformed)
                    if combo in seen_pairs:
                        continue
                    seen_pairs.add(combo)
                    ext = os.path.splitext(src_file)[1].lower()
                    if selected_exts_set and ext not in selected_exts_set:
                        continue
                    if zip_finish_exports:
                        if zf is not None:
                            zf.write(src_file, arcname=transformed)
                    else:
                        shutil.copy2(src_file, os.path.join(target_dir, transformed))
            if zf is not None:
                zf.close()

    if finish_groups:
        for finish_key, info in sorted(
            finish_groups.items(), key=lambda item: _to_str(item[1].get("label", "")).lower()
        ):
            rows = list(info.get("rows", []))
            if not rows:
                continue
            supplier = pick_supplier_for_finish(
                finish_key, db, finish_override_map, suppliers_sorted=suppliers_sorted
            )
            chosen[make_finish_selection_key(finish_key)] = supplier.supplier
            if remember_defaults and supplier.supplier not in ("", "Onbekend", NO_SUPPLIER_PLACEHOLDER):
                db.set_default_finish(finish_key, supplier.supplier)

            raw_doc_type = finish_doc_type_map.get(finish_key, "Bestelbon")
            doc_type = _to_str(raw_doc_type).strip() or "Bestelbon"
            doc_type_lower = doc_type.lower()
            is_standaard_doc = doc_type_lower.startswith("standaard")

            supplier_name_clean = _to_str(supplier.supplier).strip()
            if not supplier_name_clean and not is_standaard_doc:
                continue

            doc_num = _normalize_doc_number(
                finish_doc_num_map.get(finish_key, ""), doc_type
            )
            prefix = _prefix_for_doc_type(doc_type)
            if doc_num and prefix and doc_num.upper() == prefix.upper():
                doc_num = ""
            elif doc_num and prefix and not doc_num.upper().startswith(prefix.upper()):
                doc_num = f"{prefix}{doc_num}"
            doc_num_token = _sanitize_component(doc_num) if doc_num else ""
            num_part = f"_{doc_num_token}" if doc_num_token else ""

            folder_name = info.get("folder_name", finish_key)
            target_dir = os.path.join(dest, folder_name)
            os.makedirs(target_dir, exist_ok=True)

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

            label = _to_str(info.get("label")) or finish_key
            filename_component = info.get("filename_component") or finish_key
            delivery = finish_delivery_map.get(finish_key)
            finish_remark = (
                finish_remarks_map.get(finish_key, "") if finish_remarks_map else ""
            ).strip()
            supplier_for_docs: Supplier | None = supplier
            delivery_for_docs = delivery
            if is_standaard_doc and not supplier_name_clean:
                supplier_for_docs = None
                delivery_for_docs = None

            excel_requested = (
                f"{doc_type}{num_part}_{filename_component}_{today}.xlsx"
            )
            excel_filename = _fit_filename_within_path(target_dir, excel_requested)
            _record_path_warning(
                target_dir,
                excel_requested,
                excel_filename,
                context=f"Afwerking '{label}' – {doc_type}",
            )
            excel_path = os.path.join(target_dir, excel_filename)
            write_order_excel(
                excel_path,
                items,
                company,
                supplier_for_docs,
                delivery_for_docs,
                doc_type,
                doc_num or None,
                project_number=project_number,
                project_name=project_name,
                context_label=label,
                context_kind="Afwerking",
                order_remark=finish_remark or None,
            )

            pdf_requested = f"{doc_type}{num_part}_{filename_component}_{today}.pdf"
            pdf_filename = _fit_filename_within_path(target_dir, pdf_requested)
            _record_path_warning(
                target_dir,
                pdf_requested,
                pdf_filename,
                context=f"Afwerking '{label}' – {doc_type}",
            )
            pdf_path = os.path.join(target_dir, pdf_filename)
            try:
                generate_pdf_order_platypus(
                    pdf_path,
                    company,
                    supplier_for_docs,
                    label,
                    items,
                    doc_type=doc_type,
                    doc_number=doc_num or None,
                    footer_note=footer_note_text,
                    delivery=delivery_for_docs,
                    project_number=project_number,
                    project_name=project_name,
                    label_kind="afwerking",
                    order_remark=finish_remark or None,
                )
            except Exception as e:
                print(f"[WAARSCHUWING] PDF mislukt voor {label}: {e}", file=sys.stderr)

    # Persist any (possibly unchanged) supplier defaults so that callers can rely on
    # the database reflecting the latest state on disk.
    if export_bom:
        try:
            bom_filename = make_bom_export_filename(
                bom_source_path, today, _transform_export_name
            )
            _export_bom_workbook(
                bom_df, dest, bom_filename
            )
        except Exception as exc:  # pragma: no cover - unexpected
            raise RuntimeError(f"Kon BOM-export niet opslaan: {exc}") from exc

    if export_bom and export_related_files and bom_source_path:
        for src_file in find_related_bom_exports(bom_source_path, file_index):
            transformed = _transform_export_name(os.path.basename(src_file))
            shutil.copy2(src_file, os.path.join(dest, transformed))
            count_copied += 1

    db.save(SUPPLIERS_DB_FILE)

    return count_copied, chosen


def combine_pdfs_from_source(
    source: str,
    bom_df: pd.DataFrame,
    dest: str,
    date_str: str | None = None,
    *,
    project_number: str | None = None,
    project_name: str | None = None,
    timestamp: datetime.datetime | None = None,
    combine_per_production: bool = True,
    bom_source_path: str | None = None,
) -> CombinedPdfResult:
    """Combine PDF drawing files per production directly from ``source``.

    The BOM dataframe provides ``PartNumber`` to ``Production`` mappings.
    PDFs matching the part numbers are searched in ``source`` using
    :func:`_build_file_index` and merged per production when
    ``combine_per_production`` is :data:`True`. When the flag is :data:`False`,
    every matching PDF in the BOM is merged into a single export file. The
    resulting files are written to a newly created export directory inside
    ``dest`` whose name contains the project number, project name (slugified)
    and an ISO-like timestamp. Output filenames contain either the production
    name or ``BOM`` together with the current date. The returned
    :class:`CombinedPdfResult` provides the number of generated files and the
    absolute output directory path.
    """
    if PdfMerger is None:
        raise ModuleNotFoundError(
            "PyPDF2 must be installed to combine PDF files"
        )

    date_str = date_str or datetime.date.today().strftime("%Y-%m-%d")
    idx = _build_file_index(source, [".pdf"])

    related_bom_pdfs: List[str] = []
    if bom_source_path:
        seen_related: set[str] = set()
        for path in find_related_bom_exports(bom_source_path, idx):
            if not path.lower().endswith(".pdf"):
                continue
            if path in seen_related:
                continue
            related_bom_pdfs.append(path)
            seen_related.add(path)
    related_bom_pdfs.sort(key=lambda x: os.path.basename(x).lower())

    prod_to_files: Dict[str, List[str]] = defaultdict(list)
    for _, row in bom_df.iterrows():
        prod = (row.get("Production") or "").strip() or "_Onbekend"
        pn = str(row.get("PartNumber", ""))
        prod_to_files[prod].extend(idx.get(pn, []))

    out_dir = _create_combined_output_dir(
        dest,
        project_number,
        project_name,
        timestamp=timestamp,
    )
    count = 0

    if combine_per_production:
        for prod, files in prod_to_files.items():
            candidates = list(files)
            if not candidates and not related_bom_pdfs:
                continue
            merger = PdfMerger()
            appended: set[str] = set()
            for path in related_bom_pdfs:
                if path in appended:
                    continue
                merger.append(path)
                appended.add(path)
            for path in sorted(candidates, key=lambda x: os.path.basename(x).lower()):
                if path in appended:
                    continue
                merger.append(path)
                appended.add(path)
            if not appended:
                merger.close()
                continue
            out_name = f"{prod}_{date_str}_combined.pdf"
            safe_name = _fit_filename_within_path(out_dir, out_name)
            merger.write(os.path.join(out_dir, safe_name))
            merger.close()
            count += 1
    else:
        ordered_files: List[str] = []
        seen = set()
        for path in related_bom_pdfs:
            if path not in seen:
                ordered_files.append(path)
                seen.add(path)
        for files in prod_to_files.values():
            for path in files:
                if path not in seen:
                    ordered_files.append(path)
                    seen.add(path)
        if ordered_files:
            merger = PdfMerger()
            ordered_files.sort(key=lambda x: os.path.basename(x).lower())
            if related_bom_pdfs:
                related_sorted = sorted(related_bom_pdfs, key=lambda x: os.path.basename(x).lower())
                related_set = set(related_sorted)
                # Preserve related PDFs at the front by reordering ``ordered_files``.
                rest = [path for path in ordered_files if path not in related_set]
                ordered_files = related_sorted + rest
            for path in ordered_files:
                merger.append(path)
            out_name = f"BOM_{date_str}_combined.pdf"
            safe_name = _fit_filename_within_path(out_dir, out_name)
            merger.write(os.path.join(out_dir, safe_name))
            merger.close()
            count = 1

    return CombinedPdfResult(count=count, output_dir=out_dir)


def combine_pdfs_per_production(
    dest: str,
    date_str: str | None = None,
    *,
    project_number: str | None = None,
    project_name: str | None = None,
    timestamp: datetime.datetime | None = None,
) -> CombinedPdfResult:
    """Combine PDF drawing files per production folder into single PDFs.

    The resulting files are written to a newly created export directory inside
    ``dest`` whose name contains the project number, project name (slugified)
    and an ISO-like timestamp. Output filenames contain the production name
    and current date. The returned :class:`CombinedPdfResult` provides the
    number of generated files and the absolute output directory path.
    """
    if PdfMerger is None:
        raise ModuleNotFoundError(
            "PyPDF2 must be installed to combine PDF files"
        )

    date_str = date_str or datetime.date.today().strftime("%Y-%m-%d")
    out_dir = _create_combined_output_dir(
        dest,
        project_number,
        project_name,
        timestamp=timestamp,
    )
    out_dir_name = os.path.basename(out_dir)
    count = 0
    for prod in sorted(os.listdir(dest)):
        prod_path = os.path.join(dest, prod)
        if not os.path.isdir(prod_path):
            continue
        if prod == out_dir_name or prod.lower().startswith("combined pdf"):
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
                fallback_name = _fit_filename_within_path(prod_path, f"{prod}.zip")
                fallback = os.path.join(prod_path, fallback_name)
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
        safe_name = _fit_filename_within_path(out_dir, out_name)
        merger.write(os.path.join(out_dir, safe_name))
        merger.close()
        count += 1
    return CombinedPdfResult(count=count, output_dir=out_dir)
