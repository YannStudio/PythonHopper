from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Mapping, MutableMapping, Optional, Sequence, Tuple

import pandas as pd

from helpers import _to_str


STOCK_LENGTH_MM = 6000
LONG_STOCK_LENGTH_MM = 12000
DEFAULT_KERF_MM = 5.0


@dataclass(slots=True)
class StockScenarioResult:
    bars: int
    waste_mm: float
    waste_pct: float
    dropped_pieces: int
    cuts: int


@dataclass(slots=True)
class OpticutterPieceDetail:
    length_mm: int
    label: str


@dataclass(slots=True)
class OpticutterProfileData:
    key: Tuple[str, str, str]
    profile: str
    material: str
    production: str
    pieces: List[OpticutterPieceDetail]
    lengths_mm: List[int]
    blockers: Dict[str, set[str]]
    scenarios: Dict[str, StockScenarioResult]
    best_choice: str
    manual_length_mm: Optional[int]
    manual_choice_key: Optional[str]

    @property
    def quantity(self) -> int:
        return len(self.lengths_mm)


@dataclass(slots=True)
class OpticutterAnalysis:
    profiles: List[OpticutterProfileData]
    total_quantity: int
    kerf_mm: float
    custom_stock_mm: Optional[int]
    unparsed_lengths: List[str]
    oversized_profiles_6m: set[str]
    oversized_profiles_12m: set[str]
    aggregated_rows: List[Dict[str, Any]]
    error: Optional[str] = None


@dataclass(slots=True)
class OpticutterSelection:
    profile: OpticutterProfileData
    choice: str
    choice_label: str
    stock_length_mm: Optional[int]
    result: Optional[StockScenarioResult]
    blockers: List[str]

    @property
    def is_manual_input(self) -> bool:
        return self.choice == "input"

    @property
    def is_valid(self) -> bool:
        if self.result is None or self.is_manual_input:
            return False
        if self.result.dropped_pieces:
            return False
        return self.result.bars > 0


@dataclass(slots=True)
class OpticutterProductionExport:
    production: str
    selections: List[OpticutterSelection]


@dataclass(slots=True)
class OpticutterExportContext:
    kerf_mm: float
    custom_stock_mm: Optional[int]
    productions: Dict[str, OpticutterProductionExport]


def parse_length_to_mm(value: object) -> Optional[int]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        length = float(value)
    else:
        text = str(value).strip().lower()
        if not text:
            return None
        text = text.replace(",", ".")
        multiplier = 1.0
        if text.endswith("mm"):
            text = text[:-2]
        elif text.endswith("cm"):
            text = text[:-2]
            multiplier = 10.0
        elif text.endswith("m"):
            text = text[:-1]
            multiplier = 1000.0
        try:
            length = float(text) * multiplier
        except ValueError:
            return None
    if length <= 0:
        return None
    return int(round(length))


def _calculate_stock_scenario(
    lengths_mm: Sequence[int],
    stock_length_mm: int,
    kerf_mm: float,
) -> StockScenarioResult:
    usable = [float(length) for length in lengths_mm if length and length > 0]
    if not usable:
        return StockScenarioResult(0, 0.0, 0.0, 0, 0)

    kerf_mm = max(0.0, float(kerf_mm))
    stock_length_mm = int(stock_length_mm)
    if stock_length_mm <= 0:
        return StockScenarioResult(0, 0.0, 0.0, len(usable), 0)

    used_lengths: List[float] = []
    dropped = 0
    cuts = 0
    for length in sorted(usable, reverse=True):
        if length > stock_length_mm:
            dropped += 1
            continue

        placed = False
        for idx, used in enumerate(used_lengths):
            extra_loss = kerf_mm if used > 0 else 0.0
            if used + extra_loss + length <= stock_length_mm + 1e-6:
                used_lengths[idx] = used + extra_loss + length
                cuts += 1
                placed = True
                break
        if not placed:
            used_lengths.append(length)

    bars = len(used_lengths)
    waste_mm = float(bars * stock_length_mm - sum(used_lengths))
    waste_pct = 0.0
    if bars > 0:
        waste_pct = waste_mm / (bars * stock_length_mm) * 100.0
    return StockScenarioResult(bars, waste_mm, waste_pct, dropped, cuts)


def _normalize_production(value: str) -> str:
    text = _to_str(value).strip()
    return text or "_Onbekend"


def analyse_profiles(
    bom_df: pd.DataFrame,
    *,
    kerf_mm: float = DEFAULT_KERF_MM,
    custom_stock_mm: Optional[int] = None,
    manual_lengths: Optional[
        Mapping[Tuple[str, str, str], int]
    ] = None,
) -> OpticutterAnalysis:
    required_columns = ["PartNumber", "Profile", "Length profile", "Aantal"]
    missing_columns = [col for col in required_columns if col not in bom_df.columns]
    if missing_columns:
        return OpticutterAnalysis(
            profiles=[],
            total_quantity=0,
            kerf_mm=kerf_mm,
            custom_stock_mm=custom_stock_mm,
            unparsed_lengths=[],
            oversized_profiles_6m=set(),
            oversized_profiles_12m=set(),
            aggregated_rows=[],
            error="BOM mist profielgegevens.",
        )

    profiles_df = bom_df.loc[:, required_columns].copy()
    profiles_df["Description"] = (
        bom_df.get("Description", "").fillna("").astype(str).str.strip()
    )
    if "Material" in bom_df.columns:
        profiles_df["Material"] = (
            bom_df["Material"].fillna("").astype(str).str.strip()
        )
    elif "Materiaal" in bom_df.columns:
        profiles_df["Material"] = (
            bom_df["Materiaal"].fillna("").astype(str).str.strip()
        )
    else:
        profiles_df["Material"] = ""
    profiles_df["Production"] = (
        bom_df.get("Production", "").fillna("").astype(str).str.strip()
    )
    profiles_df["PartNumber"] = (
        profiles_df["PartNumber"].fillna("").astype(str).str.strip()
    )
    profiles_df["Profile"] = (
        profiles_df["Profile"].fillna("").astype(str).str.strip()
    )
    profiles_df["Length profile"] = (
        profiles_df["Length profile"].fillna("").astype(str).str.strip()
    )
    profiles_df["Aantal"] = (
        pd.to_numeric(profiles_df["Aantal"], errors="coerce").fillna(0).astype(int)
    )

    filtered = profiles_df[profiles_df["Profile"] != ""].copy()
    if filtered.empty:
        return OpticutterAnalysis(
            profiles=[],
            total_quantity=0,
            kerf_mm=kerf_mm,
            custom_stock_mm=custom_stock_mm,
            unparsed_lengths=[],
            oversized_profiles_6m=set(),
            oversized_profiles_12m=set(),
            aggregated_rows=[],
            error="Geen profielen gevonden in de BOM.",
        )

    filtered["Length profile mm"] = filtered["Length profile"].map(parse_length_to_mm)

    manual_lengths = dict(manual_lengths or {})

    pieces_by_profile: Dict[Tuple[str, str, str], List[int]] = defaultdict(list)
    detail_map: Dict[Tuple[str, str, str], List[OpticutterPieceDetail]] = defaultdict(list)
    blockers: Dict[Tuple[str, str, str], Dict[str, set[str]]] = defaultdict(
        lambda: {"6000": set(), "12000": set(), "custom": set()}
    )
    unparsed_lengths: List[str] = []
    oversized_profiles: set[str] = set()
    oversized_profiles_12m: set[str] = set()

    for _, row in filtered.iterrows():
        profile_name = row["Profile"]
        material_name = row.get("Material", "")
        production_name = row.get("Production", "")
        length_mm = row.get("Length profile mm")
        qty = int(row.get("Aantal", 0))
        if qty <= 0:
            continue
        key = (profile_name, material_name, production_name)

        if length_mm is None:
            length_text = row.get("Length profile", "")
            if length_text:
                unparsed_lengths.append(str(length_text))
            continue

        if length_mm > STOCK_LENGTH_MM:
            oversized_profiles.add(profile_name)
        if length_mm > LONG_STOCK_LENGTH_MM:
            oversized_profiles_12m.add(profile_name)

        pieces_by_profile[key].extend([length_mm] * qty)
        description = row.get("Description", "")
        part_number = row.get("PartNumber", "")
        length_label = row.get("Length profile", "") or f"{length_mm} mm"
        part_label = str(part_number) if part_number else "Onbekend part"
        if description:
            if part_number:
                part_label = f"{part_label} - {description}"
            else:
                part_label = description
        blocker_text = f"{part_label} ({length_label})"
        if length_mm > STOCK_LENGTH_MM:
            blockers[key]["6000"].add(blocker_text)
        if length_mm > LONG_STOCK_LENGTH_MM:
            blockers[key]["12000"].add(blocker_text)
        if custom_stock_mm is not None and length_mm > custom_stock_mm:
            blockers[key]["custom"].add(blocker_text)
        detail_map[key].append(OpticutterPieceDetail(length_mm=length_mm, label=blocker_text))

    if not pieces_by_profile:
        return OpticutterAnalysis(
            profiles=[],
            total_quantity=0,
            kerf_mm=kerf_mm,
            custom_stock_mm=custom_stock_mm,
            unparsed_lengths=unparsed_lengths,
            oversized_profiles_6m=oversized_profiles,
            oversized_profiles_12m=oversized_profiles_12m,
            aggregated_rows=[],
            error="Geen profielen gevonden in de BOM.",
        )

    aggregated = (
        filtered.groupby(
            ["PartNumber", "Profile", "Length profile", "Material", "Production"],
            as_index=False,
        )["Aantal"].sum()
        .sort_values(
            by=["Profile", "Material", "Production", "PartNumber", "Length profile"]
        )
    )

    profiles: List[OpticutterProfileData] = []

    sorted_keys = sorted(
        pieces_by_profile.keys(), key=lambda item: (item[0], item[1], item[2])
    )

    for key_tuple in sorted_keys:
        profile_name, material_name, production_name = key_tuple
        lengths = pieces_by_profile.get(key_tuple, [])
        detail_list = detail_map.get(key_tuple, [])

        scenario_6m = _calculate_stock_scenario(lengths, STOCK_LENGTH_MM, kerf_mm)
        scenario_12m = _calculate_stock_scenario(lengths, LONG_STOCK_LENGTH_MM, kerf_mm)
        scenarios: Dict[str, StockScenarioResult] = {
            "6000": scenario_6m,
            "12000": scenario_12m,
        }

        if custom_stock_mm is not None:
            scenarios["custom"] = _calculate_stock_scenario(
                lengths, custom_stock_mm, kerf_mm
            )

        manual_length = manual_lengths.get(key_tuple)
        manual_key: Optional[str] = None
        if manual_length is not None and manual_length > 0:
            scenarios_manual = _calculate_stock_scenario(lengths, manual_length, kerf_mm)
            manual_key = f"manual:{manual_length}"
            scenarios[manual_key] = scenarios_manual
            exceeding_manual = {
                piece.label
                for piece in detail_list
                if piece.length_mm > manual_length
            }
            blockers_entry = blockers[key_tuple]
            if exceeding_manual:
                blockers_entry[manual_key] = exceeding_manual
            else:
                blockers_entry.pop(manual_key, None)

        best_choice: Optional[str] = None
        best_score: Optional[Tuple[int, float]] = None
        for candidate_key, result in scenarios.items():
            if result.dropped_pieces or result.bars <= 0:
                continue
            score = (result.bars, result.waste_pct)
            if best_score is None or score < best_score:
                best_choice = candidate_key
                best_score = score
        if best_choice is None:
            best_choice = "input"

        profiles.append(
            OpticutterProfileData(
                key=key_tuple,
                profile=profile_name,
                material=material_name,
                production=production_name,
                pieces=list(detail_list),
                lengths_mm=list(lengths),
                blockers=blockers[key_tuple],
                scenarios=scenarios,
                best_choice=best_choice,
                manual_length_mm=manual_length,
                manual_choice_key=manual_key,
            )
        )

    total_qty = int(filtered["Aantal"].sum())

    aggregated_rows = (
        aggregated.to_dict(orient="records") if not aggregated.empty else []
    )

    return OpticutterAnalysis(
        profiles=profiles,
        total_quantity=total_qty,
        kerf_mm=kerf_mm,
        custom_stock_mm=custom_stock_mm,
        unparsed_lengths=unparsed_lengths,
        oversized_profiles_6m=oversized_profiles,
        oversized_profiles_12m=oversized_profiles_12m,
        aggregated_rows=aggregated_rows,
        error=None,
    )


def _choice_label(
    choice: str,
    *,
    custom_stock_mm: Optional[int],
    manual_choice_key: Optional[str],
) -> str:
    if choice == "6000":
        return "6000 mm"
    if choice == "12000":
        return "12000 mm"
    if choice == "custom":
        if custom_stock_mm is None:
            return "Custom lengte"
        return f"{custom_stock_mm} mm"
    if choice.startswith("manual:"):
        length = choice.split(":", 1)[1]
        return f"Aangepaste lengte ({length} mm)"
    if choice == "input":
        return "Input lengte – per stuk zagen"
    return choice


def _choice_stock_length(
    choice: str, custom_stock_mm: Optional[int]
) -> Optional[int]:
    if choice == "6000":
        return STOCK_LENGTH_MM
    if choice == "12000":
        return LONG_STOCK_LENGTH_MM
    if choice == "custom":
        return custom_stock_mm
    if choice.startswith("manual:"):
        try:
            return int(choice.split(":", 1)[1])
        except (ValueError, IndexError):
            return None
    return None


def prepare_opticutter_export(
    analysis: OpticutterAnalysis,
    selection_map: Optional[Mapping[Tuple[str, str, str], str]] = None,
) -> OpticutterExportContext:
    selection_map = selection_map or {}
    productions: Dict[str, OpticutterProductionExport] = {}

    for profile in analysis.profiles:
        allowed = set(profile.scenarios.keys()) | {"input"}
        choice = selection_map.get(profile.key, profile.best_choice)
        if choice not in allowed:
            choice = profile.best_choice
        if choice not in allowed:
            choice = "input"
        result = profile.scenarios.get(choice)
        stock_length = _choice_stock_length(choice, analysis.custom_stock_mm)
        blockers = sorted(profile.blockers.get(choice, set()))
        label = _choice_label(
            choice,
            custom_stock_mm=analysis.custom_stock_mm,
            manual_choice_key=profile.manual_choice_key,
        )
        selection = OpticutterSelection(
            profile=profile,
            choice=choice,
            choice_label=label,
            stock_length_mm=stock_length,
            result=result,
            blockers=blockers,
        )
        prod_key = _normalize_production(profile.production)
        if prod_key not in productions:
            productions[prod_key] = OpticutterProductionExport(
                production=prod_key,
                selections=[],
            )
        productions[prod_key].selections.append(selection)

    return OpticutterExportContext(
        kerf_mm=analysis.kerf_mm,
        custom_stock_mm=analysis.custom_stock_mm,
        productions=productions,
    )

