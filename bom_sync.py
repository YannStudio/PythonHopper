"""Helpers voor het synchroniseren van custom BOM-gegevens."""

from __future__ import annotations

from typing import Optional, Sequence, Tuple

import pandas as pd

_MAIN_STATUS_COLUMNS: Tuple[str, ...] = ("Bestanden gevonden", "Status", "Link")
_CUSTOM_TO_MAIN_ALIASES = {
    "Material": "Materiaal",
    "QTY.": "Aantal",
    "Surface Area (m²)": "Oppervlakte",
    "Weight (kg)": "Gewicht",
}
MAIN_BOM_COLUMNS: Tuple[str, ...] = (
    "PartNumber",
    "Description",
    "Profile",
    "Length profile",
    "Production",
    "Bestanden gevonden",
    "Status",
    "Materiaal",
    "Supplier",
    "Supplier code",
    "Manufacturer",
    "Manufacturer code",
    "Finish",
    "RAL color",
    "Aantal",
    "Oppervlakte",
    "Gewicht",
    "Link",
)


def _normalize_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_part_numbers(
    df: pd.DataFrame,
    column: str = "PartNumber",
    *,
    uppercase_keys: bool = False,
) -> pd.Series:
    normalized = df[column].map(_normalize_cell)
    if uppercase_keys:
        normalized = normalized.str.upper()
    return normalized


def _numeric_qty(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce").fillna(1).astype(int)
    return numeric.clip(lower=1, upper=999)


def prepare_custom_bom_for_main(
    custom_df: pd.DataFrame,
    existing_main: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """Zet custom BOM-data om naar het hoofd-BOM formaat.

    Parameters
    ----------
    custom_df:
        DataFrame afkomstig uit de Custom BOM-tab.
    existing_main:
        Optioneel: huidige hoofd-BOM om statuskolommen te behouden.

    Returns
    -------
    pandas.DataFrame
        Klaargemaakte DataFrame die direct gebruikt kan worden in de hoofd-BOM.

    Raises
    ------
    ValueError
        Wanneer er geen (geldige) kolom ``PartNumber`` beschikbaar is.
    """

    if custom_df is None or custom_df.empty:
        return pd.DataFrame(columns=MAIN_BOM_COLUMNS)

    if "PartNumber" not in custom_df.columns:
        raise ValueError("De Custom BOM bevat geen kolom 'PartNumber'.")

    normalized = custom_df.copy(deep=True)
    for column in normalized.columns:
        normalized[column] = normalized[column].map(_normalize_cell)

    part_values = _normalize_part_numbers(normalized)
    part_keys = part_values.str.upper()
    valid_mask = part_keys != ""
    normalized = normalized.loc[valid_mask].copy()
    if normalized.empty:
        raise ValueError("Er zijn geen rijen met een ingevulde 'PartNumber'.")
    normalized["PartNumber"] = part_values.loc[valid_mask]
    part_index = part_keys.loc[valid_mask]

    for source, target in _CUSTOM_TO_MAIN_ALIASES.items():
        if source not in normalized.columns:
            continue
        source_values = normalized[source]
        if target in normalized.columns:
            target_values = normalized[target]
            normalized[target] = target_values.where(
                target_values != "", source_values
            )
        else:
            normalized[target] = source_values
        normalized.drop(columns=[source], inplace=True)

    if "Aantal" in normalized.columns:
        normalized["Aantal"] = _numeric_qty(normalized["Aantal"])

    for status_col in _MAIN_STATUS_COLUMNS:
        if status_col not in normalized.columns:
            normalized[status_col] = ""

    if (
        existing_main is not None
        and not existing_main.empty
        and "PartNumber" in existing_main.columns
    ):
        existing_valid = existing_main.loc[existing_main["PartNumber"].notna()].copy()
        if not existing_valid.empty:
            existing_valid["__normalized_part__"] = _normalize_part_numbers(
                existing_valid, "PartNumber", uppercase_keys=True
            )
            existing_valid = existing_valid[existing_valid["__normalized_part__"] != ""]
            if not existing_valid.empty:
                deduped = existing_valid.drop_duplicates(
                    subset="__normalized_part__", keep="last"
                )
                existing_index = deduped.set_index("__normalized_part__")
                for status_col in _MAIN_STATUS_COLUMNS:
                    if status_col in existing_index.columns:
                        normalized[status_col] = part_index.map(
                            existing_index[status_col]
                        ).fillna("")

    result = normalized.reindex(columns=MAIN_BOM_COLUMNS, fill_value="")
    result.reset_index(drop=True, inplace=True)
    return result


__all__: Sequence[str] = [
    "MAIN_BOM_COLUMNS",
    "prepare_custom_bom_for_main",
]
