"""BOM (Bill of Materials) utilities.

This module centralizes functionality for reading BOM spreadsheets from
CSV/Excel files and normalizing their columns.
"""

from __future__ import annotations

import csv
import re
from numbers import Integral, Real
from typing import List, Optional

import pandas as pd


def read_csv_flex(filepath: str) -> pd.DataFrame:
    """Read a CSV file with flexible delimiter and encoding detection.

    The function tries a number of common delimiters and encodings and returns
    the first DataFrame that looks valid (i.e. has more than one column).
    """

    candidates = [",", ";", "\t", "|"]
    encodings = ["utf-8-sig", "utf-8", "latin1", "cp1252"]
    last_exc = None
    for enc in encodings:
        try:
            with open(filepath, "r", encoding=enc, errors="replace") as f:
                sample = f.read(8192)
            try:
                dialect = csv.Sniffer().sniff(sample)
                return pd.read_csv(
                    filepath,
                    sep=dialect.delimiter,
                    encoding=enc,
                    on_bad_lines="skip",
                )
            except Exception:
                for delim in candidates:
                    try:
                        df = pd.read_csv(
                            filepath,
                            sep=delim,
                            encoding=enc,
                            on_bad_lines="skip",
                        )
                        if df.shape[1] > 1:
                            return df
                    except Exception as e:  # pragma: no cover - best effort
                        last_exc = e
        except Exception as e:  # pragma: no cover - best effort
            last_exc = e
    raise last_exc if last_exc else Exception("CSV kon niet worden gelezen.")


def _find_col_by_regex(df: pd.DataFrame, patterns: List[str]) -> Optional[str]:
    """Find the first column whose name matches any of the regex patterns."""

    pats = [re.compile(p, re.IGNORECASE) for p in patterns]
    for c in df.columns:
        for p in pats:
            if p.search(str(c)):
                return c
    return None


def load_bom(path: str) -> pd.DataFrame:
    """Load a BOM spreadsheet and normalize expected columns.

    Returns a DataFrame with canonical column names: PartNumber, Description,
    Production, Bestanden gevonden, Status, Materiaal, Finish, RAL color,
    Aantal, Oppervlakte, Gewicht.
    """

    if path.lower().endswith(".csv"):
        df = read_csv_flex(path)
    else:
        engine = "openpyxl" if path.lower().endswith(".xlsx") else None
        df = pd.read_excel(path, engine=engine)

    def need(colname: str) -> str:
        low = {c.lower(): c for c in df.columns}
        lc = colname.lower()
        if lc not in low:
            raise ValueError(f"BOM mist kolom: {colname}")
        return low[lc]

    pn_c = need("PartNumber")
    try:
        ds_c = need("Description")
    except ValueError:
        ds_c = None
    pr_c = need("Production")

    def find_any(names: List[str]) -> Optional[str]:
        low = {c.lower(): c for c in df.columns}
        for n in names:
            if n.lower() in low:
                return low[n.lower()]
        return None

    aantal_col = find_any(["Aantal", "Qty", "Qty.", "Quantity", "Stuks"])

    opp_col = find_any(
        [
            "Oppervlakte",
            "Oppervlakte (m²)",
            "Oppervlakte (m2)",
            "Oppervlakte m²",
            "Oppervlakte m2",
            "Area",
            "Area (m2)",
            "Area (m²)",
        ]
    )
    if opp_col is None:
        opp_col = _find_col_by_regex(df, [r"\bopp\b", r"oppervl", r"\barea\b"])

    gew_col = find_any(["Gewicht", "Gewicht (kg)", "Weight", "Weight (kg)"])
    if gew_col is None:
        gew_col = _find_col_by_regex(df, [r"^gew", r"\bweight\b"])

    # Materiaal
    mat_col = find_any(["Materiaal", "Material", "Material Type", "Materia", "Mat", "Grade"])
    if mat_col is None:
        mat_col = _find_col_by_regex(df, [r"\bmaterial", r"\bmateriaal", r"\bgrade\b"])

    df["PartNumber"] = df[pn_c].astype(str).str.strip()
    if ds_c is not None:
        df["Description"] = df[ds_c].astype(str).fillna("").str.strip()
    else:
        # Fallback to PartNumber when no description column is present so that
        # downstream code keeps working with a consistent schema.
        df["Description"] = df["PartNumber"]
    df["Production"] = df[pr_c].astype(str).str.strip()

    if aantal_col is None:
        df["Aantal"] = 1
    else:
        df["Aantal"] = pd.to_numeric(df[aantal_col], errors="coerce").fillna(1).astype(int)
    df["Aantal"] = df["Aantal"].clip(lower=1, upper=999)  # max 999

    df["Oppervlakte"] = "" if opp_col is None else df[opp_col]
    df["Gewicht"] = "" if gew_col is None else df[gew_col]
    def _normalize_text_value(value: object) -> str:
        if pd.isna(value):
            return ""
        if isinstance(value, Integral):
            return str(int(value))
        if isinstance(value, Real):
            float_value = float(value)
            if float_value.is_integer():
                return str(int(float_value))
            return str(value).strip()
        return str(value).strip()

    def _text_column(col_name: Optional[str]) -> pd.Series:
        if col_name is None:
            return pd.Series([""] * len(df), index=df.index)
        series = df[col_name]
        return series.apply(_normalize_text_value)

    df["Materiaal"] = _text_column(mat_col)

    supplier_col = find_any(["Supplier"])
    supplier_code_col = find_any(["Supplier code"])
    manufacturer_col = find_any(["Manufacturer"])
    manufacturer_code_col = find_any(["Manufacturer code"])

    df["Supplier"] = _text_column(supplier_col)
    df["Supplier code"] = _text_column(supplier_code_col)
    df["Manufacturer"] = _text_column(manufacturer_col)
    df["Manufacturer code"] = _text_column(manufacturer_code_col)

    finish_col = find_any(["Finish"])
    if finish_col is None:
        finish_col = _find_col_by_regex(df, [r"\bfinish\b"])
    ral_col = find_any(["RAL color", "RAL colour"])
    if ral_col is None:
        ral_col = _find_col_by_regex(df, [r"\bral\s*(?:color|colour)\b", r"\bral\b"])

    df["Finish"] = _text_column(finish_col)
    df["RAL color"] = _text_column(ral_col)

    if "Bestanden gevonden" not in df.columns:
        df["Bestanden gevonden"] = ""
    if "Status" not in df.columns:
        df["Status"] = ""

    return df[
        [
            "PartNumber",
            "Description",
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
        ]
    ].copy()


__all__ = ["read_csv_flex", "load_bom"]

