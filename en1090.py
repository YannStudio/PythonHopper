"""Helpers for EN 1090 order requirements."""

from __future__ import annotations

import re
import unicodedata
from typing import Mapping


EN1090_NOTE_TEXT = (
    "Bestelling conform EN 1090-2:2018+A1:2024.\n"
    "Levering met materiaalcertificaat type 2.2 of 3.1 volgens EN 10204.\n"
    "Profielen in constructiestaal S235JR/S355J2 conform EN 10025-2 met toleranties volgens EN 10034."
)


def normalize_en1090_key(name: str) -> str:
    """Normalize a production name for EN 1090 lookups."""

    text = unicodedata.normalize("NFKD", str(name or ""))
    ascii_text = text.encode("ascii", "ignore").decode("ascii")
    ascii_text = ascii_text.lower()
    ascii_text = re.sub(r"[^0-9a-z]+", " ", ascii_text)
    ascii_text = re.sub(r"\s+", " ", ascii_text).strip()
    return ascii_text


_DEFAULT_ENABLED = {
    "lasercutting",
    "sheetmetal",
    "cutting",
    "tube laser",
    "tubelaser",
    "tube_laser",
    "milling",
}


def default_en1090_enabled(name: str) -> bool:
    """Return whether EN 1090 should be enabled for a production by default."""

    norm = normalize_en1090_key(name)
    if not norm:
        return False
    if norm in _DEFAULT_ENABLED:
        return True
    # Allow hyphen/space variations for tube laser.
    if norm.replace(" ", "") in _DEFAULT_ENABLED:
        return True
    return False


def should_require_en1090(
    name: str, overrides: Mapping[str, bool] | None = None
) -> bool:
    """Resolve the EN 1090 requirement for the provided production name."""

    norm = normalize_en1090_key(name)
    if overrides:
        for key, value in overrides.items():
            if normalize_en1090_key(key) == norm:
                return bool(value)
        if norm in overrides:
            return bool(overrides[norm])
    return default_en1090_enabled(name)

