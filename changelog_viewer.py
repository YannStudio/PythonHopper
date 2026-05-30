"""Changelog viewer for the GUI."""

from pathlib import Path
from typing import Optional


def load_changelog(changelog_path: Optional[Path] = None) -> str:
    """Load and return the changelog content.
    
    Parameters
    ----------
    changelog_path : Path, optional
        Path to CHANGELOG.md file. If None, looks in the project root.
    
    Returns
    -------
    str
        The changelog content, or an error message if not found.
    """
    if changelog_path is None:
        # Try to find CHANGELOG.md from the current module's location
        changelog_path = Path(__file__).resolve().parent / "CHANGELOG.md"
    
    if not changelog_path.exists():
        return "Changelog niet gevonden. CHANGELOG.md bestaat niet."
    
    try:
        return changelog_path.read_text(encoding="utf-8")
    except Exception as e:
        return f"Fout bij lezen changelog: {e}"


def get_latest_release_notes(changelog: str, max_lines: int = 5) -> str:
    """Return the most recent release notes section from the changelog."""
    lines = changelog.splitlines()
    latest = []
    in_section = False
    bullets = 0

    for line in lines:
        if line.startswith("## "):
            if in_section:
                break
            latest.append(line.strip())
            in_section = True
            continue

        if not in_section:
            continue

        if line.strip() == "":
            if bullets:
                break
            continue

        latest.append(line.rstrip())
        if line.strip().startswith("-"):
            bullets += 1
            if bullets >= max_lines:
                break

    if not latest:
        return "Geen release notes gevonden in CHANGELOG.md."

    return "\n".join(latest).strip()


def format_changelog_for_display(changelog: str) -> str:
    """Format changelog for nice display in GUI."""
    return changelog
