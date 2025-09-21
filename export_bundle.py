"""Utilities for creating export bundle directories."""

from __future__ import annotations

import datetime as _dt
import errno
import os
import re
import unicodedata
import warnings
from pathlib import Path
from typing import Iterable, Optional, Union

__all__ = ["create_export_bundle"]


class ExportBundleError(RuntimeError):
    """Raised when creating an export bundle directory fails."""


def _normalize_date(date: Optional[Union[_dt.date, _dt.datetime]]) -> _dt.date:
    if date is None:
        return _dt.date.today()
    if isinstance(date, _dt.datetime):
        return date.date()
    if isinstance(date, _dt.date):
        return date
    raise TypeError(
        "date must be a datetime.date, datetime.datetime, or None, "
        f"received {type(date)!r}",
    )


def _slugify(value: str, fallback: str, *, max_length: int = 40) -> str:
    normalized = unicodedata.normalize("NFKD", value or "")
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = ascii_text.lower()
    ascii_text = ascii_text.replace(" ", "-")
    ascii_text = re.sub(r"[^a-z0-9-]", "", ascii_text)
    ascii_text = re.sub(r"-+", "-", ascii_text).strip("-")
    if max_length > 0:
        ascii_text = ascii_text[:max_length].rstrip("-")
    if not ascii_text:
        fallback_normalized = unicodedata.normalize("NFKD", fallback)
        ascii_text = (
            fallback_normalized.encode("ascii", "ignore").decode("ascii")
            or str(fallback)
        )
        ascii_text = re.sub(r"[^a-zA-Z0-9-]", "", ascii_text)
        ascii_text = ascii_text.lower()[:max_length] or "export"
    return ascii_text


def _iter_letter_suffixes(max_attempts: int) -> Iterable[str]:
    for index in range(max_attempts):
        if index == 0:
            yield ""
            continue
        number = index
        letters = []
        while number > 0:
            number, remainder = divmod(number - 1, 26)
            letters.append(chr(ord("A") + remainder))
        yield "".join(reversed(letters))


def create_export_bundle(
    export_root: Union[str, os.PathLike],
    project_number: Union[str, int],
    project_name: str,
    *,
    date: Optional[Union[_dt.date, _dt.datetime]] = None,
    suffix_mode: str = "letters",
    dry_run: bool = False,
    max_attempts: int = 676,
    create_latest_symlink: bool = False,
) -> Path:
    """Create a timestamped export bundle directory.

    Parameters
    ----------
    export_root:
        Directory where the bundle should be created. The directory must exist.
    project_number:
        Unique identifier for the project; will be included in the bundle name.
    project_name:
        Human readable project title that will be slugified for the bundle name.
    date:
        Date to use as prefix for the bundle name. Accepts ``datetime.date`` or
        ``datetime.datetime`` and defaults to today's date when omitted.
    suffix_mode:
        Strategy for generating suffixes when the base name already exists. Only
        ``"letters"`` is supported.
    dry_run:
        If ``True`` the bundle is not created on disk. The intended path is still
        returned.
    max_attempts:
        Maximum number of name attempts (including the base name) before giving
        up.
    create_latest_symlink:
        When ``True`` a ``latest`` symlink in ``export_root`` will point to the
        newly created bundle. On Windows the operation may fail due to symlink
        restrictions; in that case a warning is emitted instead of raising.

    Returns
    -------
    pathlib.Path
        The path to the created (or intended) bundle directory.

    Raises
    ------
    ExportBundleError
        If the bundle cannot be created or a valid name cannot be found within
        ``max_attempts`` tries.
    TypeError
        If ``date`` has an unsupported type.
    """

    if max_attempts <= 0:
        raise ValueError("max_attempts must be a positive integer")
    if suffix_mode != "letters":
        raise ValueError(f"Unsupported suffix_mode: {suffix_mode!r}")

    export_path = Path(export_root)
    if not export_path.exists():
        raise ExportBundleError(f"Export root does not exist: {export_path}")
    if not export_path.is_dir():
        raise ExportBundleError(f"Export root is not a directory: {export_path}")

    bundle_date = _normalize_date(date)
    date_prefix = bundle_date.strftime("%Y-%m-%d")
    project_number_str = str(project_number).strip() or "project"
    slug = _slugify(project_name, project_number_str)
    base_name = f"{date_prefix}_{project_number_str}_{slug}"

    chosen_path: Optional[Path] = None
    for suffix in _iter_letter_suffixes(max_attempts):
        candidate_name = base_name if not suffix else f"{base_name}_{suffix}"
        candidate_path = export_path / candidate_name
        if candidate_path.exists():
            continue
        if dry_run:
            chosen_path = candidate_path
            break
        try:
            candidate_path.mkdir()
        except FileExistsError:
            continue
        except OSError as exc:  # pragma: no cover - OS dependent
            if exc.errno == errno.ENAMETOOLONG:
                raise ExportBundleError(
                    f"Path too long for export bundle: {candidate_path}"
                ) from exc
            raise ExportBundleError(
                f"Failed to create export bundle at {candidate_path}: {exc}"
            ) from exc
        else:
            chosen_path = candidate_path
            break

    if chosen_path is None:
        raise ExportBundleError(
            "Exhausted all name attempts without finding a free export bundle "
            f"slot starting from {base_name!r} inside {export_path}"
        )

    if not dry_run and create_latest_symlink:
        latest_path = export_path / "latest"
        try:
            if latest_path.exists() or latest_path.is_symlink():
                if latest_path.is_symlink() or latest_path.is_file():
                    latest_path.unlink()
                elif latest_path.is_dir():
                    raise ExportBundleError(
                        f"Cannot update latest symlink: {latest_path} is a directory"
                    )
                else:
                    latest_path.unlink()
            latest_path.symlink_to(chosen_path.name, target_is_directory=True)
        except NotImplementedError as exc:  # pragma: no cover - OS dependent
            warnings.warn(
                f"Symlinks are not supported on this platform; could not update "
                f"latest symlink in {export_path}: {exc}",
                RuntimeWarning,
            )
        except OSError as exc:  # pragma: no cover - OS dependent
            if getattr(exc, "winerror", None) is not None or os.name == "nt":
                warnings.warn(
                    "Unable to update latest symlink due to OS restrictions: "
                    f"{exc}",
                    RuntimeWarning,
                )
            else:
                raise ExportBundleError(
                    f"Failed to update latest symlink in {export_path}: {exc}"
                ) from exc

    return chosen_path


