"""Utilities for rendering STEP/STP previews as PNG thumbnails."""

from __future__ import annotations

import os
import shlex
import shutil
import subprocess
from collections import defaultdict
from typing import List, Sequence, Tuple

DEFAULT_SIZE = (300, 300)
CLI_ENV_VAR = "STEP_PREVIEW_CLI"


def _pythonocc_available() -> bool:
    """Return ``True`` when ``pythonocc-core`` can be imported."""

    try:
        __import__("OCC.Core.STEPControl")
        __import__("OCC.Display.SimpleGui")
    except Exception:
        return False
    return True


def _cli_command_template() -> str:
    """Return the CLI command template for rendering STEP previews."""

    return os.environ.get(CLI_ENV_VAR, "").strip()


def is_renderer_available() -> bool:
    """Return ``True`` when a rendering backend looks available."""

    if _pythonocc_available():
        return True
    template = _cli_command_template()
    if not template:
        return False
    try:
        first = shlex.split(template)[0]
    except ValueError:
        return False
    return bool(shutil.which(first) or first)


def _render_with_pythonocc(step_path: str, out_path: str, size: Tuple[int, int]) -> bool:
    try:
        from OCC.Display.SimpleGui import init_display  # type: ignore
        from OCC.Core.STEPControl import STEPControl_Reader  # type: ignore
        from OCC.Core.IFSelect import IFSelect_RetDone  # type: ignore
        from OCC.Core.Quantity import Quantity_TOC_RGB  # type: ignore
    except Exception as exc:  # pragma: no cover - import failure handled by caller
        raise RuntimeError("pythonocc-core niet beschikbaar") from exc

    display, _, _, _ = init_display()
    reader = STEPControl_Reader()
    status = reader.ReadFile(step_path)
    if status != IFSelect_RetDone:
        return False
    reader.TransferRoots()
    shape = reader.OneShape()
    display.DisplayShape(shape, update=True)
    display.View.SetBackgroundColor(Quantity_TOC_RGB, 1.0, 1.0, 1.0)
    display.View.MustBeResized()
    width, height = size
    display.Resize(width, height)
    display.FitAll()
    display.View.Dump(out_path)
    display.EraseAll()
    return os.path.exists(out_path)


def _render_with_cli(step_path: str, out_path: str, size: Tuple[int, int]) -> bool:
    template = _cli_command_template()
    if not template:
        return False
    width, height = size
    try:
        command_str = template.format(
            input=step_path,
            output=out_path,
            width=str(width),
            height=str(height),
        )
    except Exception:
        return False
    try:
        args = shlex.split(command_str)
    except ValueError:
        return False
    if not args:
        return False
    try:
        subprocess.run(
            args,
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except (FileNotFoundError, subprocess.CalledProcessError):
        return False
    return os.path.exists(out_path)


def render_step_thumbnail(
    step_path: str,
    out_path: str,
    size: Tuple[int, int] = DEFAULT_SIZE,
) -> bool:
    """Render a single STEP/STP file to ``out_path``.

    Returns ``True`` when the thumbnail was created.
    """

    step_path = os.path.abspath(step_path)
    if not os.path.isfile(step_path):
        return False
    out_dir = os.path.dirname(out_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    if _pythonocc_available():
        try:
            if _render_with_pythonocc(step_path, out_path, size):
                return True
        except Exception:
            return False
    if _render_with_cli(step_path, out_path, size):
        return True
    return False


def render_step_files(
    labelled_paths: Sequence[Tuple[str, str]],
    output_dir: str,
    size: Tuple[int, int] = DEFAULT_SIZE,
) -> List[dict]:
    """Render multiple STEP files and return metadata for created previews."""

    if not labelled_paths or not is_renderer_available():
        return []
    os.makedirs(output_dir, exist_ok=True)
    results: List[dict] = []
    counts: defaultdict[str, int] = defaultdict(int)
    for idx, (label, step_path) in enumerate(labelled_paths):
        step_path = os.path.abspath(step_path)
        if not os.path.isfile(step_path):
            continue
        base = os.path.splitext(os.path.basename(step_path))[0] or f"preview_{idx}"
        counts[base] += 1
        suffix = "" if counts[base] == 1 else f"_{counts[base]}"
        out_name = f"{base}{suffix}.png"
        out_path = os.path.join(output_dir, out_name)
        if render_step_thumbnail(step_path, out_path, size=size):
            results.append(
                {
                    "label": label,
                    "source": step_path,
                    "thumbnail": out_path,
                }
            )
    return results
