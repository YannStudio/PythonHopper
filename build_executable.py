"""Build standalone executables for Filehopper using PyInstaller."""

from __future__ import annotations

import argparse
import os
import platform
import subprocess
import sys
from pathlib import Path
from typing import Iterable, List

PROJECT_ROOT = Path(__file__).resolve().parent
BUILD_DIR = PROJECT_ROOT / "build" / "pyinstaller"
DIST_DIR = PROJECT_ROOT / "dist"
SPEC_DIR = BUILD_DIR / "specs"

DEFAULT_DATA_FILES = [
    "clients_db.json",
    "suppliers_db.json",
    "delivery_addresses_db.json",
    "app_settings.json",
]


class BuildError(RuntimeError):
    """Raised when the PyInstaller invocation fails."""


def _pyinstaller_cmd(entry: str, name: str, *, windowed: bool, data_files: Iterable[str]) -> List[str]:
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        entry,
        "--name",
        name,
        "--noconfirm",
        "--clean",
        "--distpath",
        str(DIST_DIR),
        "--workpath",
        str(BUILD_DIR / "work"),
        "--specpath",
        str(SPEC_DIR),
        "--collect-submodules",
        "pandastable",
    ]
    cmd.append("--windowed" if windowed else "--console")

    for filename in data_files:
        src = PROJECT_ROOT / filename
        if not src.exists():
            continue
        cmd.extend(["--add-data", f"{src}{os.pathsep}."])

    return cmd


def run_pyinstaller(entry: str, name: str, *, windowed: bool, data_files: Iterable[str]) -> None:
    cmd = _pyinstaller_cmd(entry, name, windowed=windowed, data_files=data_files)
    print("Running", " ".join(cmd))
    result = subprocess.run(cmd, check=False)
    if result.returncode != 0:
        raise BuildError(f"PyInstaller failed with exit code {result.returncode}")


def build_targets(targets: Iterable[str], *, include_gui: bool, include_cli: bool, data_files: Iterable[str]) -> None:
    data_files = list(data_files)
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    SPEC_DIR.mkdir(parents=True, exist_ok=True)

    for target in targets:
        if include_cli:
            run_pyinstaller("main.py", f"filehopper-{target}", windowed=False, data_files=data_files)
        if include_gui:
            run_pyinstaller("main.py", f"filehopper-gui-{target}", windowed=True, data_files=data_files)


def _detect_target() -> str:
    system = platform.system().lower()
    if system.startswith("darwin"):
        return "macos"
    if system.startswith("windows"):
        return "windows"
    return system or "linux"


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--target",
        dest="targets",
        action="append",
        help="Operating system target label (default: current platform)",
    )
    parser.add_argument(
        "--only-gui",
        action="store_true",
        help="Only build the windowed GUI executable",
    )
    parser.add_argument(
        "--only-cli",
        action="store_true",
        help="Only build the console/CLI executable",
    )
    parser.add_argument(
        "--data-file",
        dest="data_files",
        action="append",
        help="Additional data file to bundle (relative to the project root)",
    )

    args = parser.parse_args(argv)

    targets = args.targets or [_detect_target()]
    data_files = DEFAULT_DATA_FILES + (args.data_files or [])

    include_gui = not args.only_cli
    include_cli = not args.only_gui

    if not include_gui and not include_cli:
        parser.error("At least one of GUI or CLI builds must be enabled")

    try:
        build_targets(targets, include_gui=include_gui, include_cli=include_cli, data_files=data_files)
    except BuildError as exc:
        print(exc, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
