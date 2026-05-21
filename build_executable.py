"""Build standalone executables for Filehopper using PyInstaller."""

from __future__ import annotations

import argparse
import json
import os
import platform
import re
import subprocess
import sys
import textwrap
from pathlib import Path
from typing import Iterable, List

from app_settings import AppSettings
from app_paths import APP_NAME, APP_VERSION

PROJECT_ROOT = Path(__file__).resolve().parent
BUILD_DIR = PROJECT_ROOT / "build" / "pyinstaller"
DIST_DIR = PROJECT_ROOT / "dist"
RELEASES_DIR = PROJECT_ROOT / "releases"
SPEC_DIR = BUILD_DIR / "specs"
VERSION_FILE_DIR = BUILD_DIR / "version-files"
DATA_FILE_DIR = BUILD_DIR / "data-files"

DEFAULT_DATA_FILES = [
    "clients_db.json",
    "suppliers_db.json",
    "delivery_addresses_db.json",
    "app_settings.json",
    "order_presets.json",
    "suppliers_template.csv",
]


class BuildError(RuntimeError):
    """Raised when the PyInstaller invocation fails."""


def _windows_version_tuple(version: str) -> tuple[int, int, int, int]:
    """Convert an app version like ``3.0`` to a Windows four-part tuple."""

    parts = [int(part) for part in re.findall(r"\d+", version)]
    while len(parts) < 4:
        parts.append(0)
    return tuple(parts[:4])


def _render_windows_version_file(exe_name: str) -> str:
    """Return a PyInstaller-compatible Windows version resource file."""

    version_parts = _windows_version_tuple(APP_VERSION)
    version_text = ".".join(str(part) for part in version_parts)
    return textwrap.dedent(
        f"""\
        # UTF-8
        VSVersionInfo(
          ffi=FixedFileInfo(
            filevers={version_parts},
            prodvers={version_parts},
            mask=0x3f,
            flags=0x0,
            OS=0x40004,
            fileType=0x1,
            subtype=0x0,
            date=(0, 0)
          ),
          kids=[
            StringFileInfo(
              [
                StringTable(
                  '040904B0',
                  [
                    StringStruct('CompanyName', '{APP_NAME}'),
                    StringStruct('FileDescription', '{APP_NAME}'),
                    StringStruct('FileVersion', '{version_text}'),
                    StringStruct('InternalName', '{exe_name}'),
                    StringStruct('OriginalFilename', '{exe_name}.exe'),
                    StringStruct('ProductName', '{APP_NAME}'),
                    StringStruct('ProductVersion', '{version_text}')
                  ]
                )
              ]
            ),
            VarFileInfo([VarStruct('Translation', [1033, 1200])])
          ]
        )
        """
    )


def _prepare_windows_version_file(exe_name: str) -> Path:
    """Write and return the version resource text file for a Windows build."""

    VERSION_FILE_DIR.mkdir(parents=True, exist_ok=True)
    version_file = VERSION_FILE_DIR / f"{exe_name}.version.txt"
    version_file.write_text(_render_windows_version_file(exe_name), encoding="utf-8")
    return version_file


def _prepare_build_data_file(filename: str) -> Path:
    """Return the source path that should be bundled for ``filename``."""

    if filename == "app_settings.json":
        DATA_FILE_DIR.mkdir(parents=True, exist_ok=True)
        clean_settings_path = DATA_FILE_DIR / filename
        clean_settings = AppSettings().to_dict()
        with open(clean_settings_path, "w", encoding="utf-8") as fh:
            json.dump(clean_settings, fh, indent=2, ensure_ascii=False)
        return clean_settings_path
    return PROJECT_ROOT / filename


def _pyinstaller_cmd(
    entry: str,
    name: str,
    *,
    windowed: bool,
    onefile: bool,
    data_files: Iterable[str],
    dist_dir: Path = DIST_DIR,
) -> List[str]:
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
        str(dist_dir),
        "--workpath",
        str(BUILD_DIR / "work"),
        "--specpath",
        str(SPEC_DIR),
        "--collect-submodules",
        "pandastable",
    ]
    cmd.append("--windowed" if windowed else "--console")
    if onefile:
        cmd.append("--onefile")

    if platform.system().lower().startswith("windows"):
        cmd.extend(["--version-file", str(_prepare_windows_version_file(name))])

    for filename in data_files:
        src = _prepare_build_data_file(filename)
        if not src.exists():
            continue
        cmd.extend(["--add-data", f"{src}{os.pathsep}."])

    return cmd


def run_pyinstaller(
    entry: str,
    name: str,
    *,
    windowed: bool,
    onefile: bool,
    data_files: Iterable[str],
    dist_dir: Path = DIST_DIR,
) -> None:
    cmd = _pyinstaller_cmd(
        entry,
        name,
        windowed=windowed,
        onefile=onefile,
        data_files=data_files,
        dist_dir=dist_dir,
    )
    print("Running", " ".join(cmd))
    result = subprocess.run(cmd, check=False)
    if result.returncode != 0:
        raise BuildError(f"PyInstaller failed with exit code {result.returncode}")


def build_targets(
    targets: Iterable[str],
    *,
    include_gui: bool,
    include_cli: bool,
    onefile: bool,
    data_files: Iterable[str],
    dist_dir: Path = DIST_DIR,
) -> None:
    data_files = list(data_files)
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    dist_dir.mkdir(parents=True, exist_ok=True)
    SPEC_DIR.mkdir(parents=True, exist_ok=True)

    for target in targets:
        if include_cli:
            run_pyinstaller(
                "main.py",
                f"filehopper-{target}",
                windowed=False,
                onefile=onefile,
                data_files=data_files,
                dist_dir=dist_dir,
            )
        if include_gui:
            run_pyinstaller(
                "main.py",
                f"filehopper-gui-{target}",
                windowed=True,
                onefile=onefile,
                data_files=data_files,
                dist_dir=dist_dir,
            )


def release_dist_dir(version: str = APP_VERSION) -> Path:
    """Return the standard release output directory for an app version."""

    return RELEASES_DIR / f"{APP_NAME}-{version}"


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
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="Build a single-file executable instead of a portable folder",
    )
    parser.add_argument(
        "--release",
        action="store_true",
        help="Write build output to releases/Filehopper-<version>/",
    )
    parser.add_argument(
        "--dist-dir",
        type=Path,
        help="Custom build output directory (overrides --release)",
    )

    args = parser.parse_args(argv)

    targets = args.targets or [_detect_target()]
    data_files = DEFAULT_DATA_FILES + (args.data_files or [])

    include_gui = not args.only_cli
    include_cli = not args.only_gui
    dist_dir = args.dist_dir or (release_dist_dir() if args.release else DIST_DIR)

    if not include_gui and not include_cli:
        parser.error("At least one of GUI or CLI builds must be enabled")

    try:
        build_targets(
            targets,
            include_gui=include_gui,
            include_cli=include_cli,
            onefile=args.onefile,
            data_files=data_files,
            dist_dir=dist_dir,
        )
    except BuildError as exc:
        print(exc, file=sys.stderr)
        return 1
    print(f"Build output: {dist_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
