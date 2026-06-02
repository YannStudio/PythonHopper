"""Prepare a Filehopper release by updating version metadata."""

from __future__ import annotations

import argparse
import datetime as _dt
import re
import subprocess
import sys
from pathlib import Path

try:
    from .release_notes_generator import generate_changelog_entry
except ImportError:  # pragma: no cover - supports running as a script
    from release_notes_generator import generate_changelog_entry

PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_PATHS_FILE = PROJECT_ROOT / "app_paths.py"
CHANGELOG_FILE = PROJECT_ROOT / "CHANGELOG.md"
VERSION_FILE_DIR = PROJECT_ROOT / "build" / "pyinstaller" / "version-files"
VERSION_FILES = [
    VERSION_FILE_DIR / "filehopper-windows.version.txt",
    VERSION_FILE_DIR / "filehopper-gui-windows.version.txt",
]


def windows_version_text(version: str) -> str:
    parts = [int(part) for part in re.findall(r"\d+", version)]
    if not parts:
        raise ValueError("Version must contain at least one number")
    while len(parts) < 4:
        parts.append(0)
    return ".".join(str(part) for part in parts[:4])


def validate_version(version: str) -> str:
    clean = version.strip().lstrip("v")
    if not re.fullmatch(r"\d+(?:\.\d+){1,3}", clean):
        raise ValueError("Use a version like 3.1, 3.1.1 or 4.0.0")
    return clean


def update_app_version(version: str) -> None:
    text = APP_PATHS_FILE.read_text(encoding="utf-8")
    updated = re.sub(
        r'APP_VERSION\s*=\s*"[^"]+"',
        f'APP_VERSION = "{version}"',
        text,
        count=1,
    )
    if updated == text:
        raise RuntimeError("APP_VERSION not found in app_paths.py")
    APP_PATHS_FILE.write_text(updated, encoding="utf-8")


def update_windows_version_files(version: str) -> None:
    file_version = windows_version_text(version)
    fixed_version = ", ".join(file_version.split("."))
    for path in VERSION_FILES:
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8")
        text = re.sub(
            r"filevers=\([^)]+\)",
            f"filevers=({fixed_version})",
            text,
        )
        text = re.sub(
            r"prodvers=\([^)]+\)",
            f"prodvers=({fixed_version})",
            text,
        )
        text = re.sub(
            r"StringStruct\('FileVersion', '[^']+'\)",
            f"StringStruct('FileVersion', '{file_version}')",
            text,
        )
        text = re.sub(
            r"StringStruct\('ProductVersion', '[^']+'\)",
            f"StringStruct('ProductVersion', '{file_version}')",
            text,
        )
        path.write_text(text, encoding="utf-8")


def update_changelog(version: str, date: _dt.date) -> None:
    date_str = date.isoformat()
    entry = generate_changelog_entry(version, date_str)
    
    if CHANGELOG_FILE.exists():
        text = CHANGELOG_FILE.read_text(encoding="utf-8")
        if re.search(rf"^##\s+{re.escape(version)}\b", text, flags=re.MULTILINE):
            return
        if text.startswith("# Changelog\n"):
            text = text.replace("# Changelog\n", f"# Changelog\n\n{entry}", 1)
        else:
            text = f"# Changelog\n\n{entry}{text.lstrip()}"
    else:
        text = f"# Changelog\n\n{entry}"
    
    CHANGELOG_FILE.write_text(text, encoding="utf-8")


def run_command(command: list[str]) -> None:
    print("Running", " ".join(command))
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("version", help="New release version, for example 3.1.1")
    parser.add_argument(
        "--no-changelog",
        action="store_true",
        help="Do not create or update CHANGELOG.md",
    )
    parser.add_argument(
        "--preview-notes",
        action="store_true",
        help="Preview release notes before updating CHANGELOG.md",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="Run the test suite after updating version files",
    )
    parser.add_argument(
        "--build",
        action="store_true",
        help="Run build_executable.py --release after updating version files",
    )
    parser.add_argument(
        "--target",
        action="append",
        help="Build target to pass through when --build is used",
    )
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="Pass --onefile to build_executable.py when --build is used",
    )
    args = parser.parse_args(argv)

    try:
        version = validate_version(args.version)
        
        # Preview release notes if requested
        if args.preview_notes and not args.no_changelog:
            print("\n📋 Release Notes Preview:")
            print("─" * 70)
            run_command([sys.executable, "scripts/preview_release_notes.py", version])
            print("─" * 70 + "\n")
        
        print(f"🚀 Preparing release {version}...")
        update_app_version(version)
        print(f"  ✓ Updated APP_VERSION to {version}")
        
        update_windows_version_files(version)
        print(f"  ✓ Updated Windows version files")
        
        if not args.no_changelog:
            update_changelog(version, _dt.date.today())
            print(f"  ✓ Updated CHANGELOG.md with auto-generated release notes")
        
        if args.test:
            print("\n🧪 Running tests...")
            run_command([sys.executable, "-m", "pytest", "tests"])
            print("  ✓ All tests passed")
        
        if args.build:
            print("\n📦 Building executables...")
            build_cmd = [sys.executable, "build_executable.py", "--release"]
            for target in args.target or []:
                build_cmd.extend(["--target", target])
            if args.onefile:
                build_cmd.append("--onefile")
            run_command(build_cmd)
            print("  ✓ Build complete")
        
        print(f"\n✅ Release {version} prepared successfully!")
        print("\nNext steps:")
        print(f"  1. Review changes: git diff")
        print(f"  2. Commit: git commit -am 'Release {version}'")
        print(f"  3. Tag: git tag v{version}")
        print(f"  4. Push: git push && git push --tags")
        
    except Exception as exc:
        print(f"\n❌ Error: {exc}", file=sys.stderr)
        return 1

    print(f"Release metadata updated for Filehopper {version}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
