#!/usr/bin/env python3
"""Preview release notes that will be generated from commits."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Add scripts directory to path so we can import release_notes_generator
sys.path.insert(0, str(Path(__file__).parent))

from release_notes_generator import get_last_tag, get_commits_since_tag, generate_release_notes


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Preview release notes that will be generated from Git commits"
    )
    parser.add_argument(
        "version",
        nargs="?",
        default="X.Y.Z",
        help="Version number (for display only, not used for generation)",
    )
    parser.add_argument(
        "--since",
        help="Generate notes for commits since this tag (default: latest tag)",
    )
    args = parser.parse_args(argv)

    tag = args.since or get_last_tag()
    commits = get_commits_since_tag(tag)
    
    if not commits:
        print(f"No commits found since {tag or 'beginning'}")
        return 0
    
    print(f"📋 Release notes preview for {args.version}")
    print(f"📅 Generated from commits since: {tag or 'beginning of repository'}\n")
    print("─" * 70)
    
    release_notes = generate_release_notes(commits, args.version)
    print(f"\n## {args.version}\n{release_notes}\n")
    
    print("─" * 70)
    print(f"\n✓ Found {len(commits)} commits to include\n")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
