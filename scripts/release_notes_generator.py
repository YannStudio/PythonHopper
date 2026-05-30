"""Generate release notes from Git commits using Conventional Commits format."""

from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

PROJECT_ROOT = Path(__file__).resolve().parents[1]


@dataclass
class Commit:
    """Represents a parsed conventional commit."""
    
    hash: str
    type: str  # feat, fix, docs, style, refactor, test, chore, ci
    scope: Optional[str]  # module/component
    description: str
    breaking: bool = False
    body: str = ""
    footer: str = ""
    
    @property
    def full_message(self) -> str:
        """Return full formatted message."""
        scope_part = f"({self.scope})" if self.scope else ""
        breaking_part = "!" if self.breaking else ""
        return f"{self.type}{scope_part}{breaking_part}: {self.description}"
    
    @classmethod
    def parse(cls, commit_hash: str, commit_message: str) -> Optional[Commit]:
        """Parse a conventional commit message."""
        lines = commit_message.strip().split("\n", 1)
        header = lines[0]
        body_and_footer = lines[1] if len(lines) > 1 else ""
        
        # Parse: type(scope)!: description
        pattern = r"^(\w+)(?:\(([^)]+)\))?(!)?:\s*(.+)$"
        match = re.match(pattern, header)
        
        if not match:
            return None
        
        commit_type, scope, breaking, description = match.groups()
        
        # Check for breaking changes in footer
        has_breaking = breaking == "!" or "BREAKING CHANGE:" in body_and_footer
        
        body, footer = "", ""
        if body_and_footer:
            parts = body_and_footer.split("BREAKING CHANGE:", 1)
            body = parts[0].strip()
            if len(parts) > 1:
                footer = "BREAKING CHANGE:" + parts[1]
        
        return cls(
            hash=commit_hash[:7],
            type=commit_type.lower(),
            scope=scope,
            description=description,
            breaking=has_breaking,
            body=body,
            footer=footer,
        )


def get_commits_since_tag(tag: Optional[str] = None) -> list[Commit]:
    """Get all commits since the specified tag (or all commits if no tag)."""
    try:
        if tag:
            commit_range = f"{tag}..HEAD"
        else:
            commit_range = "HEAD"
        
        result = subprocess.run(
            ["git", "log", commit_range, "--pretty=format:%H%n%B%n---END_COMMIT---"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            check=True,
        )
        
        commits = []
        raw_commits = result.stdout.split("---END_COMMIT---")
        
        for raw_commit in raw_commits:
            if not raw_commit.strip():
                continue
            
            lines = raw_commit.strip().split("\n", 1)
            if len(lines) < 1:
                continue
            
            commit_hash = lines[0].strip()
            message = lines[1].strip() if len(lines) > 1 else ""
            
            parsed = Commit.parse(commit_hash, message)
            if parsed:
                commits.append(parsed)
        
        return commits
    
    except subprocess.CalledProcessError:
        return []


def categorize_commits(commits: list[Commit]) -> dict[str, list[Commit]]:
    """Categorize commits by type."""
    categories = {
        "breaking": [],
        "feat": [],
        "fix": [],
        "refactor": [],
        "perf": [],
        "docs": [],
        "style": [],
        "test": [],
        "chore": [],
        "ci": [],
        "other": [],
    }
    
    for commit in commits:
        if commit.breaking:
            categories["breaking"].append(commit)
        elif commit.type in categories:
            categories[commit.type].append(commit)
        else:
            categories["other"].append(commit)
    
    # Remove empty categories
    return {k: v for k, v in categories.items() if v}


def format_commit(commit: Commit) -> str:
    """Format a single commit for release notes."""
    scope_part = f"**{commit.scope}**: " if commit.scope else ""
    breaking_part = " ⚠️ BREAKING" if commit.breaking else ""
    return f"- {scope_part}{commit.description} ({commit.hash}){breaking_part}"


def generate_release_notes(commits: list[Commit], version: str) -> str:
    """Generate formatted release notes from commits."""
    if not commits:
        return "- Release voorbereid."
    
    categorized = categorize_commits(commits)
    
    # Category labels in Dutch
    category_labels = {
        "breaking": "⚠️ Breaking Changes",
        "feat": "✨ Features",
        "fix": "🐛 Bug Fixes",
        "refactor": "♻️ Refactoring",
        "perf": "⚡ Performance",
        "docs": "📚 Documentation",
        "style": "🎨 Style",
        "test": "✅ Tests",
        "chore": "🔧 Chores",
        "ci": "🤖 CI/CD",
        "other": "📝 Other Changes",
    }
    
    lines = []
    
    for category, label in category_labels.items():
        if category not in categorized:
            continue
        
        category_commits = categorized[category]
        if not category_commits:
            continue
        
        lines.append(f"\n### {label}\n")
        for commit in category_commits:
            lines.append(format_commit(commit) + "\n")
    
    return "".join(lines) if lines else "- Release voorbereid."


def get_last_tag() -> Optional[str]:
    """Get the most recent Git tag."""
    try:
        result = subprocess.run(
            ["git", "describe", "--tags", "--abbrev=0"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        
        if result.returncode == 0:
            return result.stdout.strip()
        return None
    
    except Exception:
        return None


def generate_changelog_entry(version: str, date: str) -> str:
    """Generate full changelog entry with auto-generated release notes."""
    last_tag = get_last_tag()
    commits = get_commits_since_tag(last_tag)
    release_notes = generate_release_notes(commits, version)
    
    return f"## {version} - {date}\n\n{release_notes}\n"
