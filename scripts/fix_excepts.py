#!/usr/bin/env python3
"""Small helper to replace bare 'except Exception as exc:' with 'except Exception as exc:' in .py files.

Use with care; intended for quick repo-wide lint fixes before manual review.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

for path in ROOT.rglob("*.py"):
    if "venv" in path.parts or "__pycache__" in path.parts:
        continue
    text = path.read_text(encoding="utf-8")
    if "except Exception as exc:" not in text:
        continue
    new = text.replace("except Exception as exc:", "except Exception as exc:")
    path.write_text(new, encoding="utf-8")
    print(f"Patched: {path}")
