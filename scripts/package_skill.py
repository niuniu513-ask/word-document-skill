#!/usr/bin/env python3
"""Package the skill into a distributable .skill (zip) file.

Usage:
    python scripts/package_skill.py [--output dist/word-document-skill.skill]

Includes: SKILL.md, AGENT-GUIDE.md, capabilities.json, README.md, LICENSE, scripts/, schemas/, workflows/,
references/, evals/, examples/. Excludes tests, __pycache__, .git, etc.
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent

INCLUDE_DIRS = ["scripts", "schemas", "workflows", "references", "evals", "examples", "rules"]
INCLUDE_FILES = ["SKILL.md", "AGENT-GUIDE.md", "capabilities.json", "README.md", "LICENSE", "CHANGELOG.md",
                 "requirements.txt", "requirements-notebook.txt", "pyproject.toml"]
EXCLUDE_PATTERNS = ["__pycache__", ".pyc", ".pyo", ".pytest_cache", ".git",
                    ".DS_Store", "Thumbs.db", "*.skill"]


def should_exclude(path: Path) -> bool:
    name = path.name
    for pat in EXCLUDE_PATTERNS:
        if pat.startswith("*"):
            if name.endswith(pat[1:]):
                return True
        elif pat in name or pat in str(path):
            return True
    return False


def main() -> int:
    ap = argparse.ArgumentParser(description="Package word-document-skill into .skill")
    ap.add_argument("-o", "--output", default=str(SKILL_ROOT / "dist" / "word-document-skill.skill"))
    args = ap.parse_args()

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    files_to_pack: list[tuple[Path, str]] = []

    # Top-level files
    for fname in INCLUDE_FILES:
        fpath = SKILL_ROOT / fname
        if fpath.exists():
            files_to_pack.append((fpath, fname))

    # Directories
    for dname in INCLUDE_DIRS:
        dpath = SKILL_ROOT / dname
        if not dpath.exists():
            continue
        for fpath in sorted(dpath.rglob("*")):
            if fpath.is_file() and not should_exclude(fpath):
                arcname = str(fpath.relative_to(SKILL_ROOT))
                files_to_pack.append((fpath, arcname))

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath, arcname in files_to_pack:
            zf.write(fpath, arcname)

    size_kb = out_path.stat().st_size / 1024
    print(f"=== packaged {len(files_to_pack)} files -> {out_path} ({size_kb:.1f} KB) ===")
    for _, arcname in files_to_pack:
        print(f"  {arcname}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
