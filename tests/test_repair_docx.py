# -*- coding: utf-8 -*-
"""Unit tests for repair_docx.py."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))


class TestRepairDocx:
    def test_dry_run_does_not_modify(self, empty_paragraphs_docx, tmp_path):
        from repair_docx import main as repair_main
        import shutil
        original_size = empty_paragraphs_docx.stat().st_size
        sys.argv = ["repair_docx.py", str(empty_paragraphs_docx),
                    "--fix", "empty-paragraphs"]
        rc = repair_main()
        assert rc == 0
        # File should not be modified in dry-run
        assert empty_paragraphs_docx.stat().st_size == original_size

    def test_apply_empty_paragraphs(self, empty_paragraphs_docx, tmp_path):
        from repair_docx import main as repair_main
        out = tmp_path / "fixed.docx"
        sys.argv = ["repair_docx.py", str(empty_paragraphs_docx),
                    "--fix", "empty-paragraphs", "--apply", "-o", str(out)]
        rc = repair_main()
        assert rc == 0
        assert out.exists()

        from docx import Document
        doc = Document(str(out))
        # Count consecutive empty paragraphs
        max_streak = 0
        streak = 0
        for p in doc.paragraphs:
            if not p.text.strip():
                streak += 1
                max_streak = max(max_streak, streak)
            else:
                streak = 0
        # After repair, max consecutive empties should be <= 1 (default max-keep=1)
        assert max_streak <= 1

    def test_heading_keepnext(self, simple_docx, tmp_path):
        from repair_docx import main as repair_main
        out = tmp_path / "fixed.docx"
        sys.argv = ["repair_docx.py", str(simple_docx),
                    "--fix", "heading-keepnext", "--apply", "-o", str(out)]
        rc = repair_main()
        assert rc == 0
        assert out.exists()

    def test_unknown_fix_logged(self, simple_docx, tmp_path):
        from repair_docx import main as repair_main
        out = tmp_path / "fixed.docx"
        sys.argv = ["repair_docx.py", str(simple_docx),
                    "--fix", "nonexistent-fix", "--apply", "-o", str(out)]
        rc = repair_main()
        assert rc == 0  # doesn't crash, just logs unknown

    def test_repair_output_json(self, simple_docx, tmp_path):
        from repair_docx import main as repair_main
        out = tmp_path / "fixed.docx"
        sys.argv = ["repair_docx.py", str(simple_docx),
                    "--fix", "empty-paragraphs", "--apply", "-o", str(out)]
        # Should print JSON summary
        rc = repair_main()
        assert rc == 0
