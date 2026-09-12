"""Unit tests for validate_docx.py."""
from __future__ import annotations

import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))


class TestValidateDocx:
    def test_check_simple_doc(self, simple_docx, tmp_path):
        from validate_docx import main as val_main
        out = tmp_path / "check.json"
        sys.argv = ["validate_docx.py", "--check", str(simple_docx), "-o", str(out)]
        rc = val_main()
        assert rc == 0
        assert out.exists()
        result = json.loads(out.read_text(encoding="utf-8"))
        assert result["mode"] == "check"
        assert "problem_count" in result
        assert "problems" in result

    def test_check_empty_paragraphs(self, empty_paragraphs_docx, tmp_path):
        from validate_docx import main as val_main
        out = tmp_path / "check.json"
        sys.argv = ["validate_docx.py", "--check", str(empty_paragraphs_docx), "-o", str(out)]
        val_main()
        result = json.loads(out.read_text(encoding="utf-8"))
        # Should detect consecutive empty paragraphs
        empty_issues = [p for p in result["problems"] if "空段落" in p["issue"]]
        assert len(empty_issues) >= 1

    def test_delta_identical_documents(self, simple_docx, tmp_path):
        import shutil

        from validate_docx import main as val_main
        copy = tmp_path / "copy.docx"
        shutil.copy2(simple_docx, copy)
        out = tmp_path / "delta.json"
        sys.argv = ["validate_docx.py", "--original", str(simple_docx),
                    "--final", str(copy), "-o", str(out)]
        rc = val_main()
        assert rc == 0
        result = json.loads(out.read_text(encoding="utf-8"))
        assert result["format_delta_count"] == 0

    def test_delta_requires_both_args(self, simple_docx):
        from validate_docx import main as val_main
        sys.argv = ["validate_docx.py", "--original", str(simple_docx)]
        rc = val_main()
        assert rc == 2

    def test_delta_detects_table_count_change(self, simple_docx, tmp_path):
        import shutil

        from docx import Document
        # Create a modified copy with an extra table
        copy = tmp_path / "modified.docx"
        shutil.copy2(simple_docx, copy)
        doc = Document(str(copy))
        doc.add_table(rows=1, cols=2)
        doc.save(str(copy))

        from validate_docx import main as val_main
        out = tmp_path / "delta.json"
        sys.argv = ["validate_docx.py", "--original", str(simple_docx),
                    "--final", str(copy), "-o", str(out)]
        val_main()
        result = json.loads(out.read_text(encoding="utf-8"))
        table_diffs = [d for d in result["format_delta"] if d["category"] == "tables"]
        assert len(table_diffs) >= 1
