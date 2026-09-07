# -*- coding: utf-8 -*-
"""Unit tests for analyze_docx.py."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))


class TestAnalyzeDocx:
    def test_analyze_simple_doc(self, simple_docx, tmp_path):
        from analyze_docx import main as analyze_main
        out = tmp_path / "spec.json"
        sys.argv = ["analyze_docx.py", str(simple_docx), "-o", str(out)]
        rc = analyze_main()
        assert rc == 0
        assert out.exists()
        spec = json.loads(out.read_text(encoding="utf-8"))
        # Required top-level keys
        assert "file" in spec
        assert "document" in spec
        assert "sections" in spec
        assert "styles" in spec
        assert "outline" in spec
        assert "tables" in spec
        assert "images" in spec
        assert "issues" in spec

    def test_analyze_detects_headings(self, simple_docx, tmp_path):
        from analyze_docx import main as analyze_main
        out = tmp_path / "spec.json"
        sys.argv = ["analyze_docx.py", str(simple_docx), "-o", str(out)]
        analyze_main()
        spec = json.loads(out.read_text(encoding="utf-8"))
        headings = spec["outline"]
        assert len(headings) >= 2
        texts = [h["text"] for h in headings]
        assert "测试文档" in texts
        assert "第二节" in texts

    def test_analyze_detects_tables(self, simple_docx, tmp_path):
        from analyze_docx import main as analyze_main
        out = tmp_path / "spec.json"
        sys.argv = ["analyze_docx.py", str(simple_docx), "-o", str(out)]
        analyze_main()
        spec = json.loads(out.read_text(encoding="utf-8"))
        assert spec["document"]["table_count"] == 1
        assert len(spec["tables"]) == 1
        assert spec["tables"][0]["rows"] == 2
        assert spec["tables"][0]["cols"] == 3

    def test_analyze_page_setup(self, simple_docx, tmp_path):
        from analyze_docx import main as analyze_main
        out = tmp_path / "spec.json"
        sys.argv = ["analyze_docx.py", str(simple_docx), "-o", str(out)]
        analyze_main()
        spec = json.loads(out.read_text(encoding="utf-8"))
        sec = spec["sections"][0]
        # python-docx default is Letter (21.59 x 27.94 cm), not A4
        assert sec["page"]["width_cm"] > 20
        assert sec["page"]["height_cm"] > 27
        assert sec["margins_cm"]["top"] == pytest.approx(2.54, abs=0.1)

    def test_analyze_nonexistent_file(self, tmp_path):
        from analyze_docx import main as analyze_main
        sys.argv = ["analyze_docx.py", str(tmp_path / "nonexistent.docx")]
        rc = analyze_main()
        assert rc == 2

    def test_analyze_formatting_inventory(self, simple_docx, tmp_path):
        from analyze_docx import main as analyze_main
        out = tmp_path / "spec.json"
        sys.argv = ["analyze_docx.py", str(simple_docx), "-o", str(out)]
        analyze_main()
        spec = json.loads(out.read_text(encoding="utf-8"))
        inv = spec["formatting_inventory"]
        assert "run_count" in inv
        assert "dominant" in inv
        assert "font_combinations" in inv
        assert "colors" in inv
        assert inv["run_count"] > 0
