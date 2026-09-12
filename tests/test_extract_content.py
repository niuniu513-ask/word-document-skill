"""Unit tests for extract_content.py."""
from __future__ import annotations

import json
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))


class TestExtractContent:
    def test_extract_simple_doc(self, simple_docx, tmp_path):
        from extract_content import main as extract_main
        out = tmp_path / "content.json"
        md = tmp_path / "content.md"
        sys.argv = ["extract_content.py", str(simple_docx),
                    "-o", str(out), "--markdown", str(md)]
        rc = extract_main()
        assert rc == 0
        assert out.exists()
        assert md.exists()

        content = json.loads(out.read_text(encoding="utf-8"))
        assert "blocks" in content
        assert "meta" in content
        block_types = [b["type"] for b in content["blocks"]]
        assert "heading" in block_types
        assert "paragraph" in block_types
        assert "table" in block_types

    def test_extract_markdown_view(self, simple_docx, tmp_path):
        from extract_content import main as extract_main
        md = tmp_path / "content.md"
        sys.argv = ["extract_content.py", str(simple_docx), "--markdown", str(md)]
        extract_main()
        text = md.read_text(encoding="utf-8")
        assert "# 测试文档" in text or "## 测试文档" in text

    def test_extract_nonexistent_file(self, tmp_path):
        from extract_content import main as extract_main
        sys.argv = ["extract_content.py", str(tmp_path / "nope.docx")]
        rc = extract_main()
        assert rc == 2
