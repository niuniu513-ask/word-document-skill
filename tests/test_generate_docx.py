# -*- coding: utf-8 -*-
"""Unit tests for generate_docx.py."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))


class TestGenerateDocx:
    def test_generate_from_spec(self, tmp_path):
        from generate_docx import main as gen_main
        spec = {
            "sections": [{
                "page": {"width_cm": 21.0, "height_cm": 29.7, "orientation": "portrait"},
                "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
            }],
            "conventions": {
                "body_style": {
                    "style": "Normal",
                    "font": {"ascii": "Times New Roman", "eastAsia": "宋体", "size_pt": 12},
                    "paragraph": {"line_spacing": 1.5},
                },
                "heading_styles": [
                    {"level": 1, "style": "Heading 1"},
                    {"level": 2, "style": "Heading 2"},
                ],
            },
        }
        spec_path = tmp_path / "spec.json"
        spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")

        content = {
            "blocks": [
                {"type": "heading", "level": 1, "text": "第一章 概述"},
                {"type": "paragraph", "text": "这是正文内容。"},
                {"type": "heading", "level": 2, "text": "1.1 背景"},
                {"type": "paragraph", "text": "背景说明。"},
                {"type": "table", "rows": [["项目", "数值"], ["A", "1"]], "caption": "表1 测试"},
            ]
        }
        content_path = tmp_path / "content.json"
        content_path.write_text(json.dumps(content, ensure_ascii=False), encoding="utf-8")

        out = tmp_path / "output.docx"
        sys.argv = ["generate_docx.py", "--spec", str(spec_path),
                    "--content", str(content_path), "-o", str(out)]
        rc = gen_main()
        assert rc == 0
        assert out.exists()

        from docx import Document
        doc = Document(str(out))
        texts = [p.text for p in doc.paragraphs if p.text.strip()]
        assert "第一章 概述" in texts
        assert "这是正文内容。" in texts
        assert len(doc.tables) == 1

    def test_generate_with_template_insert_mode(self, simple_docx, tmp_path):
        from generate_docx import main as gen_main
        content = {
            "blocks": [
                {"type": "paragraph", "after": "第二节", "text": "插入到第二节之后的内容。"},
            ]
        }
        content_path = tmp_path / "content.json"
        content_path.write_text(json.dumps(content, ensure_ascii=False), encoding="utf-8")

        out = tmp_path / "output.docx"
        sys.argv = ["generate_docx.py", "--template", str(simple_docx),
                    "--content", str(content_path), "-o", str(out), "--mode", "insert"]
        rc = gen_main()
        assert rc == 0
        assert out.exists()

        from docx import Document
        doc = Document(str(out))
        texts = [p.text for p in doc.paragraphs]
        assert "插入到第二节之后的内容。" in texts

    def test_generate_requires_template_or_spec(self, tmp_path):
        from generate_docx import main as gen_main
        content_path = tmp_path / "content.json"
        content_path.write_text(json.dumps({"blocks": []}), encoding="utf-8")
        out = tmp_path / "output.docx"
        sys.argv = ["generate_docx.py", "--content", str(content_path), "-o", str(out)]
        rc = gen_main()
        assert rc == 2

    def test_generate_cover_placeholders(self, tmp_path):
        """Test that {{key}} placeholders are filled."""
        from docx import Document
        # Create a template with placeholders
        doc = Document()
        doc.add_paragraph("姓名：{{name}}")
        doc.add_paragraph("学号：{{id}}")
        template = tmp_path / "template.docx"
        doc.save(str(template))

        from generate_docx import main as gen_main
        content = {
            "blocks": [
                {"type": "cover", "fields": {"name": "张三", "id": "2024001"}},
            ]
        }
        content_path = tmp_path / "content.json"
        content_path.write_text(json.dumps(content, ensure_ascii=False), encoding="utf-8")

        out = tmp_path / "output.docx"
        sys.argv = ["generate_docx.py", "--template", str(template),
                    "--content", str(content_path), "-o", str(out)]
        rc = gen_main()
        assert rc == 0

        result = Document(str(out))
        full_text = "\n".join(p.text for p in result.paragraphs)
        assert "张三" in full_text
        assert "2024001" in full_text
        assert "{{name}}" not in full_text

    def test_rebuild_preserves_fixed_template_title(self, tmp_path):
        """Rebuild must keep fixed cover text even when source content has another title."""
        from docx import Document

        template_doc = Document()
        template_doc.add_paragraph("《系统设计创新与电子系统设计实践》")
        template_doc.add_paragraph("小组拟题及方案设计报告")
        template_doc.add_heading("模板正文起点", level=1)
        template_doc.add_paragraph("旧正文，rebuild 后应被移除。")
        template = tmp_path / "protected-template.docx"
        template_doc.save(str(template))

        content = {
            "blocks": [
                {"type": "heading", "level": 1, "text": "源资料中的另一个标题"},
                {"type": "paragraph", "text": "新的正文内容。"},
            ]
        }
        content_path = tmp_path / "content.json"
        content_path.write_text(json.dumps(content, ensure_ascii=False), encoding="utf-8")
        out = tmp_path / "output.docx"

        from generate_docx import main as gen_main
        sys.argv = ["generate_docx.py", "--template", str(template),
                    "--content", str(content_path), "-o", str(out), "--mode", "rebuild"]
        assert gen_main() == 0

        result = Document(str(out))
        texts = [p.text for p in result.paragraphs]
        assert texts[:2] == ["《系统设计创新与电子系统设计实践》", "小组拟题及方案设计报告"]
        assert "源资料中的另一个标题" in texts
        assert "旧正文，rebuild 后应被移除。" not in texts
