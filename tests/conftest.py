# -*- coding: utf-8 -*-
"""Shared pytest fixtures for word-document-skill tests."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Make scripts importable
SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))


@pytest.fixture
def tmp_docx(tmp_path):
    """Return a path for a temporary .docx file."""
    return tmp_path / "test.docx"


@pytest.fixture
def simple_docx(tmp_path):
    """Create a simple docx with headings, paragraphs, and a table."""
    from docx import Document
    from docx.shared import Pt

    doc = Document()
    # Title
    title = doc.add_heading("测试文档", level=1)
    # Body paragraph
    p = doc.add_paragraph("这是第一段正文内容，用于测试。")
    p.style = doc.styles["Normal"]
    # Second heading
    doc.add_heading("第二节", level=2)
    doc.add_paragraph("第二节的正文。")
    # Table
    table = doc.add_table(rows=2, cols=3)
    table.style = "Table Grid"
    for i, text in enumerate(["列1", "列2", "列3"]):
        table.cell(0, i).text = text
    for i, text in enumerate(["a", "b", "c"]):
        table.cell(1, i).text = text
    path = tmp_path / "simple.docx"
    doc.save(str(path))
    return path


@pytest.fixture
def empty_paragraphs_docx(tmp_path):
    """Create a docx with consecutive empty paragraphs."""
    from docx import Document

    doc = Document()
    doc.add_paragraph("第一段")
    for _ in range(5):
        doc.add_paragraph("")
    doc.add_paragraph("第二段")
    path = tmp_path / "empty.docx"
    doc.save(str(path))
    return path
