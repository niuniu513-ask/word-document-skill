"""Unit tests for _docx_utils.py shared helpers."""
from __future__ import annotations

import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

from _docx_utils import (  # noqa: E402
    _fmt_num,
    _int_to_cn,
    body_blocks,
    emu_to_cm,
    first,
    half_pt_to_pt,
    is_empty_para,
    jd,
    run_font,
    style_chain,
    twips_to_cm,
    twips_to_pt,
)


class TestUnitConversions:
    def test_emu_to_cm(self):
        assert emu_to_cm(360000) == 1.0
        assert emu_to_cm(914400) == 2.54
        assert emu_to_cm(None) is None

    def test_twips_to_cm(self):
        # 1 cm ≈ 566.929 twips
        assert abs(twips_to_cm(567) - 1.0) < 0.01
        assert twips_to_cm(None) is None

    def test_twips_to_pt(self):
        assert twips_to_pt(20) == 1.0
        assert twips_to_pt(240) == 12.0
        assert twips_to_pt(None) is None

    def test_half_pt_to_pt(self):
        assert half_pt_to_pt(24) == 12.0
        assert half_pt_to_pt(None) is None


class TestJsonHelpers:
    def test_jd_chinese(self):
        result = jd({"name": "测试", "值": 42})
        assert "测试" in result
        assert "42" in result
        assert "\\u" not in result  # ensure_ascii=False

    def test_jd_nested(self):
        result = jd({"a": [1, 2, {"b": "c"}]})
        assert '"a"' in result


class TestXmlHelpers:
    def test_first_returns_none_for_none(self):
        assert first(None, "w:r") is None

    def test_first_on_element(self, simple_docx):
        from docx import Document
        doc = Document(str(simple_docx))
        p = doc.paragraphs[0]
        # first paragraph should have a pPr or runs
        result = first(p._p, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        assert result is not None


class TestBodyBlocks:
    def test_yields_paragraphs_and_tables(self, simple_docx):
        from docx import Document
        doc = Document(str(simple_docx))
        blocks = list(body_blocks(doc))
        kinds = [k for k, _ in blocks]
        assert "p" in kinds
        assert "tbl" in kinds

    def test_body_blocks_order(self, simple_docx):
        from docx import Document
        doc = Document(str(simple_docx))
        blocks = list(body_blocks(doc))
        # First block should be a paragraph (heading)
        assert blocks[0][0] == "p"


class TestIsEmptyPara:
    def test_empty_paragraph(self, simple_docx):
        from docx import Document
        doc = Document(str(simple_docx))
        # Add an empty paragraph
        p = doc.add_paragraph("")
        assert is_empty_para(p) is True

    def test_nonempty_paragraph(self, simple_docx):
        from docx import Document
        doc = Document(str(simple_docx))
        p = doc.paragraphs[0]
        assert is_empty_para(p) is False

    def test_whitespace_only(self):
        from docx import Document
        doc = Document()
        p = doc.add_paragraph("   ")
        assert is_empty_para(p) is True


class TestChineseNumbering:
    def test_int_to_cn_single(self):
        assert _int_to_cn(1) == "一"
        assert _int_to_cn(5) == "五"
        assert _int_to_cn(9) == "九"
        assert _int_to_cn(10) == "十"

    def test_int_to_cn_teens(self):
        assert _int_to_cn(11) == "十一"
        assert _int_to_cn(19) == "十九"

    def test_int_to_cn_tens(self):
        assert _int_to_cn(20) == "二十"
        assert _int_to_cn(25) == "二十五"
        assert _int_to_cn(99) == "九十九"

    def test_fmt_num_decimal(self):
        assert _fmt_num("decimal", 5) == "5"

    def test_fmt_num_chinese(self):
        assert _fmt_num("chineseCounting", 3) == "三"

    def test_fmt_num_upper_letter(self):
        assert _fmt_num("upperLetter", 1) == "A"
        assert _fmt_num("upperLetter", 26) == "Z"

    def test_fmt_num_lower_letter(self):
        assert _fmt_num("lowerLetter", 1) == "a"


class TestStyleChain:
    def test_style_chain_no_cycle(self, simple_docx):
        from docx import Document
        doc = Document(str(simple_docx))
        normal = doc.styles["Normal"]
        chain = list(style_chain(normal))
        assert len(chain) >= 1
        assert chain[0] == normal


class TestRunFont:
    def test_theme_font_and_color_metadata(self):
        from docx import Document
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        doc = Document()
        run = doc.add_paragraph().add_run("Theme formatted")
        rpr = run._r.get_or_add_rPr()
        fonts = OxmlElement("w:rFonts")
        fonts.set(qn("w:asciiTheme"), "minorHAnsi")
        fonts.set(qn("w:eastAsiaTheme"), "majorEastAsia")
        rpr.append(fonts)
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "4472C4")
        color.set(qn("w:themeColor"), "accent1")
        color.set(qn("w:themeTint"), "99")
        rpr.append(color)

        detected = run_font(run._r)
        assert detected["ascii_theme"] == "minorHAnsi"
        assert detected["eastAsia_theme"] == "majorEastAsia"
        assert detected["color"] == "4472C4"
        assert detected["color_theme"] == "accent1"
        assert detected["color_theme_tint"] == "99"
