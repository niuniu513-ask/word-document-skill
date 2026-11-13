"""Tests for the MoonBit engine bridge (``scripts/mbt_engine.py``).

The engine in ``engine/`` is an optional second implementation of the skill's
contract. These tests skip whenever the MoonBit toolchain is not installed, so a
plain Python install still runs the whole suite.
"""
from __future__ import annotations

import mbt_engine as mbt
import pytest

pytestmark = pytest.mark.skipif(
    not mbt.available(),
    reason="MoonBit toolchain or engine sources are not available",
)


@pytest.fixture(scope="module")
def case_docx(tmp_path_factory):
    """A small report with a real heading, a body line, a fake heading and a blank."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    doc.add_heading("一、引言", level=1)
    doc.add_paragraph("这是正文段落，用来确认正文没有被误判成标题。")
    fake = doc.add_paragraph("研究背景")
    fake.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fake.runs[0].bold = True
    doc.add_paragraph("【待补充】")
    path = tmp_path_factory.mktemp("mbt") / "case.docx"
    doc.save(str(path))
    return path


def test_doctor_reports_a_usable_engine():
    report = mbt.doctor()
    assert report["available"] is True
    assert report["ok"] is True, report.get("smoke")


def test_analyze_produces_a_specification(case_docx):
    spec = mbt.analyze(case_docx)
    assert spec["engine"] == "docx-mbt"
    document = spec["document"]
    assert document["paragraph_count"] == 4
    assert document["table_count"] == 0
    assert document["part_count"] >= 10
    kinds = {issue["kind"] for issue in spec["issues"]}
    assert "unfilled_placeholders" in kinds
    assert "manual_heading_numbering" in kinds


def test_fake_heading_is_detected(case_docx):
    spec = mbt.analyze(case_docx)
    fake = [i for i in spec["issues"] if i["kind"] == "fake_heading"]
    assert fake, "centered bold body paragraph should be reported as a fake heading"
    assert fake[0]["count"] == 1


def test_fill_then_validate_keeps_formatting(case_docx, tmp_path):
    out = tmp_path / "filled.docx"
    result = mbt.fill(case_docx, {"【待补充】": "已确认"}, out)
    assert result["replaced"] == 1
    assert out.exists()

    delta = mbt.validate(case_docx, out)
    assert delta["drift_count"] == 0
    assert delta["ok"] is True

    # Read the result back with a different stack: the text must really be there.
    from docx import Document

    text = "\n".join(p.text for p in Document(str(out)).paragraphs)
    assert "已确认" in text
    assert "【待补充】" not in text
