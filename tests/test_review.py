"""review.py（合规体检）的测试。

这些测试用 python-docx 现造文档，因此不依赖任何外部样本文件；
每个用例只断言"哪条规则应当/不应当触发"，避免把分数写死。
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import review  # noqa: E402


def add_page_number(section) -> None:
    """插入真正的 PAGE 域；手打页码不算。"""
    p = section.footer.paragraphs[0]
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    fld.append(run)
    p._p.append(fld)


def build(tmp_path: Path, sections_=("一、工程概况", "二、编制依据", "三、施工方案"),
          body="正文内容。" * 20, number=True, placeholders=False, jump=False,
          extra="", name="doc.docx") -> Path:
    doc = Document()
    doc.sections[0].page_width = Cm(21.0)
    doc.sections[0].page_height = Cm(29.7)
    if number:
        add_page_number(doc.sections[0])
    for i, title in enumerate(sections_):
        level = 1 if not (jump and i == 1) else 3
        doc.add_heading(title, level=level)
        p = doc.add_paragraph("本节说明文字。" * 6)
        p.paragraph_format.first_line_indent = Pt(24)
    doc.add_paragraph(body)
    if placeholders:
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "验收情况"
        table.cell(0, 1).text = "【待补充:验收情况】"
    if extra:
        doc.add_paragraph(extra)
    path = tmp_path / name
    doc.save(str(path))
    return path


def run(path: Path, profile: str = "generic") -> review.Report:
    return review.review_file(path, review.load_profile(profile))


def rules_hit(rep: review.Report) -> set[str]:
    return {f.rule for f in rep.findings}


def test_clean_document_scores_full_marks(tmp_path):
    rep = run(build(tmp_path, sections_=("一、工程概况", "二、编制依据")))
    assert rep.score == 100
    assert not rep.blocked
    assert rules_hit(rep) == set()


def test_missing_required_sections_is_a_blocker(tmp_path):
    rep = run(build(tmp_path, sections_=("一、工程概况",)), profile="railway-plan")
    assert "required_sections" in rules_hit(rep)
    assert rep.blocked
    finding = next(f for f in rep.findings if f.rule == "required_sections")
    assert "编制依据" in finding.detail


def test_placeholder_in_table_cell_is_found(tmp_path):
    rep = run(build(tmp_path, placeholders=True))
    finding = next(f for f in rep.findings if f.rule == "placeholders")
    assert "1 处" in finding.title
    assert "表 1" in finding.detail


def test_heading_level_jump_is_reported(tmp_path):
    rep = run(build(tmp_path, jump=True))
    assert "heading_level_jump" in rules_hit(rep)


def test_missing_page_number_field_is_reported(tmp_path):
    rep = run(build(tmp_path, number=False), profile="railway-plan")
    assert "page_number" in rules_hit(rep)


def test_forbidden_phrases_are_reported(tmp_path):
    rep = run(build(tmp_path, extra="具体时间待定，另行通知。"))
    assert "forbidden_phrases" in rules_hit(rep)


def test_profile_choice_changes_the_verdict(tmp_path):
    path = build(tmp_path, sections_=("一、工程概况",))
    generic = run(path, "generic")
    railway = run(path, "railway-plan")
    assert not generic.blocked
    assert railway.blocked
    assert railway.score < generic.score


def test_report_renders_markdown_and_html(tmp_path):
    path = build(tmp_path, placeholders=True)
    rep = run(path, "generic")
    md = review.render_markdown(rep)
    html = review.render_html(rep)
    assert "# 文档合规体检报告" in md
    assert "怎么改" in md
    assert "<!doctype html>" in html and "问题清单" in html


def test_cli_exit_code_reflects_blockers(tmp_path, capsys):
    path = build(tmp_path, placeholders=True)
    argv = sys.argv
    try:
        sys.argv = ["review.py", str(path)]
        assert review.main() == 1
    finally:
        sys.argv = argv
