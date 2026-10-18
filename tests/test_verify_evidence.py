"""Unit tests for verify_evidence.py."""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))


def _solid(path: Path, size=(200, 120), color=(18, 19, 20)) -> Path:
    from PIL import Image

    Image.new("RGB", size, color).save(path)
    return path


def _text_like(path: Path, size=(600, 220), margin=40) -> Path:
    """Dark background with light blocks in the middle (stands in for code)."""
    from PIL import Image, ImageDraw

    im = Image.new("RGB", size, (18, 19, 20))
    draw = ImageDraw.Draw(im)
    for y in range(30, size[1] - 30, 26):
        draw.rectangle((margin, y, size[0] - margin, y + 12), fill=(190, 190, 190))
    im.save(path)
    return path


def _wide_bar(path: Path, size=(1200, 140)) -> Path:
    """Content only in the left 12% of a very wide image."""
    from PIL import Image, ImageDraw

    im = Image.new("RGB", size, (18, 19, 20))
    draw = ImageDraw.Draw(im)
    for y in range(30, size[1] - 30, 24):
        draw.rectangle((30, y, 150, y + 12), fill=(190, 190, 190))
    im.save(path)
    return path


def _build(tmp_path: Path, steps) -> Path:
    """Build a docx from [(label, image_path, width_cm, height_cm), ...]."""
    from docx import Document
    from docx.shared import Cm

    doc = Document()
    doc.add_paragraph("实验报告")
    for item in steps:
        label, img, w, h = item
        if label:
            doc.add_paragraph(f"{label}：")
        if img:
            kw = {"width": Cm(w)}
            if h:
                kw["height"] = Cm(h)
            doc.add_picture(str(img), **kw)
    path = tmp_path / "report.docx"
    doc.save(str(path))
    return path


def test_counts_labels_and_images(tmp_path):
    from verify_evidence import verify_docx

    good = _text_like(tmp_path / "code.png")
    doc = _build(tmp_path, [
        ("代码截图", good, 15.2, None),
        ("实验结果截图", good, 8.0, None),
    ])
    rep = verify_docx(doc, ["代码截图", "实验结果截图"], 0.35, 8.0)
    assert rep.label_counts == {"代码截图": 1, "实验结果截图": 1}
    assert len(rep.images) == 2
    assert rep.counts("error") == 0


def test_label_without_image_is_error(tmp_path):
    from verify_evidence import verify_docx

    good = _text_like(tmp_path / "code.png")
    doc = _build(tmp_path, [("代码截图", None, 0, None), ("实验结果截图", good, 8.0, None)])
    rep = verify_docx(doc, ["代码截图", "实验结果截图"], 0.35, 8.0)
    assert any(f.code == "label_without_image" for f in rep.findings)


@pytest.mark.parametrize("qualifier", ["（VS Code）", "(VS Code)", " （Jupyter） "])
def test_editor_qualified_captions_pair_with_images(tmp_path, qualifier):
    from verify_evidence import verify_docx

    good = _text_like(tmp_path / "code.png")
    doc = _build(tmp_path, [
        (f"代码截图{qualifier}", good, 15.2, None),
        (f"实验结果截图{qualifier}", good, 8.0, None),
    ])
    rep = verify_docx(doc, ["代码截图", "实验结果截图"], 0.35, 8.0)
    assert rep.label_counts == {"代码截图": 1, "实验结果截图": 1}
    assert all(i.label for i in rep.images)
    assert not any(f.code == "orphan_image" for f in rep.findings)


def test_orphan_image_reported_once(tmp_path):
    from verify_evidence import verify_docx

    good = _text_like(tmp_path / "code.png")
    doc = _build(tmp_path, [(None, good, 10.0, None), (None, good, 10.0, None)])
    rep = verify_docx(doc, ["代码截图"], 0.35, 8.0)
    orphans = [f for f in rep.findings if f.code == "orphan_image"]
    assert len(orphans) == 1
    assert "2 张图" in orphans[0].detail

    quiet = verify_docx(doc, ["代码截图"], 0.35, 8.0, allow_orphan=True)
    assert not any(f.code == "orphan_image" for f in quiet.findings)


def test_blank_image_flagged(tmp_path):
    from verify_evidence import verify_docx

    blank = _solid(tmp_path / "blank.png")
    doc = _build(tmp_path, [("代码截图", blank, 10.0, None)])
    rep = verify_docx(doc, ["代码截图"], 0.35, 8.0)
    assert any(f.code == "blank_image" for f in rep.findings)
    assert rep.counts("error") >= 1


def test_wide_bar_flagged(tmp_path):
    from verify_evidence import verify_docx

    bar = _wide_bar(tmp_path / "bar.png")
    doc = _build(tmp_path, [("实验结果截图", bar, 15.2, None)])
    rep = verify_docx(doc, ["实验结果截图"], 0.35, 8.0)
    assert any(f.code == "wide_bar" for f in rep.findings)


def test_distorted_image_flagged(tmp_path):
    from verify_evidence import verify_docx

    good = _text_like(tmp_path / "code.png")
    doc = _build(tmp_path, [("代码截图", good, 12.0, 2.0)])
    rep = verify_docx(doc, ["代码截图"], 0.35, 8.0)
    assert any(f.code == "distorted" for f in rep.findings)


def test_too_tall_flagged(tmp_path):
    from PIL import Image, ImageDraw
    from verify_evidence import verify_docx

    p = tmp_path / "tall.png"
    im = Image.new("RGB", (200, 3000), (18, 19, 20))
    draw = ImageDraw.Draw(im)
    for y in range(20, 2980, 30):
        draw.rectangle((20, y, 180, y + 10), fill=(190, 190, 190))
    im.save(p)
    doc = _build(tmp_path, [("代码截图", p, 10.0, None)])
    rep = verify_docx(doc, ["代码截图"], 0.35, 8.0)
    assert any(f.code == "too_tall" for f in rep.findings)


def test_cli_writes_json_and_exit_code(tmp_path):
    from verify_evidence import main

    good = _text_like(tmp_path / "code.png")
    doc = _build(tmp_path, [("代码截图", good, 15.2, None)])
    out = tmp_path / "evidence.json"
    rc = main([str(doc), "-o", str(out), "--label", "代码截图",
               "--expect", "代码截图=1", "--quiet"])
    assert rc == 0
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["labels"]["代码截图"] == 1
    assert data["images"]["total"] == 1
    assert data["summary"]["passed"] is True


def test_cli_expect_mismatch_warns(tmp_path):
    from verify_evidence import main

    good = _text_like(tmp_path / "code.png")
    doc = _build(tmp_path, [("代码截图", good, 15.2, None)])
    out = tmp_path / "evidence.json"
    rc = main([str(doc), "-o", str(out), "--label", "代码截图",
               "--expect", "代码截图=3", "--strict", "--quiet"])
    assert rc == 1
    data = json.loads(out.read_text(encoding="utf-8"))
    assert any(f["code"] == "unexpected_count" for f in data["findings"])


def test_cli_rejects_unknown_suffix(tmp_path):
    from verify_evidence import main

    bad = tmp_path / "a.txt"
    bad.write_text("x", encoding="utf-8")
    with pytest.raises(SystemExit):
        main([str(bad), "--quiet"])
