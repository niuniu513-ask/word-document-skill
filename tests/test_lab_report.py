"""Workflow invariants. Fixture drawings are test-only, never screenshot evidence."""

import argparse

import pytest
from docx import Document
from docx.shared import Cm
from lab_report import audit, build, crop, digest, prepare, read, write
from PIL import Image, ImageDraw
from verify_evidence import main as verify_main
from verify_evidence import verify_docx


# The lab-report workflow needs nbformat; a plain install without the notebook
# extras should skip these tests instead of failing collection.
nbformat = pytest.importorskip("nbformat")


def test_nested_frame_evidence_counted_once(tmp_path):
    image = tmp_path / "fixture.png"
    im = Image.new("RGB", (300, 100), "white")
    ImageDraw.Draw(im).rectangle((20, 20, 280, 80), fill="blue")
    im.save(image)
    d = Document()
    t = d.add_table(rows=1, cols=2)
    cell = t.cell(0, 0).merge(t.cell(0, 1))
    p = cell.paragraphs[0]
    p.add_run("代码截图：")
    p.add_run().add_picture(str(image), width=Cm(10))
    inner = cell.add_table(rows=1, cols=1)
    p = inner.cell(0, 0).paragraphs[0]
    p.add_run("实验结果截图：")
    p.add_run().add_picture(str(image), width=Cm(10))
    path = tmp_path / "report.docx"
    d.save(path)
    r = verify_docx(path, ["代码截图", "实验结果截图"], 0.35, 8)
    assert r.label_counts == {"代码截图": 1, "实验结果截图": 1}
    assert len(r.images) == 2 and not r.findings
    assert verify_main([str(path), "--expect", "代码截图=2", "--strict", "--quiet"]) == 1


def test_prepare_skips_markdown_and_rejects_unexecuted(tmp_path):
    nb = nbformat.v4.new_notebook(
        cells=[
            nbformat.v4.new_markdown_cell("任务"),
            nbformat.v4.new_code_cell("x=1", execution_count=1, outputs=[]),
        ]
    )
    source = tmp_path / "input.ipynb"
    nbformat.write(nb, source)
    args = argparse.Namespace(notebook=str(source), execute=False, out=str(tmp_path / "evidence"))
    prepare(args)
    plan = read(tmp_path / "evidence/capture-plan.json")
    assert plan["cells"][0]["cell"] == 1 and not plan["cells"][0]["hasOutput"]
    assert nbformat.read(plan["cells"][0]["evidence_notebook"], 4).cells[0] == nb.cells[1]
    nb.cells[1].execution_count = None
    nbformat.write(nb, source)
    with pytest.raises(ValueError, match="unexecuted"):
        prepare(args)


def test_build_rejects_wrong_cover_and_missing_analysis(tmp_path):
    d = Document()
    d.add_paragraph("班级")
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "固定文字"
    t.cell(0, 0).add_paragraph("可填区")
    template = tmp_path / "template.docx"
    d.save(template)
    shots = tmp_path / "shots.json"
    write(shots, [{"cell": 0}])
    cfg = {
        "template": str(template),
        "output": str(tmp_path / "out.docx"),
        "container": {"kind": "cell", "table": 0, "row": 0, "column": 0},
        "keep_prefix": 1,
        "prototypes": {"heading": 1, "caption": 1, "analysis": 1},
        "screenshots": str(shots),
        "groups": [],
        "cover_edits": [{"paragraph": 0, "run": 0, "expected": "姓名", "text": "test"}],
    }
    p = tmp_path / "cfg.json"
    write(p, cfg)
    with pytest.raises(ValueError, match="Cover run mismatch"):
        build(argparse.Namespace(config=str(p)))
    cfg["cover_edits"] = []
    cfg["groups"] = [{"title": "题目", "items": [{"cell": 0}]}]
    img = tmp_path / "fixture.png"
    Image.new("RGB", (100, 60), "blue").save(img)
    write(shots, [{"cell": 0, "codeCrop": str(img), "hasOutput": False}])
    write(p, cfg)
    cfg["max_width_cm"] = 10
    write(p, cfg)
    with pytest.raises(ValueError, match="Missing analysis"):
        build(argparse.Namespace(config=str(p)))


def test_crop_preserves_pale_content_and_actual_pixel_scale(tmp_path):
    raw = tmp_path / "raw.png"
    im = Image.new("RGB", (400, 200), "white")
    ImageDraw.Draw(im).rectangle((370, 20, 390, 40), fill=(225, 225, 225))
    im.save(raw)
    plan = {
        "cells": [
            {
                "cell": 0,
                "hasOutput": False,
                "env": {"w": 200, "h": 100, "d": 1},
                "codeFile": str(raw),
                "codeBox": {"x": 0, "y": 0, "width": 200, "height": 100},
            }
        ]
    }
    p = tmp_path / "shots.json"
    write(p, plan)
    crop(argparse.Namespace(manifest=str(p), out=str(tmp_path / "crops")))
    item = read(tmp_path / "crops/screenshots-cropped.json")["cells"][0]
    assert Image.open(item["codeCrop"]).size == (400, 200)
    assert Image.open(item["codeCrop"]).getpixel((380, 30)) == (225, 225, 225)
    assert item["codeAudit"]["raw_sha256"] == digest(raw)
    # Padded Electron canvas must use the explicitly calibrated rendered scale.
    plan["cells"][0]["screenshotScale"] = {"x": 1, "y": 1}
    write(p, plan)
    crop(argparse.Namespace(manifest=str(p), out=str(tmp_path / "calibrated")))
    item = read(tmp_path / "calibrated/screenshots-cropped.json")["cells"][0]
    assert Image.open(item["codeCrop"]).size == (200, 100)


def test_audit_accepts_offscreen_manifest_and_joint_mode(tmp_path):
    """离屏清单没有 evidence_notebook：audit 改用 source 对账；joint 模式只插一张图。"""
    d = Document()
    d.add_paragraph("班级")
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "固定文字"
    t.cell(0, 0).add_paragraph("可填区")
    template = tmp_path / "template.docx"
    d.save(template)
    nb = nbformat.v4.new_notebook(
        cells=[
            nbformat.v4.new_code_cell(
                "print(1)",
                execution_count=1,
                outputs=[nbformat.v4.new_output("stream", name="stdout", text="1\n")],
            )
        ]
    )
    nbpath = tmp_path / "executed.ipynb"
    nbformat.write(nb, nbpath)
    img = tmp_path / "fixture.png"
    Image.new("RGB", (300, 120), "white").save(img)
    h = digest(img)
    rec = {
        "cell": 0,
        "hasOutput": True,
        "sourceFirstLine": "print(1)",
        "source": "print(1)",
        "codeCrop": str(img),
        "outputCrop": str(img),
        "cellCrop": str(img),
        "codeAudit": {"output_sha256": h},
        "outputAudit": {"output_sha256": h},
    }
    shots = tmp_path / "shots.json"
    write(shots, {"cells": [rec]})
    cfg = {
        "template": str(template),
        "notebook": str(nbpath),
        "output": str(tmp_path / "out.docx"),
        "container": {"kind": "cell", "table": 0, "row": 0, "column": 0},
        "keep_prefix": 1,
        "prototypes": {"heading": 1, "caption": 1, "analysis": 1},
        "screenshots": str(shots),
        "groups": [{"title": "题目", "items": [{"cell": 0, "analysis": "看输出。"}]}],
        "cover_edits": [],
        "max_width_cm": 10,
        "image_mode": "joint",
    }
    p = tmp_path / "cfg.json"
    write(p, cfg)
    build(argparse.Namespace(config=str(p)))
    doc = Document(tmp_path / "out.docx")
    assert len(doc.inline_shapes) == 1
    texts = [x.text for tb in doc.tables for c in tb._cells for x in c.paragraphs]
    assert any(x.startswith("代码与运行结果截图：") for x in texts)
    audit(argparse.Namespace(config=str(p), out=str(tmp_path / "audit.json")))
    assert read(tmp_path / "audit.json")["passed"] is True
    # 清单指向了错误的 cell 时必须失败，而不是静默通过
    rec["sourceFirstLine"] = "print(2)"
    write(shots, {"cells": [rec]})
    with pytest.raises(AssertionError, match="different cell"):
        audit(argparse.Namespace(config=str(p), out=str(tmp_path / "audit2.json")))
