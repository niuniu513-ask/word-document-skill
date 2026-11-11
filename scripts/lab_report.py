"""Deterministic helpers for strict-template notebook reports. See workflows/lab-report-runbook.md."""

import argparse
import hashlib
import json
import math
from copy import deepcopy
from pathlib import Path


def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write(path, data):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf8")


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def inventory(args):
    from docx import Document

    d = Document(args.template)

    def paragraphs(ps):
        return [
            {
                "index": i,
                "text": p.text,
                "style": p.style.name,
                "pPr": p._p.pPr.xml if p._p.pPr is not None else None,
                "runs": [
                    {
                        "index": j,
                        "text": r.text,
                        "size_pt": r.font.size.pt if r.font.size else None,
                        "font": r.font.name,
                        "underline": r.underline,
                        "rPr": r._r.rPr.xml if r._r.rPr is not None else None,
                    }
                    for j, r in enumerate(p.runs)
                ],
            }
            for i, p in enumerate(ps)
        ]

    data = {"body": paragraphs(d.paragraphs), "tables": []}
    for ti, t in enumerate(d.tables):
        seen = set()
        for ri, row in enumerate(t.rows):
            for ci, c in enumerate(row.cells):
                if c._tc in seen:
                    continue
                seen.add(c._tc)
                data["tables"].append(
                    {
                        "table": ti,
                        "row": ri,
                        "column": ci,
                        "paragraphs": paragraphs(c.paragraphs),
                        "tblPr": t._tbl.tblPr.xml,
                    }
                )
    write(args.out, data)
    print(args.out)


def frames(args):
    import fitz

    def measure(path):
        result = []
        with fitz.open(path) as pdf:
            for i, p in enumerate(pdf):
                xs = []
                for drawing in p.get_drawings():
                    r = drawing["rect"]
                    # constrain to known frame x values, excluding chart axes / illustration lines
                    if r.height >= args.min_height and r.width <= 1:
                        x = (r.x0 + r.x1) / 2
                        if min(abs(x - args.left), abs(x - args.right)) <= args.tolerance:
                            xs.append(round(x, 3))
                result.append({"page": i + 1, "vertical_x": sorted(set(xs))})
        return result

    data = {"template": measure(args.template_pdf), "final": measure(args.final_pdf)}
    failures = []
    for page in data["final"]:
        for target in [args.left, args.right]:
            if not any(abs(x - target) <= args.tolerance for x in page["vertical_x"]):
                failures.append({"page": page["page"], "missing_x": target})
    data["failures"] = failures
    write(args.out, data)
    if failures:
        raise ValueError(
            "Frame line not detected on some pages; inspect geometry and actual page images"
        )
    print(args.out)


def prepare(args):
    import nbformat

    nb = nbformat.read(args.notebook, 4)
    if args.execute:
        from nbclient import NotebookClient

        sources = [c.source for c in nb.cells]
        NotebookClient(
            nb,
            timeout=args.timeout,
            kernel_name=args.kernel,
            resources={"metadata": {"path": str(Path(args.notebook).resolve().parent)}},
        ).execute()
        assert sources == [c.source for c in nb.cells]
        executed = Path(args.out) / "executed.ipynb"
        executed.parent.mkdir(parents=True, exist_ok=True)
        nbformat.write(nb, executed)
    else:
        executed = Path(args.notebook).resolve()
    root = Path(args.out).resolve()
    root.mkdir(parents=True, exist_ok=True)
    records = []
    for index, c in enumerate(nb.cells):
        if c.cell_type != "code" or not c.source.strip():
            continue
        if c.execution_count is None or any(o.output_type == "error" for o in c.outputs):
            raise ValueError(f"Cell {index}: unexecuted or error; repair before capture")
        path = root / f"cell-{index:03}.ipynb"
        evidence = nbformat.v4.new_notebook(
            metadata=deepcopy(nb.metadata),
            cells=[deepcopy(c), nbformat.v4.new_markdown_cell("END OF EVIDENCE CELL")],
        )
        nbformat.write(evidence, path)
        records.append(
            {
                "cell": index,
                "source": c.source,
                "execution_count": c.execution_count,
                "hasOutput": bool(c.outputs),
                "evidence_notebook": str(path),
            }
        )
    write(
        root / "capture-plan.json",
        {"notebook": str(executed), "notebook_sha256": digest(executed), "cells": records},
    )
    print(root / "capture-plan.json")


def crop(args):
    from PIL import Image

    data = read(args.manifest)
    root = Path(args.out).resolve()
    root.mkdir(parents=True, exist_ok=True)
    records = data["cells"] if isinstance(data, dict) else data
    for item in records:
        for kind in ["code", "output"]:
            if kind == "output" and not item["hasOutput"]:
                continue
            raw = Path(item[kind + "File"])
            im = Image.open(raw)
            b = item[kind + "Box"]
            env = item["env"]
            # Explicit calibrated scale wins: Electron can add white canvas padding.
            scale = item.get("screenshotScale", {})
            sx = scale.get("x", im.width / env["w"])
            sy = scale.get("y", im.height / env["h"])
            box = [
                math.floor(b["x"] * sx),
                math.floor(b["y"] * sy),
                math.ceil((b["x"] + b["width"]) * sx),
                math.ceil((b["y"] + b["height"]) * sy),
            ]
            if (
                box[0] < 0
                or box[1] < 0
                or box[2] > im.width
                or box[3] > im.height
                or box[2] <= box[0]
                or box[3] <= box[1]
            ):
                raise ValueError(f"Invalid capture box: {box}")
            out = root / f"{item['cell']:03}-{kind}.png"
            # Conservative crop: no threshold trimming of dim unused imports or pale plot labels.
            im.crop(box).save(out)
            item[kind + "Crop"] = str(out)
            item[kind + "Audit"] = {
                "raw_sha256": digest(raw),
                "output_sha256": digest(out),
                "crop_ltrb_px": box,
            }
    write(root / "screenshots-cropped.json", data)
    print(root / "screenshots-cropped.json")


def container(doc, spec):
    if spec["kind"] == "body":
        return doc
    if spec["kind"] == "cell":
        return doc.tables[spec["table"]].cell(spec["row"], spec["column"])
    raise ValueError("container.kind must be body or cell")


def build(args):
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Cm
    from docx.text.paragraph import Paragraph
    from PIL import Image

    cfg = read(args.config)
    d = Document(cfg["template"])
    target = container(d, cfg["container"])
    originals = [deepcopy(p._p) for p in target.paragraphs]
    section_xml = d._element.body.sectPr.xml
    for edit in cfg.get("cover_edits", []):
        p = d.paragraphs[edit["paragraph"]]
        r = p.runs[edit["run"]]
        if r.text != edit["expected"]:
            raise ValueError(f"Cover run mismatch: {edit}")
        r.text = edit["text"]
    keep = cfg["keep_prefix"]
    fixed = [p.text for p in target.paragraphs[:keep]]
    if cfg["container"]["kind"] == "cell":
        if (
            len(d.tables[cfg["container"]["table"]].rows) != 1
            or len(d.tables[cfg["container"]["table"]].columns) != 1
        ):
            raise ValueError("Strict helper only supports a single-cell frame")
        if target._tc.xpath("./w:tbl"):
            raise ValueError("Nested data table: use anchor editing instead")
    elif d._element.body.xpath("./w:tbl"):
        raise ValueError("Mixed body with tables: use anchor editing instead")
    for p in list(target.paragraphs)[keep:]:
        p._p.getparent().remove(p._p)
    styles = cfg["prototypes"]

    def add(text, role, keep_next=False):
        proto = originals[styles[role]]
        el = deepcopy(proto)
        for child in list(el):
            if child.tag != qn("w:pPr"):
                el.remove(child)
        if cfg["container"]["kind"] == "body":
            d._element.body.insert(len(d._element.body) - 1, el)
        else:
            target._tc.append(el)
        p = Paragraph(el, target)
        r = p.add_run(text)
        rp = proto.find(".//" + qn("w:rPr"))
        if rp is not None:
            r._r.insert(0, deepcopy(rp))
        if role in cfg.get("clear_instruction_format", []):
            for tag in ["color", "highlight", "shd"]:
                if r._r.rPr is not None:
                    for e in list(r._r.rPr.findall(qn("w:" + tag))):
                        r._r.rPr.remove(e)
        p.paragraph_format.keep_with_next = keep_next
        return p

    manifest = read(cfg["screenshots"])
    records = manifest["cells"] if isinstance(manifest, dict) else manifest
    shots = {s["cell"]: s for s in records}
    used = []
    joint = cfg.get("image_mode") == "joint"
    joint_label = cfg.get("joint_label", "代码与运行结果截图：")

    def picture(label, path):
        p = add(label, "caption")
        p.paragraph_format.keep_together = True
        p.add_run().add_break()
        w, h = Image.open(path).size
        width = min(
            cfg["max_width_cm"],
            w / cfg.get("pixels_per_cm", 55),
            cfg.get("max_height_cm", 9) * w / h,
        )
        p.add_run().add_picture(path, width=Cm(width))

    for group in cfg["groups"]:
        add(group["title"], "heading", True)
        for task in group["items"]:
            s = shots[task["cell"]]
            used.append(task["cell"])
            if task.get("subtitle"):
                add(task["subtitle"], "heading", True)
            if joint:
                # 连体图：代码与输出在同一张里，一个单元格一个标签
                if not s.get("cellCrop"):
                    raise ValueError(f"cell {task['cell']}: image_mode=joint 但清单没有 cellCrop")
                picture(joint_label, s["cellCrop"])
            else:
                picture("代码截图：", s["codeCrop"])
                if s["hasOutput"]:
                    picture("实验结果截图：", s["outputCrop"])
            if not task.get("analysis", "").strip():
                raise ValueError("Missing analysis")
            add("结果分析：" + task["analysis"], "analysis")
    for question in cfg.get("questions", []):
        if question.get("section"):
            add(question["section"], "section", True)
        if "template_paragraph" in question:
            el = deepcopy(originals[question["template_paragraph"]])
            if cfg["container"]["kind"] == "cell":
                target._tc.append(el)
            else:
                d._element.body.insert(len(d._element.body) - 1, el)
        else:
            add(question["text"], "question", True)
        if not question.get("answer", "").strip():
            raise ValueError("Missing answer")
        add("答：" + question["answer"], "analysis")
    if len(used) != len(set(used)) or set(used) != set(shots):
        raise ValueError("Cell coverage incomplete/duplicated")
    assert [p.text for p in target.paragraphs[:keep]] == fixed
    assert d._element.body.sectPr.xml == section_xml
    output = Path(cfg["output"])
    output.parent.mkdir(parents=True, exist_ok=True)
    d.save(output)
    write(
        output.with_suffix(".build.json"),
        {
            "template_sha256": digest(cfg["template"]),
            "config_sha256": digest(args.config),
            "images": len(d.inline_shapes),
            "cells": used,
            "fixed_prefix_unchanged": True,
        },
    )
    print(output)


def audit(args):
    import nbformat
    from docx import Document

    cfg = read(args.config)
    t = Document(cfg["template"])
    d = Document(cfg["output"])
    a = container(t, cfg["container"])
    b = container(d, cfg["container"])
    n = cfg["keep_prefix"]
    assert t._element.body.sectPr.xml == d._element.body.sectPr.xml, "section changed"
    assert [p.text for p in a.paragraphs[:n]] == [p.text for p in b.paragraphs[:n]], (
        "fixed text changed"
    )
    if cfg["container"]["kind"] == "cell":
        assert (
            t.tables[cfg["container"]["table"]]._tbl.tblPr.xml
            == d.tables[cfg["container"]["table"]]._tbl.tblPr.xml
        ), "frame changed"
    m = read(cfg["screenshots"])
    records = m["cells"] if isinstance(m, dict) else m
    nb = nbformat.read(cfg["notebook"], 4)
    for s in records:
        c = nb.cells[s["cell"]]
        if s.get("evidence_notebook"):
            e = nbformat.read(s["evidence_notebook"], 4).cells[0]
            assert c.source == e.source == "".join(s["source"])
            assert c.outputs == e.outputs
        else:
            # 离屏渲染直接用执行后的 notebook，没有证据副本：用首行+源码文本对账
            assert (
                s.get("sourceFirstLine", "").strip() == (c.source or "").strip().split("\n")[0]
            ), "manifest points at a different cell"
            if s.get("source") is not None:
                assert s["source"].strip() == (c.source or "").strip(), (
                    "manifest source differs from notebook"
                )
            assert bool(c.outputs) == bool(s["hasOutput"]), "hasOutput mismatch"
        assert c.execution_count is not None
        assert not any(o.output_type == "error" for o in c.outputs)
        for kind in ["code", "output"]:
            if kind == "output" and not s["hasOutput"]:
                continue
            assert digest(s[kind + "Crop"]) == s[kind + "Audit"]["output_sha256"]
            if s.get(kind + "File"):
                assert digest(s[kind + "File"]) == s[kind + "Audit"]["raw_sha256"]
    if cfg.get("image_mode") == "joint":
        expected = {cfg.get("joint_label", "代码与运行结果截图："): len(records)}
    else:
        expected = {"代码截图": len(records), "实验结果截图": sum(s["hasOutput"] for s in records)}
    actual = {label: 0 for label in expected}
    for p in b.paragraphs:
        for label in actual:
            probe = label if label.endswith("：") else label + "："
            if p.text.startswith(probe):
                actual[label] += 1
                assert p._p.xpath(".//w:drawing"), "label separated from image"
    assert actual == expected and len(d.inline_shapes) == sum(expected.values())
    write(
        args.out,
        {
            "passed": True,
            "labels": actual,
            "images": len(d.inline_shapes),
            "source_output_hashes_verified": True,
            "template_structure_preserved": True,
            "visual_review": "REQUIRED: automated checks cannot assess all screenshot readability/pagination",
        },
    )
    print(args.out)


def contacts(args):
    from PIL import Image, ImageDraw

    files = sorted(Path(args.directory).glob(args.pattern))
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    if not files:
        raise ValueError("No images found")
    for start in range(0, len(files), 4):
        canvas = Image.new("RGB", (1400, 2060), "#dddddd")
        draw = ImageDraw.Draw(canvas)
        for j, path in enumerate(files[start : start + 4]):
            im = Image.open(path)
            im.thumbnail((690, 990))
            x = j % 2 * 700
            y = j // 2 * 1030
            canvas.paste(im, (x, y + 25))
            draw.text((x + 8, y + 6), path.name, fill="black")
        canvas.save(out / f"contact-{start // 4 + 1:03}.jpg")
    print(f"{len(files)} images; inspect every contact sheet and zoom uncertain pages")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="command", required=True)
    p = sub.add_parser("inventory")
    p.add_argument("template")
    p.add_argument("--out", required=True)
    p.set_defaults(func=inventory)
    p = sub.add_parser("frames")
    p.add_argument("template_pdf")
    p.add_argument("final_pdf")
    p.add_argument("--left", type=float, required=True)
    p.add_argument("--right", type=float, required=True)
    p.add_argument("--tolerance", type=float, default=0.5)
    p.add_argument("--min-height", type=float, default=15)
    p.add_argument("--out", required=True)
    p.set_defaults(func=frames)
    p = sub.add_parser("prepare")
    p.add_argument("notebook")
    p.add_argument("--out", required=True)
    p.add_argument("--execute", action="store_true")
    p.add_argument("--timeout", type=int, default=600)
    p.add_argument("--kernel", default="python3")
    p.set_defaults(func=prepare)
    p = sub.add_parser("crop")
    p.add_argument("manifest")
    p.add_argument("--out", required=True)
    p.set_defaults(func=crop)
    for mode in ["build", "audit"]:
        p = sub.add_parser(mode)
        p.add_argument("config")
        p.set_defaults(func=build if mode == "build" else audit)
        if mode == "audit":
            p.add_argument("--out", required=True)
    p = sub.add_parser("contacts")
    p.add_argument("directory")
    p.add_argument("--pattern", default="page-*.png")
    p.add_argument("--out", required=True)
    p.set_defaults(func=contacts)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
