# -*- coding: utf-8 -*-
"""Detect and repair common layout problems in a .docx.

Dry-run by default; pass --apply to write the fixed file (-o out.docx or --in-place).

Fixes:
  empty-paragraphs   collapse >=2 consecutive empty body paragraphs to --max-keep
  heading-keepnext   headings get keepNext+keepLines (no orphan headings)
  image-aspect       restore distorted images to native aspect ratio
  manual-pagenum     digits-only footer paragraphs -> real PAGE field
  fake-headings      bold+centered short paragraphs -> apply --fake-heading-style

Usage:
  python repair_docx.py input.docx --fix empty-paragraphs,heading-keepnext [--apply] [-o out.docx]
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import (  # noqa: E402
    all_, body_blocks, dump_json, first, is_empty_para, iter_all,
    setup_stdout, style_paragraph,
)
from docx import Document  # noqa: E402

NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

HEADING_RE = re.compile(r"(heading|标题)", re.I)


def fix_empty_paragraphs(doc, max_keep: int) -> list[str]:
    log = []
    body = doc.element.body
    streak = []
    removed = 0
    for child in list(body.iterchildren()):
        if child.tag == qn("w:p"):
            p = Paragraph(child, None)
            if is_empty_para(p) and not all_(child, f".//{qn('w:drawing')}"):
                streak.append(child)
                continue
            if len(streak) > max_keep:
                for el in streak[max_keep:]:
                    body.remove(el)
                    removed += 1
            streak = []
        elif child.tag == qn("w:tbl"):
            if len(streak) > max_keep:
                for el in streak[max_keep:]:
                    body.remove(el)
                    removed += 1
            streak = []
    # trailing empties at end of body
    children = [c for c in body.iterchildren() if c.tag == qn("w:p")]
    tail = 0
    for child in reversed(children):
        if is_empty_para(Paragraph(child, None)) and not all_(child, f".//{qn('w:drawing')}"):
            tail += 1
        else:
            break
    if tail > max_keep:
        for el in children[-tail:][-tail + max_keep:]:
            body.remove(el)
            removed += 1
    log.append(f"empty-paragraphs: removed {removed} empty paragraphs (kept max {max_keep} per gap)")
    return log


def is_heading_para(p: Paragraph) -> bool:
    try:
        name = p.style.name or ""
    except Exception:
        return False
    if HEADING_RE.search(name):
        return True
    try:
        ol = style_paragraph(p.style).get("outline_level")
        return ol is not None and 0 <= ol <= 8
    except Exception:
        return False


def fix_heading_keepnext(doc) -> list[str]:
    log = []
    count = 0
    for kind, obj in body_blocks(doc):
        if kind != "p":
            continue
        if is_heading_para(obj):
            ppr = obj._p.get_or_add_pPr()
            for tag in ("w:keepNext", "w:keepLines"):
                if first(ppr, qn(tag)) is None:
                    el = OxmlElement(tag)
                    ppr.insert(0, el)
                    count += 1
    log.append(f"heading-keepnext: set keepNext/keepLines on {count} properties")
    return log


def fix_image_aspect(doc) -> list[str]:
    log = []
    fixed = 0
    try:
        from PIL import Image
        import io
    except ImportError:
        return ["image-aspect: PIL unavailable, skipped"]
    for drawing in iter_all(doc.element.body, qn("w:drawing")):
        blip = first(drawing, f".//{{{NS_A}}}blip")
        ext = first(drawing, f".//{{{NS_WP}}}extent")
        xfrm_ext = first(drawing, f".//{{{NS_A}}}xfrm/{{{NS_A}}}ext")
        if blip is None or ext is None:
            continue
        rid = blip.get(f"{{{NS_R}}}embed")
        if not rid:
            continue
        try:
            part = doc.part.rels[rid].target_part
            img = Image.open(io.BytesIO(part.blob))
            native_ar = img.width / img.height
        except Exception:
            continue
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        if cx <= 0 or cy <= 0:
            continue
        disp_ar = cx / cy
        if abs(native_ar - disp_ar) / native_ar <= 0.02:
            continue
        new_cy = int(round(cx / native_ar))
        ext.set("cy", str(new_cy))
        if xfrm_ext is not None:
            xfrm_ext.set("cy", str(new_cy))
        fixed += 1
    log.append(f"image-aspect: corrected {fixed} distorted images (width kept, height recomputed)")
    return log


def fix_manual_pagenum(doc) -> list[str]:
    log = []
    count = 0
    for sec in doc.sections:
        for hf_holder in (sec.footer, sec.first_page_footer, sec.even_page_footer):
            try:
                if hf_holder is None or hf_holder.is_linked_to_previous:
                    continue
            except Exception:
                continue
            for p in hf_holder.paragraphs:
                t = (p.text or "").strip()
                if t.isdigit() and not para_fields_safe(p):
                    for r in list(p._p.findall(qn("w:r"))):
                        p._p.remove(r)
                    fld = OxmlElement("w:fldSimple")
                    fld.set(qn("w:instr"), " PAGE   \\* MERGEFORMAT ")
                    run = OxmlElement("w:r")
                    rpr = OxmlElement("w:rPr")
                    run.append(rpr)
                    fld.append(run)
                    p._p.append(fld)
                    count += 1
    log.append(f"manual-pagenum: converted {count} hand-typed page numbers to PAGE fields")
    return log


def para_fields_safe(p):
    from _docx_utils import para_fields
    try:
        return para_fields(p._p)
    except Exception:
        return []


def detect_fake_headings(doc) -> list[str]:
    """Bold + centered + short, not a heading style, followed by body text."""
    out = []
    blocks = [b for b in body_blocks(doc)]
    for i, (kind, obj) in enumerate(blocks):
        if kind != "p":
            continue
        p = obj
        txt = (p.text or "").strip()
        if not txt or len(txt) > 25 or is_heading_para(p):
            continue
        ppr = first(p._p, qn("w:pPr"))
        jc = first(ppr, qn("w:jc")) if ppr is not None else None
        if jc is None or jc.get(qn("w:val")) != "center":
            continue
        bold = any(r.font.bold for r in p.runs if r.font.bold is not None)
        if not bold:
            continue
        nxt = None
        for j in range(i + 1, len(blocks)):
            if blocks[j][0] == "p":
                t2 = (blocks[j][1].text or "").strip()
                if t2:
                    nxt = t2
                break
            break
        if nxt:
            out.append(f"body#{i}: {txt}  (后接正文: {nxt[:20]}...)")
    return out


def fix_fake_headings(doc, style_name: str) -> list[str]:
    log = []
    try:
        st = doc.styles[style_name]
    except KeyError:
        return [f"fake-headings: 样式 {style_name} 不存在,未修改"]
    count = 0
    blocks = list(body_blocks(doc))
    for i, (kind, obj) in enumerate(blocks):
        if kind != "p":
            continue
        p = obj
        txt = (p.text or "").strip()
        if not txt or len(txt) > 25 or is_heading_para(p):
            continue
        ppr = first(p._p, qn("w:pPr"))
        jc = first(ppr, qn("w:jc")) if ppr is not None else None
        if jc is None or jc.get(qn("w:val")) != "center":
            continue
        if not any(r.font.bold for r in p.runs if r.font.bold is not None):
            continue
        # look ahead: must be followed by non-empty body paragraph
        has_next_body = False
        for j in range(i + 1, len(blocks)):
            if blocks[j][0] == "p":
                has_next_body = bool((blocks[j][1].text or "").strip())
                break
            break
        if has_next_body:
            p.style = st
            count += 1
    log.append(f"fake-headings: applied style '{style_name}' to {count} paragraphs")
    return log


def main() -> int:
    setup_stdout()
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--fix", default="empty-paragraphs,heading-keepnext",
                    help="comma list: empty-paragraphs,heading-keepnext,image-aspect,manual-pagenum,fake-headings")
    ap.add_argument("--max-keep", type=int, default=1)
    ap.add_argument("--fake-heading-style", default=None)
    ap.add_argument("--apply", action="store_true", help="actually write the fixed file")
    ap.add_argument("-o", "--output", default=None)
    ap.add_argument("--in-place", action="store_true")
    args = ap.parse_args()

    doc = Document(args.input)
    fixes = [f.strip() for f in args.fix.split(",") if f.strip()]
    logs: list[str] = []
    candidates = detect_fake_headings(doc)  # always report

    for fix in fixes:
        if fix == "empty-paragraphs":
            logs.extend(fix_empty_paragraphs(doc, args.max_keep))
        elif fix == "heading-keepnext":
            logs.extend(fix_heading_keepnext(doc))
        elif fix == "image-aspect":
            logs.extend(fix_image_aspect(doc))
        elif fix == "manual-pagenum":
            logs.extend(fix_manual_pagenum(doc))
        elif fix == "fake-headings":
            if args.fake_heading_style:
                logs.extend(fix_fake_headings(doc, args.fake_heading_style))
            else:
                logs.append("fake-headings: 需要 --fake-heading-style 指定要应用的样式,未修改")
        else:
            logs.append(f"unknown fix: {fix}")

    print(f"=== repair plan for {args.input} ===")
    for line in logs:
        print(f"  {line}")
    if candidates:
        print("疑似假标题(如需转换,使用 --fix fake-headings --fake-heading-style '标题 2'):")
        for c in candidates[:10]:
            print(f"  {c}")
    if args.apply:
        out = args.output or (args.input if args.in_place else
                              str(Path(args.input).with_name(Path(args.input).stem + "-fixed.docx")))
        doc.save(out)
        print(f"[applied] saved -> {out}")
    else:
        print("[dry-run] 未写入任何文件;确认后加 --apply")
    dump_json({"logs": logs, "fake_heading_candidates": candidates, "applied": bool(args.apply)}, None)
    return 0


if __name__ == "__main__":
    sys.exit(main())
