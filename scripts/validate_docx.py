# -*- coding: utf-8 -*-
"""Validate docx outputs.

Mode 1 -- Format Delta (guard "only change what the user asked for"):
    python validate_docx.py --original template.docx --final out.docx [-o delta.json]

Mode 2 -- Structural sanity check of a single document:
    python validate_docx.py --check final.docx [-o report.json]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import (  # noqa: E402
    all_, body_blocks, dump_json, first, is_empty_para, package_flags,
    para_fields, setup_stdout, style_font, style_paragraph,
)
from docx import Document  # noqa: E402

CM_TOL = 0.06  # cm tolerance for margin/page comparisons


def cm(v):
    return round(int(v) / 360000.0, 2) if v is not None else None


def near(a, b, tol=CM_TOL):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def section_fingerprint(doc) -> list[dict]:
    out = []
    for sec in doc.sections:
        out.append({
            "page": (cm(sec.page_width), cm(sec.page_height)),
            "margins": (cm(sec.top_margin), cm(sec.bottom_margin), cm(sec.left_margin), cm(sec.right_margin)),
            "hf_dist": (cm(sec.header_distance), cm(sec.footer_distance)),
            "first_page_diff": bool(sec.different_first_page_header_footer),
        })
    return out


def header_footer_fingerprint(doc) -> list[dict]:
    out = []
    for sec in doc.sections:
        entry = {}
        for kind in ("header", "footer"):
            try:
                hf = getattr(sec, kind)
                texts = []
                for p in hf.paragraphs:
                    t = (p.text or "").strip()
                    f = para_fields(p._p)
                    if t or f:
                        texts.append((t, tuple(f)))
            except Exception:
                texts = []
            entry[kind] = texts
        out.append(entry)
    return out


def styles_fingerprint(doc) -> dict:
    out = {}
    for st in doc.styles:
        try:
            if st.type is None:
                continue
            key = st.style_id
            entry = {"name": st.name}
            if str(st.type).endswith("PARAGRAPH") or str(st.type).endswith("CHARACTER"):
                entry["font"] = style_font(st)
            if str(st.type).endswith("PARAGRAPH"):
                entry["paragraph"] = style_paragraph(st)
            out[key] = entry
        except Exception:
            continue
    return out


def body_text_lines(doc) -> list[str]:
    lines = []
    for kind, obj in body_blocks(doc):
        if kind == "p":
            t = (obj.text or "").strip()
            if t:
                lines.append(t)
        else:
            try:
                for r in obj.rows:
                    lines.append("\t".join(c.text.strip() for c in r.cells))
            except Exception:
                pass
    return lines


def compare_format(orig: Document, final: Document) -> dict:
    diffs = []
    so, sf = section_fingerprint(orig), section_fingerprint(final)
    if len(so) != len(sf):
        diffs.append({"category": "sections", "what": f"节数 {len(so)} -> {len(sf)}"})
    for i, (a, b) in enumerate(zip(so, sf)):
        for key in ("page", "margins", "hf_dist"):
            if not near_all(a[key], b[key]):
                diffs.append({"category": "sections", "what": f"节{i + 1} {key}: {a[key]} -> {b[key]}"})
        if a["first_page_diff"] != b["first_page_diff"]:
            diffs.append({"category": "sections", "what": f"节{i + 1} 首页不同设置改变"})

    ho, hf = header_footer_fingerprint(orig), header_footer_fingerprint(final)
    for i, (a, b) in enumerate(zip(ho, hf)):
        if a != b:
            diffs.append({"category": "headers_footers",
                          "what": f"节{i + 1} 页眉/页脚内容改变: {a} -> {b}"})

    sto, stf = styles_fingerprint(orig), styles_fingerprint(final)
    for sid, before in sto.items():
        after = stf.get(sid)
        if after is None:
            diffs.append({"category": "styles", "what": f"样式被删除: {before.get('name')}"})
            continue
        for part in ("font", "paragraph"):
            if before.get(part) != after.get(part):
                changed = {k: (before[part].get(k), after[part].get(k))
                           for k in set(before[part]) | set(after[part])
                           if before[part].get(k) != after[part].get(k)}
                if changed:
                    diffs.append({"category": "styles", "what": f"样式 {before.get('name')} {part} 改变: {changed}"})
    for sid, after in stf.items():
        if sid not in sto:
            diffs.append({"category": "styles", "what": f"新增样式: {after.get('name')}"})

    to, tf = len(orig.tables), len(final.tables)
    if to != tf:
        diffs.append({"category": "tables", "what": f"表格数 {to} -> {tf}"})
    for i, (a, b) in enumerate(zip(orig.tables, final.tables)):
        try:
            if (len(a.rows), len(a.columns)) != (len(b.rows), len(b.columns)):
                diffs.append({"category": "tables",
                              "what": f"表格{i + 1} 尺寸 {len(a.rows)}x{len(a.columns)} -> {len(b.rows)}x{len(b.columns)}"})
        except Exception:
            pass

    io_ = media_count(orig)
    if io_ != media_count(final):
        diffs.append({"category": "images", "what": f"图片文件数 {io_} -> {media_count(final)}"})

    return {"format_delta_count": len(diffs), "format_delta": diffs}


def near_all(a: tuple, b: tuple) -> bool:
    return len(a) == len(b) and all(near(x, y) for x, y in zip(a, b))


def media_count(doc):
    return len([n for n in doc.part.package.iter_parts() if "media/" in str(n.partname)])


def check_document(doc) -> dict:
    problems = []
    blocks = [
        {"kind": k, "obj": o, "p": o if k == "p" else None, "index": i}
        for i, (k, o) in enumerate(body_blocks(doc))
    ]
    streak = 0
    for b in blocks:
        if b["kind"] == "p" and is_empty_para(b["p"]):
            streak += 1
            if streak == 3:
                problems.append({"severity": "medium", "issue": "连续>=3个空段落(用回车排版)",
                                 "where": f"body index {b['index']}"})
        else:
            streak = 0
    # distorted images
    from analyze_docx import collect_images
    imgs = collect_images(doc, [{"kind": "p", "p": b["p"], "index": b["index"]} if b["kind"] == "p" else {"kind": b["kind"], "obj": b["obj"], "index": b["index"]} for b in blocks])
    for img in imgs:
        if img["distorted"]:
            problems.append({"severity": "high", "issue": f"图片#{img['index']}显示比例失真",
                             "where": f"paragraph {img['paragraph_index']}"})
    # headings without keep_with_next
    for b in blocks:
        if b["kind"] != "p":
            continue
        p = b["p"]
        name = ""
        ol = None
        try:
            name = p.style.name or ""
            ol = style_paragraph(p.style).get("outline_level")
        except Exception:
            continue
        import re as _re
        if (ol is not None and 0 <= ol <= 8) or _re.search(r"(heading|标题)", name, _re.I):
            kwn = first(p._p, f"{qn('w:pPr')}/{qn('w:keepNext')}") is not None
            skwn = style_paragraph(p.style).get("keep_with_next")
            if not kwn and not skwn:
                problems.append({"severity": "medium", "issue": f"标题缺少 keep_with_next(可能孤立在页尾): {p.text[:24]}",
                                 "where": f"body index {b['index']}"})
    # manual page numbers in footers
    for si, sec in enumerate(doc.sections):
        try:
            for p in sec.footer.paragraphs:
                t = (p.text or "").strip()
                if t.isdigit() and not para_fields(p._p):
                    problems.append({"severity": "high",
                                     "issue": f"页脚使用手打页码 '{t}' 而非 PAGE 域",
                                     "where": f"section {si}"})
        except Exception:
            pass
    # tables wider than content area
    for si, sec in enumerate(doc.sections):
        content_w = (cm(sec.page_width) or 21) - (cm(sec.left_margin) or 3.18) - (cm(sec.right_margin) or 3.18)
        break
    else:
        content_w = 14.6
    for i, t in enumerate(doc.tables):
        try:
            width = sum(cm(gc.get(qn("w:w"))) or 0 for gc in all_(first(t._tbl, qn("w:tblGrid")), qn("w:gridCol")))
            if width and width > content_w + 0.3:
                problems.append({"severity": "medium", "issue": f"表格{i + 1}总宽 {width}cm 超出版心 {content_w}cm",
                                 "where": f"table {i}"})
        except Exception:
            pass
    return {"problem_count": len(problems), "problems": problems}


def main() -> int:
    setup_stdout()
    ap = argparse.ArgumentParser()
    ap.add_argument("--original")
    ap.add_argument("--final")
    ap.add_argument("--check")
    ap.add_argument("-o", "--output", default=None)
    args = ap.parse_args()

    if args.check:
        doc = Document(args.check)
        result = {"mode": "check", "file": args.check, **check_document(doc)}
        print(f"=== check: {args.check} ===")
        print(f"problems: {result['problem_count']}")
        for p in result["problems"]:
            print(f"  [{p['severity']}] {p['issue']} ({p['where']})")
        dump_json(result, args.output)
        return 0

    if not (args.original and args.final):
        print("ERROR: 需要 --original + --final,或 --check")
        return 2
    orig, final = Document(args.original), Document(args.final)
    result = {"mode": "delta", "original": args.original, "final": args.final,
              **compare_format(orig, final)}
    po, pf = body_text_lines(orig), body_text_lines(final)
    result["content_delta"] = {
        "lines_original": len(po), "lines_final": len(pf),
        "note": "lines are non-empty paragraphs and table rows",
    }

    print(f"=== format delta: {args.original} -> {args.final} ===")
    print(f"formatting differences: {result['format_delta_count']}")
    for d in result["format_delta"][:20]:
        print(f"  [{d['category']}] {d['what']}")
    if result["format_delta_count"] > 20:
        print(f"  ... and {result['format_delta_count'] - 20} more")
    print(f"body lines: {len(po)} -> {len(pf)}")
    dump_json(result, args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
