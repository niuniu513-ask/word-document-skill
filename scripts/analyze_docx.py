"""Analyze a .docx and emit a Document Specification (JSON).

This is the "understand Word first" step of the skill. Run it on EVERY docx
before generating/editing anything. Pure python (no Word COM), fast and safe.

Usage:
    python analyze_docx.py <input.docx> [-o spec.json] [--max-paras 250]

Prints a short human-readable summary; full JSON goes to -o or stdout.
"""
from __future__ import annotations

import argparse
import collections
import datetime
import re
import sys
from pathlib import Path

from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import (  # noqa: E402
    all_,
    body_blocks,
    doc_defaults,
    dump_json,
    emu_to_cm,
    first,
    is_empty_para,
    iter_all,
    load_numbering,
    numbering_counters,
    package_flags,
    para_fields,
    para_first_run_font,
    para_numpr,
    run_font,
    setup_stdout,
    style_font,
    style_paragraph,
    twips_to_cm,
)
from docx import Document  # noqa: E402

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

HEADING_NAME_RE = re.compile(r"(heading|标题|标题\s*\d)", re.I)
MANUAL_NUM_RE = re.compile(
    r"^\s*(?:(\d+(?:\.\d+){0,4})|([一二三四五六七八九十百]+)|((?:\(|（)[一二三四五六七八九十]+(?:\)|）))|(第\s*[一二三四五六七八九十百\d]+\s*[章节]))[、.．,，:：\s]?"
)
LABEL_RE = re.compile(r"(学\s*校|大\s*学|学\s*院|院\s*系|系\s*别|专\s*业|班\s*级|姓\s*名|学\s*号|指导教师|完成日期|日\s*期|题\s*目|课\s*题|课程名称)")


def analyze_section(sec) -> dict:
    def cm(v):
        return emu_to_cm(v) if v is not None else None

    out = {
        "index": None,
        "page": {
            "width_cm": cm(sec.page_width),
            "height_cm": cm(sec.page_height),
            "orientation": str(sec.orientation).split(".")[-1].split(" ")[0] if sec.orientation is not None else None,
        },
        "margins_cm": {
            "top": cm(sec.top_margin), "bottom": cm(sec.bottom_margin),
            "left": cm(sec.left_margin), "right": cm(sec.right_margin),
            "gutter": cm(sec.gutter),
            "header": cm(sec.header_distance), "footer": cm(sec.footer_distance),
        },
        "first_page_different": bool(sec.different_first_page_header_footer),
        "headers": {}, "footers": {},
        "page_number": {},
    }
    for kind in ("header", "footer"):
        for variant in ("", "first_page_", "even_page_"):
            try:
                hf = getattr(sec, f"{variant}{kind}")
            except Exception:
                continue
            key = (variant[:-1] or "default") + ("" if not variant else f"_{kind}")
            key = variant.replace("_", "-") + kind if variant else kind
            info = {"linked_to_previous": bool(hf.is_linked_to_previous)}
            paras = []
            try:
                for p in hf.paragraphs:
                    t = (p.text or "").strip()
                    fields = para_fields(p._p)
                    if t or fields:
                        paras.append({"text": t, "fields": fields, "alignment": str(p.alignment).split(".")[-1].split(" ")[0] if p.alignment is not None else None})
            except Exception:
                pass
            info["content"] = paras
            out[kind + "s"][key] = info
    # page number format from sectPr pgNumType
    try:
        pgnum = first(sec._sectPr, qn("w:pgNumType"))
        if pgnum is not None:
            out["page_number"] = {
                "format": pgnum.get(qn("w:fmt")),
                "start": pgnum.get(qn("w:start")),
                "chapter_sep": pgnum.get(qn("w:chapStyle")),
            }
    except Exception:
        pass
    return out


def collect_images(doc, blocks) -> list[dict]:
    images = []
    media_cache: dict[str, tuple | None] = {}

    def native_size(rid: str):
        if rid in media_cache:
            return media_cache[rid]
        try:
            part = doc.part.rels[rid].target_part
            import io

            from PIL import Image
            img = Image.open(io.BytesIO(part.blob))
            res = (img.width, img.height)
        except Exception:
            res = None
        media_cache[rid] = res
        return res

    for b in blocks:
        if b["kind"] != "p":
            continue
        p_el = b["p"]._p
        for drawing in iter_all(p_el, qn("w:drawing")):
            ext = first(drawing, f".//{{{NS_WP}}}extent")
            inline = first(drawing, f"{{{NS_WP}}}inline") is not None
            blip = first(drawing, f".//{{{NS_A}}}blip")
            rid = blip.get(f"{{{NS_R}}}embed") if blip is not None else None
            entry = {
                "index": len(images), "type": "inline" if inline else "anchored",
                "paragraph_index": b["index"],
                "width_cm": emu_to_cm(ext.get("cx")) if ext is not None else None,
                "height_cm": emu_to_cm(ext.get("cy")) if ext is not None else None,
                "native_px": native_size(rid) if rid else None,
                "distorted": None, "caption_guess": None,
            }
            if entry["width_cm"] and entry["height_cm"] and entry["native_px"]:
                w, h = entry["native_px"]
                if w and h:
                    native_ar = w / h
                    disp_ar = entry["width_cm"] / entry["height_cm"]
                    entry["distorted"] = abs(native_ar - disp_ar) / native_ar > 0.02
            images.append(entry)
        for _pict in all_(p_el, qn("w:pict")):
            images.append({
                "index": len(images), "type": "legacy-vml", "paragraph_index": b["index"],
                "width_cm": None, "height_cm": None, "native_px": None,
                "distorted": None, "caption_guess": None,
            })
    # caption guess: next non-empty paragraph "图N ..." / "Fig" / caption style
    for img in images:
        for j in range(img["paragraph_index"] + 1, len(blocks)):
            if blocks[j]["kind"] == "p" and not is_empty_para(blocks[j]["p"]):
                nxt = blocks[j]["p"]
                txt = (nxt.text or "").strip()
                stylename = nxt.style.name if nxt.style is not None else ""
                if re.match(r"^[图表]\s*\d+", txt) or re.match(r"^(Figure|Fig\.?|Table)\s*\d+", txt, re.I) or "caption" in stylename.lower() or "题注" in stylename:
                    img["caption_guess"] = txt[:60]
                break
    return images


def detect_cover(blocks, images) -> dict:
    """Heuristic cover-page detection from the first blocks."""
    score = 0
    fields = set()
    placeholders = []
    examined = 0
    for b in blocks[:30]:
        if b["kind"] != "p":
            continue
        p = b["p"]
        if is_empty_para(p):
            continue
        examined += 1
        if examined > 18:
            break
        txt = (p.text or "").strip()
        f = para_first_run_font(p) or {}
        size = f.get("size_pt") or 0
        if size >= 16 and len(txt) <= 40:
            score += 2
        if LABEL_RE.search(txt):
            score += 2
            for m in re.finditer(LABEL_RE, txt):
                fields.add(m.group(1))
        if re.search(r"[_＿]{3,}", txt):
            score += 1
            placeholders.append(re.sub(r"\s+", "", txt)[:40])
        if re.search(r"[{{【]\s*(name|title|学校|姓名|[A-Za-z_]+)\s*[}}】]", txt, re.I):
            score += 1
            placeholders.append(txt[:40])
    return {
        "present": score >= 4,
        "confidence": min(score / 10.0, 1.0),
        "fields_detected": sorted(fields),
        "placeholders_preview": placeholders[:8],
    }


def detect_toc(doc, blocks, styles_by_id) -> dict:
    present = False
    via = []
    for b in blocks:
        if b["kind"] != "p":
            continue
        codes = para_fields(b["p"]._p)
        if any(c.startswith("TOC") for c in codes):
            present = True
            via.append("TOC field")
            break
    for st in doc.styles:
        try:
            if st.type is not None and re.match(r"^(TOC|目录)\s*\d$", st.name or "", re.I):
                via.append(f"style {st.name}")
        except Exception:
            pass
    for b in blocks:
        if b["kind"] != "p":
            continue
        sid_el = first(b["p"]._p, f"{qn('w:pPr')}/{qn('w:pStyle')}")
        if sid_el is not None:
            st = styles_by_id.get(sid_el.get(qn("w:val")))
            if st is not None and re.match(r"^(TOC|目录)\s*\d$", st.name or "", re.I):
                present = True
                via.append("paragraphs styled as TOC entries")
                break
    entries = []
    for b in blocks:
        if b["kind"] == "p":
            t = (b["p"].text or "").strip()
            if re.match(r"^.+\s[.\u2026·\s]{3,}\s*\d+\s*$", t) or re.match(r"^(TOC|目录)", t, re.I):
                entries.append(t[:60])
        if len(entries) >= 8:
            break
    return {"present": present, "detected_via": sorted(set(via)), "entries_preview": entries[:8]}


FONT_KEYS = (
    "ascii", "eastAsia", "hAnsi", "cs",
    "ascii_theme", "eastAsia_theme", "hAnsi_theme", "cs_theme",
    "size_pt", "bold", "italic", "underline",
    "color", "color_theme", "color_theme_tint", "color_theme_shade",
)


def effective_run_font(p, run_el) -> dict:
    """Resolve direct run formatting over the paragraph style chain."""
    direct = run_font(run_el)
    inherited = style_font(p.style) if p.style is not None else {}
    out = {}
    for key in FONT_KEYS:
        value = direct.get(key)
        if value is None:
            value = inherited.get(key)
        out[key] = value
    return out


def iter_format_paragraphs(doc):
    """Yield body, table-cell, header and footer paragraphs for font inventory."""
    def add(p, location):
        # Do not use id(lxml_element) here: lxml may create a new proxy for
        # the same XML node. Duplicate linked headers are harmless for a
        # usage inventory and are preferable to silently dropping formats.
        yield p, location

    for p in doc.paragraphs:
        yield from add(p, "body")
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    yield from add(p, f"table[{ti}].cell[{ri},{ci}]")
    for si, sec in enumerate(doc.sections):
        for kind in ("header", "first_page_header", "even_page_header", "footer", "first_page_footer", "even_page_footer"):
            try:
                for p in getattr(sec, kind).paragraphs:
                    yield from add(p, f"section[{si}].{kind}")
            except Exception:
                continue


def font_inventory(doc) -> dict:
    """Summarize fonts and text colors actually used by text runs."""
    profiles = collections.Counter()
    font_counts = collections.Counter()
    color_counts = collections.Counter()
    examples = {}
    run_count = 0
    for p, location in iter_format_paragraphs(doc):
        for run_el in iter_all(p._p, qn("w:r")):
            text = "".join(t.text or "" for t in run_el.iter(qn("w:t")))
            if not text.strip():
                continue
            run_count += 1
            profile = effective_run_font(p, run_el)
            key = tuple(profile.get(k) for k in FONT_KEYS)
            profiles[key] += 1
            font_key = {
                "ascii": profile.get("ascii"), "eastAsia": profile.get("eastAsia"),
                "hAnsi": profile.get("hAnsi"), "cs": profile.get("cs"),
                "ascii_theme": profile.get("ascii_theme"),
                "eastAsia_theme": profile.get("eastAsia_theme"),
                "hAnsi_theme": profile.get("hAnsi_theme"),
                "cs_theme": profile.get("cs_theme"),
                "size_pt": profile.get("size_pt"),
            }
            font_counts[tuple(font_key.items())] += 1
            if profile.get("color_theme"):
                color = f"theme:{profile['color_theme']}"
                if profile.get("color_theme_tint"):
                    color += f";tint:{profile['color_theme_tint']}"
                if profile.get("color_theme_shade"):
                    color += f";shade:{profile['color_theme_shade']}"
                if profile.get("color") and profile["color"].lower() != "auto":
                    color += f";fallback:{profile['color']}"
            else:
                color = profile.get("color") or "auto/default"
            color_counts[color] += 1
            examples.setdefault(key, {"text": text[:40], "location": location})

    def top(counter, fields):
        out = []
        for key, count in counter.most_common():
            item = dict(key) if isinstance(key, tuple) and key and isinstance(key[0], tuple) else {"value": key}
            item["run_count"] = count
            if isinstance(key, tuple) and key in examples:
                item["example"] = examples[key]
            out.append(item)
        return out

    dominant = None
    if profiles:
        key, count = profiles.most_common(1)[0]
        dominant = {k: v for k, v in zip(FONT_KEYS, key, strict=True)}
        dominant["run_count"] = count
        dominant["example"] = examples.get(key)
    return {
        "run_count": run_count,
        "dominant": dominant,
        "font_combinations": top(font_counts, FONT_KEYS),
        "colors": [{"color": color, "run_count": count} for color, count in color_counts.most_common()],
        "mixed_fonts": len(font_counts) > 1,
        "mixed_colors": len(color_counts) > 1,
    }


def manual_number_label(text: str) -> str | None:
    m = MANUAL_NUM_RE.match(text or "")
    if not m:
        return None
    return next(g for g in m.groups() if g)


def main() -> int:
    setup_stdout()
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", default=None)
    ap.add_argument("--max-paras", type=int, default=250)
    args = ap.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"ERROR: file not found: {path}")
        return 2
    doc = Document(str(path))

    styles_by_id = {}
    styles_out = []
    used_style_counts: collections.Counter = collections.Counter()
    for st in doc.styles:
        try:
            entry = {
                "name": st.name,
                "style_id": st.style_id,
                "type": str(st.type).split(".")[-1].split(" ")[0] if st.type is not None else None,
                "builtin": st.builtin,
                "based_on": st.base_style.name if st.base_style is not None else None,
            }
            if entry["type"] in ("PARAGRAPH", "CHARACTER"):
                entry["font"] = style_font(st)
            if entry["type"] == "PARAGRAPH":
                entry["paragraph"] = style_paragraph(st)
            styles_out.append(entry)
            styles_by_id[st.style_id] = st
        except Exception:
            continue

    numbering_defs = load_numbering(doc)
    next_num = numbering_counters()

    flags = package_flags(str(path))

    blocks = []
    for bi, (kind, obj) in enumerate(body_blocks(doc)):
        blocks.append({"kind": kind, "obj": obj, "index": bi, "p": obj if kind == "p" else None})

    # ---- walk paragraphs
    paragraphs = []
    headings = []
    field_types: collections.Counter = collections.Counter()
    empty_streak = 0
    max_empty_streak = 0
    empty_runs_ge3 = 0
    body_style_counts: collections.Counter = collections.Counter()
    for b in blocks:
        if b["kind"] != "p":
            continue
        p = b["p"]
        text = p.text or ""
        stylename = ""
        style_id = None
        outline = None
        try:
            if p.style is not None:
                stylename = p.style.name or ""
                style_id = p.style.style_id
        except Exception:
            pass
        ppr = first(p._p, qn("w:pPr"))
        pstyle_el = first(ppr, qn("w:pStyle")) if ppr is not None else None
        if pstyle_el is not None and pstyle_el.get(qn("w:val")) in styles_by_id:
            outline = style_paragraph(styles_by_id[pstyle_el.get(qn("w:val"))]).get("outline_level")
        empty = not text.strip()
        if empty:
            empty_streak += 1
            if empty_streak >= 3:
                empty_runs_ge3 += 1
            max_empty_streak = max(max_empty_streak, empty_streak)
        else:
            empty_streak = 0

        codes = para_fields(p._p)
        for c in codes:
            field_types[c] += 1

        numinfo = para_numpr(p._p, styles_by_id)
        num_text = None
        if numinfo:
            num_text = next_num(numinfo["num_id"], numinfo["ilvl"], numbering_defs)

        entry = {
            "index": b["index"],
            "style": stylename,
            "empty": empty,
            "text": (text.strip()[:80] if not empty else ""),
            "fields": codes,
        }
        if not empty:
            used_style_counts[style_id or stylename] += 1
            is_heading = bool(
                (outline is not None and 0 <= outline <= 8)
                or HEADING_NAME_RE.search(stylename or "")
            )
            if is_heading and not numinfo:
                pass
            if is_heading:
                lvl = (outline + 1) if outline is not None else (1 if re.search(r"1", stylename) else 2)
                headings.append({
                    "level": lvl, "style": stylename, "text": text.strip()[:80],
                    "native_numbering": num_text, "manual_number": manual_number_label(text.strip()),
                    "index": b["index"],
                })
            else:
                if len(text.strip()) > 15:
                    body_style_counts[style_id or stylename] += 1
        if len(paragraphs) < args.max_paras:
            f = para_first_run_font(p)
            al = None
            ind = first(ppr, qn("w:ind")) if ppr is not None else None
            jc = first(ppr, qn("w:jc")) if ppr is not None else None
            al = jc.get(qn("w:val")) if jc is not None else None
            entry.update({
                "direct": {
                    "alignment": al,
                    "first_line_indent_cm": (twips_to_cm(ind.get(qn("w:firstLine"))) if ind is not None and ind.get(qn("w:firstLine")) else None),
                    "left_indent_cm": (twips_to_cm(ind.get(qn("w:left"))) if ind is not None and ind.get(qn("w:left")) else None),
                },
                "run_font": f,
                "numbering": ({"num_id": numinfo["num_id"], "ilvl": numinfo["ilvl"], "text": num_text} if numinfo else None),
            })
            paragraphs.append(entry)

    # ---- tables
    tables = []
    for b in blocks:
        if b["kind"] != "tbl":
            continue
        t = b["obj"]
        try:
            nrows, ncols = len(t.rows), len(t.columns)
        except Exception:
            nrows = ncols = None
        merged = len(t._tbl.findall(f".//{qn('w:vMerge')}")) + len(t._tbl.findall(f".//{qn('w:hMerge')}"))
        grid = [twips_to_cm(gc.get(qn("w:w"))) for gc in all_(first(t._tbl, qn("w:tblGrid")), qn("w:gridCol"))]
        tstyle = None
        tblpr = first(t._tbl, qn("w:tblPr"))
        if tblpr is not None:
            tsel = first(tblpr, qn("w:tblStyle"))
            tstyle = tsel.get(qn("w:val")) if tsel is not None else None
        hdr_repeat = bool(all_(t._tbl, f"{qn('w:trPr')}/{qn('w:tblHeader')}"))
        preview = []
        try:
            for r in t.rows[:2]:
                preview.append([c.text.strip()[:20] for c in r.cells])
        except Exception:
            pass
        borders = first(tblpr, qn("w:tblBorders")) if tblpr is not None else None
        tables.append({
            "index": len(tables), "body_index": b["index"], "rows": nrows, "cols": ncols,
            "style": tstyle, "col_widths_cm": grid, "merged_cells": merged,
            "header_row_repeat": hdr_repeat,
            "borders": "defined" if borders is not None else "default/none",
            "preview": preview,
        })

    # ---- images, cover, toc
    images = collect_images(doc, blocks)
    cover = detect_cover(blocks, images)
    toc = detect_toc(doc, blocks, styles_by_id)
    formatting_inventory = font_inventory(doc)

    sections = []
    for si, sec in enumerate(doc.sections):
        s = analyze_section(sec)
        s["index"] = si
        sections.append(s)

    # manual page number in footers?
    manual_pagenum = False
    for s in sections:
        for _k, v in s["footers"].items():
            for para in v.get("content", []):
                if re.fullmatch(r"\d+", para["text"]) and not para["fields"]:
                    manual_pagenum = True

    # ---- conventions
    body_style = body_style_counts.most_common(1)
    body_style_entry = None
    if body_style:
        sid = body_style[0][0]
        st = styles_by_id.get(sid)
        if st is not None:
            body_style_entry = {
                "style": st.name, "font": style_font(st), "paragraph": style_paragraph(st),
                "used_by_paragraphs": body_style[0][1],
            }

    issues = []
    if max_empty_streak >= 3:
        issues.append(f"发现连续 {max_empty_streak} 个空段落(疑似用回车排版,共 {empty_runs_ge3} 处)")
    distorted = [i for i in images if i["distorted"]]
    if distorted:
        issues.append(f"{len(distorted)} 张图片显示比例与原图不一致(被拉伸变形)")
    fake = 0
    for b in blocks:
        if b["kind"] != "p":
            continue
        p = b["p"]
        txt = (p.text or "").strip()
        if not txt or len(txt) > 30:
            continue
        sid_el = first(first(p._p, qn("w:pPr")), qn("w:pStyle")) if first(p._p, qn("w:pPr")) is not None else None
        st = styles_by_id.get(sid_el.get(qn("w:val"))) if sid_el is not None else None
        if st is not None and (HEADING_NAME_RE.search(st.name or "") or (style_paragraph(st).get("outline_level") or -1) >= 0):
            continue
        ppr = first(p._p, qn("w:pPr"))
        jc = first(ppr, qn("w:jc")) if ppr is not None else None
        f = para_first_run_font(p) or {}
        if jc is not None and jc.get(qn("w:val")) == "center" and f.get("bold") and len(txt) <= 25:
            fake += 1
    if fake:
        issues.append(f"{fake} 个疑似'假标题'(居中加粗但未使用标题样式)")

    spec = {
        "file": {
            "path": str(path.resolve()),
            "name": path.name,
            "size_kb": round(path.stat().st_size / 1024, 1),
            "analyzed_at": datetime.datetime.now().isoformat(timespec="seconds"),
        },
        "document": {
            "package_flags": flags,
            "field_types_used": dict(field_types),
            "body_block_count": len(blocks),
            "paragraph_count": sum(1 for b in blocks if b["kind"] == "p"),
            "table_count": len(tables),
            "image_count": len(images),
            "section_count": len(sections),
        },
        "doc_defaults": doc_defaults(doc),
        "sections": sections,
        "styles": styles_out,
        "numbering": {"definitions_available": len(numbering_defs)},
        "outline": headings,
        "paragraphs_preview": paragraphs,
        "paragraphs_truncated": sum(1 for b in blocks if b["kind"] == "p") > len(paragraphs),
        "tables": tables,
        "images": images,
        "cover": cover,
        "toc": toc,
        "formatting_inventory": formatting_inventory,
        "conventions": {
            "body_style": body_style_entry,
            "heading_styles": [
                {"level": h["level"], "style": h["style"],
                 "native_numbering": h["native_numbering"], "sample": h["text"][:30]}
                for h in headings[:12]
            ],
            "uses_native_numbering": any(h["native_numbering"] for h in headings),
            "manual_page_numbers_in_footer": manual_pagenum,
        },
        "issues": issues,
    }

    # ---- human summary
    sec0 = sections[0] if sections else {}
    print(f"=== Document Specification: {path.name} ===")
    print(f"blocks: {spec['document']['body_block_count']} (paragraphs {spec['document']['paragraph_count']}, tables {spec['document']['table_count']}, images {spec['document']['image_count']}), sections: {len(sections)}")
    if sec0:
        pg, mg = sec0["page"], sec0["margins_cm"]
        print(f"page: {pg['width_cm']}x{pg['height_cm']}cm {pg['orientation']}; margins T{mg['top']} B{mg['bottom']} L{mg['left']} R{mg['right']}")
    if body_style_entry:
        bf = body_style_entry["font"]
        print(f"body style: {body_style_entry['style']} | eastAsia={bf['eastAsia']} ascii={bf['ascii']} {bf['size_pt']}pt")
    dominant = formatting_inventory.get("dominant")
    if dominant:
        color = dominant.get("color") or (f"theme:{dominant['color_theme']}" if dominant.get("color_theme") else "auto/default")
        print(f"observed dominant run format: eastAsia={dominant.get('eastAsia')} ascii={dominant.get('ascii')} {dominant.get('size_pt')}pt color={color} (runs={dominant.get('run_count')})")
    if formatting_inventory.get("mixed_colors"):
        observed_colors = ", ".join(
            f"{item['color']}×{item['run_count']}"
            for item in formatting_inventory["colors"][:6]
        )
        print(f"observed text colors: {observed_colors}")
    hs = spec["conventions"]["heading_styles"]
    if hs:
        samples = "; ".join(f"L{h['level']} {h['style']}" + (f"[{h['native_numbering']}]" if h["native_numbering"] else "") for h in hs[:5])
        print(f"headings: {samples}")
    print(f"cover: {cover['present']} (conf {cover['confidence']}), toc: {toc['present']}, comments: {flags['has_comments']}, revisions: {flags['has_revisions']}")
    for it in issues:
        print(f"[issue] {it}")
    print(f"(paragraphs preview capped at {args.max_paras}; full structure in JSON)")
    dump_json(spec, args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
