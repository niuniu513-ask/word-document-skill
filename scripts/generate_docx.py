# -*- coding: utf-8 -*-
"""Generate a .docx INSIDE an existing format system (template or format spec).

Two modes:
  --template t.docx   Copy the template (keeps styles/sections/headers/theme),
                      fill placeholders, then rebuild the body from content.json
                      using the template's own styles.
  --spec doc-spec.json  Build a new document from a Document Specification
                      (the JSON emitted by analyze_docx.py) and content.json.

content.json follows schemas/content-model.json:
  cover | heading | paragraph | table | table_fill | image | caption | page_break

Usage:
  python generate_docx.py --template t.docx --content c.json -o out.docx [--mode auto|fill|rebuild]
  python generate_docx.py --spec spec.json --content c.json -o out.docx
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import (  # noqa: E402
    all_, body_blocks, dump_json, first, setup_stdout, style_paragraph,
)

HEADING_NAME_RE = re.compile(r"^(heading|标题)\s*(\d)", re.I)
MANUAL_NUM_RE = re.compile(
    r"^\s*(?:(?:\d+(?:\.\d+){0,4})|(?:[一二三四五六七八九十百]+)|(?:\([一二三四五六七八九十]+\))|(?:（[一二三四五六七八九十]+）)|(?:第\s*[一二三四五六七八九十百\d]+\s*[章节]))[、.．,，:：\s]?\s*"
)


# ------------------------------------------------------------------ helpers
def style_id_by_outline(doc, level: int):
    """Find the template's paragraph style carrying outline level (level-1)."""
    best = None
    for st in doc.styles:
        try:
            if st.type != WD_STYLE_TYPE.PARAGRAPH:
                continue
        except Exception:
            continue
        pf = style_paragraph(st)
        if pf.get("outline_level") == level - 1:
            name = st.name or ""
            if HEADING_NAME_RE.search(name):
                return st
            best = best or st
    # fallback: name convention
    for st in doc.styles:
        try:
            if st.type == WD_STYLE_TYPE.PARAGRAPH and HEADING_NAME_RE.match(st.name or ""):
                m = HEADING_NAME_RE.match(st.name)
                if m and int(m.group(2)) == level:
                    return st
        except Exception:
            continue
    return best


def most_used_body_style(doc) -> str | None:
    counts: dict[str, int] = {}
    try:
        for p in doc.paragraphs:
            t = (p.text or "").strip()
            if len(t) > 15 and p.style is not None:
                name = p.style.name or ""
                if HEADING_NAME_RE.match(name) or (style_paragraph(p.style).get("outline_level") or -1) >= 0:
                    continue
                counts[name] = counts.get(name, 0) + 1
    except Exception:
        pass
    if counts:
        return max(counts, key=counts.get)
    return "Normal"


def most_used_table_style(doc):
    counts: dict[str, int] = {}
    for t in doc.tables:
        el = first(t._tbl, f"{qn('w:tblPr')}/{qn('w:tblStyle')}")
        if el is not None:
            sid = el.get(qn("w:val"))
            counts[sid] = counts.get(sid, 0) + 1
    if not counts:
        return None
    sid = max(counts, key=counts.get)
    for st in doc.styles:
        if st.style_id == sid:
            return st
    return None


def find_caption_style(doc):
    for st in doc.styles:
        try:
            n = st.name or ""
            if st.type == WD_STYLE_TYPE.PARAGRAPH and ("caption" in n.lower() or "题注" in n):
                return st
        except Exception:
            continue
    return None


def para_full_text(p) -> str:
    return "".join(t.text or "" for t in p._p.iter(qn("w:t")))


def replace_token_in_runs(p, token: str, value: str) -> bool:
    """Replace token possibly spanning runs; keeps the run where it starts."""
    runs = list(p._p.iter(qn("w:r")))
    texts = ["".join(t.text or "" for t in r.findall(qn("w:t"))) for r in runs]
    full = "".join(texts)
    idx = full.find(token)
    if idx < 0:
        return False
    end = idx + len(token)
    pos = 0
    touched = []
    for r, txt in zip(runs, texts):
        rs, re_ = pos, pos + len(txt)
        pos = re_
        if re_ <= idx or rs >= end:
            continue
        touched.append((r, txt, rs, re_))
    if not touched:
        return False
    first_r, first_txt, first_rs, _ = touched[0]
    last_txt = touched[-1][1]
    prefix = first_txt[: idx - first_rs]
    suffix = last_txt[end - touched[-1][2]:]
    newtext = prefix + value + suffix
    # clear all w:t in touched runs, then write newtext into first
    for i, (r, txt, rs, re_) in enumerate(touched):
        for t in r.findall(qn("w:t")):
            r.remove(t)
        if i == 0:
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = newtext
            r.append(t)
    return True


def _key_body_pattern(key: str) -> str:
    r"""学院 -> 学\s*院 (tolerates wide-spaced labels like 学    院)."""
    chars = [re.escape(c) for c in key if not c.isspace()]
    return r"\s*".join(chars)


def fill_placeholders(doc, fields: dict) -> tuple[int, list[str]]:
    """Fill {{key}}, 【key】 and label+underscores placeholders. Returns (filled, missing)."""
    filled = 0
    missing = []
    targets = list(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                targets.extend(cell.paragraphs)
    for key, value in fields.items():
        v = str(value)
        done = False
        for p in targets:
            for token in (f"{{{{{key}}}}}", f"【{key}】", f"{{ {key} }}"):
                if token in para_full_text(p):
                    if replace_token_in_runs(p, token, v):
                        done = True
            if not done:
                # label + underscores: 学    院：______
                text = para_full_text(p)
                pat = re.compile(rf"({_key_body_pattern(key)}\s*[:：]?\s*)[_＿]{{2,}}")
                if pat.search(text):
                    new_full = pat.sub(lambda m: m.group(1) + v, text)
                    # keep first run formatting, wipe the rest
                    runs = list(p._p.iter(qn("w:r")))
                    for i, r in enumerate(runs):
                        for t in r.findall(qn("w:t")):
                            r.remove(t)
                        if i == 0:
                            t = OxmlElement("w:t")
                            t.set(qn("xml:space"), "preserve")
                            t.text = new_full
                            r.append(t)
                    done = True
            if done:
                break
        if done:
            filled += 1
        else:
            missing.append(key)
    return filled, missing


def has_page_break(p) -> bool:
    return bool(all_(p._p, f"{qn('w:r')}/{qn('w:br')}[@{qn('w:type')}='page']"))


def add_page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)


def set_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tbl_pr.append(borders)


def set_cell_eastasia(cell, eastasia: str | None, ascii_font: str | None):
    if not (eastasia or ascii_font):
        return
    for p in cell.paragraphs:
        for r in p.runs:
            rpr = r._r.get_or_add_rPr()
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is None:
                fonts = OxmlElement("w:rFonts")
                rpr.insert(0, fonts)
            if eastasia:
                fonts.set(qn("w:eastAsia"), eastasia)
            if ascii_font:
                fonts.set(qn("w:ascii"), ascii_font)
                fonts.set(qn("w:hAnsi"), ascii_font)


# ------------------------------------------------------------------ block renderers
class Renderer:
    def __init__(self, doc, opts, warnings: list):
        self.doc = doc
        self.o = opts
        self.warn = warnings
        self.heading_cache: dict[int, object] = {}
        self.body_style_name = opts.body_style or most_used_body_style(doc)
        self.table_style = most_used_table_style(doc)
        self.caption_style = find_caption_style(doc)
        self.default_font = None  # (eastAsia, ascii) for new table cells

    def heading_style(self, level):
        if level not in self.heading_cache:
            self.heading_cache[level] = style_id_by_outline(self.doc, level)
        return self.heading_cache[level]

    def add_heading(self, blk):
        st = self.heading_style(blk.get("level", 1))
        text = blk.get("text", "")
        if blk.get("numbering") == "native":
            text = MANUAL_NUM_RE.sub("", text).strip()
        p = self.doc.add_paragraph()
        if st is not None:
            p.style = st
        else:
            self.warn.append(f"未找到 level {blk.get('level')} 的标题样式,已用默认段落")
        p.add_run(text)
        return [p._p]

    def add_paragraph(self, blk):
        p = self.doc.add_paragraph()
        try:
            if self.body_style_name:
                p.style = self.doc.styles[self.body_style_name]
        except KeyError:
            pass
        p.add_run(blk.get("text", ""))
        return [p._p]

    def add_caption(self, text, kind="figure"):
        p = self.doc.add_paragraph()
        if self.caption_style is not None:
            p.style = self.caption_style
            p.add_run(text)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text)
            r.font.size = Pt(9)
        return p

    def add_table(self, blk):
        rows = blk.get("rows") or []
        if not rows:
            self.warn.append("遇到空表格,已跳过")
            return [None]
        ncols = max(len(r) for r in rows)
        style = None
        if blk.get("style"):
            for st in self.doc.styles:
                try:
                    if st.name == blk["style"] or st.style_id == blk["style"]:
                        style = st
                        break
                except Exception:
                    continue
            if style is None:
                self.warn.append(f"表格样式 {blk['style']} 不存在,使用文档默认表格样式")
        if style is None:
            style = self.table_style
        tbl = self.doc.add_table(rows=len(rows), cols=ncols)
        styled = False
        if style is not None:
            try:
                tbl.style = style.name
                styled = True
            except Exception:
                try:
                    tblpr = tbl._tbl.tblPr
                    tsel = OxmlElement("w:tblStyle")
                    tsel.set(qn("w:val"), style.style_id)
                    tblpr.insert(0, tsel)
                    styled = True
                except Exception:
                    pass
        if not styled and not self.o.no_table_borders:
            set_table_borders(tbl)
        header = blk.get("header_row", True)
        for ri, row in enumerate(rows):
            for ci in range(ncols):
                cell = tbl.cell(ri, ci)
                cell.text = str(row[ci]) if ci < len(row) else ""
                if ri == 0 and header:
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for r in p.runs:
                            r.font.bold = True
        widths = blk.get("widths_cm")
        if widths:
            tbl.autofit = False
            for ci, w in enumerate(widths[:ncols]):
                for ri in range(len(rows)):
                    tbl.cell(ri, ci).width = Cm(w)
        els = [tbl._tbl]
        if blk.get("caption"):
            cap = self.add_caption(blk["caption"], "table")
            els.append(cap._p)
        return els

    def fill_table(self, blk):
        idx = blk.get("table_index", 0)
        try:
            tbl = self.doc.tables[idx]
        except IndexError:
            self.warn.append(f"table_fill: 表格索引 {idx} 不存在(共 {len(self.doc.tables)} 个表格)")
            return []
        for addr, value in (blk.get("cells") or {}).items():
            try:
                ri, ci = (int(x) for x in str(addr).split(","))
                cell = tbl.cell(ri, ci)
                # write into first paragraph, preserve its formatting
                paras = cell.paragraphs
                target = paras[0]
                for extra in paras[1:]:
                    extra._p.getparent().remove(extra._p)
                for r in list(target.runs):
                    r._r.getparent().remove(r._r)
                target.add_run(str(value))
            except Exception as e:
                self.warn.append(f"table_fill 单元格 {addr} 写入失败: {e}")
        return []

    def add_image(self, blk):
        src = blk.get("path")
        if not src or not Path(src).exists():
            self.warn.append(f"图片不存在: {src}")
            return []
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        try:
            if blk.get("width_cm"):
                run.add_picture(str(src), width=Cm(blk["width_cm"]))
            else:
                run.add_picture(str(src))  # native size, aspect preserved
        except Exception as e:
            self.warn.append(f"图片插入失败 {src}: {e}")
            return []
        els = [p._p]
        if blk.get("caption"):
            cap = self.add_caption(blk["caption"], "figure")
            els.append(cap._p)
        return els

    def render_block(self, blk):
        """Dispatch one content block; return list of created lxml elements."""
        t = blk.get("type")
        if t == "heading":
            return self.add_heading(blk)
        if t == "paragraph":
            return self.add_paragraph(blk)
        if t == "table":
            return self.add_table(blk)
        if t == "table_fill":
            return self.fill_table(blk)
        if t == "image":
            return self.add_image(blk)
        if t == "caption":
            return [self.add_caption(blk.get("text", ""))._p]
        if t == "page_break":
            add_page_break(self.doc)
            return []
        self.warn.append(f"未知块类型: {t}")
        return []


# ------------------------------------------------------------------ body surgery
def first_heading_index(blocks) -> int | None:
    for i, (kind, obj) in enumerate(blocks):
        if kind == "p":
            p = obj
            name = ""
            try:
                name = p.style.name or ""
            except Exception:
                pass
            if HEADING_NAME_RE.search(name):
                return i
            try:
                if p.style is not None and (style_paragraph(p.style).get("outline_level") or -1) >= 0:
                    return i
            except Exception:
                continue
        elif kind == "tbl":
            continue
    return None


def delete_body_from(doc, start_index, blocks):
    """Delete body blocks from start_index onward (keeps final sectPr)."""
    removed = 0
    for i, (kind, obj) in enumerate(blocks):
        if i < start_index:
            continue
        el = obj._p if kind == "p" else obj._tbl
        if el.getparent() is not None:
            el.getparent().remove(el)
            removed += 1
    return removed


def ensure_styles_for_spec(doc, spec, warnings):
    """Patch/ensure styles referenced by the spec's conventions (spec mode)."""
    conv = spec.get("conventions", {})
    body = conv.get("body_style") or {}
    def ensure(name, font, para, outline=None):
        if not name:
            return None
        try:
            st = doc.styles[name]
        except (KeyError, Exception):
            try:
                st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            except Exception:
                return None
        try:
            f = font or {}
            if f.get("ascii"):
                st.font.name = f["ascii"]
            if f.get("eastAsia"):
                rpr = st.element.get_or_add_rPr()
                fonts = rpr.find(qn("w:rFonts"))
                if fonts is None:
                    fonts = OxmlElement("w:rFonts")
                    rpr.insert(0, fonts)
                fonts.set(qn("w:eastAsia"), f["eastAsia"])
            if f.get("size_pt"):
                st.font.size = Pt(f["size_pt"])
            if f.get("bold"):
                st.font.bold = True
            p = para or {}
            from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
            amap = {"center": AL.CENTER, "left": AL.LEFT, "right": AL.RIGHT, "justify": AL.JUSTIFY, "both": AL.JUSTIFY}
            if p.get("alignment") in amap:
                st.paragraph_format.alignment = amap[p["alignment"]]
            if p.get("line_spacing"):
                st.paragraph_format.line_spacing = p["line_spacing"]
            if p.get("before_pt") is not None:
                st.paragraph_format.space_before = Pt(p["before_pt"])
            if p.get("after_pt") is not None:
                st.paragraph_format.space_after = Pt(p["after_pt"])
            if p.get("first_line_indent_cm") is not None:
                st.paragraph_format.first_line_indent = Cm(p["first_line_indent_cm"])
            if outline is not None and outline >= 0:
                ppr = st.element.get_or_add_pPr()
                ol = ppr.find(qn("w:outlineLvl"))
                if ol is None:
                    ol = OxmlElement("w:outlineLvl")
                    ppr.append(ol)
                ol.set(qn("w:val"), str(outline))
        except Exception as e:
            warnings.append(f"样式 {name} 应用规格失败: {e}")
        return st
    ensure(body.get("style"), body.get("font"), body.get("paragraph"))
    for h in conv.get("heading_styles", [])[:9]:
        ensure(h.get("style"), None, None, outline=h.get("level", 1) - 1)


def apply_spec_page_setup(doc, spec, warnings):
    try:
        sec0 = spec["sections"][0]
    except (KeyError, IndexError):
        return
    sec = doc.sections[0]
    pg, mg = sec0.get("page", {}), sec0.get("margins_cm", {})
    try:
        if pg.get("width_cm"):
            sec.page_width = Cm(pg["width_cm"])
        if pg.get("height_cm"):
            sec.page_height = Cm(pg["height_cm"])
        if pg.get("orientation") == "landscape":
            from docx.enum.section import WD_ORIENT
            sec.orientation = WD_ORIENT.LANDSCAPE
        for k in ("top", "bottom", "left", "right", "header", "footer"):
            if mg.get(k) is not None:
                setattr(sec, {"top": "top_margin", "bottom": "bottom_margin", "left": "left_margin",
                              "right": "right_margin", "header": "header_distance",
                              "footer": "footer_distance"}[k], Cm(mg[k]))
    except Exception as e:
        warnings.append(f"页面设置失败: {e}")


def normalize_heading_text(s: str) -> str:
    s = re.sub(r"\s+", "", s or "")
    return MANUAL_NUM_RE.sub("", s)


def find_heading_anchor(doc, anchor: str):
    """Find a heading paragraph matching anchor text (exact > contains)."""
    target = normalize_heading_text(anchor)
    for p in doc.paragraphs:
        name = ""
        ol = None
        try:
            name = p.style.name or ""
            ol = style_paragraph(p.style).get("outline_level")
        except Exception:
            continue
        if not (HEADING_NAME_RE.search(name) or (ol is not None and 0 <= ol <= 8)):
            continue
        ptxt = normalize_heading_text(p.text or "")
        if ptxt == target:
            return p
    for p in doc.paragraphs:
        ptxt = normalize_heading_text(p.text or "")
        if target and target in ptxt:
            return p
    return None


# ------------------------------------------------------------------ main
def main() -> int:
    setup_stdout()
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", default=None)
    ap.add_argument("--spec", default=None, help="document-spec.json from analyze_docx.py")
    ap.add_argument("--content", required=True, help="content-model.json")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--mode", choices=["auto", "fill", "rebuild", "insert"], default="auto")
    ap.add_argument("--body-style", default=None, help="override body paragraph style name")
    ap.add_argument("--no-table-borders", action="store_true")
    args = ap.parse_args()

    if not args.template and not args.spec:
        print("ERROR: 需要 --template 或 --spec 之一")
        return 2
    content = json.loads(Path(args.content).read_text(encoding="utf-8"))
    blocks_in = content.get("blocks", content if isinstance(content, list) else [])
    warnings: list[str] = []

    if args.template:
        shutil.copyfile(args.template, args.output)
        doc = Document(args.output)
    else:
        spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
        doc = Document()
        apply_spec_page_setup(doc, spec, warnings)
        ensure_styles_for_spec(doc, spec, warnings)

    opts = argparse.Namespace(body_style=args.body_style, no_table_borders=args.no_table_borders)
    renderer = Renderer(doc, opts, warnings)

    # 1. cover fields first (works in both modes; no-op when no placeholders)
    cover_blocks = [b for b in blocks_in if b.get("type") == "cover"]
    filled, missing = 0, []
    for cb in cover_blocks:
        f_, m_ = fill_placeholders(doc, cb.get("fields", {}))
        filled += f_
        missing.extend(m_)

    content_blocks = [b for b in blocks_in if b.get("type") != "cover"]
    anchored: dict[str, list] = {}
    floating: list = []
    for b in content_blocks:
        anchor = b.get("after")
        if isinstance(anchor, str) and anchor.strip():
            anchored.setdefault(anchor.strip(), []).append(b)
        else:
            floating.append(b)

    # 2. table_fill BEFORE any deletion (they target original template tables)
    for blk in [b for b in floating if b.get("type") == "table_fill"]:
        renderer.fill_table(blk)
    floating = [b for b in floating if b.get("type") != "table_fill"]

    # 3. decide body strategy
    blocks = list(body_blocks(doc))
    mode = args.mode
    if mode == "auto":
        if anchored:
            mode = "insert"
        elif first_heading_index(blocks) is not None:
            mode = "rebuild"
        else:
            mode = "fill"

    if mode == "rebuild" and args.template:
        blocks = list(body_blocks(doc))
        hidx = first_heading_index(blocks)
        start = hidx if hidx is not None else len(blocks)
        removed = delete_body_from(doc, start, blocks)
        kept = blocks[:start]
        tail_pb = any(kind == "p" and has_page_break(obj) for kind, obj in kept[-2:]) if kept else False
        if kept and not tail_pb:
            add_page_break(doc)
        if anchored:
            warnings.append("rebuild 模式忽略 after 锚点(原正文已删除),相关块已追加到文末")
        print(f"[rebuild] 保留封面/前置块 {start} 个,移除原正文 {removed} 个块")
    elif mode == "insert":
        print(f"[insert] 保留全部原文档,按 after 锚点插入 {sum(len(v) for v in anchored.values())} 个块")
    else:
        print(f"[fill] 保留全部原文档内容,追加 {len(floating)} 个块")

    # 4. anchored insertion (insert mode)
    placed = 0
    for anchor_text, blks in anchored.items():
        if mode != "insert":
            floating.extend(blks)
            continue
        anchor_p = find_heading_anchor(doc, anchor_text)
        if anchor_p is None:
            warnings.append(f"锚点未匹配到标题: '{anchor_text}',块已追加到文末")
            floating.extend(blks)
            continue
        cursor = anchor_p._p
        for blk in blks:
            els = renderer.render_block(blk)
            # move freshly created elements to sit right after cursor;
            # els holds direct references, addnext reparents each one
            for el in els:
                if el is None:
                    continue
                cursor.addnext(el)
                cursor = el
                placed += 1

    # 5. floating blocks append at end
    added = {"heading": 0, "paragraph": 0, "table": 0, "image": 0, "caption": 0, "page_break": 0}
    for blk in floating:
        t = blk.get("type")
        if t in added:
            added[t] += 1
        renderer.render_block(blk)

    doc.save(args.output)

    print(f"=== generate_docx -> {args.output} ===")
    print(f"mode: {mode} | cover fields filled: {filled}/{filled + len(missing)} | anchored blocks placed: {placed}")
    print(f"floating blocks: {added}")
    if missing:
        print(f"[warn] 以下占位符未在文档中找到: {missing}")
    for w in warnings:
        print(f"[warn] {w}")
    summary = {"output": args.output, "mode": mode, "filled": filled, "missing_placeholders": missing,
               "anchored_placed": placed, "blocks_added": added, "warnings": warnings}
    dump_json(summary, None)
    return 0


if __name__ == "__main__":
    sys.exit(main())
