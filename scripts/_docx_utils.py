"""Shared helpers for word-document-skill scripts.

All scripts in this directory import from here. Keep this module dependency-light:
python-docx, lxml, PIL are required; everything else is optional.
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------- console/IO
def setup_stdout() -> None:
    """Windows consoles default to GBK; force UTF-8 so Chinese text never crashes."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def jd(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2, default=str)


def dump_json(obj, out_path: str | None) -> None:
    text = jd(obj)
    if out_path:
        Path(out_path).write_text(text, encoding="utf-8")
        print(f"[saved] {out_path}")
    else:
        print(text)


def load_json(path: str):
    return json.loads(Path(path).read_text(encoding="utf-8"))


# ---------------------------------------------------------------- units
def emu_to_cm(v) -> float | None:
    return round(int(v) / 360000.0, 2) if v is not None else None


def twips_to_cm(v) -> float | None:
    return round(int(v) / 566.929, 2) if v is not None else None


def twips_to_pt(v) -> float | None:
    return round(int(v) / 20.0, 1) if v is not None else None


def half_pt_to_pt(v) -> float | None:
    return round(int(v) / 2.0, 1) if v is not None else None


# ---------------------------------------------------------------- XML helpers
def el_text(el) -> str:
    """Concatenate all w:t under an element."""
    if el is None:
        return ""
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def first(el, path: str):
    return None if el is None else el.find(path)


def all_(el, path: str):
    return [] if el is None else el.findall(path)


def iter_all(el, path: str):
    return [] if el is None else el.iter(path)


# ---------------------------------------------------------------- fields
FIELD_RE = re.compile(r"^\s*([A-Z0-9_]+)\s")


def para_fields(p_el) -> list[str]:
    """Return field instructions (PAGE, NUMPAGES, TOC, ...) used in a paragraph."""
    codes = []
    for instr in iter_all(p_el, qn("w:instrText")):
        m = FIELD_RE.match(instr.text or "")
        if m:
            codes.append(m.group(1))
    for fld in iter_all(p_el, qn("w:fldSimple")):
        instr = fld.get(qn("w:instr")) or ""
        m = FIELD_RE.match(instr)
        if m:
            codes.append(m.group(1))
    return codes


# ---------------------------------------------------------------- document model
def load_document(path: str) -> Document:
    return Document(path)


def body_blocks(doc: Document):
    """Yield ('p', Paragraph) / ('tbl', Table) in document body order.

    Covers everything python-docx's doc.paragraphs/doc.tables miss:
    ordering, tables between paragraphs, SDT-wrapped blocks are skipped by
    design (reported separately).
    """
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield "p", Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield "tbl", Table(child, doc)


def is_empty_para(p: Paragraph) -> bool:
    return not (p.text or "").strip()


def para_has_drawing(p_el) -> bool:
    return bool(
        first(p_el, f".//{qn('w:drawing')}") is not None
        or all_(p_el, f".//{qn('w:pict')}")
    )


# ---------------------------------------------------------------- style chain
def style_chain(style):
    """Yield style, its base, its base's base... (guards against cycles)."""
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def rpr_of_style(style):
    return first(style.element, qn("w:rPr"))


def ppr_of_style(style):
    return first(style.element, qn("w:pPr"))


def style_font(style) -> dict:
    """Effective font resolved through the basedOn chain."""
    out = {
        "ascii": None, "eastAsia": None, "hAnsi": None, "cs": None,
        "ascii_theme": None, "eastAsia_theme": None,
        "hAnsi_theme": None, "cs_theme": None,
        "size_pt": None, "bold": None, "italic": None, "underline": None,
        "color": None, "color_theme": None,
        "color_theme_tint": None, "color_theme_shade": None,
    }
    for st in style_chain(style):
        rpr = rpr_of_style(st)
        if rpr is None:
            continue
        fonts = first(rpr, qn("w:rFonts"))
        if fonts is not None:
            for key in ("ascii", "eastAsia", "hAnsi", "cs"):
                if out[key] is None:
                    out[key] = fonts.get(qn(f"w:{key}"))
            for xml_key, out_key in (
                ("asciiTheme", "ascii_theme"),
                ("eastAsiaTheme", "eastAsia_theme"),
                ("hAnsiTheme", "hAnsi_theme"),
                ("csTheme", "cs_theme"),
            ):
                if out[out_key] is None:
                    out[out_key] = fonts.get(qn(f"w:{xml_key}"))
        if out["size_pt"] is None:
            out["size_pt"] = half_pt_to_pt(first(rpr, qn("w:sz")).get(qn("w:val"))) if first(rpr, qn("w:sz")) is not None else None
        for key, tag in (("bold", "w:b"), ("italic", "w:i"), ("underline", "w:u")):
            if out[key] is None:
                el = first(rpr, qn(tag))
                if el is not None:
                    val = el.get(qn("w:val"))
                    out[key] = False if val in ("false", "0", "none") else True
        if out["color"] is None and out["color_theme"] is None:
            cel = first(rpr, qn("w:color"))
            if cel is not None:
                out["color"] = cel.get(qn("w:val"))
                out["color_theme"] = cel.get(qn("w:themeColor"))
                out["color_theme_tint"] = cel.get(qn("w:themeTint"))
                out["color_theme_shade"] = cel.get(qn("w:themeShade"))
    return out


def style_paragraph(style) -> dict:
    """Effective paragraph format resolved through the basedOn chain."""
    out = {
        "alignment": None, "line_spacing": None, "line_rule": None,
        "before_pt": None, "after_pt": None,
        "first_line_indent_cm": None, "left_indent_cm": None, "right_indent_cm": None,
        "keep_with_next": None, "keep_lines": None, "page_break_before": None,
        "outline_level": None,
    }
    for st in style_chain(style):
        ppr = ppr_of_style(st)
        if ppr is None:
            continue
        jc = first(ppr, qn("w:jc"))
        if out["alignment"] is None and jc is not None:
            out["alignment"] = jc.get(qn("w:val"))
        sp = first(ppr, qn("w:spacing"))
        if sp is not None:
            if out["line_spacing"] is None and sp.get(qn("w:line")):
                raw = int(sp.get(qn("w:line")))
                out["line_spacing"] = round(raw / 240.0, 2) if sp.get(qn("w:lineRule"), "auto") == "auto" else twips_to_pt(sp.get(qn("w:line")))
                out["line_rule"] = sp.get(qn("w:lineRule"), "auto")
            if out["before_pt"] is None and sp.get(qn("w:before")):
                out["before_pt"] = twips_to_pt(sp.get(qn("w:before")))
            if out["after_pt"] is None and sp.get(qn("w:after")):
                out["after_pt"] = twips_to_pt(sp.get(qn("w:after")))
        ind = first(ppr, qn("w:ind"))
        if ind is not None:
            if out["first_line_indent_cm"] is None:
                fl = ind.get(qn("w:firstLine")) or ind.get(qn("w:hanging"))
                flcm = twips_to_cm(fl) if fl else None
                if flcm is not None and ind.get(qn("w:hanging")):
                    flcm = -flcm
                out["first_line_indent_cm"] = flcm
            if out["left_indent_cm"] is None:
                out["left_indent_cm"] = twips_to_cm(ind.get(qn("w:left")) or ind.get(qn("w:start")))
            if out["right_indent_cm"] is None:
                out["right_indent_cm"] = twips_to_cm(ind.get(qn("w:right")) or ind.get(qn("w:end")))
        for key, tag in (("keep_with_next", "w:keepNext"), ("keep_lines", "w:keepLines"), ("page_break_before", "w:pageBreakBefore")):
            if out[key] is None and first(ppr, qn(tag)) is not None:
                out[key] = True
        ol = first(ppr, qn("w:outlineLvl"))
        if out["outline_level"] is None and ol is not None:
            out["outline_level"] = int(ol.get(qn("w:val")))
    return out


def para_numpr(p_el, styles_by_id: dict):
    """Resolve a paragraph's numbering reference, incl. inheritance from style."""
    ppr = first(p_el, qn("w:pPr"))
    if ppr is not None:
        numpr = first(ppr, qn("w:numPr"))
        if numpr is not None:
            numid_el, ilvl_el = first(numpr, qn("w:numId")), first(numpr, qn("w:ilvl"))
            if numid_el is not None and numid_el.get(qn("w:val")) not in (None, "0"):
                return {"num_id": int(numid_el.get(qn("w:val"))), "ilvl": int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0, "via": "direct"}
    # style inheritance
    if ppr is not None:
        pstyle = first(ppr, qn("w:pStyle"))
        sid = pstyle.get(qn("w:val")) if pstyle is not None else None
        st = styles_by_id.get(sid)
        while st is not None:
            ppr_s = first(st.element, qn("w:pPr"))
            if ppr_s is not None:
                numpr = first(ppr_s, qn("w:numPr"))
                if numpr is not None:
                    numid_el, ilvl_el = first(numpr, qn("w:numId")), first(numpr, qn("w:ilvl"))
                    if numid_el is not None and numid_el.get(qn("w:val")) not in (None, "0"):
                        return {"num_id": int(numid_el.get(qn("w:val"))), "ilvl": int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0, "via": "style"}
            st = st.base_style
    return None


# ---------------------------------------------------------------- numbering resolver
CN_DIGITS = "零一二三四五六七八九"


def _int_to_cn(n: int) -> str:
    if n <= 10:
        return CN_DIGITS[n] if n < 10 else "十"
    if n < 20:
        return "十" + CN_DIGITS[n % 10]
    if n < 100:
        return CN_DIGITS[n // 10] + "十" + (CN_DIGITS[n % 10] if n % 10 else "")
    return str(n)


def _fmt_num(fmt: str, val: int) -> str:
    try:
        if fmt in ("decimal", None):
            return str(val)
        if fmt == "chineseCounting" or fmt == "chineseCountingThousand":
            return _int_to_cn(val)
        if fmt == "chineseLegalSimplified":
            return _int_to_cn(val)
        if fmt == "upperLetter":
            return chr(64 + val) if 0 < val < 27 else str(val)
        if fmt == "lowerLetter":
            return chr(96 + val) if 0 < val < 27 else str(val)
        if fmt == "upperRoman":
            return str(val)
        if fmt == "decimalEnclosedCircle":
            return f"({val})"
        return str(val)
    except Exception:
        return str(val)


def load_numbering(doc: Document) -> dict:
    """Build {num_id: {"levels": {ilvl: {"fmt","lvl_text","start"}}}} from numbering.xml."""
    out: dict = {}
    # numbering part may not exist
    try:
        numbering_el = None
        for rel in doc.part.rels.values():
            if "numbering" in rel.reltype:
                numbering_el = rel.target_part._element if hasattr(rel.target_part, "_element") else None
                if numbering_el is None and hasattr(rel.target_part, "blob"):
                    from lxml import etree
                    numbering_el = etree.fromstring(rel.target_part.blob)
                break
        if numbering_el is None:
            return out
        abstract = {}
        for a in all_(numbering_el, qn("w:abstractNum")):
            aid = a.get(qn("w:abstractNumId"))
            levels = {}
            for lvl in all_(a, qn("w:lvl")):
                ilvl = int(lvl.get(qn("w:ilvl")))
                fmt_el, txt_el, start_el = first(lvl, qn("w:numFmt")), first(lvl, qn("w:lvlText")), first(lvl, qn("w:start"))
                levels[ilvl] = {
                    "fmt": fmt_el.get(qn("w:val")) if fmt_el is not None else "decimal",
                    "lvl_text": txt_el.get(qn("w:val")) if txt_el is not None else "",
                    "start": int(start_el.get(qn("w:val"))) if start_el is not None else 1,
                }
            abstract[aid] = levels
        for n in all_(numbering_el, qn("w:num")):
            nid = n.get(qn("w:numId"))
            ref = first(n, qn("w:abstractNumId"))
            if ref is not None:
                out[int(nid)] = {"levels": abstract.get(ref.get(qn("w:val")), {})}
    except Exception:
        pass
    return out


def numbering_counters():
    """Stateful counter helper: call next_numbering(num_id, ilvl, defs) per paragraph."""
    state: dict = {}

    def next_numbering(num_id: int, ilvl: int, defs: dict) -> str | None:
        d = defs.get(num_id)
        if not d:
            return None
        lvl = d["levels"].get(ilvl)
        if not lvl:
            return None
        # reset deeper levels
        for deeper in [k for k in state if k[0] == num_id and k[1] > ilvl]:
            del state[deeper]
        key = (num_id, ilvl)
        state[key] = state.get(key, lvl["start"] - 1) + 1
        text = lvl["lvl_text"]
        for slot in set(re.findall(r"%(\d)", text)):
            idx = int(slot) - 1
            if idx == ilvl:
                text = text.replace(f"%{slot}", _fmt_num(lvl["fmt"], state[key]))
            elif (num_id, idx) in state:
                sub = d["levels"].get(idx, {})
                text = text.replace(f"%{slot}", _fmt_num(sub.get("fmt", "decimal"), state[(num_id, idx)]))
        return text

    return next_numbering


# ---------------------------------------------------------------- run fonts
def run_font(run_el) -> dict:
    rpr = first(run_el, qn("w:rPr"))
    out = {
        "ascii": None, "eastAsia": None, "hAnsi": None, "cs": None,
        "ascii_theme": None, "eastAsia_theme": None,
        "hAnsi_theme": None, "cs_theme": None,
        "size_pt": None, "bold": None, "italic": None, "underline": None,
        "color": None, "color_theme": None,
        "color_theme_tint": None, "color_theme_shade": None,
    }
    if rpr is None:
        return out
    fonts = first(rpr, qn("w:rFonts"))
    if fonts is not None:
        out["ascii"] = fonts.get(qn("w:ascii"))
        out["eastAsia"] = fonts.get(qn("w:eastAsia"))
        out["hAnsi"] = fonts.get(qn("w:hAnsi"))
        out["cs"] = fonts.get(qn("w:cs"))
        out["ascii_theme"] = fonts.get(qn("w:asciiTheme"))
        out["eastAsia_theme"] = fonts.get(qn("w:eastAsiaTheme"))
        out["hAnsi_theme"] = fonts.get(qn("w:hAnsiTheme"))
        out["cs_theme"] = fonts.get(qn("w:csTheme"))
    sz = first(rpr, qn("w:sz"))
    if sz is not None:
        out["size_pt"] = half_pt_to_pt(sz.get(qn("w:val")))
    for key, tag in (("bold", "w:b"), ("italic", "w:i"), ("underline", "w:u")):
        el = first(rpr, qn(tag))
        if el is not None:
            out[key] = el.get(qn("w:val")) not in ("false", "0", "none")
    color = first(rpr, qn("w:color"))
    if color is not None:
        out["color"] = color.get(qn("w:val"))
        out["color_theme"] = color.get(qn("w:themeColor"))
        out["color_theme_tint"] = color.get(qn("w:themeTint"))
        out["color_theme_shade"] = color.get(qn("w:themeShade"))
    return out


def para_first_run_font(p: Paragraph) -> dict | None:
    for r in all_(p._p, qn("w:r")):
        if (r.find(qn("w:t")) is not None) or (r.find(qn("w:br")) is not None):
            f = run_font(r)
            if any(f.get(key) is not None for key in f):
                return f
    return None


# ---------------------------------------------------------------- doc defaults
def doc_defaults(doc: Document) -> dict:
    out = {"font": {"ascii": None, "eastAsia": None, "size_pt": None}, "para": {}}
    try:
        styles_el = doc.styles.element
        rpr_default = first(first(styles_el, qn("w:docDefaults")), qn("w:rPrDefault"))
        rpr = first(rpr_default, qn("w:rPr")) if rpr_default is not None else None
        if rpr is not None:
            fonts = first(rpr, qn("w:rFonts"))
            if fonts is not None:
                out["font"]["ascii"] = fonts.get(qn("w:ascii"))
                out["font"]["eastAsia"] = fonts.get(qn("w:eastAsia"))
            sz = first(rpr, qn("w:sz"))
            if sz is not None:
                out["font"]["size_pt"] = half_pt_to_pt(sz.get(qn("w:val")))
    except Exception:
        pass
    return out


# ---------------------------------------------------------------- parts & flags
def package_flags(path: str) -> dict:
    """Inspect the raw zip: comments, revisions, settings, media inventory."""
    flags = {
        "has_comments": False, "comments_count": 0,
        "has_revisions": False, "has_hyperlinks": False,
        "even_and_odd_headers": False, "has_textboxes": False,
        "has_content_controls": False, "media_count": 0,
        "has_macros": False, "has_embedded_objects": False,
    }
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "word/comments.xml" in names:
                flags["has_comments"] = True
                try:
                    from lxml import etree
                    root = etree.fromstring(z.read("word/comments.xml"))
                    flags["comments_count"] = len(all_(root, qn("w:comment")))
                except Exception:
                    pass
            docxml = z.read("word/document.xml") if "word/document.xml" in names else b""
            if docxml:
                from lxml import etree
                root = etree.fromstring(docxml)
                if all_(root, qn("w:ins")) or all_(root, qn("w:del")):
                    flags["has_revisions"] = True
                if all_(root, qn("w:hyperlink")):
                    flags["has_hyperlinks"] = True
                if all_(root, qn("w:txbxContent")):
                    flags["has_textboxes"] = True
                if all_(root, qn("w:sdt")):
                    flags["has_content_controls"] = True
            if "word/settings.xml" in names:
                from lxml import etree
                sroot = etree.fromstring(z.read("word/settings.xml"))
                if first(sroot, qn("w:evenAndOddHeaders")) is not None:
                    flags["even_and_odd_headers"] = True
            flags["media_count"] = len([n for n in names if n.startswith("word/media/")])
            flags["has_macros"] = any(n.endswith(".bin") and "vbaProject" in n for n in names)
    except Exception:
        pass
    return flags
