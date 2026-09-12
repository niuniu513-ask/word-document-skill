"""Extract the CONTENT of a .docx into a structured content-model JSON.

Content and formatting are separated: this script returns WHAT the document
says (blocks: headings, paragraphs, tables, images, cover fields), with only
lightweight format hints needed for faithful re-generation.

Usage:
    python extract_content.py <input.docx> [-o content.json] [--markdown out.md]
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import (  # noqa: E402
    all_,
    body_blocks,
    dump_json,
    first,
    is_empty_para,
    load_numbering,
    numbering_counters,
    para_fields,
    para_numpr,
    setup_stdout,
)
from docx import Document  # noqa: E402

HEADING_RE = re.compile(r"^(heading|标题)\s*(\d)", re.I)


def table_to_rows(tbl) -> list[list[str]]:
    rows = []
    try:
        seen_spans = set()
        for r in tbl.rows:
            row = []
            for c in r.cells:
                key = id(c._tc)
                row.append("" if key in seen_spans else c.text.strip())
                seen_spans.add(key)
            rows.append(row)
    except Exception:
        pass
    return rows


def merged_map(tbl) -> list[dict]:
    out = []
    try:
        for ri, r in enumerate(tbl.rows):
            for ci, c in enumerate(r.cells):
                tc = c._tc
                vm = first(tc, f"{qn('w:tcPr')}/{qn('w:vMerge')}")
                gs = first(tc, f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
                if (vm is not None and vm.get(qn("w:val")) == "continue") or gs is not None:
                    out.append({"row": ri, "col": ci,
                                "vmerge": vm is not None,
                                "grid_span": int(gs.get(qn("w:val"))) if gs is not None else 1})
    except Exception:
        pass
    return out


def main() -> int:
    setup_stdout()
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", default=None)
    ap.add_argument("--markdown", default=None, help="also emit a readable markdown view")
    args = ap.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"ERROR: file not found: {path}")
        return 2
    doc = Document(str(path))
    numbering_defs = load_numbering(doc)
    next_num = numbering_counters()

    blocks = []
    para_count = 0
    for kind, obj in body_blocks(doc):
        if kind == "p":
            para_count += 1
            p = obj
            text = (p.text or "").strip()
            stylename = ""
            outline = None
            try:
                stylename = p.style.name or ""
                if p.style is not None:
                    from _docx_utils import style_paragraph
                    outline = style_paragraph(p.style).get("outline_level")
            except Exception:
                pass
            if is_empty_para(p) and not (first(p._p, f".//{qn('w:drawing')}") is not None or all_(p._p, qn("w:pict"))):
                if blocks and blocks[-1]["type"] == "empty":
                    continue
                blocks.append({"type": "empty", "index": len(blocks)})
                continue
            codes = para_fields(p._p)
            if any(c.startswith("TOC") for c in codes):
                blocks.append({"type": "toc", "index": len(blocks), "note": "native TOC field"})
                continue
            has_drawing = first(p._p, f".//{qn('w:drawing')}") is not None or bool(all_(p._p, qn("w:pict")))
            if has_drawing:
                blocks.append({"type": "image", "index": len(blocks), "paragraph_style": stylename,
                               "text_around": text[:60] or None})
                continue
            numinfo = para_numpr(p._p, {})
            num_text = next_num(numinfo["num_id"], numinfo["ilvl"], numbering_defs) if numinfo else None
            m = HEADING_RE.match(stylename)
            if m:
                blocks.append({"type": "heading", "level": int(m.group(2)),
                               "text": text, "style": stylename,
                               "native_numbering": num_text})
            elif outline is not None and 0 <= outline <= 8:
                blocks.append({"type": "heading", "level": outline + 1,
                               "text": text, "style": stylename,
                               "native_numbering": num_text})
            elif re.match(r"^[图表]\s*\d+", text) and len(text) < 80:
                blocks.append({"type": "caption", "kind": "figure" if text.startswith("图") else "table",
                               "text": text, "style": stylename})
            else:
                blocks.append({"type": "paragraph", "text": text, "style": stylename,
                               "native_numbering": num_text,
                               "fields": codes or None})
        else:
            rows = table_to_rows(obj)
            blocks.append({
                "type": "table", "index": len(blocks),
                "rows": rows, "merged": merged_map(obj) or None,
                "header_row_is_data": None,
            })

    content = {
        "source": str(path.resolve()),
        "meta": {"paragraph_count": para_count, "block_count": len(blocks)},
        "blocks": blocks,
    }

    # markdown view
    if args.markdown:
        lines = []
        for b in blocks:
            if b["type"] == "heading":
                lines.append("#" * min(b["level"], 6) + " " + b["text"])
            elif b["type"] == "paragraph":
                prefix = (b.get("native_numbering") + " ") if b.get("native_numbering") else ""
                lines.append(prefix + b["text"])
            elif b["type"] == "table":
                rows = b["rows"]
                if rows:
                    lines.append("| " + " | ".join(c.replace("\n", " ") or " " for c in rows[0]) + " |")
                    lines.append("|" + "---|" * len(rows[0]))
                    for r in rows[1:]:
                        lines.append("| " + " | ".join(c.replace("\n", " ") or " " for c in r) + " |")
                lines.append("")
            elif b["type"] == "image":
                lines.append("[图片]" + (f" {b['text_around']}" if b.get("text_around") else ""))
            elif b["type"] == "toc":
                lines.append("[目录]")
            elif b["type"] == "caption":
                lines.append(f"*{b['text']}*")
        Path(args.markdown).write_text("\n".join(lines), encoding="utf-8")
        print(f"[saved] {args.markdown}")

    h = [b for b in blocks if b["type"] == "heading"]
    print(f"=== Content model: {path.name} ===")
    print(f"blocks: {len(blocks)} (headings {len(h)}, tables {sum(1 for b in blocks if b['type']=='table')}, images {sum(1 for b in blocks if b['type']=='image')})")
    for b in h[:8]:
        print(f"  H{b['level']}: {b['text'][:50]}")
    if len(h) > 8:
        print(f"  ... {len(h) - 8} more headings")
    dump_json(content, args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
