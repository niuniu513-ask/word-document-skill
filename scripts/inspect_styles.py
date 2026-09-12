"""Dump a detailed, human-readable style report for a .docx.

Use when analyze_docx.py's style list is not enough and you need the full
basedOn chain, latent styles, and effective formatting of each style.

Usage:
    python inspect_styles.py <input.docx> [-o styles.json] [--text out.txt] [--style "Heading"]
"""
from __future__ import annotations

import argparse
import fnmatch
import sys
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import dump_json, setup_stdout, style_font, style_paragraph  # noqa: E402


def main() -> int:
    setup_stdout()
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--output", default=None)
    ap.add_argument("--text", default=None)
    ap.add_argument("--style", default=None, help="glob filter, e.g. 'Heading*' or '标题*'")
    args = ap.parse_args()

    doc = Document(args.input)
    rows = []
    for st in doc.styles:
        try:
            entry = {
                "name": st.name,
                "style_id": st.style_id,
                "type": str(st.type).split(".")[-1].split(" ")[0] if st.type is not None else None,
                "builtin": st.builtin,
                "based_on": st.base_style.name if st.base_style is not None else None,
                "locked": st.locked if hasattr(st, "locked") else None,
                "quick_style": st.hidden is False if hasattr(st, "hidden") else None,
            }
            if entry["type"] in ("PARAGRAPH", "CHARACTER"):
                entry["font"] = style_font(st)
            if entry["type"] == "PARAGRAPH":
                entry["paragraph"] = style_paragraph(st)
            rows.append(entry)
        except Exception:
            continue
    if args.style:
        rows = [r for r in rows if fnmatch.fnmatch(r["name"] or "", args.style)]

    lines = [f"=== Styles in {Path(args.input).name} ({len(rows)}) ==="]
    for r in rows:
        line = f"[{r['type']}] {r['name']} (id={r['style_id']}, builtin={r['builtin']})"
        if r.get("based_on"):
            line += f"  based_on: {r['based_on']}"
        lines.append(line)
        f = r.get("font") or {}
        fparts = [f"{k}={v}" for k, v in f.items() if v is not None]
        if fparts:
            lines.append("    font: " + ", ".join(fparts))
        p = r.get("paragraph") or {}
        pparts = [f"{k}={v}" for k, v in p.items() if v is not None]
        if pparts:
            lines.append("    para: " + ", ".join(pparts))
    text = "\n".join(lines)
    print(text[:4000])
    if args.text:
        Path(args.text).write_text(text, encoding="utf-8")
        print(f"[saved] {args.text}")
    dump_json({"style_count": len(rows), "styles": rows}, args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
