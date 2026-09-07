# -*- coding: utf-8 -*-
"""Render a .docx to PDF and per-page PNGs for visual QA.

Backends (auto): MS Word COM (most faithful, Windows) -> LibreOffice soffice.
Also updates fields (TOC page numbers) before exporting when --update-fields.

Usage:
    python render_docx.py <input.docx> -o <outdir> [--dpi 110] [--update-fields] [--pages 1-5]
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _docx_utils import dump_json, setup_stdout  # noqa: E402


def render_with_word(docx: Path, pdf: Path, update_fields: bool) -> tuple[bool, str]:
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False, "pywin32 not available"
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.Options.UpdateFieldsAtPrint = True
        doc = word.Documents.Open(str(docx.resolve()), ReadOnly=False, AddToRecentFiles=False)
        if update_fields:
            try:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    doc.TablesOfContents(i).Update()
            except Exception:
                pass
            try:
                doc.Fields.Update()
            except Exception:
                pass
            try:
                doc.Repaginate()
            except Exception:
                pass
            doc.SaveAs2(str(docx.resolve()), FileFormat=16)
        doc.ExportAsFixedFormat(str(pdf.resolve()), ExportFormat=17)  # wdExportFormatPDF
        return True, "word-com"
    except Exception as e:
        return False, f"word-com failed: {e}"
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def render_with_soffice(docx: Path, outdir: Path) -> tuple[bool, str]:
    soffice = None
    for cand in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice", "/usr/local/bin/soffice",
    ):
        if Path(cand).exists():
            soffice = cand
            break
    if soffice is None:
        try:
            subprocess.run(["soffice", "--version"], capture_output=True, check=True)
            soffice = "soffice"
        except Exception:
            return False, "LibreOffice not found"
    try:
        cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf",
               "--outdir", str(outdir), str(docx.resolve())]
        subprocess.run(cmd, capture_output=True, timeout=180, check=True)
        pdf = outdir / (docx.stem + ".pdf")
        if pdf.exists():
            return True, "soffice"
        return False, "soffice produced no pdf"
    except Exception as e:
        return False, f"soffice failed: {e}"


def parse_pages(spec: str | None, total: int) -> list[int]:
    if not spec:
        return list(range(1, total + 1))
    out = []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            out.extend(range(int(a), min(int(b), total) + 1))
        else:
            if int(part) <= total:
                out.append(int(part))
    return sorted(set(out))


def main() -> int:
    setup_stdout()
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("-o", "--outdir", required=True)
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--update-fields", action="store_true",
                    help="update TOC/PAGE fields via Word before export (writes the docx)")
    ap.add_argument("--pages", default=None, help="e.g. 1-3,7")
    ap.add_argument("--pdf-only", action="store_true")
    args = ap.parse_args()

    docx = Path(args.input)
    if not docx.exists():
        print(f"ERROR: file not found: {docx}")
        return 2
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = outdir / (docx.stem + ".pdf")

    t0 = time.time()
    ok, backend = render_with_word(docx, pdf, args.update_fields)
    if not ok:
        print(f"[render] {backend}; trying LibreOffice...", file=sys.stderr)
        ok, backend = render_with_soffice(docx, outdir)
    if not ok or not pdf.exists():
        print(f"ERROR: render failed. last: {backend}")
        return 3

    import fitz
    pdoc = fitz.open(pdf)
    total = pdoc.page_count
    result = {"pdf": str(pdf), "backend": backend, "pages": total,
              "duration_s": round(time.time() - t0, 1), "images": []}
    if not args.pdf_only:
        for pno in parse_pages(args.pages, total):
            page = pdoc[pno - 1]
            pix = page.get_pixmap(dpi=args.dpi)
            img = outdir / f"page-{pno:03d}.png"
            pix.save(img)
            result["images"].append(str(img))
    pdoc.close()

    print(f"=== rendered {docx.name} -> {total} pages (backend: {backend}, {result['duration_s']}s) ===")
    for img in result["images"]:
        print(f"  {img}")
    if args.update_fields:
        print("[note] fields updated and saved into the docx (TOC page numbers refresh)")
    dump_json(result, None)
    return 0


if __name__ == "__main__":
    sys.exit(main())
