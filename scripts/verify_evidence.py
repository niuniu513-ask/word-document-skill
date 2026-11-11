"""Evidence verification for screenshot-backed documents.

Checks that a report's screenshot evidence is complete and readable:

* label counts (e.g. "代码截图:" / "实验结果截图:") match expectations
* every label is followed by an image, and there are no stray images
* images are tightly cropped (not a wide black bar with 10% content)
* images are not blank, not extremely tall, and not distorted

Usage:
    python verify_evidence.py report.docx -o evidence.json
    python verify_evidence.py report.docx --label 代码截图 --expect 代码截图=29 --strict
    python verify_evidence.py report.pdf --label 代码截图     # PDF: counts only
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from PIL import Image

EMU_PER_CM = 360000
DEFAULT_LABELS = ("代码截图", "实验结果截图")

# Findings that mean "the evidence chain is broken".
ERROR_CODES = {
    "label_without_image",
    "image_missing",
    "blank_image",
    "unreadable_image",
}
WARNING_CODES = {
    "wide_bar",
    "too_tall",
    "distorted",
    "orphan_image",
    "unexpected_count",
}


@dataclass
class Finding:
    level: str
    code: str
    detail: str
    where: str = ""

    def as_dict(self) -> dict:
        return {"level": self.level, "code": self.code, "detail": self.detail, "where": self.where}


@dataclass
class ImageInfo:
    index: int
    part_name: str
    width_px: int = 0
    height_px: int = 0
    width_cm: float = 0.0
    height_cm: float = 0.0
    content_ratio: float = 0.0
    fill_ratio: float = 0.0
    label: str = ""

    @property
    def aspect(self) -> float:
        return self.height_px / self.width_px if self.width_px else 0.0

    def as_dict(self) -> dict:
        return {
            "index": self.index,
            "part": self.part_name,
            "label": self.label,
            "size_px": [self.width_px, self.height_px],
            "size_cm": [round(self.width_cm, 2), round(self.height_cm, 2)],
            "content_ratio": round(self.content_ratio, 3),
            "fill_ratio": round(self.fill_ratio, 3),
        }


@dataclass
class Report:
    path: str
    kind: str
    label_counts: dict = field(default_factory=dict)
    images: list = field(default_factory=list)
    findings: list = field(default_factory=list)

    def add(self, level: str, code: str, detail: str, where: str = "") -> None:
        self.findings.append(Finding(level, code, detail, where))

    def counts(self, level: str) -> int:
        return sum(1 for f in self.findings if f.level == level)

    def as_dict(self) -> dict:
        return {
            "file": self.path,
            "kind": self.kind,
            "labels": self.label_counts,
            "images": {
                "total": len(self.images),
                "labeled": sum(1 for i in self.images if i.label),
                "orphan": sum(1 for i in self.images if not i.label),
                "detail": [i.as_dict() for i in self.images],
            },
            "findings": [f.as_dict() for f in self.findings],
            "summary": {
                "errors": self.counts("error"),
                "warnings": self.counts("warning"),
                "passed": self.counts("error") == 0,
            },
        }


def _label_regex(label: str) -> re.Pattern:
    # Editor qualifiers are part of captions, not a different evidence label.
    return re.compile(rf"{re.escape(label)}\s*(?:[（(][^()（）\r\n]*[)）]\s*)?[:：]")


def _measure(blob: bytes) -> tuple[int, int, float, float]:
    """Return (w, h, content_ratio, fill_ratio) using only Pillow.

    content_ratio: share of columns that carry content (detects wide black bars)
    fill_ratio: share of pixels that differ from the background (detects blank images)
    """
    import io

    with Image.open(io.BytesIO(blob)) as im:
        w, h = im.size
        small = im.convert("L")
        small.thumbnail((300, 4000))
        sw, sh = small.size
        data = small.tobytes()

    if not data:
        return w, h, 0.0, 0.0

    # background = most common gray level; robust for both dark and light themes
    hist = [0] * 256
    for v in data:
        hist[v] += 1
    bg = max(range(256), key=lambda v: hist[v])

    tol = 18
    col_hits = [0] * sw
    row_hits = [0] * sh
    filled = 0
    for y in range(sh):
        base = y * sw
        for x in range(sw):
            if abs(data[base + x] - bg) > tol:
                col_hits[x] += 1
                row_hits[y] += 1
                filled += 1

    # a column counts as content only if a few pixels in it differ
    min_hits = max(1, sh * 0.002)
    content_cols = sum(1 for c in col_hits if c > min_hits)
    content_rows = sum(1 for r in row_hits if r > max(1, sw * 0.002))
    content_ratio = content_cols / sw if sw else 0.0
    fill_ratio = filled / (sw * sh) if sw and sh else 0.0
    if content_rows == 0:
        fill_ratio = 0.0
    return w, h, content_ratio, fill_ratio


def _walk_body(doc):
    """Walk XML once, including nested tables; merged cells are not duplicated."""
    for paragraph in doc.element.body.iter(qn("w:p")):
        yield "p", paragraph


def _paragraph_images(p) -> list[tuple[str, float, float]]:
    """Return [(rId, width_cm, height_cm)] for images anchored in a paragraph."""
    out = []
    for drawing in p.findall(".//" + qn("w:drawing")):
        extent = drawing.find(".//" + qn("wp:extent"))
        cx = float(extent.get("cx")) if extent is not None else 0.0
        cy = float(extent.get("cy")) if extent is not None else 0.0
        for blip in drawing.findall(".//" + qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            if rid:
                out.append((rid, cx / EMU_PER_CM, cy / EMU_PER_CM))
    return out


def verify_docx(path: Path, labels: list[str], min_content_ratio: float,
                max_aspect: float, allow_orphan: bool = False) -> Report:
    doc = Document(str(path))
    rep = Report(str(path), "docx")
    rep.label_counts = {label: 0 for label in labels}
    patterns = {label: _label_regex(label) for label in labels}

    pending: str | None = None      # a label waiting for its image
    pending_at = ""
    orphan = 0
    orphan_at = ""
    idx = 0
    para_no = 0

    for kind, element in _walk_body(doc):
        if kind != "p":
            continue
        para_no += 1
        text = "".join(element.itertext()).strip()
        hit = None
        for label in labels:
            if patterns[label].search(text):
                hit = label
                rep.label_counts[label] += 1
                # a paragraph may carry one label; take the first match
                break
        if hit is not None:
            if pending is not None:
                rep.add("error", "label_without_image",
                        f"标签「{pending}」后面没有图片", pending_at)
            pending = hit
            pending_at = f"段落 {para_no}"

        for rid, wcm, hcm in _paragraph_images(element):
            idx += 1
            info = ImageInfo(index=idx, part_name=rid, width_cm=wcm, height_cm=hcm)
            try:
                part = doc.part.related_parts[rid]
                info.part_name = str(getattr(part, "partname", rid))
                w, h, content_ratio, fill_ratio = _measure(part.blob)
                info.width_px, info.height_px = w, h
                info.content_ratio, info.fill_ratio = content_ratio, fill_ratio
            except Exception as exc:  # unreadable media
                rep.add("error", "image_missing", f"图片 {rid} 无法读取: {exc}")
                rep.images.append(info)
                continue

            if pending is not None:
                info.label = pending
                pending = None
                pending_at = ""
            else:
                orphan += 1
                if not orphan_at:
                    orphan_at = f"段落 {para_no}"

            rep.images.append(info)

            if info.fill_ratio < 0.002:
                rep.add("error", "blank_image",
                        f"第 {idx} 张图几乎是空白(填充率 {info.fill_ratio:.3%})")
            elif info.content_ratio < min_content_ratio:
                rep.add("warning", "wide_bar",
                        f"第 {idx} 张图是宽黑条(内容仅占宽度 {info.content_ratio:.0%})")
            if info.aspect > max_aspect:
                rep.add("warning", "too_tall",
                        f"第 {idx} 张图过高({info.width_px}×{info.height_px}),A4 上读不清")
            if info.width_cm and info.height_cm and info.width_px:
                shown = info.width_cm / info.height_cm
                real = info.width_px / info.height_px
                if abs(shown - real) / real > 0.02:
                    rep.add("warning", "distorted",
                            f"第 {idx} 张图被拉伸(显示比 {shown:.3f} vs 原始比 {real:.3f})")

    if pending is not None:
        rep.add("error", "label_without_image",
                f"标签「{pending}」后面没有图片", pending_at)
    if orphan and not allow_orphan:
        rep.add("warning", "orphan_image",
                f"{orphan} 张图之前没有截图标签(首个在{orphan_at})"
                f";若这些图另有题号标题,可加 --allow-orphan 忽略")
    return rep


def _read_pdf_page_images(page) -> int:
    """Count placed image objects on a page (deduplicated per page)."""
    seen = set()
    for img in page.get_images(full=True):
        seen.add(img[0])
    return len(seen)


def verify_pdf(path: Path, labels: list[str]) -> Report:
    import fitz

    rep = Report(str(path), "pdf")
    rep.label_counts = {label: 0 for label in labels}
    patterns = {label: _label_regex(label) for label in labels}
    total = 0
    with fitz.open(str(path)) as doc:
        for page in doc:
            text = page.get_text()
            for label in labels:
                n = len(patterns[label].findall(text))
                rep.label_counts[label] += n
            total += _read_pdf_page_images(page)
    rep.add("warning", "unexpected_count",
            f"PDF 模式仅统计数量,不做配对与尺寸检查;图片对象 {total} 个")
    return rep


def parse_expect(values: list[str]) -> dict[str, int]:
    out: dict[str, int] = {}
    for raw in values or []:
        if "=" not in raw:
            raise SystemExit(f"--expect 需要 LABEL=N 形式,收到: {raw}")
        key, num = raw.split("=", 1)
        out[key.strip()] = int(num)
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="核对截图证据的完整性与可读性")
    ap.add_argument("path", help="待检查的 .docx 或 .pdf")
    ap.add_argument("-o", "--output", help="写出 JSON 报告")
    ap.add_argument("--label", action="append", dest="labels",
                    help=f"截图标签前缀,可重复;默认 {list(DEFAULT_LABELS)}")
    ap.add_argument("--expect", action="append", default=[],
                    help="期望数量,如 --expect 代码截图=29;可重复")
    ap.add_argument("--min-content-ratio", type=float, default=0.35,
                    help="内容宽度占比下限,低于视为宽黑条(默认 0.35)")
    ap.add_argument("--max-aspect", type=float, default=8.0,
                    help="高宽比上限(默认 8)")
    ap.add_argument("--allow-orphan", action="store_true",
                    help="允许没有截图标签的图(如按题号排版的整屏图)")
    ap.add_argument("--strict", action="store_true", help="有 warning 也返回非零")
    ap.add_argument("--quiet", action="store_true", help="不打印摘要")
    args = ap.parse_args(argv)

    path = Path(args.path)
    if not path.exists():
        raise SystemExit(f"文件不存在: {path}")
    labels = args.labels or list(DEFAULT_LABELS)

    if path.suffix.lower() == ".pdf":
        rep = verify_pdf(path, labels)
    elif path.suffix.lower() == ".docx":
        rep = verify_docx(path, labels, args.min_content_ratio, args.max_aspect,
                          args.allow_orphan)
    else:
        raise SystemExit("只支持 .docx 与 .pdf")

    for label, want in parse_expect(args.expect).items():
        got = rep.label_counts.get(label, 0)
        if got != want:
            rep.add("warning", "unexpected_count",
                    f"标签「{label}」数量 {got},期望 {want}")

    data = rep.as_dict()
    if args.output:
        Path(args.output).write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    if not args.quiet:
        print(f"[{data['kind']}] {path.name}")
        for label, n in data["labels"].items():
            print(f"  标签 {label}: {n}")
        print(f"  图片 {data['images']['total']} 张"
              f"(已标注 {data['images']['labeled']},孤立 {data['images']['orphan']})")
        for f in rep.findings:
            mark = "ERROR" if f.level == "error" else "warn "
            loc = f" @{f.where}" if f.where else ""
            print(f"  {mark} {f.code}: {f.detail}{loc}")
        verdict = "通过" if data["summary"]["passed"] else "未通过"
        print(f"  结论: {verdict}({data['summary']['errors']} error,"
              f" {data['summary']['warnings']} warning)")

    if rep.counts("error"):
        return 1
    if args.strict and rep.counts("warning"):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
