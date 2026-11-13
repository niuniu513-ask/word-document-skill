#!/usr/bin/env python3
"""review.py — 一键文档合规体检。

给一份 .docx 和一套规则包（rules/*.json），输出中文体检报告：
得分、等级、问题清单、每条问题的位置与修复建议。

设计取舍：
- 结构、样式、页面这类"文档骨架"事实取自 analyze_docx.py 生成的 Document Specification，
  避免两套解析逻辑各说各话；
- 占位符、禁用语、空段落这类"正文文字"事实直接读 python-docx，因为要点名到段落/单元格。

用法：
    python scripts/review.py 方案.docx
    python scripts/review.py 方案.docx --profile laboratory -o review.md --html review.html
    python scripts/review.py 方案.docx --baseline 原件.docx   # 附带格式漂移检查
退出码：0 = 通过（无阻断项），1 = 需修改，2 = 用法/文件错误。
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document

HERE = Path(__file__).resolve().parent
RULES_DIR = HERE.parent / "rules"

PENALTY = {"blocker": 15, "major": 6, "minor": 2}
LEVEL_CN = {"blocker": "阻断", "major": "严重", "minor": "提示"}
GRADE_CUT = [(90, "A", "可直接报审"), (75, "B", "小改后可用"), (60, "C", "需补充要件"), (0, "D", "退回重做")]


@dataclass
class Finding:
    rule: str
    level: str
    title: str
    detail: str = ""
    where: str = ""
    fix: str = ""

    def as_dict(self) -> dict:
        return {
            "rule": self.rule,
            "level": self.level,
            "level_cn": LEVEL_CN.get(self.level, self.level),
            "title": self.title,
            "detail": self.detail,
            "where": self.where,
            "fix": self.fix,
        }


@dataclass
class Report:
    path: Path
    profile: dict
    findings: list[Finding] = field(default_factory=list)
    facts: dict = field(default_factory=dict)

    @property
    def score(self) -> int:
        return max(0, 100 - sum(PENALTY.get(f.level, 2) for f in self.findings))

    @property
    def grade(self) -> tuple[str, str]:
        for cut, letter, hint in GRADE_CUT:
            if self.score >= cut:
                return letter, hint
        return "D", GRADE_CUT[-1][2]

    def count(self, level: str) -> int:
        return sum(1 for f in self.findings if f.level == level)

    @property
    def blocked(self) -> bool:
        return self.count("blocker") > 0


# ---------------------------------------------------------------- 事实提取


def load_profile(name: str) -> dict:
    path = Path(name)
    if not path.exists():
        path = RULES_DIR / f"{name}.json"
    if not path.exists():
        raise SystemExit(f"ERROR: 找不到规则包 {name}（可选：{[p.stem for p in sorted(RULES_DIR.glob('*.json'))]}）")
    return json.loads(path.read_text(encoding="utf-8"))


def run_analyzer(path: Path) -> dict:
    """调用 analyze_docx.py 拿到 Document Specification。"""
    out = Path(tempfile_tmp()) / "spec.json"
    proc = subprocess.run(
        [sys.executable, str(HERE / "analyze_docx.py"), str(path), "-o", str(out)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if not out.exists():
        raise SystemExit(f"ERROR: 文档分析失败：{(proc.stderr or proc.stdout or '').strip()[-400:]}")
    return json.loads(out.read_text(encoding="utf-8"))


def tempfile_tmp() -> str:
    import tempfile

    return tempfile.mkdtemp(prefix="docx-review-")


def iter_texts(doc: Document):
    """正文段落与表格单元格的文字，带可定位的位置标签。"""
    for i, p in enumerate(doc.paragraphs):
        if p.text and p.text.strip():
            yield f"第 {i + 1} 段", p.text.strip()
    for ti, table in enumerate(doc.tables, 1):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                text = cell.text.strip()
                if text:
                    yield f"表 {ti} 第 {ri + 1} 行第 {ci + 1} 列", text


def has_field(section, part: str, code: str) -> bool:
    try:
        obj = getattr(section, part)
    except Exception:
        return False
    return code in obj._element.xml


# ---------------------------------------------------------------- 检查项


def check_required_sections(rep: Report, spec: dict) -> None:
    required = rep.profile.get("required_sections") or []
    if not required:
        return
    titles = [str(e.get("text") or "") for e in spec.get("outline") or []]
    joined = "\n".join(titles)
    missing = []
    for sec in required:
        aliases = sec.get("aliases") or [sec["key"]]
        if not any(a and a in joined for a in aliases):
            missing.append(sec["key"])
    rep.facts["required_sections"] = len(required)
    rep.facts["missing_sections"] = missing
    if missing:
        rep.findings.append(
            Finding(
                "required_sections",
                "blocker",
                f"缺少 {len(missing)} 个必备章节",
                "报审时这几节最容易被要求补件：" + "、".join(missing),
                where="全文大纲",
                fix="按规则包补齐小节标题（用标题样式，不要用加粗段落伪装），或说明该节不适用的理由。",
            )
        )


def check_heading_levels(rep: Report, spec: dict) -> None:
    rule = rep.profile.get("heading") or {}
    outline = spec.get("outline") or []
    rep.facts["heading_count"] = len(outline)
    prev = 0
    jumps = []
    for entry in outline:
        level = int(entry.get("level") or 0)
        if prev and rule.get("forbid_level_jump", True) and level - prev > 1:
            jumps.append(f"「{entry.get('text')}」从 {prev} 级跳到 {level} 级")
        prev = level or prev
    max_level = int(rule.get("max_level") or 4)
    deep = [str(e.get("text")) for e in outline if int(e.get("level") or 0) > max_level]
    if jumps:
        rep.findings.append(
            Finding("heading_level_jump", "major", f"{len(jumps)} 处标题层级跳级", "；".join(jumps[:4]),
                    where="全文大纲", fix="补齐中间层级的标题，或把该标题降/升到相邻层级。")
        )
    if deep:
        rep.findings.append(
            Finding("heading_depth", "minor", f"{len(deep)} 个标题层级超过 {max_level} 级",
                    "过深的层级会让目录臃肿：" + "、".join(deep[:3]),
                    where="全文大纲", fix="把五级及以下标题改为正文加粗段落，或用列表表达。")
        )


def check_heading_numbering(rep: Report, spec: dict) -> None:
    rule = rep.profile.get("heading") or {}
    mode = rule.get("numbering", "any")
    if mode == "any":
        return
    outline = spec.get("outline") or []
    manual = [e for e in outline if e.get("manual_number") and not e.get("native_numbering")]
    auto = [e for e in outline if e.get("native_numbering")]
    rep.facts["manual_numbered_headings"] = len(manual)
    rep.facts["auto_numbered_headings"] = len(auto)
    if manual and auto:
        rep.findings.append(
            Finding("mixed_numbering", "major", "自动编号与手打编号混用",
                    f"{len(auto)} 个标题用 Word 自动编号，{len(manual)} 个是手打的（如「{manual[0].get('text')}」）",
                    where="全文大纲",
                    fix="统一成一种：要么全用多级列表自动编号，要么全部手打并删除自动编号，否则插入/删除章节后编号会错位。")
        )
    elif manual and mode == "auto":
        rep.findings.append(
            Finding("manual_numbering", "major", f"{len(manual)} 个标题的编号是手打的",
                    "手打编号在增删章节后不会自动更新，是排版返工的主要来源。",
                    where="全文大纲", fix="改用多级列表自动编号。")
        )


def check_placeholders(rep: Report, doc: Document) -> None:
    cfg = rep.profile.get("placeholders") or {}
    patterns = [re.compile(p) for p in (cfg.get("patterns") or [r"【[^】]{0,60}】"])]
    hits: list[tuple[str, str]] = []
    for where, text in iter_texts(doc):
        if any(pat.search(text) for pat in patterns):
            hits.append((where, text[:60]))
    rep.facts["placeholders"] = len(hits)
    if hits and cfg.get("forbid", True):
        sample = "；".join(f"{w}：{t}" for w, t in hits[:4])
        rep.findings.append(
            Finding("placeholders", "blocker", f"仍有 {len(hits)} 处占位符未填写",
                    sample + ("" if len(hits) <= 4 else " …"),
                    where="正文与表格",
                    fix="逐处补入真实内容；确实暂时没有的，保留【待补充:说明】并在提交说明里列出，不要删掉占位符假装完成。")
        )


def check_forbidden(rep: Report, doc: Document) -> None:
    phrases = rep.profile.get("forbidden_phrases") or []
    if not phrases:
        return
    hits: list[tuple[str, str, str]] = []
    for where, text in iter_texts(doc):
        for phrase in phrases:
            if phrase and phrase in text:
                hits.append((where, phrase, text[:50]))
    rep.facts["forbidden_phrases"] = len(hits)
    if hits:
        rep.findings.append(
            Finding("forbidden_phrases", "major", f"{len(hits)} 处出现不该出现在正式文稿里的字眼",
                    "；".join(f"{w} 含「{p}」" for w, p, _ in hits[:4]),
                    where="正文与表格", fix="替换为确定表述；未定的内容写成待确认事项单独列出，别留在正文里。")
        )


def check_empty_paragraphs(rep: Report, doc: Document) -> None:
    limit = int((rep.profile.get("body") or {}).get("max_consecutive_empty", 2))
    longest, streak, at = 0, 0, 0
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            streak = 0
        else:
            streak += 1
            if streak > longest:
                longest, at = streak, i + 1
    rep.facts["max_consecutive_empty"] = longest
    if longest > limit:
        rep.findings.append(
            Finding("empty_paragraphs", "minor", f"有 {longest} 个连续空段落",
                    f"出现在第 {at} 段附近；用空回车撑版面会让任何内容改动都错位。",
                    where=f"第 {at} 段附近", fix="删掉空段落，改用段前段后间距或分页符。")
        )


def resolve_font(combo: dict, defaults: dict) -> str | None:
    """把主题字体名（minorEastAsia）还原成真实字体名，避免报告里出现无意义的主题标识。"""
    direct = combo.get("eastAsia")
    if direct:
        return str(direct)
    theme = str(combo.get("eastAsia_theme") or "")
    if not theme:
        return None
    fallback = (defaults or {}).get("eastAsia") or (defaults or {}).get("ascii")
    if "EastAsia" in theme or "ea" in theme:
        return str(fallback or "（主题字体）")
    return str(fallback or theme)


def check_fonts(rep: Report, spec: dict) -> None:
    body = rep.profile.get("body") or {}
    allowed_sizes = body.get("allowed_sizes_pt") or []
    allowed_fonts = body.get("allowed_east_asia_fonts") or []
    inv = spec.get("formatting_inventory") or {}
    dominant = inv.get("dominant") or {}
    defaults = (spec.get("doc_defaults") or {}).get("font") or {}
    size = dominant.get("size_pt")
    font = resolve_font(dominant, defaults)
    rep.facts["dominant_size_pt"] = size
    rep.facts["dominant_font"] = font
    rep.facts["run_count"] = inv.get("run_count")
    if allowed_sizes and size is not None and not any(abs(size - s) < 0.05 for s in allowed_sizes):
        rep.findings.append(
            Finding("font_size", "major", f"正文字号 {size} pt 不在规则要求的范围内",
                    "允许：" + "、".join(f"{s} pt" for s in allowed_sizes),
                    where="正文", fix="统一整篇正文字号；只有一个字号，目录和行距才会稳定。")
        )
    combos = inv.get("font_combinations") or []
    stray = []
    seen: set[str] = set()
    for combo in combos:
        name = resolve_font(combo, defaults)
        if not name or name in seen:
            continue
        if allowed_fonts and not any(f.lower() in name.lower() for f in allowed_fonts):
            seen.add(name)
            stray.append(f"{name}（{combo.get('run_count', 0)} 处）")
    rep.facts["stray_fonts"] = stray
    if stray:
        rep.findings.append(
            Finding("font_family", "minor", f"{len(stray)} 种规定以外的中文字体",
                    "、".join(stray[:4]),
                    where="正文", fix="改用规则包允许的字体；混排字体在打印稿上最显眼。")
        )


def check_page(rep: Report, doc: Document, spec: dict) -> None:
    cfg = rep.profile.get("page") or {}
    sections = doc.sections
    rep.facts["section_count"] = len(sections)
    if cfg.get("require_page_number"):
        no_number = [i + 1 for i, s in enumerate(sections) if not (has_field(s, "footer", "PAGE") or has_field(s, "header", "PAGE"))]
        if no_number:
            rep.findings.append(
                Finding("page_number", "major", "缺少页码域",
                        "这些节里没有 PAGE 域（手打的页码同样不算）：第 " + "、".join(map(str, no_number)) + " 节",
                        where="页脚", fix="插入页码域（PAGE），不要手打数字。")
            )
    if cfg.get("require_header"):
        empty = [i + 1 for i, s in enumerate(sections) if not s.header.paragraphs or not s.header.paragraphs[0].text.strip()]
        if empty:
            rep.findings.append(
                Finding("header", "minor", "页眉为空", "第 " + "、".join(map(str, empty)) + " 节", where="页眉",
                        fix="按模板补页眉（单位名/文件编号）。")
            )
    allowed = cfg.get("allowed_page_sizes") or []
    sizes = []
    for sec in spec.get("sections") or []:
        page = sec.get("page") or sec
        w = page.get("width_cm") or sec.get("page_width_cm")
        h = page.get("height_cm") or sec.get("page_height_cm")
        if w and h:
            sizes.append((round(float(w), 1), round(float(h), 1)))
    rep.facts["page_sizes_cm"] = sizes
    known = {"A4": (21.0, 29.7), "A3": (29.7, 42.0), "B5": (17.6, 25.0)}
    if allowed:
        bad = []
        for w, h in sizes:
            ok = any(abs(w - known[a][0]) < 0.4 and abs(h - known[a][1]) < 0.4 for a in allowed if a in known)
            if not ok:
                bad.append(f"{w}×{h} cm")
        if bad:
            rep.findings.append(
                Finding("page_size", "major", "页面尺寸不符合要求",
                        "检出 " + "、".join(sorted(set(bad))) + "；要求 " + "/".join(allowed),
                        where="页面设置", fix="改为规定纸张（A4 21×29.7 cm），否则打印装订会错位。")
            )


def check_tables(rep: Report, doc: Document) -> None:
    cfg = rep.profile.get("tables") or {}
    max_cols = int(cfg.get("max_cols") or 20)
    wide, empty_header = [], []
    for i, table in enumerate(doc.tables, 1):
        cols = max((len(row.cells) for row in table.rows), default=0)
        if cols > max_cols:
            wide.append(f"表 {i}（{cols} 列）")
        if table.rows:
            first = [c.text.strip() for c in table.rows[0].cells]
            if all(not t for t in first):
                empty_header.append(f"表 {i}")
    rep.facts["table_count"] = len(doc.tables)
    if wide:
        rep.findings.append(
            Finding("table_width", "minor", f"{len(wide)} 张表列数过多", "、".join(wide),
                    where="表格", fix="拆分或用纵向表头（指标名放行、取值放列）。")
        )
    if empty_header:
        rep.findings.append(
            Finding("table_header", "minor", f"{len(empty_header)} 张表缺少表头文字", "、".join(empty_header),
                    where="表格", fix="首行写清各列含义；跨页时设置「重复标题行」。")
        )


def check_body_length(rep: Report, doc: Document) -> None:
    minimum = int((rep.profile.get("body") or {}).get("min_body_paragraphs", 0))
    count = sum(1 for p in doc.paragraphs if p.text.strip())
    rep.facts["body_paragraphs"] = count
    rep.facts["char_count"] = sum(len(p.text) for p in doc.paragraphs)
    if minimum and count < minimum:
        rep.findings.append(
            Finding("body_length", "major", f"正文只有 {count} 段，明显少于 {minimum} 段",
                    "内容量不足通常意味着要素没写全，而不只是写得短。",
                    where="正文", fix="对照必备章节逐节补写。")
        )


def check_baseline(rep: Report, original: Path, final: Path) -> None:
    sys.path.insert(0, str(HERE))
    from validate_docx import compare_format  # noqa: PLC0415

    diff = compare_format(Document(str(original)), Document(str(final)))
    rep.facts["format_delta"] = {
        "summary": diff.get("summary"),
        "drift_count": len(diff.get("drift") or []),
    }
    if diff.get("drift"):
        first = diff["drift"][0]
        rep.findings.append(
            Finding("format_drift", "blocker", f"与原件相比有 {len(diff['drift'])} 处格式漂移",
                    f"例如 {first.get('field')}: {first.get('original')} → {first.get('final')}",
                    where=str(first.get("where") or "全文"),
                    fix="只改内容不改格式：回到原件操作，或用 template-fill 流程在原文档上填。")
        )


# ---------------------------------------------------------------- 渲染


def review_file(path: Path, profile: dict, baseline: Path | None = None) -> Report:
    """跑完整套检查，返回 Report。CLI 与测试共用这一条路径。"""
    rep = Report(path=path, profile=profile)
    doc = Document(str(path))
    spec = run_analyzer(path)
    check_required_sections(rep, spec)
    check_heading_levels(rep, spec)
    check_heading_numbering(rep, spec)
    check_placeholders(rep, doc)
    check_forbidden(rep, doc)
    check_empty_paragraphs(rep, doc)
    check_fonts(rep, spec)
    check_page(rep, doc, spec)
    check_tables(rep, doc)
    check_body_length(rep, doc)
    if baseline is not None:
        check_baseline(rep, baseline, path)
    return rep


def render_markdown(rep: Report) -> str:
    letter, hint = rep.grade
    lines = [
        "# 文档合规体检报告",
        "",
        f"**文件**：{rep.path.name}",
        f"**规则包**：{rep.profile.get('name')}（{rep.profile.get('id')}）",
        f"**结论**：{'🔴 需修改后提交' if rep.blocked else ('🟡 小改即可' if rep.findings else '🟢 通过')}",
        f"**得分**：{rep.score} / 100　**等级**：{letter}（{hint}）",
        f"**问题**：阻断 {rep.count('blocker')}　严重 {rep.count('major')}　提示 {rep.count('minor')}",
        "",
        "## 一、检查项",
        "",
        "| 检查项 | 结果 | 说明 |",
        "| --- | --- | --- |",
    ]
    checks = [
        ("必备章节", f"{rep.facts.get('required_sections', '—')} 节", rep.facts.get("missing_sections")),
        ("标题数量", f"{rep.facts.get('heading_count', '—')} 个", None),
        ("正文段落", f"{rep.facts.get('body_paragraphs', '—')} 段", None),
        ("表格", f"{rep.facts.get('table_count', '—')} 张", None),
        ("占位符", f"{rep.facts.get('placeholders', 0)} 处", None),
        ("正文字号", f"{rep.facts.get('dominant_size_pt', '—')} pt", None),
        ("页面", "、".join(f"{w}×{h}" for w, h in rep.facts.get("page_sizes_cm") or []) or "—", None),
    ]
    for name, value, bad in checks:
        mark = "❌" if bad else "✅"
        note = "缺：" + "、".join(bad) if bad else "—"
        lines.append(f"| {name} | {mark} {value} | {note} |")
    lines += ["", "## 二、问题清单", ""]
    if not rep.findings:
        lines.append("未发现问题。可以提交。")
    else:
        order = {"blocker": 0, "major": 1, "minor": 2}
        for f in sorted(rep.findings, key=lambda x: order.get(x.level, 3)):
            lines += [
                f"### [{LEVEL_CN[f.level]}] {f.title}",
                "",
                f"- 位置：{f.where or '—'}",
                f"- 详情：{f.detail or '—'}",
                f"- 怎么改：{f.fix or '—'}",
                "",
            ]
    lines += ["## 三、文档事实", "", "```json", json.dumps(rep.facts, ensure_ascii=False, indent=2), "```", ""]
    return "\n".join(lines)


HTML_CSS = """
body{font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;margin:0;background:#f5f6f8;color:#1f2329}
.wrap{max-width:860px;margin:0 auto;padding:32px 28px 56px}
h1{font-size:26px;margin:0 0 4px}
.meta{color:#646a73;font-size:13px;margin-bottom:20px}
.hero{background:#fff;border:1px solid #e3e6ea;border-radius:12px;padding:20px 22px;margin-bottom:20px}
.score{font-size:44px;font-weight:700;line-height:1}
.badge{display:inline-block;padding:3px 10px;border-radius:999px;font-size:12px;font-weight:600}
.b-red{background:#fde8e8;color:#c62828}.b-amber{background:#fff4e0;color:#a55a00}.b-green{background:#e6f4ea;color:#1e7b34}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e3e6ea;border-radius:12px;overflow:hidden;font-size:14px}
th,td{padding:9px 12px;border-bottom:1px solid #eef0f3;text-align:left}
th{background:#fafbfc;color:#646a73;font-weight:600}
.card{background:#fff;border:1px solid #e3e6ea;border-left-width:4px;border-radius:10px;padding:14px 18px;margin:12px 0}
.card.blocker{border-left-color:#c62828}.card.major{border-left-color:#a55a00}.card.minor{border-left-color:#2b6cb0}
.card h3{margin:0 0 6px;font-size:16px}
.k{color:#646a73;font-size:12px}
"""


def render_html(rep: Report) -> str:
    letter, hint = rep.grade
    badge = "b-red" if rep.blocked else ("b-amber" if rep.findings else "b-green")
    state = "需修改后提交" if rep.blocked else ("小改即可" if rep.findings else "通过")
    rows = [
        ("必备章节", f"{rep.facts.get('required_sections', '—')} 节", bool(rep.facts.get("missing_sections"))),
        ("标题数量", f"{rep.facts.get('heading_count', '—')} 个", False),
        ("正文段落", f"{rep.facts.get('body_paragraphs', '—')} 段", False),
        ("表格", f"{rep.facts.get('table_count', '—')} 张", False),
        ("占位符", f"{rep.facts.get('placeholders', 0)} 处", bool(rep.facts.get("placeholders"))),
        ("正文字号", f"{rep.facts.get('dominant_size_pt', '—')} pt", False),
    ]
    trs = "\n".join(
        f"<tr><td>{n}</td><td>{'❌' if bad else '✅'} {v}</td></tr>" for n, v, bad in rows
    )
    order = {"blocker": 0, "major": 1, "minor": 2}
    cards = ""
    for f in sorted(rep.findings, key=lambda x: order.get(x.level, 3)):
        cards += (
            f'<div class="card {f.level}"><h3>{f.title}'
            f'<span class="badge {badge}">{LEVEL_CN[f.level]}</span></h3>'
            f'<div class="k">位置</div><div>{f.where or "—"}</div>'
            f'<div class="k">详情</div><div>{f.detail or "—"}</div>'
            f'<div class="k">怎么改</div><div>{f.fix or "—"}</div></div>'
        )
    if not cards:
        cards = '<div class="card minor"><h3>未发现问题</h3><div>可以提交。</div></div>'
    return f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>文档合规体检报告 · {rep.path.name}</title><style>{HTML_CSS}</style></head><body><div class="wrap">
<h1>文档合规体检报告</h1>
<div class="meta">{rep.path.name}　·　规则包 {rep.profile.get('name')}</div>
<div class="hero">
  <div class="score">{rep.score}<span style="font-size:16px;color:#646a73"> / 100</span></div>
  <div style="margin-top:8px"><span class="badge {badge}">{letter} · {state}</span>
  <span class="badge b-amber" style="margin-left:8px">阻断 {rep.count('blocker')} · 严重 {rep.count('major')} · 提示 {rep.count('minor')}</span></div>
  <div class="meta" style="margin:10px 0 0">{hint}</div>
</div>
<table><thead><tr><th>检查项</th><th>结果</th></tr></thead><tbody>{trs}</tbody></table>
<h2 style="font-size:18px;margin:24px 0 4px">问题清单</h2>{cards}
</div></body></html>"""


def main() -> int:
    from _docx_utils import setup_stdout  # noqa: PLC0415

    setup_stdout()
    ap = argparse.ArgumentParser(description="一键文档合规体检")
    ap.add_argument("input")
    ap.add_argument("--profile", default="railway-plan", help="规则包 id 或 JSON 路径")
    ap.add_argument("-o", "--output", default=None, help="Markdown 报告输出路径")
    ap.add_argument("--html", default=None, help="同时输出一份 HTML 报告（用于截图/评审）")
    ap.add_argument("--json", dest="json_out", default=None, help="输出机器可读结果")
    ap.add_argument("--baseline", default=None, help="原件；给出时额外检查格式漂移")
    args = ap.parse_args()

    path = Path(args.input)
    if not path.exists():
        print(f"ERROR: 文件不存在：{path}")
        return 2

    rep = review_file(path, load_profile(args.profile), Path(args.baseline) if args.baseline else None)

    md = render_markdown(rep)
    if args.output:
        Path(args.output).write_text(md, encoding="utf-8")
        print(f"[saved] {args.output}")
    else:
        print(md)
    if args.html:
        Path(args.html).write_text(render_html(rep), encoding="utf-8")
        print(f"[saved] {args.html}")
    if args.json_out:
        Path(args.json_out).write_text(
            json.dumps(
                {
                    "file": str(path),
                    "profile": rep.profile.get("id"),
                    "score": rep.score,
                    "grade": rep.grade[0],
                    "blocked": rep.blocked,
                    "findings": [f.as_dict() for f in rep.findings],
                    "facts": rep.facts,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"[saved] {args.json_out}")

    letter, _ = rep.grade
    print(f"[review] 得分 {rep.score}/100 等级 {letter} 阻断 {rep.count('blocker')} 严重 {rep.count('major')} 提示 {rep.count('minor')}")
    return 1 if rep.blocked else 0


if __name__ == "__main__":
    sys.exit(main())
