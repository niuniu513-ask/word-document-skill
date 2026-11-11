#!/usr/bin/env python
"""把已执行的 notebook 渲染成独立 HTML，供无头浏览器截图。

用途：报告需要"真实代码与运行结果截图"时，离屏渲染比驱动 VS Code 窗口更稳：
不占屏幕、一次跑完、每个单元格的边界由 DOM 直接给出，不需要滚动推算。

产物结构（每个 code cell 一个块，data-idx 是原 notebook 的 cell 索引）：

    <div class="cell" data-idx="12" data-seq="5" data-has-output="1">
      <div class="input"><pre class="code">...</pre></div>     代码区
      <div class="output">...</div>                            输出区（文本/图片）
    </div>

配合 scripts/shoot_nb_cells.cjs 截图：分别对 .input / .output / .cell 取整块，
得到代码图、输出图、以及"代码+输出"连体图。

    python -X utf8 scripts/render_nb_html.py NB.ipynb -o EVIDENCE/nb.html
    node scripts/shoot_nb_cells.cjs EVIDENCE/nb.html EVIDENCE/shots
"""
from __future__ import annotations

import argparse
import html
import json
import sys
from pathlib import Path

# 页面宽度与字号共同决定截图插入 Word 后的实际字高：
# 插入宽度 fixed_cm 时，字高 ≈ font_size / width * fixed_cm。
# 880px/14.5px 在 15.2cm 版心上约 6.9pt，属于"缩到版心仍可读"的区间。
DEFAULT_WIDTH = 880
DEFAULT_FONT_SIZE = 14.5
DEFAULT_INSERT_CM = 15.2

CSS = """* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ width:{width}px; background:#fff; padding:0 0 12px 0;
       font-family:"Cascadia Mono","Consolas","Courier New",monospace;
       font-size:{font}px; line-height:1.6; color:#1f2328;
       -webkit-font-smoothing:antialiased; }}
.cell {{ padding:0 14px 14px 14px; }}
.input {{ background:#f6f8fa; border-left:3px solid #4a7bb7; padding:8px 12px;
         white-space:pre; overflow:visible; }}
pre.code {{ font-family:inherit; font-size:{font}px; line-height:1.6; white-space:pre; }}
.output {{ padding:8px 0 2px 0; }}
pre.out-txt {{ font-family:inherit; font-size:{font_out}px; line-height:1.5; white-space:pre;
              color:#24292f; padding:2px 2px 0 2px; }}
pre.out-err {{ font-family:inherit; font-size:{font_out}px; line-height:1.5; white-space:pre;
              color:#b31d28; background:#ffeef0; padding:4px 6px; margin:2px 0; }}
.out-img {{ display:block; max-width:100%; margin:10px auto 4px auto; background:#fff; }}
"""


def _text_of(value) -> str:
    """nbformat 的 text 可能是 str 或 list[str]。"""
    if isinstance(value, list):
        return "".join(value)
    return value or ""


def _cell_outputs(cell: dict) -> list[tuple[str, str]]:
    """抽取 (kind, payload)：kind ∈ {text, png, err}。"""
    out: list[tuple[str, str]] = []
    for o in cell.get("outputs", []):
        otype = o.get("output_type")
        if otype == "stream":
            txt = _text_of(o.get("text"))
            if txt:
                out.append(("text", txt))
        elif otype in ("execute_result", "display_data"):
            data = o.get("data", {}) or {}
            for mime, val in data.items():
                if mime == "image/png":
                    out.append(("png", _text_of(val).replace("\n", "")))
            if "image/png" not in data and "text/plain" in data:
                out.append(("text", _text_of(data["text/plain"])))
        elif otype == "error":
            body = f"{o.get('ename', '')}: {o.get('evalue', '')}"
            tb = o.get("traceback") or []
            if tb:
                body += "\n" + "\n".join(tb)
            out.append(("err", body))
    return out


def build_html(nb: dict, width: int = DEFAULT_WIDTH,
               font_size: float = DEFAULT_FONT_SIZE,
               with_placeholders: bool = False) -> tuple[str, int, int]:
    """返回 (html, code_cell_count, error_cell_count)。"""
    blocks: list[str] = []
    seq = 0
    errors = 0
    for raw_idx, cell in enumerate(nb.get("cells", [])):
        if cell.get("cell_type") != "code":
            continue
        source = _text_of(cell.get("source")).rstrip()
        if not source.strip():
            continue
        seq += 1
        outs = _cell_outputs(cell)
        if any(k == "err" for k, _ in outs):
            errors += 1

        body: list[str] = []
        for kind, payload in outs:
            if kind == "text":
                body.append(f"<pre class='out-txt'>{html.escape(payload.rstrip())}</pre>")
            elif kind == "err":
                body.append(f"<pre class='out-err'>{html.escape(payload.rstrip())}</pre>")
            else:
                body.append(f"<img class='out-img' src='data:image/png;base64,{payload}'>")
        if not body:
            if not with_placeholders:
                body.append("")
            else:
                body.append("<pre class='out-txt'>(无输出)</pre>")

        has_out = 0 if not outs else 1
        blocks.append(
            f"<div class='cell' data-idx='{raw_idx}' data-seq='{seq}' "
            f"data-has-output='{has_out}'>"
            f"<div class='input'><pre class='code'>{html.escape(source)}</pre></div>"
            f"<div class='output'>{''.join(body)}</div></div>"
        )

    css = CSS.format(width=width, font=font_size, font_out=round(font_size - 0.5, 1))
    doc = (
        '<!DOCTYPE html><html><head><meta charset="utf-8">'
        f"<style>{css}</style></head><body>\n{chr(10).join(blocks)}\n</body></html>"
    )
    return doc, seq, errors


def main() -> int:
    ap = argparse.ArgumentParser(description="notebook → 可截图的独立 HTML")
    ap.add_argument("notebook", type=Path)
    ap.add_argument("-o", "--out", type=Path, required=True)
    ap.add_argument("--width", type=int, default=DEFAULT_WIDTH,
                    help="渲染页宽 px（默认 880；越小则插入 Word 后字越大）")
    ap.add_argument("--font-size", type=float, default=DEFAULT_FONT_SIZE,
                    help="代码字号 px（默认 14.5）")
    ap.add_argument("--insert-cm", type=float, default=DEFAULT_INSERT_CM,
                    help="仅供提示：报告里的图片插入宽度 cm（默认 15.2）")
    ap.add_argument("--allow-empty-output", action="store_true",
                    help="无输出的 cell 也画占位文字（默认留空）")
    ap.add_argument("--allow-error", action="store_true",
                    help="允许 notebook 含报错输出（默认拒绝，证据不能带报错）")
    args = ap.parse_args()

    if not args.notebook.exists():
        print(f"找不到 notebook: {args.notebook}", file=sys.stderr)
        return 2

    nb = json.loads(args.notebook.read_text(encoding="utf-8"))
    doc, cells, errors = build_html(nb, args.width, args.font_size, args.allow_empty_output)
    if cells == 0:
        print("notebook 里没有非空 code cell", file=sys.stderr)
        return 2
    if errors and not args.allow_error:
        print(
            f"notebook 含 {errors} 个报错 cell；先修好再截图，或用 --allow-error",
            file=sys.stderr,
        )
        return 3

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(doc, encoding="utf-8")

    pt = args.insert_cm / 2.54 * 72 * args.font_size / args.width
    print(f"cells={cells} errors={errors} -> {args.out}")
    print(
        f"预估字高: 插入 {args.insert_cm:.1f}cm 宽时约 {pt:.1f}pt"
        "（<5pt 说明需要调小 --width 或调大 --font-size）"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
