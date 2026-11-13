#!/usr/bin/env python3
"""生成参赛用的三页方案 PPT（每页一屏）。

按赛事模板：① 场景与问题 ② 技能与实现 ③ 效果与复用。
姓名、专业班级、试用数据留成占位符，请在 `PEOPLE` 里改完再重新生成：

    python submission/build_ppt.py

依赖 python-pptx；不需要模板文件，版式全部用图形画。
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).resolve().parent
ASSETS = HERE / "assets"

# ── 待填 ────────────────────────────────────────────────────────────
PEOPLE = {
    "name": "【姓名】",
    "class_name": "【专业班级】",
    "skill": "STDU2026-方案文档合规初审助手",
}

INK = RGBColor(0x1F, 0x23, 0x29)
MUTED = RGBColor(0x64, 0x6A, 0x73)
BLUE = RGBColor(0x0B, 0x63, 0xF6)
ORANGE = RGBColor(0xF5, 0x74, 0x1A)
CARD = RGBColor(0xF4, 0xF6, 0xFA)
LINE = RGBColor(0xE1, 0xE5, 0xEA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x1E, 0x7B, 0x34)

W, H = Inches(13.333), Inches(7.5)


def textbox(slide, x, y, w, h, text="", size=12, bold=False, color=INK, align=PP_ALIGN.LEFT,
            line_spacing=1.25, anchor=MSO_ANCHOR.TOP, font="微软雅黑"):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.line_spacing = line_spacing
        run = para.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = font
    return box


def card(slide, x, y, w, h, fill=CARD, radius=0.045):
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.adjustments[0] = radius
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = fill
    shape.shadow.inherit = False
    return shape


def rule(slide, x, y, w, color=LINE, weight=1.0):
    from pptx.enum.shapes import MSO_SHAPE

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Pt(weight))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def kicker(slide, text, color=BLUE):
    textbox(slide, Inches(0.7), Inches(0.42), Inches(11.9), Inches(0.3), text, size=12, bold=True, color=color)


def title(slide, text, size=27):
    textbox(slide, Inches(0.7), Inches(0.78), Inches(11.9), Inches(0.95), text, size=size, bold=True, color=INK,
            line_spacing=1.15)


def footer(slide, text):
    rule(slide, Inches(0.7), Inches(6.86), Inches(11.93))
    textbox(slide, Inches(0.7), Inches(6.95), Inches(11.9), Inches(0.3), text, size=9, color=MUTED)


def slide1(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "① 场景与问题（为什么做）")
    title(s, "方案报审被打回，往往不是技术不行，而是\n格式与要件没对齐")

    cards = [
        ("谁在用", "项目技术负责人、施工单位资料员、\n课程设计／实验报告的指导教师与学生。\n\n共同点：都要在**别人的模板**里交东西。"),
        ("什么场景", "天窗点前夜赶施工方案；\n报监理前的最后一轮自查；\n课程报告交给老师前的“格式过一遍”。\n\n时间紧、模板长、条款分散在七八个文件里。"),
        ("原来卡在哪", "人工逐条对照模板：15–20 分钟起；\n缺节漏项全靠眼睛，编号错位看不出来；\n想改格式又怕整篇版式被改乱，只好重排。\n\n退回一次 = 至少一天返工。"),
    ]
    xs = [Inches(0.7), Inches(4.85), Inches(9.0)]
    for (head, body), x in zip(cards, xs):
        card(s, x, Inches(1.95), Inches(3.9), Inches(2.55), WHITE)
        rule(s, x, Inches(1.95), Inches(3.9), BLUE, 3.0)
        textbox(s, x + Inches(0.28), Inches(2.14), Inches(3.34), Inches(0.3), head, size=15, bold=True, color=BLUE)
        textbox(s, x + Inches(0.28), Inches(2.52), Inches(3.34), Inches(1.8), body.replace("**", ""), size=11.5,
                color=INK, line_spacing=1.4)

    textbox(s, Inches(0.7), Inches(4.78), Inches(11.9), Inches(0.3), "真实样本里的三个硬伤", size=13, bold=True,
            color=INK)
    chips = [
        ("13", "处【待补充】留在表格单元格里\n（正文看着是齐的）"),
        ("10", "个必备章节全缺\n（工程概况、编制依据、施工方案…）"),
        ("0", "个页码域\n（页脚是空的，手打的不算）"),
    ]
    xs = [Inches(0.7), Inches(4.85), Inches(9.0)]
    for (num, label), x in zip(chips, xs):
        card(s, x, Inches(5.15), Inches(3.9), Inches(1.15))
        textbox(s, x + Inches(0.24), Inches(5.3), Inches(0.78), Inches(0.8), num, size=32, bold=True, color=ORANGE)
        textbox(s, x + Inches(1.05), Inches(5.32), Inches(2.6), Inches(0.9), label, size=10.5, color=INK,
                line_spacing=1.3)

    footer(s, "数据来源：D:/report-skill/outputs/第四小组-基于C51的独居老人智能监护系统-方案设计报告.docx（48 段 / 9 表 / 12 个包内文件），"
              "由 review.py 实测；不是估计值。")


def slide2(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "② 技能与实现（怎么做）", ORANGE)
    textbox(s, Inches(0.7), Inches(0.78), Inches(11.9), Inches(0.5), PEOPLE["skill"], size=27, bold=True, color=INK)
    textbox(s, Inches(0.7), Inches(1.32), Inches(11.9), Inches(0.4),
            "一句话定位：把“这类文档交上去必须满足什么”写成可执行的规则包，一条命令给出分数、等级与逐条修复建议。",
            size=12.5, color=MUTED)

    feats = [
        ("01 一键体检", "输入：一份 .docx\n输出：得分／等级／问题清单\n（每条带位置、详情、怎么改）",
         "0.6 秒跑完 48 段 9 表；\n退出码 0/1，可直接做 CI 卡口；\n--html 产出可截图的报告。"),
        ("02 规则包", "输入：rules/*.json\n输出：该文种的检查口径",
         "内置铁路施工方案（10 个必备章节）、\n实验报告（七要素）、通用体检；\n换单位只改 JSON，不改代码。"),
        ("03 零漂移修复", "输入：原件 + 问题清单\n输出：只改该改的定稿",
         "复用 repair_docx.py 自动修编号／空段落／页码域；\nvalidate_docx.py 逐项对比格式指纹，\n证明“没有动别人的版式”。"),
    ]
    xs = [Inches(0.7), Inches(4.85), Inches(9.0)]
    for (head, io, detail), x in zip(feats, xs):
        card(s, x, Inches(1.95), Inches(3.9), Inches(2.75), WHITE)
        rule(s, x, Inches(1.95), Inches(3.9), ORANGE, 3.0)
        textbox(s, x + Inches(0.26), Inches(2.12), Inches(3.4), Inches(0.3), head, size=14, bold=True, color=ORANGE)
        textbox(s, x + Inches(0.26), Inches(2.48), Inches(3.4), Inches(1.0), io, size=10.5, color=INK, line_spacing=1.35)
        rule(s, x + Inches(0.26), Inches(3.58), Inches(3.36), LINE)
        textbox(s, x + Inches(0.26), Inches(3.68), Inches(3.4), Inches(0.9), detail, size=10, color=MUTED,
                line_spacing=1.35)

    textbox(s, Inches(0.7), Inches(4.92), Inches(11.9), Inches(0.3), "完整流程", size=12, bold=True, color=INK)
    steps = ["原文档", "分析 spec", "体检打分", "按清单修复", "复检 + 格式证明"]
    x = Inches(0.7)
    for i, step in enumerate(steps):
        card(s, x, Inches(5.22), Inches(2.0), Inches(0.46), CARD)
        textbox(s, x, Inches(5.28), Inches(2.0), Inches(0.35), step, size=11, bold=True, color=INK,
                align=PP_ALIGN.CENTER)
        x += Inches(2.0)
        if i < len(steps) - 1:
            textbox(s, x, Inches(5.28), Inches(0.33), Inches(0.3), "→", size=13, bold=True, color=ORANGE,
                    align=PP_ALIGN.CENTER)
            x += Inches(0.33)

    tech = ("技术路径：Python 3.10+ / python-docx + lxml 直接读写 OOXML；Document Specification 作为唯一事实来源；"
            "规则包 JSON 驱动判定；PAGE 域、样式链、编号定义均解析原始 XML，不依赖 Word 进程。\n"
            "差异化：不做“帮你写一篇”，而是“判你这份能不能交”——确定性规则 + 可量化结论 + 不动版式；"
            "已用 82 个 pytest 用例与 CI 守住回归。")
    textbox(s, Inches(0.7), Inches(5.86), Inches(11.9), Inches(0.9), tech, size=10, color=MUTED, line_spacing=1.4)
    footer(s, "可运行证据：CI 全绿（GitHub Actions）；本地 pytest 82 passed；review.py / validate_docx.py 均有命令行实测截图。")


def slide3(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "③ 效果与复用（做得怎么样）", GREEN)
    textbox(s, Inches(0.7), Inches(0.78), Inches(11.9), Inches(0.55),
            "同一份方案：73 分 C 级 → 100 分 A 级", size=26, bold=True, color=INK)
    textbox(s, Inches(0.7), Inches(1.45), Inches(11.9), Inches(0.35),
            "初稿与定稿逐字可比：只补了缺的那一节、把“待定”写实、加了页码域；正文样式一个字没动。",
            size=11.5, color=MUTED)

    rows = [
        ("判断依据", "凭印象：\"看着差不多\"", "73 分 C 级 → 100 分 A 级，附 3 条到 0 条问题清单"),
        ("耗时", "人工逐条对照 15–20 分钟", "0.6 秒 / 份，可批量"),
        ("漏项", "缺一节、漏一个页码域，肉眼很难发现", "缺章节、占位符、PAGE 域、层级跳级、字号越界自动列出"),
        ("返工风险", "为改格式重排，版式全乱", "格式差异 1 项，且正是新增的页码域（其余 20 段未变）"),
    ]
    top = Inches(1.95)
    col = [Inches(0.7), Inches(2.5), Inches(6.05)]
    head_w = [Inches(1.8), Inches(3.55), Inches(5.9)]
    for j, head in enumerate(["", "以前", "现在"]):
        textbox(s, col[j], top, head_w[j], Inches(0.3), head, size=11, bold=True,
                color=MUTED if j else INK)
    y = top + Inches(0.38)
    for label, before, after in rows:
        rule(s, Inches(0.7), y - Inches(0.07), Inches(11.25))
        textbox(s, col[0], y, head_w[0], Inches(0.6), label, size=11, bold=True, color=INK)
        textbox(s, col[1], y, head_w[1], Inches(0.6), before, size=10.5, color=MUTED, line_spacing=1.3)
        textbox(s, col[2], y, head_w[2], Inches(0.6), after, size=10.5, color=INK, line_spacing=1.3)
        y += Inches(0.62)

    card(s, Inches(0.7), Inches(4.9), Inches(5.5), Inches(1.55), WHITE)
    textbox(s, Inches(0.94), Inches(5.05), Inches(5.0), Inches(0.3), "能否复用", size=12, bold=True, color=GREEN)
    textbox(s, Inches(0.94), Inches(5.37), Inches(5.0), Inches(1.0),
            "换单位／换文种 = 复制一份 rules/*.json 改字段：\n必备章节、字号白名单、页面要求都在 JSON 里，\n代码零改动；新增一套约 30 分钟。",
            size=10, color=INK, line_spacing=1.35)

    card(s, Inches(6.45), Inches(4.9), Inches(5.5), Inches(1.55), WHITE)
    textbox(s, Inches(6.69), Inches(5.05), Inches(5.0), Inches(0.3), "下一步", size=12, bold=True, color=GREEN)
    textbox(s, Inches(6.69), Inches(5.37), Inches(5.0), Inches(1.0),
            "① 补编号一致性、图表题注、术语统一三类规则；\n② 体检结果接 repair 自动修，做成“一键出定稿”；\n③ 征集真实方案样本，建立回归基准集。",
            size=10, color=INK, line_spacing=1.35)

    textbox(s, Inches(0.7), Inches(6.52), Inches(11.9), Inches(0.3),
            f"验证者：{PEOPLE['name']}（{PEOPLE['class_name']}）　试用【待补充：N 人 / M 次，及反馈原文要点】　"
            "—— 这一行必须换成真实试用记录再提交。", size=9.5, color=MUTED)
    footer(s, "截图见随作品提交的 docs/screenshots/review-draft.png 与 review-final.png（同一命令生成的 HTML 报告）。")


def main() -> int:
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    slide1(prs)
    slide2(prs)
    slide3(prs)
    out = HERE / f"{PEOPLE['skill']}-{PEOPLE['name']}-{PEOPLE['class_name']}.pptx"
    prs.save(str(out))
    print(f"[ok] {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
