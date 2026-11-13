#!/usr/bin/env python3
"""生成参赛方案 PPT —— 默认五页（封面 + 三页正文 + 落地路线），也可只出赛事要求的三页。

作品名：一稿过（OneDraft）· Word 智能体指南书
作品定位：一份给智能体的 Word 工作指南书（AGENT-GUIDE.md + capabilities.json + docx_agent.py），
审查只是它的一个能力，覆盖 分析→生成→修复→渲染→体检→验证 的完整闭环。

    python submission/build_ppt.py            # 五页（封面 / 场景 / 实现 / 效果 / 路线）
    python submission/build_ppt.py --pages 3  # 只出赛事模板要求的三页正文

姓名、专业班级、试用数据留成占位符，请在 `PEOPLE` 里改完再重新生成。依赖 python-pptx。
"""

from __future__ import annotations

import argparse
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
    "school": "石家庄铁道大学",
    "contact": "【手机 / 邮箱】",
    "skill": "STDU2026-一稿过",
}
BRAND = "一稿过"
BRAND_EN = "OneDraft"

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
    for i, line in enumerate(text.split("\n")):
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


def cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "石家庄铁道大学 · 天翼智创杯 TeleAgent 智能体创新竞赛　｜　选题：铁路与交通（工务 / 信号施工方案）")
    textbox(s, Inches(0.7), Inches(1.55), Inches(7.4), Inches(1.7), BRAND, size=104, bold=True, color=INK,
            line_spacing=1.0)
    textbox(s, Inches(0.75), Inches(3.28), Inches(6.6), Inches(0.4), f"{BRAND_EN}　·　Word 智能体指南书", size=20,
            bold=True, color=BLUE)
    textbox(s, Inches(0.7), Inches(3.92), Inches(11.4), Inches(0.9),
            "让智能体改 Word 时，格式不动、数据不编、改完能验 —— 把“怎么改才算对”写成它能执行的判断。",
            size=22, bold=True, color=INK, line_spacing=1.35)

    stats = [("94", "个自动化测试全绿"), ("0.6", "秒 / 份合规体检"), ("0", "目标改动之外的格式漂移")]
    xs = [Inches(0.7), Inches(4.85), Inches(9.0)]
    for (num, label), x in zip(stats, xs):
        card(s, x, Inches(5.05), Inches(3.9), Inches(0.95), WHITE)
        textbox(s, x + Inches(0.24), Inches(5.17), Inches(0.86), Inches(0.5), num, size=24, bold=True, color=BLUE)
        textbox(s, x + Inches(1.16), Inches(5.28), Inches(2.5), Inches(0.55), label, size=11, color=INK,
                line_spacing=1.3)

    textbox(s, Inches(0.7), Inches(6.28), Inches(11.9), Inches(0.34),
            f"参赛人：{PEOPLE['name']}　·　{PEOPLE['class_name']}　·　{PEOPLE['school']}　·　{PEOPLE['contact']}",
            size=12, bold=True, color=INK)
    footer(s, "作品本体：word-document-skill 仓库（AGENT-GUIDE.md / capabilities.json / scripts/docx_agent.py），MIT 许可，代码与测试公开可复现")
    return s


def slide1(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "① 场景与问题（为什么做）")
    title(s, "智能体改 Word 翻车，几乎从不是“写不出来”，\n而是改了别人的格式、编了不知道的数据、改完没验", size=24)

    cards = [
        ("谁在用", "接 Word 任务的智能体（Claude / Cursor / Coze）+ 资料员、教师、学生。\n\n"
                   "共同点：都要在别人的模板里交东西，且格式不能变。"),
        ("什么场景", "施工方案报监理前的最后一轮自查；\n课程报告交给老师前的“格式过一遍”；\n一份模板套 20 份数据的批量生成。\n\n"
                     "模板长、条款散、每次都赶时间。"),
        ("原来卡在哪", "每个新会话都要重新教一遍“这份文档的样式叫什么”；\n改完说不清动了哪些格式；\n"
                       "缺数据时模型顺手编一个，评审一眼看出。\n\n退回一次 = 至少一天返工。"),
    ]
    xs = [Inches(0.7), Inches(4.85), Inches(9.0)]
    for (head_t, body), x in zip(cards, xs):
        card(s, x, Inches(1.95), Inches(3.9), Inches(2.6), WHITE)
        rule(s, x, Inches(1.95), Inches(3.9), BLUE, 3.0)
        textbox(s, x + Inches(0.28), Inches(2.14), Inches(3.34), Inches(0.3), head_t, size=15, bold=True, color=BLUE)
        textbox(s, x + Inches(0.28), Inches(2.52), Inches(3.34), Inches(1.9), body, size=11, color=INK,
                line_spacing=1.4)

    textbox(s, Inches(0.7), Inches(4.8), Inches(11.9), Inches(0.3), "真实样本里的三个硬伤（review.py 实测，不是估计）",
            size=13, bold=True, color=INK)
    chips = [
        ("13", "处【待补充】藏在表格单元格里\n（正文看着是齐的）"),
        ("10", "个必备章节全缺\n（工程概况、编制依据、施工方案…）"),
        ("0", "个页码域\n（页脚是空的，手打的不算）"),
    ]
    for (num, label), x in zip(chips, xs):
        card(s, x, Inches(5.2), Inches(3.9), Inches(1.1))
        textbox(s, x + Inches(0.24), Inches(5.33), Inches(0.78), Inches(0.8), num, size=30, bold=True, color=ORANGE)
        textbox(s, x + Inches(1.02), Inches(5.36), Inches(2.66), Inches(0.85), label, size=10.5, color=INK,
                line_spacing=1.3)

    footer(s, "样本：outputs/…方案设计报告.docx（48 段 / 9 表 / 12 个包内文件），体检 64/100 C；复现命令见 submission/README.md")


def slide2(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "② 技能与实现（怎么做）", ORANGE)
    textbox(s, Inches(0.7), Inches(0.78), Inches(11.9), Inches(0.5), PEOPLE["skill"], size=26, bold=True, color=INK)
    textbox(s, Inches(0.7), Inches(1.30), Inches(11.9), Inches(0.34),
            "一本可执行的指南书：判断写进 AGENT-GUIDE.md，清单写进 capabilities.json，动作收敛到一个入口 docx_agent.py。",
            size=12.5, color=MUTED)

    feats = [
        ("01 读得懂：分析", "输入：任意 .docx → 输出：spec.json\n（页面 / 样式 / 编号 / 页眉页脚 / 域 / 表格）",
         "样式链按 basedOn 层层解析，主题字体还原成真实字体；编号定义取自 numbering.xml。不靠肉眼猜格式。"),
        ("02 写得进：生成与修复", "输入：模板 + content.json → 输出：形制不变的 docx\n（insert 按锚点插入 / fill 只填占位符）",
         "不新增样式、不改页面设置；编号混用、空段落、居中加粗的假标题由 repair 按问题清单批量修掉。"),
        ("03 验得住：验证与体检", "输入：修改前后两份 docx → 输出：差异 + 分数 + 逐页图\n（格式指纹 / 三级清单 / PNG 复核）",
         "差异可指名到段落与属性；规则包给阻断·严重·提示三级结论，每条都带位置与修法，退出码可直接做卡口。"),
    ]
    xs = [Inches(0.7), Inches(4.85), Inches(9.0)]
    for (head_t, io, detail), x in zip(feats, xs):
        card(s, x, Inches(1.78), Inches(3.9), Inches(2.42), WHITE)
        rule(s, x, Inches(1.78), Inches(3.9), ORANGE, 3.0)
        textbox(s, x + Inches(0.26), Inches(1.96), Inches(3.4), Inches(0.3), head_t, size=13.5, bold=True, color=ORANGE)
        textbox(s, x + Inches(0.26), Inches(2.36), Inches(3.4), Inches(0.66), io, size=9.5, color=INK,
                line_spacing=1.35)
        textbox(s, x + Inches(0.26), Inches(3.12), Inches(3.4), Inches(0.95), detail, size=9, color=MUTED,
                line_spacing=1.4)

    textbox(s, Inches(0.7), Inches(4.44), Inches(11.9), Inches(0.26),
            "命令行实录 —— 同一份方案，改前 73 分 C、改后 100 分 A（文本取自真实运行输出）", size=11, bold=True, color=INK)
    s.shapes.add_picture(str(ASSETS / "terminal-review.png"), Inches(0.7), Inches(4.74), width=Inches(11.93))

    textbox(s, Inches(0.7), Inches(6.44), Inches(11.9), Inches(0.26),
            "分析 spec  →  规划改动  →  执行 / 修复  →  验证 + 渲染  →  交付说明（改了什么 / 没改什么 / 哪些【待补充】）",
            size=10, bold=True, color=INK)
    footer(s, "技术路径：Python 3.10+，python-docx + lxml 直接读写 OOXML，不依赖 Word 进程　｜　"
              "可运行证据：CI 全绿 · pytest 94 passed · docx_agent.py --doctor 一键自检依赖与清单一致性")


def slide3(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "③ 效果与复用（值不值）", GREEN)
    title(s, "把“怎么改 Word 才算对”写成智能体能执行的判断", size=24)

    stats = [("94", "个自动化测试"), ("0.6", "秒/份体检"), ("0", "目标改动之外的漂移"), ("14", "个任务，一条命令入口")]
    xs = [Inches(0.7), Inches(3.745), Inches(6.79), Inches(9.835)]
    for (num, label), x in zip(stats, xs):
        card(s, x, Inches(1.72), Inches(2.85), Inches(1.0), WHITE)
        textbox(s, x + Inches(0.22), Inches(1.85), Inches(1.2), Inches(0.55), num, size=26, bold=True, color=GREEN)
        textbox(s, x + Inches(0.22), Inches(2.41), Inches(2.45), Inches(0.28), label, size=9.5, color=MUTED)

    card(s, Inches(0.7), Inches(2.95), Inches(5.85), Inches(2.35), WHITE)
    textbox(s, Inches(0.95), Inches(3.08), Inches(5.3), Inches(0.3), "以前", size=12, bold=True, color=MUTED)
    textbox(s, Inches(0.95), Inches(3.42), Inches(5.35), Inches(1.75),
            "· 逐条对照模板，15–20 分钟/份\n"
            "· 缺节漏项靠眼睛，编号错位看不出\n"
            "· 改格式怕整篇版式被改乱，只好重排\n"
            "· 每个新会话都重新教一遍脚本用法\n"
            "· 缺数据时模型顺手编，评审一眼看出",
            size=10, color=MUTED, line_spacing=1.45)

    card(s, Inches(6.78), Inches(2.95), Inches(5.85), Inches(2.35), WHITE)
    rule(s, Inches(6.78), Inches(2.95), Inches(5.85), GREEN, 3.0)
    textbox(s, Inches(7.03), Inches(3.08), Inches(5.3), Inches(0.3), "现在", size=12, bold=True, color=GREEN)
    textbox(s, Inches(7.03), Inches(3.42), Inches(5.35), Inches(1.75),
            "· 一条命令 0.6 秒出分数与逐条修法\n"
            "· 格式指纹逐项比对，差异可指名\n"
            "· 只改目标段落，版式一字不动\n"
            "· 读 capabilities.json 自动路由 14 个任务\n"
            "· 缺数据写【待补充】，不进正文",
            size=10, color=INK, line_spacing=1.45)

    textbox(s, Inches(0.7), Inches(5.45), Inches(11.9), Inches(0.3), "复用：换单位 / 换文种 / 换智能体，各只需改一处", size=12,
            bold=True, color=INK)
    reuse = [("换单位", "改 rules/*.json 的必备章节与字号白名单"), ("换文种", "加一个规则包，代码零改动"),
             ("换智能体", "把 capabilities.json 喂给它，任务表即接口")]
    xs = [Inches(0.7), Inches(4.6), Inches(8.5)]
    for (head_t, body), x in zip(reuse, xs):
        card(s, x, Inches(5.8), Inches(3.75), Inches(0.75), CARD)
        textbox(s, x + Inches(0.22), Inches(5.88), Inches(3.3), Inches(0.26), head_t, size=10.5, bold=True, color=INK)
        textbox(s, x + Inches(0.22), Inches(6.16), Inches(3.3), Inches(0.3), body, size=9, color=MUTED)

    textbox(s, Inches(0.7), Inches(6.62), Inches(11.9), Inches(0.3),
            "验证者：试用【待补充：N 人 / M 次，及反馈原文要点】　　下一步：接入实验室算例，补工务/信号两套真实方案做回归集",
            size=9.5, color=ORANGE)
    footer(s, "全部数字可用仓库命令复现：pytest -q（94 passed）、review.py（0.56–0.61 s）、validate_docx.py（format_delta_count）。")


def closing(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    kicker(s, "④ 落地路线与团队（怎么继续）", GREEN)
    title(s, "从“我这边能跑”到“别人敢拿去用”", size=26)

    cols = [
        ("接下来两周", BLUE, [
            "补工务、信号各 2 套真实方案做回归集",
            "规则包扩到 6 套（施工日志、技术交底、月度总结）",
            "第 3 页「验证者」换成真实试用记录",
        ]),
        ("这学期", ORANGE, [
            "在班级 / 实验室做 20 人试用，收集退回原因",
            "把 AGENT-GUIDE.md 出英文版，接 Coze / Cursor",
            "TOC 与 numbering 规则补齐，覆盖率到 90% 检查项",
        ]),
        ("往下一年", GREEN, [
            "接电信侧文档场景：工单规范、运维手册、交底纪要",
            "做评测基准：30 份带标准答案的文档，公开跑分",
            "把入口封成 MCP 工具，任何智能体可挂载",
        ]),
    ]
    xs = [Inches(0.7), Inches(4.85), Inches(9.0)]
    for (head_t, color, items), x in zip(cols, xs):
        card(s, x, Inches(1.85), Inches(3.9), Inches(2.95), WHITE)
        rule(s, x, Inches(1.85), Inches(3.9), color, 3.0)
        textbox(s, x + Inches(0.26), Inches(2.02), Inches(3.4), Inches(0.3), head_t, size=14, bold=True, color=color)
        body = "\n\n".join(f"· {item}" for item in items)
        textbox(s, x + Inches(0.26), Inches(2.44), Inches(3.4), Inches(2.2), body, size=10, color=INK,
                line_spacing=1.45)

    card(s, Inches(0.7), Inches(5.02), Inches(11.93), Inches(1.15), CARD)
    textbox(s, Inches(0.95), Inches(5.15), Inches(11.4), Inches(0.28), "为什么是我们（不是“再做一个知识库问答”）", size=12,
            bold=True, color=INK)
    textbox(s, Inches(0.95), Inches(5.48), Inches(11.4), Inches(0.6),
            "把“格式不变、数据不编、改完必验”写成智能体可执行的判断与验收卡口 —— 已经有 94 个测试和一个 CI 在守它；"
            "换单位、换文种、换智能体，都只需改一处。",
            size=10.5, color=INK, line_spacing=1.45)

    textbox(s, Inches(0.7), Inches(6.28), Inches(11.9), Inches(0.34),
            f"{BRAND}（{BRAND_EN}）· {PEOPLE['name']} · {PEOPLE['class_name']} · {PEOPLE['contact']}",
            size=12, bold=True, color=INK)
    footer(s, "仓库：github.com/niuniu513-ask/word-document-skill　｜　MIT　｜　全部数字可用仓库命令复现，未做过的试用不写数字")
    return s


def main() -> int:
    parser = argparse.ArgumentParser(description="生成参赛 PPT")
    parser.add_argument("--pages", type=int, choices=(3, 5), default=5, help="3=赛事模板三页正文；5=封面+三页正文+路线（默认）")
    parser.add_argument("--out", type=Path, default=None)
    args = parser.parse_args()

    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    if args.pages == 5:
        cover(prs)
    for build in (slide1, slide2, slide3):
        build(prs)
    if args.pages == 5:
        closing(prs)

    out = args.out or HERE / f"{PEOPLE['skill']}-{PEOPLE['name']}-{PEOPLE['class_name']}.pptx"
    prs.save(out)
    print(f"[ok] {out}  ({args.pages} pages)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
