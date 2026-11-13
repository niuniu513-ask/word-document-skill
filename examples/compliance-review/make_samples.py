#!/usr/bin/env python3
"""生成两份用于演示合规体检的施工方案样本。

方案-初稿.docx  故意留三类问题：缺"应急预案"一节、正文写了"待定"、页脚没有页码域。
方案-定稿.docx  同样的内容，把这三处改好。

两边正文逐字相同，所以 validate_docx.py 能证明"只改了该改的"。
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

OUT = Path(__file__).resolve().parent

SECTIONS = [
    ("一、工程概况", "本工程为京广线某区间信号设备大修，涉及 3 个站场、12 组道岔。天窗点每日 90 分钟，施工与行车交叉作业。"),
    ("二、编制依据", "依据《铁路信号工程施工质量验收标准》TB 10419、《铁路技术管理规程》及设计文件、施工合同编制。"),
    ("三、施工组织", "设项目经理 1 人、技术负责人 1 人、安全员 2 人，分三个作业班组，按站场流水作业。"),
    ("四、施工方案", "先室外设备安装、后室内联锁调试；道岔安装采用预装配整体推移，减少天窗占用时间。"),
    ("五、进度计划", "总工期 45 天，其中室外 25 天、室内调试 15 天、联调验收 5 天。"),
    ("六、资源配置", "主要机具：道岔安装车 2 台、电缆敷设机 1 台、仪器仪表 6 套；材料按周计划进场。"),
    ("七、安全保证措施", "严格执行天窗点登记与销记，驻站联络员与现场防护员双岗确认，作业前进行安全预想。"),
    ("八、质量保证措施", "关键工序设置质量控制点，隐蔽工程留存影像资料，实行三检制与工序交接检。"),
    ("九、应急预案", "针对设备故障、人身伤害、恶劣天气分别制定处置流程，配备应急物资并每季度演练一次。"),
    ("十、文明施工", "施工场地围挡封闭，材料分类堆放，作业结束做到工完料尽场地清。"),
]

FOOTER = "第 PAGE 页"


def add_page_number(section) -> None:
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("第 ")
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    inner = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    inner.append(text)
    fld.append(inner)
    p._p.append(fld)
    p.add_run(" 页")


def build(path: Path, *, drop_section: str | None = None, vague: bool = False, numbering: bool = True) -> None:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "仿宋"
    style.font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")

    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(3.17)
    sec.right_margin = Cm(3.17)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    if numbering:
        # 页码域：初稿故意不加
        pass
    if numbering:
        add_page_number(sec)

    title = doc.add_paragraph("京广线某区间信号设备大修施工方案")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True
    title.runs[0].font.size = Pt(16)

    for head, body in SECTIONS:
        if drop_section and drop_section in head:
            continue
        doc.add_heading(head, level=1)
        p = doc.add_paragraph(body)
        p.paragraph_format.first_line_indent = Pt(24)
    if vague:
        doc.add_paragraph("具体进场时间待定，由项目部另行通知。")
    doc.save(str(path))


def main() -> int:
    build(OUT / "方案-初稿.docx", drop_section="应急预案", vague=True, numbering=False)
    build(OUT / "方案-定稿.docx")
    print("[ok] 方案-初稿.docx / 方案-定稿.docx")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
