"""Builds tests/fixtures/mini.docx, a small document with known properties.

Deterministic on purpose: the integration test asserts on these numbers, so the
fixture must not drift. Regenerate with `python scripts/make_fixture.py`.
"""
import os
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "tests", "fixtures", "mini.docx")

d = docx.Document()
normal = d.styles["Normal"]
normal.font.name = "仿宋"
normal.font.size = Pt(12)

d.add_heading("一、方案描述", level=1)
p = d.add_paragraph("本系统以 C51 单片机为核心，实现独居老人状态监护。")
p.paragraph_format.first_line_indent = Pt(24)

d.add_heading("二、作品需要的材料", level=1)
d.add_paragraph("")            # deliberate empty paragraph
d.add_paragraph("")

t = d.add_table(rows=2, cols=3)
t.style = "Table Grid"
t.rows[0].cells[0].text = "器件"
t.rows[0].cells[1].text = "型号"
t.rows[0].cells[2].text = "数量"
t.rows[1].cells[0].text = "单片机"
t.rows[1].cells[1].text = "STC89C52"
t.rows[1].cells[2].text = "【待补充】"

d.add_heading("三、仿真过程及结果", level=1)
d.add_paragraph("仿真结果见附录，照片预留：____（待插入）")

d.save(OUT)
print("wrote", os.path.normpath(OUT))
