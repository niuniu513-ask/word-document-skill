# 案例：在一份真实课程报告上跑通"分析 → 填充 → 校验"

样本：`第四小组-基于C51的独居老人智能监护系统-方案设计报告.docx`
（12 个 zip 部件，114,828 字节，48 个正文段落，9 张表，3,560 字）

## 1. 分析

```bash
moon run cmd/docx --target js -- analyze report.docx --out spec.json
moon run cmd/docx --target js -- issues report.docx
```

得到的关键事实：

| 指标 | 值 |
| --- | --- |
| 部件数 | 12 |
| 正文段落 / 含表格内段落 | 48 / 142 |
| 表格 | 9 |
| 页面 | 21.00 × 29.70 cm 纵向，上下边距 2.54 cm，左右 3.175 cm |
| 正文样式 | `1`（Normal），默认中文字体随 run 继承 |
| 文本 run | 141 个，字号以继承为主（111 个），另见 10.5pt ×15、11pt ×6、9pt ×4、12pt ×3、14pt ×1 |
| 占位符 | **19 处，全部位于封面表格单元格内** |
| 规则命中 | `manual_heading_numbering` ×6、`unfilled_placeholders` ×19 |

第一条有价值的发现：用户口头说的"正文里的占位符"，实际全在表格里。
只扫 `word/document.xml` 的正文段落（python-docx 的 `Document.paragraphs` 就是这样）会一个都看不到 ——
这也是本项目把表格单元格纳入扫描范围的原因。

第二条：六个一级标题把编号写进了文字（"一、方案描述"），样式上却已经是 `heading 1`。
这类文档一旦生成目录，会出现"1. 一、方案描述"的重复编号，是真实存在的排版事故。

## 2. 填充

```bash
moon run cmd/docx --target js -- fill report.docx \
  --set "【待补充】=已通过验收" \
  --set "【待补充:具体分工】=张三" \
  --out filled.docx
```

输出 `replaced 18 occurrence(s)`。注意占位符在 XML 里常常被拆成多个 run
（本样本中 `【待补充:验收情况】` 就跨了 run），逐 run 替换会漏掉它们，
所以编辑器做的是"合并 run 文本 → 字符级定位 → 映射回 run → 从右向左改写"。

## 3. 校验

```bash
moon run cmd/docx --target js -- validate report.docx filled.docx
# summary formatting delta 0: no drift detected
```

校验覆盖：页面几何（纸张/方向/五类页边距/栏数/页码格式）、文档默认字体与字号、
每个段落的 16 项属性（样式、大纲级别、对齐、缩进、段前段后、行距与行距规则、
编号、keepNext、段前分页）、表格行列数与样式。

## 4. 交叉验证（用另一套技术栈确认没自说自话）

```python
from docx import Document
d = Document("filled.docx")
print(len(d.paragraphs), len(d.tables))                     # 48 9
print(sum("已通过验收" in c.text for t in d.tables for r in t.rows for c in r.cells))  # 13
```

段落数、表格数不变，13 个单元格出现新文本，3 处 `【待补充` 仍未填（因为没提供对应键）——
与引擎自报的 18 次替换一致。

## 5. 这个案例说明了什么

1. **"格式零漂移"可以是可执行的断言**，不是口号：`validate` 进 CI，改坏格式会让构建失败。
2. **真实文档的脏活在于边界**：占位符在表格里、文本跨 run、标题编号既在文字里又在样式上。
3. **跨语言交叉验证是低成本的正确性证据**：用 python-docx 读回来，避免"自己写、自己测、自己判"。
