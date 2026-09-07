# Word 样式与字体规则(Word Style Rules)

操作样式层时的硬规则与常识。新建、重建、格式迁移、修字体都必须先读这里。

## 1. 三轨道字体

一个 run 的字体由三部分组成,**永远分开设置**:

| 轨道 | XML 属性 | 典型值(中文文档) |
|---|---|---|
| 中文 | `w:rFonts/@w:eastAsia` | 宋体、黑体、楷体、仿宋 |
| 西文 | `w:rFonts/@w:ascii`(及 `w:hAnsi`) | Times New Roman |
| 复杂文字 | `w:rFonts/@w:cs` | 一般不动 |

python-docx 的 `run.font.name = "Times New Roman"` 只设了 ascii/hAnsi;eastAsia 必须补 XML:

```python
from docx.oxml.ns import qn
rpr = run._r.get_or_add_rPr()
rfonts = rpr.get_or_add_rFonts()
rfonts.set(qn("w:eastAsia"), "宋体")
```

诊断字体问题时用 analyze spec 的 `conventions.body_style.font` 与 `doc_defaults.font`:如果 eastAsia 或 ascii 为 None,说明沿用 theme/默认值,不要臆断成具体字体。

## 2. 字号对照(中文出版习惯)

初号42 / 小初36 / 一号26 / 小一24 / 二号22 / 小二18 / 三号16 / 小三15 / 四号14 / **小四12** / 五号10.5 / **小五9** (单位 pt)。常见搭配:正文宋体小四或五号;一级标题黑体三号或小三;封面标题黑体/宋体加粗一号到二号。

## 3. 优先复用 Style,其次才新建

- 修改/填充场景:新段落一律挂文档已有样式(通常 `conventions.body_style.style`,如 Normal;标题用 outline 对应的 Heading/标题 N)。**禁止对段落做大量直接格式覆盖**——那样文档后续没法统一调整。
- 确需新样式(格式迁移、spec 模式):基于最接近的现有样式创建(basedOn),只覆写差异属性。
- 直接格式(跑在 run/paragraph 上的覆盖)只用于一次性的局部修正。

## 4. 段落格式要点

- **首行缩进**:中文正文习惯 2 字符。用 `w:ind/@w:firstLineChars="200"`(字符单位,随字号缩放)优于固定 cm;python-docx 的 `paragraph_format.first_line_indent = Cm(0.74)` 是厘米单位,小四(12pt)2 字符 ≈ 0.85cm。
- **行距**:`w:spacing/@w:line` + `@w:lineRule`。`auto` 时 240=单倍、360=1.5倍、480=2倍;`exact`/`atLeast` 时单位是 twips(20 分之一磅)。1.5 倍行距正文最常见;固定值行距会裁剪大号公式与图片,慎用。
- **段前段后**:标题通常段前 12-24pt、段后 6-12pt;正文段前段后 0,靠行距。
- **分页规则**:标题样式应带 keepNext + keepLines(防孤立);`w:pageBreakBefore` 用于强制章前分页。**禁止用连续空段/空格制造分页**。

## 5. 标题与编号体系

识别到的编号形态(spec 的 `outline[].native_numbering / manual_number`):

- 原生编号(numPr 绑定 numbering.xml):插入新标题时**只写标题文字**,编号自动出现;content.json 里给 `numbering: "native"` 让脚本剥离手写编号。
- 手工编号(文字里带"一、"/"1.1"):新增标题照抄同一编号风格,手动续号;先数清现有序号。
- 混用是文档的既有事实,沿用文档当前做法,不要"统一修正"——除非用户明确要求整理编号。

## 6. 页面、页眉页脚与域

- 页边距、纸张大小属于 section;改一个 section 不影响其他。中文文档常见 A4(21×29.7cm),上下 2.54、左右 3.17cm(Word 默认)或左右 2.6-3.0cm。
- 页码必须是域:`PAGE` / `NUMPAGES`,格式(第 X 页 共 Y 页、- X -)用域外的文字包裹。第 1 节(封面/目录)常隐藏页码:分节后 `w:pgNumType/@w:start` 重新从 1 起。
- 页眉横线是 Header 样式的下边框;删内容别删样式。
- 目录(TOC 域)在内容变动后必须更新:render_docx.py --update-fields(经 Word COM),失败则手动提示用户在 Word 里 F9。

## 7. 中西文混排细节

- 中文与数字/英文之间,排版规范建议留约 1/4 空隙(Word"自动调整中文与数字的间距"默认开启,无需手动加空格;**不要**在正文里手写空格模拟)。
- 数字用什么字体由 ascii 轨道决定;全文改成"中文用宋体、数字用 Times New Roman"是常见的真实需求——在样式层设置,不要逐 run 改。
