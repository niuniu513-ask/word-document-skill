# 文档分析(Document Analysis)

Phase 1 的操作手册。目标:拿到 Document Specification,并真正读懂它。

## 为什么必须先分析

用户递来的 .docx 承载了他已有的排版决定:封面怎么做、标题用几号字、行距多少、表格长什么样。直接生成内容等于无视这些决定。分析一步把文档的格式体系显式化,后续每个操作都以此为依据——这就是"建立 Document Specification"的含义。

## 标准动作

```bash
python <skill>/scripts/analyze_docx.py <input.docx> -o spec.json
```

纯 Python 解析(zipfile + python-docx + lxml),不启动 Word,秒级完成,对损坏文件也会尽量给出部分结果。一个输入文件跑一次;多个输入文件(模板 + 资料)各跑一次,分别产出 spec。

需要更细的样式信息(比如确认"标题 2"的编号绑定):

```bash
python <skill>/scripts/inspect_styles.py <input.docx> --style "*标题*" --text styles.txt
```

需要内容全貌(判断结构、撰写映射)时:

```bash
python <skill>/scripts/extract_content.py <input.docx> -o content.json --markdown content.md
```

## 怎么读 spec.json(按用途抓重点)

- **准备填充模板** → `cover.placeholders_preview` + `cover.fields_detected`(有哪些空要填)、`outline`(标题结构 = 内容的骨架)、`tables`(哪些表格待填,记下每个表的 `body_index` 即 table_index)、`conventions.body_style`(新插入正文用什么样式)。
- **准备重建/迁移格式** → `sections`(页面与页边距)、`doc_defaults` + `conventions.body_style`(默认字体三轨道)、`conventions.heading_styles`(各级标题样式与编号)、`styles`(全部样式定义,注意 based_on 链)。
- **准备编辑现有文档** → `document.package_flags`(有没有修订/批注/文本框——有修订要先问用户要不要保留修订痕迹)、`outline` + `paragraphs_preview`(定位要改的段落)、`issues`(顺手要报告的问题)。
- **准备修排版** → 直接看 `issues`,再配合 `--check` 模式的 validate_docx.py。
- **识别字体与字色** → 看 `styles[].font` 获取样式继承后的字体、字号、粗斜体和显式颜色；看 `paragraphs_preview[].run_font` 获取段落首个实际文字运行；看 `formatting_inventory` 获取正文、表格、页眉页脚中实际使用格式的频次、主要组合和颜色清单。颜色可能是 RGB（如 `365F91`）、`auto/default` 或 `theme:accent1` 等主题色。

字体识别必须区分 `eastAsia`（中文）、`ascii/hAnsi`（西文与数字）和 `cs`（复杂文字）三个轨道。迁移或生成时，优先沿用 `formatting_inventory.dominant` 与对应样式，而不是把所有文字强行设成一种字体；如果 `mixed_fonts` 或 `mixed_colors` 为 `true`，说明模板存在多种真实格式，需按样式/段落角色分别保留。

## 汇报规范

分析完,用 3-5 行向用户概括(见 SKILL.md 示例)。必须点出:节/页面设置、正文样式与字体、标题体系(是否原生编号)、封面与占位符、目录有无、issues 列表。这个概括同时是你自己的"操作契约"——后面 Phase 4 验证时,格式差异要以它为基线。

## 脚本失败时的兜底

analyze_docx.py 崩溃或字段明显缺失时,按顺序降级:

1. 用 `--check` 跑 validate_docx.py 拿基础信息;
2. 直接解包看 XML(文档的真相在 XML 里):
   ```bash
   python - <<'PY'
   import zipfile
   z = zipfile.ZipFile("文档.docx")
   print(z.namelist())          # 部件清单:document.xml / styles.xml / numbering.xml / media/
   xml = z.read("word/document.xml").decode("utf-8")
   print(xml[:3000])
   PY
   ```
3. 关键 XML 位置:页面设置在 `w:body/w:sectPr`(w:pgSz、w:pgMar);样式在 `word/styles.xml` 的 `w:style`;编号在 `word/numbering.xml`;文本框内容在 `w:txbxContent`(python-docx 不能直接编辑,需要 XML 级操作)。
4. 全部失败:如实告知用户文件可能损坏,请其另存/修复后重试。禁止在无分析的情况下凭猜测动手。

## 多文档任务

- 模板 + 资料:两个都分析;内容从资料 spec/extract 来,格式以模板 spec 为准。
- 内容迁移:源文档用 extract_content.py 取内容,目标模板用 analyze_docx.py 取格式。
- 批量:模板分析一次,数据逐份映射(见 workflows/batch-generate.md)。
