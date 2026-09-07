---
name: word-document-skill
description: Word 智能文档处理技能:先完整分析任何 .docx 的结构、样式与格式并建立 Document Specification,再严格在原文档格式体系内完成生成、修改、填充、润色、排版修复、格式迁移与批量生成,最后经过格式验证、PDF 渲染逐页视觉检查与自动修复闭环再交付。凡任务涉及 Word 文档、docx、模板、实验报告、实习报告、课程设计、实践报告、课程作业、通知、总结、会议纪要、申请书、填表、排版、格式修复、按模板生成、内容迁移等场景,必须使用本 Skill——只要需要读取、创建、修改、填充或排版 Word 文件,即使用户没有明确提到"Word",也应触发本 Skill。
compatibility: 需要 Python 3.10+;python-docx、lxml、Pillow 必装;渲染检查依赖 pywin32 + Microsoft Word(Windows)或 LibreOffice,以及 PyMuPDF
---

# Word 智能文档处理(Word Document Intelligence)

核心理念:**先理解 Word,再操作 Word。**

不是"套模板",而是"理解文档"。任何 .docx 都是一套精密的格式体系:页面、样式、编号、页眉页脚、域、表格规则。用户把它交给你,是要求你**在它的体系内**工作,而不是重新设计一个你认为好看的文档。本 Skill 的全部机制都围绕这一点。

## 0. 铁律(先读这个)

**优先级链**(冲突时,排在前面的赢):

1. 用户明确要求
2. 用户提供的 Word 文档的结构与格式(既有 Style、编号体系、页面设置)
3. 用户提供的其他资料(数据、素材)
4. 用户指定的格式要求
5. 通用 Word 排版规范
6. 你自己的推断

**七条禁令:**

- 禁止在未分析用户 Word 文档的情况下直接生成或重新设计文档。拿到 .docx,第一步永远是分析(见 Phase 1)。
- 禁止因为"我觉得这样更好看"而改动用户原始格式。默认 **Formatting Delta ≈ 0**:只改变用户要求改变的东西。
- 禁止虚构数据与经历。实验数据、测量结果、实习经历、企业信息、日期、人名——缺了就写 `【待补充:具体说明】`;用户明确要求编造时,必须标注 `【模拟数据】`。
- 禁止用空格、空回车、手打页码、手打编号模拟排版。用 Style、段落设置、分页符、PAGE 域、原生编号。
- 禁止把整篇文档设成一种字体。中文字体(eastAsia)、西文字体(ascii)、数字字体是三个轨道,必须分开处理。
- 禁止破坏标题体系。原文档用"一、/(一)/1." 还是原生编号,就沿用哪套,不得强行改写编号体系。
- 禁止交付前不做验证。生成/修改后必须:格式验证 → 渲染成 PDF 逐页看 → 发现问题修复 → 再渲染,直到通过(最多 3 轮)。

**内容与格式分离:** 文档 = Content(文字、数据、表格内容、图片)+ Formatting(页面、样式、字体、行距、页眉页脚、分页)。改内容不得动格式;改格式不得动内容。所有脚本都按此设计。

## 1. 六阶段工作流(总览)

```
Phase 0  任务识别    用户要什么?→ 选定任务类型与工作流
Phase 1  文档分析    analyze_docx.py → Document Specification(必有 .docx 输入时)
Phase 2  方案制定    内容规划 / 内容映射;产出 content.json(生成类任务)
Phase 3  执行        按 Phase 0 选定的工作流操作(workflows/*.md)
Phase 4  验证闭环    validate_docx.py + render_docx.py 逐页视觉检查 + repair_docx.py
Phase 5  交付        最终 docx + 格式差异报告 + 待补充清单
```

用 TodoWrite 追踪这六个阶段;Phase 4 是硬性关口,不可跳过。

## 2. Phase 0 —— 任务识别

先判断任务类型,再选工作流。类型有歧义时,选择对用户原文档侵入最小的一种,并在最终汇报里说明你的假设。

| 任务类型 | 判断特征 | 工作流 |
|---|---|---|
| 模板填充 | 给了模板 + 资料,要"按模板写" | workflows/template-fill.md |
| 新建文档 | 没给模板,从零生成 | workflows/generate.md |
| 修改/润色/重写 | 给了现有文档,改内容 | workflows/edit.md |
| 格式迁移 | "按这个 Word 的格式做一份新的" | workflows/format-migration.md |
| 内容迁移 | "把这个文档的内容放进那个模板" | workflows/content-migration.md |
| 排版修复 | "帮我把排版修好" | workflows/edit.md(修复模式)+ references/repair-rules.md |
| 批量生成 | 一个模板 + 多份数据 | workflows/batch-generate.md |
| 仅分析 | "这个文档是什么结构/格式" | workflows/analyze.md |

报告类文档(实验/实习/课程设计/实践报告)只是模板填充与新建文档的常见场景,默认结构见 references/report-writing.md;**用户模板有明确结构时,以模板为准**。

## 3. Phase 1 —— 文档分析(机制核心)

拿到任何 .docx,先运行:

```bash
python <skill>/scripts/analyze_docx.py 输入.docx -o spec.json
python <skill>/scripts/inspect_styles.py 输入.docx --text styles.txt   # 需要细看样式时
python <skill>/scripts/extract_content.py 输入.docx -o content.json --markdown content.md
```

analyze_docx.py 产出 **Document Specification**(schemas/document-spec.json):文件基本信息、节数与页面设置(尺寸/方向/页边距/页眉页脚距)、全部 Style 的有效字体(中/西文分开)与段落格式、实际文字运行的字体/字号/字色频次(`formatting_inventory`)、标题大纲(层级 + 原生编号 + 手打编号)、表格(行列/列宽/合并/表头重复)、图片(尺寸/比例/失真/题注)、页眉页脚内容与页码域、封面检测与占位符、目录检测、批注/修订/文本框,以及 `issues` 列表(连续空段落、图片变形、疑似假标题、手打页码等)。

读 spec 后,向用户用 3-5 行概括:"这是一份 A4 纵向、单节的实验报告模板,封面含 7 个待填字段,正文用 Normal(宋体/Times New Roman 小四)+ Heading 1-2,表格用 TableGrid,页脚居中 PAGE 域。识别到正文主要字体为宋体/Times New Roman、主要字色为自动色；发现:封面字段未填、无目录。"

没有 .docx 输入(纯新建)时跳过分析,但必须读 references/word-style-rules.md 再动手。

## 4. Phase 2 —— 方案制定

生成/填充类任务,把内容写成 **content.json**(schemas/content-model.json),让脚本负责格式、你负责内容:

```json
{"blocks": [
  {"type": "cover", "fields": {"姓名": "张三", "学号": "2024010101"}},
  {"type": "paragraph", "after": "一、实验目的", "text": "……"},
  {"type": "table_fill", "table_index": 0, "cells": {"1,0": "米尺", "1,1": "0-100cm"}},
  {"type": "table", "rows": [["摆长","周期"]], "caption": "表1 原始数据"},
  {"type": "image", "after": "五、数据记录与处理", "path": "fig.png", "width_cm": 10, "caption": "图1 曲线"},
  {"type": "heading", "level": 2, "text": "六、结果分析", "numbering": "native"}
]}
```

要点:`after` 锚点让块精确插入到匹配标题之后(模板里已有标题时,内容块全部挂锚点);表格数据放进 `rows`,模板已有空表用 `table_fill` 按原始表格序号填单元格;图片 `path` 相对于运行目录;缺失内容写 `【待补充:…】`。内容中的事实必须全部来自用户资料,禁止编造。

## 5. Phase 3 —— 执行

按 Phase 0 的路由读取对应 workflows/*.md 后再动手。常用命令:

```bash
# 模板填充/插入(insert:锚点插入,保留模板一切;rebuild:保留封面重建正文;fill:只填占位符)
python <skill>/scripts/generate_docx.py --template 模板.docx --content content.json -o 输出.docx
# 按分析规格新建文档
python <skill>/scripts/generate_docx.py --spec spec.json --content content.json -o 输出.docx
```

深规则在 references/:样式与字体规则(word-style-rules.md)、模板与占位符细节(template-rules.md)、编辑保格式手法(document-editing.md)、报告结构默认值(report-writing.md)。

## 6. Phase 4 —— 验证闭环(不可跳过)

```bash
# 1) 格式差异:原文档 vs 产出,格式差异应为 0(或仅含用户要求的变化)
python <skill>/scripts/validate_docx.py --original 原文档.docx --final 输出.docx -o delta.json
# 2) 结构体检
python <skill>/scripts/validate_docx.py --check 输出.docx
# 3) 渲染:Word COM(最忠实)→ LibreOffice 兜底;--update-fields 更新目录/页码域
python <skill>/scripts/render_docx.py 输出.docx -o 渲染目录 --dpi 110 --update-fields
```

然后用 Read 逐页查看渲染出的 page-XXX.png,按下表检查(完整版见 references/visual-qa.md):

空白页/大面积空白 · 标题孤立页尾 · 图片变形 · 表格溢出版心或断裂异常 · 文字溢出 · 页眉页脚丢失 · 页码异常 · 边距异常 · 字体行距异常 · 目录页码错误。

发现问题 → references/repair-rules.md 选修法:

```bash
python <skill>/scripts/repair_docx.py 输出.docx --fix empty-paragraphs,heading-keepnext,image-aspect --apply -o 修复.docx
```

修复 → 重新渲染 → 重新检查,最多 3 轮;仍不过关就把问题如实报告给用户。**改过内容后必须重跑 validate 的 delta 检查**,确认没有顺手改坏别的东西。

## 7. Phase 5 —— 交付

最终答复必须包含:

1. 产出文件路径(docx;渲染 PDF/PNG 目录)
2. 格式差异结论:"与原文档相比,格式变化:N 处(逐条列出或确认为 0)"
3. 内容来源说明:哪些来自用户资料;哪些是 `【待补充】`;哪些是用户要求生成的 `【模拟数据】`
4. 视觉检查结论:渲染页数、发现并修复了什么、是否还有遗留问题

## 8. 脚本速查

| 脚本 | 用途 | 关键参数 |
|---|---|---|
| analyze_docx.py | 生成 Document Specification | `-o spec.json` |
| extract_content.py | 提取内容模型/Markdown 视图 | `-o content.json --markdown out.md` |
| inspect_styles.py | 样式全量详查 | `--style "Head*" --text styles.txt` |
| generate_docx.py | 模板/spec 内生成 | `--template/--spec --content -o --mode` |
| render_docx.py | docx→PDF→逐页 PNG | `-o 目录 --dpi 110 --update-fields --pages 1-5` |
| validate_docx.py | 格式差异 + 结构体检 | `--original --final` 或 `--check` |
| repair_docx.py | 自动修复(默认 dry-run) | `--fix a,b --apply -o` |

依赖:python-docx、lxml、Pillow 必需;PyMuPDF + (pywin32 或 LibreOffice)用于渲染。缺什么先 `pip install` 再继续;渲染后端不可用时,如实告知用户视觉检查受限,不得谎称已检查。

## 9. 高频错误(前人踩过的坑)

- **改小地方却重写全文档**——用户让改一处日期,你重新生成了整份。改内容用 edit 工作流的"定位-替换"手法,别动其他任何段落。
- **锚点失配**——content.json 的 `after` 写了"实验目的",文档里是"一、实验目的"。脚本会做编号宽容匹配,但锚点最好直接复制 analyze 大纲里的标题原文。
- **table_fill 用在 rebuild 之后**——rebuild 会删掉原正文表格,table_fill 必须针对保留的表格;拿不准就用 insert 模式。
- **图片拉伸**——用 `width_cm` 单边设定让高度自适应;发现失真跑 `repair_docx.py --fix image-aspect`。
- **忘了 --update-fields**——插入了正文却没更新目录页码,视觉检查时目录页码全是旧的。
- **整篇换字体**——只允许在 style 层面改;eastAsia/ascii 分开设,禁止对 run 全量覆盖。
- **渲染完不看图**——render 成功 ≠ 排版正确。必须逐页 Read PNG。
