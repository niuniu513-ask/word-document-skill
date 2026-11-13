---
name: word-document-skill
description: Word 智能文档处理技能:先完整分析任何 .docx 的结构、样式与格式并建立 Document Specification,再严格在原文档格式体系内完成生成、修改、填充、润色、排版修复、格式迁移与批量生成,最后经过格式验证、PDF 渲染逐页视觉检查与自动修复闭环再交付。另提供一键合规体检(review.py + rules/ 规则包):给文档打分、列出缺章节/占位符/编号/页码等阻断项并给修复建议。凡任务涉及 Word 文档、docx、模板、实验报告、实习报告、课程设计、实践报告、课程作业、通知、总结、会议纪要、申请书、填表、排版、格式修复、按模板生成、内容迁移,或问"这份方案能不能报审/能不能交/格式合不合规"等文档体检与合规初审场景,必须使用本 Skill——只要需要读取、创建、修改、填充或排版 Word 文件,即使用户没有明确提到"Word",也应触发本 Skill。报告要求"真实代码截图/运行结果截图/VS Code 截图"时,同时按 workflows/lab-report-runbook.md 采集并核对证据。
metadata:
  compatibility: 需要 Python 3.10+;python-docx、lxml、Pillow 必装;渲染检查依赖 pywin32 + Microsoft Word(Windows)或 LibreOffice,以及 PyMuPDF
---

# Word 智能文档处理(Word Document Intelligence)

> **一稿过 / OneDraft** —— 本仓库参加石家庄铁道大学「天翼智创杯」TeleAgent 智能体创新竞赛的对外名称，
> 指向同一个能力：把「怎么改 Word 才算对」写成智能体能执行的判断。打包分发时用 `STDU2026-一稿过.skill`。

> 本技能是一本**可执行的 Word 指南书**：决策地图在 `AGENT-GUIDE.md`，机器可读清单在 `capabilities.json`，
> 唯一入口是 `python scripts/docx_agent.py <task>`（`list` 看全部 14 个任务，`--dry-run` 先看命令再执行）。
> 不确定该用哪个脚本时，先跑 `docx_agent.py list`，不要凭记忆猜脚本名。

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

**九条禁令:**

- 禁止在未分析用户 Word 文档的情况下直接生成或重新设计文档。拿到 .docx,第一步永远是分析(见 Phase 1)。
- 禁止因为"我觉得这样更好看"而改动用户原始格式。默认 **Formatting Delta ≈ 0**:只改变用户要求改变的东西。
- 禁止虚构数据与经历。实验数据、测量结果、实习经历、企业信息、日期、人名——缺了就写 `【待补充:具体说明】`;用户明确要求编造时,必须标注 `【模拟数据】`。
- 禁止用空格、空回车、手打页码、手打编号模拟排版。用 Style、段落设置、分页符、PAGE 域、原生编号。
- 禁止把整篇文档设成一种字体。中文字体(eastAsia)、西文字体(ascii)、数字字体是三个轨道,必须分开处理。
- 禁止破坏标题体系。原文档用"一、/(一)/1." 还是原生编号,就沿用哪套,不得强行改写编号体系。
- 禁止擅自改写模板固定文本。封面标题、课程名、校名、固定说明、页眉页脚和既有章节标题默认受保护;只填明确占位符/字段/空白单元格或用户明确要求替换的内容。模板与源资料标题冲突时以模板为准,并在交付说明中提示。
- 禁止伪造截图证据。用户要"真实的代码截图/运行结果截图"时,图必须来自真实运行的窗口;不许用代码绘制模拟界面、不许把两次运行的输出拼在一起(同一窗口连续滚动拼接除外)、不许给没有输出的单元格插空白图。实测数字必须能追溯到对应截图或执行输出;理论值和推导值应明确区分。保留原屏,报告内允许分别裁出成对的代码和输出。
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

用可用的任务记录工具追踪这六个阶段;Phase 4 是硬性关口,不可跳过。

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
| 带截图的报告 | "要**真实代码截图**/运行结果截图/实验报告" | workflows/lab-report-runbook.md(端到端) |
| 只要 VS Code 界面截图 | 明确说"VS Code 窗口/IDE 界面" | runbook §3.2 + workflows/lab-report-screenshots.md(后备) |
| 仅分析 | "这个文档是什么结构/格式" | workflows/analyze.md |

报告类文档(实验/实习/课程设计/实践报告)只是模板填充与新建文档的常见场景,默认结构见 references/report-writing.md;**用户模板有明确结构时,以模板为准**。
截图证据的采集、裁切与比例换算是独立的一层,见 references/screenshot-evidence.md。
用户同时要求“严格原模板 / notebook / 真实运行截图”时，逐关执行 [可执行实验报告工作流](workflows/lab-report-runbook.md)：截图默认走 §3.1 **离屏渲染**（`scripts/render_nb_html.py` + `scripts/shoot_nb_cells.cjs`），只有明确要求 IDE 界面才走 §3.2 的 `scripts/capture_vscode_cells.cjs`；配置字段见 [实验报告配置](references/lab-report/lab-report-config.md)，构建与审计用 `scripts/lab_report.py`。原模板优先于历史成稿；保留原大框容器和固定章节，不额外增加目的、环境、总结。自动校验必须覆盖框内图片并传入预期数量；最后逐页检查标签与图片跨页情况。
实验报告这一类任务的做法集中在 `references/lab-report/`（含“通用规则 / 本类方法 / 个案记录”的分档说明），不要把其中的个案数值当成通用常量。
模板呈现为正文外大框、单格容器或用户纠正“不是表格”时,先读 [大框与分页规则](references/template-frames.md),分别记录视觉布局和底层 XML 结构,不要按题拆成多行表格。

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

深规则在 references/:样式与字体规则(word-style-rules.md)、模板与占位符细节(template-rules.md)、编辑保格式手法(document-editing.md)、报告结构默认值(report-writing.md)、截图证据链(screenshot-evidence.md)。实验报告类任务另见 references/lab-report/(配置字段、离屏采集与分档说明)。

## 6. Phase 4 —— 验证闭环(不可跳过)

```bash
# 1) 格式差异:原文档 vs 产出,格式差异应为 0(或仅含用户要求的变化)
python <skill>/scripts/validate_docx.py --original 原文档.docx --final 输出.docx -o delta.json
# 2) 结构体检
python <skill>/scripts/validate_docx.py --check 输出.docx
# 3) 渲染:Word COM(最忠实)→ LibreOffice 兜底;--update-fields 更新目录/页码域
python <skill>/scripts/render_docx.py 输出.docx -o 渲染目录 --dpi 110 --update-fields
# 4) 证据对账(含截图的文档必跑):标签数 / 图片数 / 空图 / 宽黑条 / 变形
python <skill>/scripts/verify_evidence.py 输出.docx --label 代码截图 --label 实验结果截图 -o evidence.json
```

然后用 Read 逐页查看渲染出的 page-XXX.png,按下表检查(完整版见 references/visual-qa.md):

空白页/大面积空白 · 标题孤立页尾 · 图片变形 · 表格溢出版心或断裂异常 · 文字溢出 · 页眉页脚丢失 · 页码异常 · 边距异常 · 字体行距异常 · 目录页码错误。
含截图的文档另查:标签下紧跟着图(不是错位)· 图里字与正文字号相当 · 输出图不是宽黑条或纯黑块 · 实测数字能追溯到对应图或执行输出,推导值注明依据。

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
| review.py | 一键合规体检(打分+问题清单) | `--profile railway-plan\|lab-report\|generic` `--html --json --baseline` |
| repair_docx.py | 自动修复(默认 dry-run) | `--fix a,b --apply -o` |
| verify_evidence.py | 截图证据对账 | `--label 代码截图 --expect 代码截图=29 --strict` |
| crop_evidence.py | 明确坐标裁图并保存来源哈希 | `raw.png crop.png --box X Y W H --identity cell5/code` |
| capture_notebook.py | 滚动窗口采集 + 拼接(Windows) | `--list-windows` / `--title` / `--stitch` |

依赖:python-docx、lxml、Pillow 必需;PyMuPDF + (pywin32 或 LibreOffice)用于渲染。缺什么先 `pip install` 再继续;渲染后端不可用时,如实告知用户视觉检查受限,不得谎称已检查。

## 9. 高频错误(前人踩过的坑)

- **改小地方却重写全文档**——用户让改一处日期,你重新生成了整份。改内容用 edit 工作流的"定位-替换"手法,别动其他任何段落。
- **锚点失配**——content.json 的 `after` 写了"实验目的",文档里是"一、实验目的"。脚本会做编号宽容匹配,但锚点最好直接复制 analyze 大纲里的标题原文。
- **table_fill 用在 rebuild 之后**——rebuild 会删掉原正文表格,table_fill 必须针对保留的表格;拿不准就用 insert 模式。
- **图片拉伸**——用 `width_cm` 单边设定让高度自适应;发现失真跑 `repair_docx.py --fix image-aspect`。
- **忘了 --update-fields**——插入了正文却没更新目录页码,视觉检查时目录页码全是旧的。
- **整篇换字体**——只允许在 style 层面改;eastAsia/ascii 分开设,禁止对 run 全量覆盖。
- **渲染完不看图**——render 成功 ≠ 排版正确。必须逐页 Read PNG。
- **截图全部撑满版心**——短图字大如标题、长图字小如注释。按 `width_cm ≈ px × 0.0245` 换算,让每张图的字符大小一致。
- **从拼接长图切单元格**——接缝会把一行代码拦腰切断。在长图上定位,回**单帧原图**裁像素。
- **给没输出的单元格插空图**——纯 `import`/赋值没有输出,应删掉该题的输出标签,而不是插一张黑图。
- **分析里的数字凭记忆写**——`np.prod` 写成"累积乘积"、0.9981 其实是 0.7364。数字必须从输出复制。
- **采集途中执行 cell**——输出块增减,行高模型全错。执行与采集必须分开。
- **框宽和下划线只凭印象恢复**——实测原模板框线坐标,计入边框距离与缩进;保留封面填充值 runs 的下划线,逐字段渲染核对。
- **把大框识别成题目表格**——单个大容器不代表按题分行;先渲染原模板,再按用户认定的视觉结构处理。
- **缩放后沿用旧裁剪框**——重新采集同一帧的坐标和变换,核对首末代码行与输出,重新生成裁图并替换最终 DOCX 的图片。
- **只核对图片数量**——92 张图也可能放错题;逐题核对代码、输出、分析三者,数量检查不能证明语义正确。

## 10. 唯一入口：docx_agent.py

```
python scripts/docx_agent.py list              # 14 个任务：何时用 / 命令 / 产出 / 工作流
python scripts/docx_agent.py --capabilities    # 机器可读清单（capabilities.json）
python scripts/docx_agent.py <task> [参数...]  # 直接路由到对应脚本，退出码原样透传
python scripts/docx_agent.py <task> --dry-run  # 只打印将执行的命令，不产生副作用
python scripts/docx_agent.py --doctor          # 依赖 / 脚本 / 规则包 / 文档一致性自检
python scripts/docx_agent.py --guide           # AGENT-GUIDE.md 的章节目录
```

中文别名可用：`分析`→analyze、`体检/审查/合规`→review、`验证/格式差异`→validate、`修复`→repair、`渲染/出图`→render。
路由表与脚本永远不会漂移——`tests/test_docx_agent.py` 会检查清单里写的每个脚本、工作流、文档是否真的存在。
