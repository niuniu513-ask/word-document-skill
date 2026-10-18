# 工作流:模板填充(Template Fill)

触发:用户给了 模板.docx(+ 资料),要"按模板写"。最高频工作流。

## 步骤

1. **双分析**:
   ```bash
   analyze_docx.py 模板.docx -o spec-模板.json
   analyze_docx.py 资料.docx -o spec-资料.json      # 资料是 docx 时;纯文本/图片直接读
   extract_content.py 资料.docx -o content-资料.json --markdown 资料.md
   ```
   同时渲染原模板首页及边框跨页位置,记录“大框/表格”的视觉用途;发现单格大容器或用户要求大框时,按 `references/template-frames.md` 处理。`.doc` 先用 Word/LibreOffice 转为 `.docx` 副本,不能直接交给 python-docx。
2. **规划映射**:模板大纲(analyze 的 outline)是骨架;把资料内容分配到各标题下,产出 content.json:
   - 封面:cover.fields(analyze 报的占位符字段名)
   - 每节正文:{"type": "paragraph", "after": "<模板标题原文>", "text": "..."}
   - 模板空表:{"type": "table_fill", "table_index": <analyze 的表序号>, "cells": {...}}
   - 新增表/图:table / image 块,同样挂 after 锚点
   - 模板缺章节而资料需要:heading 块(numbering 按模板体系,参照 references/report-writing.md)
   - 模板已有封面标题、课程名、校名或固定说明时,将其作为受保护文本保留;不要用资料标题或普通 `heading` 块覆盖。
3. **生成**(insert 为主):
   ```bash
   generate_docx.py --template 模板.docx --content content.json -o 输出.docx
   ```
4. **核对 warnings**:missing_placeholders 逐个处理;锚点未匹配的块改锚点重跑(锚点直接从 analyze outline 复制)。
5. **验证闭环**:delta(相对模板,允许项=新增内容表格与填充文本)+ 渲染逐页检查,重点:封面填齐没溢出、每个模板标题下都有内容且位置正确、表格数据行数撑版、目录页码已更新。若使用 `rebuild`,逐字核对封面/前置块固定文本仍与模板一致;标题不一致属于警告,不是自动改写理由。

## 资料不足时

按 report-writing.md 的结构判断每节能否成文;缺数据的节写引导句 + `【待补充:××数据】`,交付说明列出全部待补充项。禁止为了"看起来完整"编造。

## 常见变体

- 模板只有封面(正文无标题):写 content.json 不带锚点,--mode rebuild 会保留封面重建正文;或先问用户正文结构。
- 用户给的是"上一届的成稿"当模板:把它当模板用 insert 填,或当内容源走 content-migration.md;先问清"格式学它,内容换成我的新资料"属于哪种。
