# 工作流:新建文档(Generate)

触发:用户要从零生成一份 Word,没有提供模板。提供了模板 → 走 template-fill.md。

## 步骤

1. **定规格**:用户给了格式要求(字体/页边距/结构)就照做;没给时按文档类型套用规范——中文报告类默认:A4 纵向,边距上下 2.54 / 左右 3.17cm,正文宋体+Times New Roman 小四 1.5 倍行距首行缩进 2 字符,标题黑体(一级三号、二级四号、三级小四)加粗,页脚居中页码。把选定规格在答复开头告诉用户(一句即可)。
2. **建 Document Specification**:写一个 spec.json(与 analyze_docx.py 输出同构,至少含 sections[0] 与 conventions),内容如下——generate_docx.py --spec 直接吃这个格式:
   ```json
   {"sections": [{"page": {"width_cm": 21.0, "height_cm": 29.7, "orientation": "portrait"},
                  "margins_cm": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17}}],
    "conventions": {
      "body_style": {"style": "Normal",
                     "font": {"ascii": "Times New Roman", "eastAsia": "宋体", "size_pt": 12},
                     "paragraph": {"line_spacing": 1.5, "first_line_indent_cm": 0.85}},
      "heading_styles": [{"level": 1, "style": "Heading 1"},
                         {"level": 2, "style": "Heading 2"}]}}
   ```
   spec 模式下脚本会把 Normal/Heading 样式按此字体与段落规格设置好(eastAsia 会正确写入)。
3. **写内容**:content.json 按 schemas/content-model.json;报告类结构参考 references/report-writing.md;事实红线照旧。
4. **生成**:`generate_docx.py --spec spec.json --content content.json -o 输出.docx`
5. **页眉页脚与页码**:脚本不建页眉页脚,需要时用 python-docx 补:`section.footer.paragraphs[0]` 居中 + PAGE 域(参考 fixtures 的做法);页眉横线用 Header 样式下边框。
6. **验证闭环**:Phase 4 全流程(渲染逐页看:新建文档重点看标题层级是否呈现、编号是否连续、页码是否从正文开始)。

## 注意

- 新建 ≠ 随意:用户脑子里往往有隐含规范(学校/单位惯例),spec 里没覆盖的细节在交付说明里声明"按 XX 规范,如需调整告诉我"。
- 目录:文档超过 5 页且分章时,主动问用户是否要目录;要就加 TOC 域并 --update-fields。
