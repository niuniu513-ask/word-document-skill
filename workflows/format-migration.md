# 工作流:格式迁移(Format Migration)

触发:"按照这个 Word 的格式,把下面内容做成新的 Word" / "照它的样子排一份"。输入 = 格式源文档 + 目标内容(文本/资料/另一文档)。

## 步骤

1. **格式源分析**:`analyze_docx.py 格式源.docx -o spec.json`,必要时 inspect_styles 细看。这份 spec 就是 Formatting Specification。
2. **读目标内容**:文本直接读;docx 用 extract_content.py。
3. **结构映射**:把目标内容组织成 content.json:
   - 标题层级对齐源文档的 heading_styles(level ↔ 样式)
   - 正文走 body_style(字体行距缩进自动继承)
   - 表格:源文档表格什么风格(边框/表头加粗/三线表),content 块就用什么(blk.style 写源表格样式名;三线表用无边框+手动上下框线时在交付说明注明)
   - 图片:宽度不超过源文档版心
4. **生成**:`generate_docx.py --spec spec.json --content content.json -o 输出.docx`
5. **补页眉页脚**:spec 模式不复制页眉页脚。源文档有页眉页脚/页码时,用 python-docx 按源 spec 的 headers_footers 描述复刻(文字 + PAGE 域 + 对齐);复杂艺术页眉(图片/文本框)如实告知用户需在 Word 中手动带入。
6. **验证闭环**:
   - `validate_docx.py --original 格式源.docx --final 输出.docx`:页面/样式/字体类 delta 应为 0(表格数、正文行数不同是内容差异,不算格式差异)
   - 渲染逐页对比源文档:标题视觉层级、正文字体行距、表格风格应"看不出是两个人排的"

## 底线(必须尽可能复刻)

页面尺寸/方向/边距、正文字体三轨道与行距缩进、各级标题字体字号、表格样式、页码格式。做不到的项(如主题配色、艺术字)在交付说明里逐条声明。
