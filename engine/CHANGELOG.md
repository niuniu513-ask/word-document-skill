# Changelog

## 0.1.0 — 2026-11

首个可用版本：读、分析、编辑、写回 Word `.docx`，并校验格式漂移。

### 新增

- `xml`：命名空间感知的 XML 1.0 解析器（迭代式，显式栈）、序列化器（自动补齐
  新建元素的 `xmlns` 声明）、元素树查询/修改 API（按命名空间 URI 匹配）。
- `ooxml`：`WordPackage` zip 容器读写；单位换算（twips/半磅/EMU/fiftieth-cm）；
  `RunProps` / `ParaProps` 读取与层叠；`styles.xml` 样式链解析（`basedOn`，循环保护）；
  `Docx` / `Paragraph` / `Run` / `Table` 文档模型。
- `docx`：Document Specification 分析（document / doc_defaults / sections / styles /
  outline / paragraphs_preview / tables / images / placeholders / formatting_inventory /
  conventions / issues）；规则引擎（手工编号、未填占位符、连续空段落、标题跳级、
  缺失节属性、字体混杂、无正文）；跨 run 文本替换、整段改写、段落插入；
  格式漂移 diff（页面几何、文档默认值、逐段落 16 属性、表格形状）。
- `cmd/docx`：`info` / `text` / `outline` / `parts` / `styles` / `paragraphs` /
  `tables` / `analyze` / `issues` / `fill` / `validate`。
- `examples/demo`：端到端示例（分析 → 填充 → 校验 → 写回）。
- 34 个测试；`scripts/make_fixture.py` 生成确定性 fixture。

### 修复

- 新建元素序列化时丢失命名空间前缀（`<t>` 而非 `<w:t>`），导致填充后的文本被 Word 与
  python-docx 视为空——writer 现在按作用域补 `xmlns:` 声明，并有回归测试。
- `parse_fragment` 返回了合成的 `<fragment>` 包装节点。
- `read_decl` 从 `<?` 起读，导致声明解析偏移。
