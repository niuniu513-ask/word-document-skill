# 工作流:内容迁移(Content Migration)

触发:"把这个 Word 的内容放进这个模板"。输入 = 内容源文档 + 目标模板。与 template-fill 的区别:内容不是用户口头给的,而是装在另一个 docx 里。

## 步骤

1. **分离两层**:
   ```bash
   extract_content.py 内容源.docx -o source-content.json --markdown source.md
   analyze_docx.py 模板.docx -o spec-模板.json
   ```
   Source Content 来自源文档(只要内容,不要格式);Target Formatting 来自模板(只要格式,不要内容)。
2. **结构对齐**:通读 source.md,把源内容块映射到模板大纲:
   - 源标题 → 模板对应级别标题(模板没有的章节:追加 heading 块,编号沿用模板体系)
   - 源正文段 → paragraph 块(after 锚到模板标题)
   - 源表格 → table 块(数据原样;若模板同位置有空表,优先 table_fill)
   - 源图片 → 提取 media(解包 word/media/)后按原位插入,题注跟随
   - 源文档的封面/页眉页脚/目录:**丢弃**(模板自带;源里有而模板没有的信息,提醒用户)
3. **生成**:`generate_docx.py --template 模板.docx --content content.json -o 输出.docx`(insert/rebuild 视模板结构,同 template-rules.md §1)。
4. **验证闭环**:对模板做 delta(格式 delta=0);渲染逐页检查内容完整(源文档章节清单 vs 输出章节清单逐一对照,不许丢节);--update-fields 刷目录。

## 陷阱

- 源文档的"内容"里藏着格式语义(红色加粗=重点):迁移时把语义转成文字标注(如【重点】)或问用户,不要把源格式带进模板。
- 源表格有合并单元格:extract 的 rows 已按"左上角有值、被合并处为空"展开;重建时内容不丢即可,复杂合并版式不值得复刻(模板风格优先)。
- 图片提取:
  ```python
  import zipfile; zipfile.ZipFile("源.docx").extractall("解包")   # word/media/ 下即图片
  ```
  按 source-content.json 里 image 块的出现顺序对应 media 编号不可靠时,回源文档 XML 确认 rId→文件名。
