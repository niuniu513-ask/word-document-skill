# 文档编辑(Document Editing)

修改、润色、扩写、缩写、批量替换、插入图片/表格时读这里。核心手法:**定位 → 最小侵入修改**,不做全文重写。

## 1. 定位

用 analyze spec 的 `paragraphs_preview` / `outline` 找目标段落的 `index`(body 顺序号),或用 extract 的 markdown 通读后定位。定位后一切修改用 python-docx 按 index 取段:

```python
from docx import Document
doc = Document("文档.docx")
paras = doc.paragraphs          # 注意:不含表格内段落
p = paras[i]
```

需要跨表格顺序定位时,用 `_docx_utils.body_blocks`。改完必须 `validate --original 修改前 --final 修改后`,确认 delta 只含你要改的东西。

## 2. 改文字不改格式(run 级替换)

段落的格式一半在段落属性,一半在 run 属性。替换文字的正确姿势:

- **整段重写**且段落格式在样式层:保留段落,只重建 runs——把新文本写进第一个 run(继承其字符格式),删除其余 runs:
  ```python
  runs = p.runs
  runs[0].text = 新文本
  for r in runs[1:]:
      r._r.getparent().remove(r._r)
  ```
  注意:若原段落里部分文字特意加粗(如"注意:"前缀),整段重写会丢掉它——这种情况改用下面的局部替换。
- **局部替换**(推荐):在 runs 里找包含目标片段的 run,直接改 `run.text`;片段跨 run 时(常见于中英混排、拼写检查切分),把涉及的连续 runs 合并处理:计算跨 run 的偏移,把替换后的整段重新分配到首尾 run,或先 `select_text` 式合并再替换。
- **批量查找替换**:遍历所有段落(含表格 cell.paragraphs、页眉页脚 paragraphs),对每个段落先比较全文是否含目标,再做局部替换;替换词长度不同不影响此法。禁止"导出全文→字符串替换→重建文档"——那会摧毁格式。

## 3. 润色/扩写/缩写

- 逐段处理:extract_content.md 通读 → 你改写文本 → 按 §2 手法回写,段落样式、编号、缩进一律不碰。
- 润色只动表达不动事实;扩写不得新增"事实"(可以新增解释、推理,不能新增数据/经历)。
- 改动较多时,交付说明里给"修改段落清单",方便用户核对。

## 4. 插入内容

- 插入段落到指定标题后:优先用 generate_docx.py 的 insert 模式(content.json 全部带 after 锚点)。手写 XML 场景(python-docx 无 API 时):`anchor._p.addnext(new_p)`。
- 插入图片:新段落居中 + `run.add_picture(路径, width=Cm(x))`(单边指定保比例)+ 题注段落(用模板 Caption/题注样式;没有则居中小号字)。
- 插入表格:先看模板有没有表格样式(analyze 的 tables[].style),有就 `tbl.style = 样式名`;列宽按版心(页宽-左右边距)分配。
- 编号:文档用原生编号时,新列表段挂同一 numId;手工编号时手动续号。

## 5. 删除内容

- 删段落:`p._p.getparent().remove(p._p)`;删除后检查前后空段与分页是否变得突兀(渲染确认)。
- 删表格:`tbl._tbl.getparent().remove(tbl._tbl)`,注意题注段落要一起处理。
- 删多段造成的章节空洞,在交付说明里指出,不要擅自"补写填充"。

## 6. 排版修复模式

用户说"帮我把排版修好":先 `analyze`(看 issues)+ `validate --check`(拿问题清单)→ 按 references/repair-rules.md 逐项修 → 渲染对比修复前后 → 交付"修复了什么"清单。修复只处理清单上的问题,不顺手重排全文档。
