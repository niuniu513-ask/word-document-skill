# docx-mbt

**一个纯 MoonBit 实现的 Word（.docx / OOXML）智能引擎：读取、分析、编辑、写回，零原生依赖。**

用 MoonBit 写的 Word 文档处理库目前几乎没有，而"用程序改 Word 文档"恰恰是格式最容易崩掉的场景：
Word 会把一句话拆进任意多个 `w:r`（run），每个 run 可能带不同字体；命名空间前缀各家不同；
一旦序列化时丢了 `xmlns`，文档看起来正常、实际内容全部失效。
`docx-mbt` 把这些问题当作核心问题解决，并给出可验证的质量承诺：

> **Formatting Delta ≈ 0** —— 改完文字后，段落样式、缩进、行距、页面几何、表格结构一个字节都不许动，并且可以当场验证。

```
moonbit  0.1.20261104        moon 包管理 + 编译
目标      --target js / wasm-gc （native 需要 MSVC/clang）
依赖      仅 moonbitlang/core（zip / json / utf8）与 moonbitlang/x（fs）
测试      34 个单元 + 集成测试，全部通过
```

---

## 1. 它解决什么问题

Agent 或脚本要往 Word 报告里填数据时，常见做法是"重新生成一份文档"——
样式、编号、页眉页脚、表格边框全部重来，交上去的东西一眼就不是原来那份。
本项目反过来：**在文档自己的格式体系内编辑**。

| 场景 | 常见做法 | docx-mbt |
| --- | --- | --- |
| 填占位符 | 重建段落 | 原地替换，保留字符格式；跨 run 也能命中 |
| 分析结构 | 调用 python-docx | 纯 MoonBit 解析 OOXML，输出 JSON 规范 |
| 判断"有没有改坏" | 人工目检 | `validate` 输出逐属性 diff，delta 应为 0 |
| 运行环境 | 需要 Python/Office | 编译成 js 或 wasm-gc 直接跑 |

真实样例（12 个 zip 部件、48 个正文段落、9 张表格的方案设计报告）：
识别出 19 处 `【待补充…】` 占位符（其中 19 处全在表格单元格里，`python-docx` 的
`Document.paragraphs` 根本看不到它们），填充 18 处后 **formatting delta = 0**。

---

## 2. 快速开始

```bash
# 需要 MoonBit 工具链：https://www.moonbitlang.com/download/
git clone https://github.com/niuniu513-ask/docx-mbt
cd docx-mbt

moon check                     # 类型检查
moon test --target js          # 34 个测试
moon run cmd/docx --target js -- info tests/fixtures/mini.docx
```

### 命令行

```
docx <command> <file.docx>

  info        包结构与文档概要（部件数、段落、表格、样式、字数）
  text        正文纯文本
  outline     标题层级结构
  parts       zip 部件列表
  styles      文档内定义的样式
  paragraphs  带样式 id 的段落清单
  tables      表格尺寸与首行
  analyze     完整 Document Specification（JSON，可 --out FILE）
  issues      仅问题清单（人类可读）
  fill        替换占位符并写出新文件
  validate    比较两份文件的格式差异
```

实战：

```bash
# 1. 看清文档长什么样、有什么问题
moon run cmd/docx --target js -- issues 报告.docx
[warn] 有 6 个标题把编号写进了文字（如「一、」），自动编号与目录会因此错乱
[warn] 仍有 19 处占位符未填写

# 2. 导出机器可读规范
moon run cmd/docx --target js -- analyze 报告.docx --out spec.json

# 3. 填充并写回（可多次 --set）
moon run cmd/docx --target js -- fill 报告.docx \
    --set "【待补充】=已通过验收" \
    --set "【待补充:具体分工】=张三" \
    --out 报告-filled.docx
replaced 18 occurrence(s)
wrote 报告-filled.docx

# 4. 证明没改坏格式
moon run cmd/docx --target js -- validate 报告.docx 报告-filled.docx
summary formatting delta 0: no drift detected
formatting delta 0 - the edit did not move any formatting
```

`validate --json` 会给出逐条 drift（`field / before / after / where`），差别精确到 twips。

---

## 3. 库用法

```moonbit
let bytes = @fs.read_file_to_bytes("报告.docx")
let doc = @ooxml.Docx::open(bytes)          // 解析 document.xml + styles.xml

// 读：结构、样式、大纲
for para in doc.body_paragraphs() {
  match para.heading_level() {
    Some(level) => println("H" + level.to_string() + " " + para.text())
    None => ()
  }
}

// 分析：Document Specification（JSON）
let spec = @docx.analyze(doc, @docx.AnalyzeOptions::default())

// 写：跨 run 的原地替换，返回替换次数
let count = @docx.fill(doc, [("【待补充】", "STC89C52")])

// 校验：把"格式没动"变成断言
let delta = @docx.format_delta(before, doc)
```

可运行示例见 `examples/demo`：

```bash
moon run examples/demo --target js            # 用内置 fixture
moon run examples/demo --target js -- 报告.docx
```

---

## 4. 包结构

```
xml/        XML 1.0 解析与序列化（命名空间感知）
  tree.mbt      Element / Attr / XmlNode，按命名空间 URI 而非前缀匹配
  parser.mbt    迭代式解析器（显式栈，无递归），保留注释/CDATA/PI，实体解码
  writer.mbt    序列化；新建元素会自动补 xmlns 声明
ooxml/      OOXML 容器与 Word 文档模型
  pack.mbt      WordPackage：zip 部件读写（基于 moonbitlang/core/archive/zip）
  units.mbt     twips / 半磅 / EMU / fiftieth-cm 换算
  props.mbt     RunProps / ParaProps：rPr、pPr 的读取与层叠
  styles.mbt    styles.xml：样式链（basedOn）解析与合并
  doc.mbt       Docx / Paragraph / Run / Table：可查询可编辑的文档模型
docx/       应用层
  analyze.mbt   Document Specification 生成（14 个顶层字段）
  issues.mbt    规则引擎：手工编号、占位符、空段落、标题跳级、字体混杂…
  edit.mbt      跨 run 替换、段落插入、整段改写
  validate.mbt  格式漂移 diff
cmd/docx/   命令行
examples/   可运行示例
```

---

## 5. 关键设计

**1）按命名空间 URI 匹配，不按前缀。**
Word 允许 `w:` / `ns0:` / 任意前缀绑定同一个命名空间。解析时就把前缀解析成 URI 存进元素，
之后所有查询（`is_w`、`child(ns, local)`、`find_all_w`）都走 URI，生产者怎么写都不影响。

**2）序列化时补声明。**
新建的 `w:t` 如果前缀在当前作用域里没有绑定，writer 会在该元素上补 `xmlns:w="…"`。
没有这一步，生成的 `<t>TBD</t>` 会落进"无命名空间"，Word 与 python-docx 都会当作空文本——
本项目就是先踩了这个坑（19 个单元格填完显示为空），才有这条不变式。

**3）跨 run 替换。**
先把段落所有 run 的文本拼成一条字符串，在字符级别查找，再把命中位置映射回
"第一个 run + 偏移 / 最后一个 run + 偏移"，从右往左应用：
首 run 保留前缀 + 替换值，末 run 保留后缀，中间 run 清空并摘除。
这样 `【待补` + `充】` 这种被 Word 拆开的占位符也能命中，且只产生一处替换。

**4）样式链而不是"当前值"。**
`pPr`/`rPr` 里没写的属性来自样式链（`basedOn` 逐级向上）再到 `docDefaults`。
`resolve_para` / `resolve_run` 按"基类在前、派生在后"的顺序层叠，循环引用有深度保护。

**5）把承诺变成测试。**
`format_delta` 对比两份文档的页面几何、文档默认值、逐段落 16 个属性、逐表格形状，
输出 drift 列表；集成测试断言纯文本替换后 `drift_count == 0`。

---

## 6. 测试与质量

```bash
moon test --target js
Total tests: 34, passed: 34, failed: 0
```

| 包 | 内容 |
| --- | --- |
| `xml` (13) | 声明、嵌套前缀、默认命名空间、属性命名空间、实体、转义往返、注释/CDATA/PI、非法标签、新建元素的 xmlns 补齐 |
| `docx` (21) | 占位符识别、手工编号识别、截断；单 run / 跨 run / 多处替换；整段改写保留格式；真实 docx 的部件数、段落数、大纲、占位符（含表格内）、issue 规则；fill→save→重开 的往返；命名空间不被破坏；delta = 0 |

fixture 由 `scripts/make_fixture.py` 生成（结构固定，含 1 张表格、跨 2 个空段落、3 个手工编号标题）。

---

## 7. 与其他工作的关系

- **参考来源**：本项目与作者此前的 Python 技能 [word-document-skill](https://github.com/niuniu513-ask/word-document-skill)（MIT）同源，
  沿用了它的"先分析成 Document Specification，再在文档自身格式体系内编辑"的方法论与 `issues` 词汇表
  （`manual_heading_numbering`、`unfilled_placeholders`、`heading_level_jump` 等）。
  **但实现是全新的**：Python 版基于 `python-docx` + `lxml`，本项目从 XML 解析器、zip 容器封装、
  样式链解析到编辑与校验全部用 MoonBit 重写，没有移植任何 Python 代码行。
- 解析器、序列化器、文档模型、分析器、编辑器、校验器均为本项目原创实现。
- 依赖仅为 MoonBit 官方核心库（`archive/zip`、`json`、`encoding/utf8`）与 `moonbitlang/x` 的文件系统封装。

## 8. 路线图

- [x] XML 解析/序列化（命名空间感知）
- [x] zip 容器读写（WordPackage）
- [x] 样式链与段落/字符属性层叠
- [x] Document Specification 分析（正文 + 表格占位符）
- [x] 跨 run 替换、段落插入、整段改写
- [x] 格式漂移校验（Formatting Delta）
- [ ] 目录（TOC）与编号定义（numbering.xml）解析
- [ ] 图片抽取与替换（r:embed → media）
- [ ] 批注 / 修订追踪保留策略
- [ ] `--target native` 打包（需要 MSVC 环境验证）

## 9. 许可证

MIT，见 [LICENSE](LICENSE)。

---

## English overview

`docx-mbt` is a pure-MoonBit engine for Word `.docx` files: it parses OOXML with a
namespace-aware XML reader, exposes a document model (`Docx`/`Paragraph`/`Run`/`Table`)
that resolves formatting through the style chain, analyzes a file into a JSON
*Document Specification* with rule-based findings, edits text **in place across run
splits** while preserving character and paragraph formatting, and finally verifies the
result with a formatting-drift diff that should report zero.

Highlights:

- Namespace-URI based matching, prefix-independent; the writer re-declares `xmlns` for
  elements it creates so generated markup never falls into the null namespace.
- Cross-run replacement: joins runs, matches at character level, maps the hit back onto
  the covered runs, applies right-to-left, detaches emptied runs.
- `format_delta` compares page geometry, document defaults, 16 paragraph properties and
  table shapes — the "Formatting Delta ≈ 0" promise is a test, not a slogan.
- 34 tests, runs on `--target js` and `--target wasm-gc`; no native toolchain needed.
