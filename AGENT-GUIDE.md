# Word 文档智能体指南书

> 一句话：**把「怎么改一份 Word 才算对」写成智能体能执行的判断与工具**，而不是一堆提示词。
> 读者：接 Word 任务的智能体（首选）、带学生做文档的助教、审方案的工程师。

这份指南是作品的正面：`SKILL.md` 是给模型看的操作手册，本文件是**决策地图 + 验收标准 + 踩坑台账**，
`capabilities.json` 是同一份知识的机器可读版本，`scripts/docx_agent.py` 是它的唯一入口。

---

## 1 三条铁律

1. **不在文档自己的格式系统之外造格式。** 先 `analyze` 拿到 Document Specification，
   再在既有样式/页面设置里写；新建样式名、改字号、改页边距都属于「动了别人的格式」，默认禁止。
2. **不编造数据。** 缺什么写 `【待补充：…】`，并在交付说明里逐条列出；宁可少写，不可假写。
3. **改完必须验证。** 格式差异 ≈ 0（只含用户要求的变化）+ 逐页渲染核对，最多 3 轮修复；
   3 轮还不行就停下来报告，不要靠反复重试掩盖问题。

违反任何一条，产出即使「看起来对」也不算完成——这份指南的所有工具都围绕这三条设计。

---

## 2 任务路由表：我在说什么 → 该跑什么

| 用户的话（触发特征） | task | 命令（`python scripts/docx_agent.py …`） | 产出 |
|---|---|---|---|
| 「这文档什么结构 / 格式怎么定义的」 | `analyze` | `analyze a.docx -o spec.json` | Document Specification（JSON） |
| 「把内容搬进那个模板」 | `extract` | `extract a.docx -o content.json --markdown a.md` | 内容模型 + Markdown |
| 「为什么一级标题长得不对」 | `styles` | `styles a.docx --style "Head*" --text styles.txt` | 样式详查（人读） |
| 「照这个模板写一份 / 填占位符 / 插一节」 | `generate` | `generate tpl.docx --content content.json -o out.docx --mode fill` | 新 docx |
| 「我改完有没有动到别的版式」 | `validate` | `validate --original before.docx --final after.docx` | 格式差异清单 |
| 「这份能不能报审 / 合不合规」 | `review` | `review a.docx --profile railway-plan -o r.md --html r.html` | 分数+等级+逐条修法 |
| 「编号混用、空段落、假标题帮我修掉」 | `repair` | `repair a.docx --fix numbering,empty --apply -o fixed.docx` | 修复后 docx |
| 「导 PDF / 每页出图 / 目录页码没更新」 | `render` | `render a.docx -o pages/ --dpi 110 --update-fields` | PDF + 逐页 PNG |
| 「报告里的截图数量对不对 / 有没有空图」 | `evidence` | `evidence a.docx --label 代码截图 --expect 代码截图=29 --strict` | 对账报告 |
| 「截图只留这一块」 | `crop` | `crop raw.png crop.png --box X Y W H --identity cell5/code` | 裁剪图 + 来源哈希 |
| 「要真实 IDE/终端截图」 | `capture` | `capture --title "Visual Studio Code" --stitch --out frames/` | 拼接长图 |
| 「notebook 要出截图」 | `notebook-html` | `notebook-html nb.ipynb -o nb.html --width 1100` | 独立 HTML |
| 「按学校模板做实验报告」 | `lab-report` | `lab-report --help` | 报告 docx + 证据清单 |
| 「打包成 .skill 发给别人」 | `package` | `package -o dist/word-document-skill.skill` | .skill 压缩包 |

自然语言别名也可直接当 task 用：`审查`/`体检`/`合规` → `review`，`生成`/`填充`/`插入` → `generate`，
`验证`/`格式差异` → `validate`，`修复` → `repair`，`渲染`/`出图` → `render`。

---

## 3 五阶段闭环（每一步都有验收标准）

| 阶段 | 做什么 | 验收标准（不达标就别进下一步） |
|---|---|---|
| ① 分析 | `analyze` + `extract` + `styles` | spec 里 sections/styles/outline 都有内容；能说出「正文样式名、字号、页面尺寸、页边距」四个事实 |
| ② 规划 | 写清改哪些段落、用什么样式、是否新增内容 | 用户确认过「改哪里、不改哪里」；需要新数据的地方已标 `【待补充】` |
| ③ 执行 | `generate` / `repair` / 直接改 XML | 只动目标段落；未新增样式；未改页面设置 |
| ④ 验证 | `validate` → `review` → `render` → `evidence` | `format_delta` 只含目标改动；体检无新增阻断；逐页看过受影响的页及前后页 |
| ⑤ 交付 | 交付说明 | 列出：改了什么、没改什么、哪些是 `【待补充】`、验证结论（附数字） |

**顺序不能反**：先 `review` 再改再 `validate`。反过来会把自己引入的问题算成存量问题，说不清责任。

---

## 4 输入输出契约

| task | 必要输入 | 产出 | 可重跑 | 幂等 |
|---|---|---|---|---|
| `analyze` | docx | `spec.json` | 是 | 是（同输入同输出，JSON 键序稳定） |
| `extract` | docx | `content.json` / `.md` | 是 | 是 |
| `styles` | docx | 文本报告 | 是 | 是 |
| `generate` | 模板/spec + content.json | docx | 是 | 否（`insert` 模式重复跑会重复插入） |
| `validate` | 两份 docx | 差异清单 + 退出码 | 是 | 是 |
| `review` | docx + 规则包 | md/html/json + 退出码 0/1/2 | 是 | 是（单份 0.6 秒） |
| `repair` | docx | docx（默认 dry-run 不落盘） | 是 | 是（`--apply` 后不再有可修项） |
| `render` | docx + Word COM 或 LibreOffice | pdf + png | 是 | 是 |
| `evidence` | docx（含图） | 对账报告 + 退出码 | 是 | 是 |
| `crop` | 原图 + 坐标 | 裁剪图 + 来源哈希 | 是 | 是 |
| `capture` | 窗口标题 | 拼接长图 | 是 | 否（滚动时序相关） |
| `notebook-html` | ipynb | html | 是 | 是 |
| `lab-report` | 模板 + notebook | docx + 指纹 | 是 | 否 |
| `package` | 仓库 | .skill | 是 | 是 |

退出码统一：`0` 通过 / `1` 有阻断 / `2` 用法或输入错误——可直接当 CI 卡口。

---

## 5 失败模式速查（全是真踩过的）

| 现象 | 诊断 | 修法 |
|---|---|---|
| 终端打不出 🟢 / ✓，报 `UnicodeEncodeError: 'gbk' codec can't encode` | Windows 控制台默认 GBK | 脚本入口调 `_docx_utils.setup_stdout()`；或 `PYTHONIOENCODING=utf-8` |
| 占位符体检说 0 处，明明文档里一堆 | 占位符在**表格单元格**里，`Document.paragraphs` 看不到 | 用 `body_paragraphs + tables 网格` 双路扫描（`review.py` 已内置） |
| `page_size` 误报 | python-docx 新建文档默认 Letter 21.6×27.9 cm | 显式设 A4（`page_width = Cm(21.0)`）再写内容 |
| 字体名报成 `minorEastAsia` | 主题字体引用，不是字体名 | `resolve_font()` 用 `docDefaults` 还原成真实字体（如宋体） |
| `format_delta` 不为 0 | 三成因：重建了段落（丢 run 级 rPr）、样式继承断链、section 属性没复制 | 用「克隆锚点 pPr/rPr」的方式插入；只改文字时保持 run 结构 |
| 标题看起来是标题，样式却是正文 | 居中 + 加粗伪装的「假标题」 | `review` 的 `fake_heading` 规则会点名；`repair --fake-heading-style` 可整批抢救 |
| 目录/页码是死文本 | 缺 PAGE/TOC 域 | `render --update-fields` 更新域；`review` 会判「缺页码域」 |
| 两份文档一页数不一样 | 空段落堆积、行距/段前段后不同 | `repair --fix empty,spacing`；对比 spec 的 `left_twips/space_before` |
| 图在 Word 里变形 | 宽高比与原图不一致 | `evidence --max-aspect` 卡口；插入时按原始 EMU 尺寸等比缩放 |
| LibreOffice 与 Word 渲染不同 | 排版引擎差异 | 以 Word COM 为准（`render_docx.py` 优先 COM），LibreOffice 只作兜底 |
| 生成的 XML 少命名空间前缀，文字看不见 | 手工造元素时 qname 未带前缀 | 只通过库 API 造元素；自造节点必须带 `w:` 前缀与命名空间声明 |

---

## 6 让智能体自己找路

```bash
python scripts/docx_agent.py --capabilities   # 机器可读任务表（id/何时用/命令/输入/产出）
python scripts/docx_agent.py list             # 人读任务表
python scripts/docx_agent.py analyze --dry-run a.docx   # 只打印将执行的命令，不执行
python scripts/docx_agent.py --doctor         # 依赖、脚本、规则包、文档一致性自检
python scripts/docx_agent.py --guide          # 本文件的章节目录
```

编排建议（多智能体）：**分析器**只调 `analyze/extract/styles`；**写手**只调 `generate`；
**验收员**只调 `validate/review/render`，且不允许写文档。三者用文件交接，天然可审计。

---

## 7 质量门：怎么证明「没退化」

- `pytest -q`：82 个用例，覆盖分析、生成、修复、体检、对账、路由与清单一致性。
- CI：`.github/workflows/ci.yml` 跑 ruff + pytest；规则包与 `capabilities.json` 的路径一致性由测试守住
  （清单里写的脚本或工作流文件不存在，测试直接红）。
- 真实样本回归：`outputs/` 里那份 48 段 9 表的方案报告，体检 64/100 C、演示样本 73 C → 100 A，
  数字写进 README，改代码后必须仍然成立。

---

## 8 边界（不做什么）

- 不替代人做**技术判断**：只判「要件齐不齐、格式合不合规、有没有编数据」，不判方案对不对。
- 不做联网检索、不调大模型打分：所有结论来自文档 XML 的确定性规则，同一份文档永远同一结论。
- 不接管用户的 Word 会话：全部离线处理，`render` 才用到 Word COM 或 LibreOffice。
