# word-document-skill

> **Word Document Intelligence** — 先理解任何 .docx 的结构与格式，再在原有格式体系内完成生成、修改、填充、排版、迁移与批量处理。

[![CI](https://github.com/niuniu513-ask/word-document-skill/actions/workflows/ci.yml/badge.svg)](https://github.com/niuniu513-ask/word-document-skill/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Tests](https://img.shields.io/badge/tests-73%20passed-brightgreen)]()

核心理念：**先理解 Word，再操作 Word。不是"套模板"，而是"理解文档"。**

任何 .docx 都是一套精密的格式体系——页面、样式、编号、页眉页脚、域、表格规则。本 Skill 的全部机制都围绕"在它的体系内工作"而设计，默认 **Formatting Delta ≈ 0**（只改变用户要求改变的东西）。

---

## 架构总览

```
用户提供 Word
    ↓
┌─────────────────────────────────────────────────────┐
│  Phase 1  Document Analyzer                          │
│  analyze_docx.py → Document Specification (JSON)     │
│  页面/样式三轨道字体/大纲编号/表格/图片失真/页眉页脚域  │
└──────────────────────┬──────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────┐
│  Phase 2  Content Planning                           │
│  content.json (内容与格式分离)                        │
│  cover | heading | paragraph | table | image | ...   │
└──────────────────────┬──────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────┐
│  Phase 3  Execution (8 workflows)                    │
│  generate_docx.py: insert / fill / rebuild           │
│  template-fill | generate | edit | format-migration  │
│  content-migration | batch-generate | analyze        │
│  lab-report-screenshots | capture_notebook.py        │
└──────────────────────┬──────────────────────────────┘
                       ↓
┌─────────────────────────────────────────────────────┐
│  Phase 4  Validation Loop (≤3 rounds)                │
│  validate_docx.py: Format Delta + structural check   │
│  render_docx.py: Word COM → PDF → per-page PNG       │
│  verify_evidence.py: 截图标签/图片/空图/变形 对账      │
│  repair_docx.py: auto-fix issues                     │
└──────────────────────┬──────────────────────────────┘
                       ↓
              docx + 格式差异报告 + 待补充清单
```

## 功能特性

| 能力 | 脚本 | 说明 |
|---|---|---|
| 文档分析 | `analyze_docx.py` | 完整 Document Specification：页面、样式、大纲、表格、图片、页眉页脚、封面/目录检测、显式/主题字体与 RGB/主题字色清单、问题列表 |
| 内容提取 | `extract_content.py` | 结构化 content-model JSON + Markdown 视图 |
| 样式详查 | `inspect_styles.py` | basedOn 链全量样式报告 |
| 文档生成 | `generate_docx.py` | 三模式：insert（锚点插入，最安全）/ fill（仅占位符）/ rebuild（保留封面重建正文） |
| 格式验证 | `validate_docx.py` | Format Delta（原文档 vs 产出）+ 结构体检 |
| 视觉渲染 | `render_docx.py` | Word COM（最忠实）→ LibreOffice 兜底；PDF + 逐页 PNG |
| 自动修复 | `repair_docx.py` | 空段落/标题keepNext/图片比例/手打页码/假标题；默认 dry-run |
| 证据对账 | `verify_evidence.py` | 截图标签数、图片数、空图、宽黑条、图片变形与超高 |
| 截图裁剪 | `crop_evidence.py` | 显式坐标比例、越界拒绝、PNG 保真与来源哈希；语义匹配另行核对 |
| 离屏采集 | `render_nb_html.py` + `shoot_nb_cells.cjs` | notebook → HTML → 无头浏览器整页截图 + 逐 cell 精确裁剪；不占屏幕、无滚动推算（默认路线） |
| 窗口采集 | `capture_notebook.py` | 真实前台滚动采集 + 二维模板匹配拼接（Windows，后备） |
| 实验报告流水线 | `lab_report.py` | inventory / prepare / crop / contacts / build / audit / frames；配置驱动、显式映射、不猜题号 |

字体分析区分中文、西文/数字和复杂文字轨道，并保留 Word 主题字体引用；字色分析区分显式 RGB、自动色和主题色，同时记录 tint、shade 与备用 RGB，便于精确复刻模板。

## 安装

```bash
# 核心依赖
pip install python-docx lxml Pillow

# 渲染检查（推荐）
pip install PyMuPDF
# Windows + Microsoft Word（最忠实的渲染后端）
pip install pywin32
# 或安装 LibreOffice 作为兜底渲染后端
```

或一键安装：

```bash
pip install -r requirements.txt
```

## 快速上手

```bash
S=scripts

# 1. 分析：拿到文档的完整格式体系
python $S/analyze_docx.py 模板.docx -o spec.json

# 2. 规划内容：写 content.json（见 examples/ 和 schemas/content-model.json）
# 3. 生成：在模板格式体系内插入/填充/重建
python $S/generate_docx.py --template 模板.docx --content content.json -o 输出.docx

# 4. 验证 + 渲染 + 修复（闭环，最多 3 轮）
python $S/validate_docx.py --original 模板.docx --final 输出.docx -o delta.json
python $S/render_docx.py 输出.docx -o 渲染/ --update-fields
python $S/repair_docx.py 输出.docx --fix empty-paragraphs,heading-keepnext --apply
```

### content.json 示例

```json
{"blocks": [
  {"type": "cover", "fields": {"姓名": "张三", "学号": "2024010101"}},
  {"type": "heading", "level": 1, "text": "一、实验目的", "numbering": "native"},
  {"type": "paragraph", "after": "一、实验目的", "text": "掌握单摆法测量重力加速度。"},
  {"type": "table", "rows": [["摆长","周期"],["50cm","1.41s"]], "caption": "表1 数据"},
  {"type": "image", "path": "fig.png", "width_cm": 10, "caption": "图1 曲线"}
]}
```

更多示例见 [`examples/`](examples/)。

## 目录结构

```
word-document-skill/
├── SKILL.md                 # Agent 入口：六阶段工作流 + 铁律 + 任务路由
├── README.md                # 本文件
├── LICENSE                  # MIT
├── CHANGELOG.md             # 版本变更记录
├── CONTRIBUTING.md          # 贡献指南
├── pyproject.toml           # 包配置 + 工具配置
├── requirements.txt         # 运行时依赖
├── requirements-dev.txt     # 开发依赖
├── scripts/                 # 16 个 Python 脚本 + 2 个 Node 采集脚本
│   ├── _docx_utils.py       # 共享工具（XML解析、样式链、编号解析等）
│   ├── analyze_docx.py      # 文档分析 → Document Specification
│   ├── generate_docx.py     # 模板/spec 内生成
│   ├── validate_docx.py     # 格式差异 + 结构体检
│   ├── render_docx.py       # docx → PDF → 逐页 PNG
│   ├── repair_docx.py       # 自动修复（dry-run 默认）
│   ├── verify_evidence.py   # 截图证据对账（标签数/图片数/空图/宽黑条/变形）
│   ├── lab_report.py        # 实验报告主力：inventory / prepare / crop / contacts / build / audit / frames
│   ├── render_nb_html.py    # notebook → 可截图 HTML（离屏渲染）
│   ├── shoot_nb_cells.cjs   # 无头浏览器整页截图 + 逐 cell 精确裁剪
│   ├── capture_vscode_cells.cjs # 真实 VS Code 窗口采集（CDP，仅明确要求 IDE 界面时）
│   ├── capture_notebook.py  # 滚动窗口采集 + 拼接（Windows，后备）
│   ├── crop_evidence.py     # 显式坐标裁图 + 哈希旁车记录
│   ├── smoke_lab_standalone.py # 独立安装自检（不读任何课程资料）
│   ├── extract_content.py   # 内容提取 → content-model + Markdown
│   └── inspect_styles.py    # 样式全量详查
├── schemas/                 # JSON Schema 定义
│   ├── document-spec.json   # analyze 输出格式
│   ├── content-model.json   # generate 输入格式
│   └── task-model.json      # 任务识别模型
├── workflows/               # 9 种任务类型的工作流指南
├── references/              # 9 篇通用深度规则文档
│   └── lab-report/          # 实验报告专用：配置字段、离屏采集、通用/本类/个案分档说明
├── evals/                   # 集成测试用例（4 个场景）
├── tests/                   # 单元测试（pytest，73 项）
├── examples/                # content.json 示例
└── .github/                 # CI + Issue Templates
```

## 设计原则

1. **分析先于操作** — 没分析过用户的 docx，不生成任何内容。
2. **原文档格式优先** — 优先级链：用户要求 > 用户文档格式 > 用户资料 > 用户格式要求 > 通用规范 > AI 推断。
3. **内容/格式分离** — content.json 管内容，脚本管格式，人机分工明确。
4. **原生结构** — Styles、段落分页规则、PAGE 域、原生编号；拒绝空格回车模拟排版。
5. **真实性** — 不虚构数据与经历；缺失标 `【待补充】`，模拟标 `【模拟数据】`。
6. **验证闭环** — 格式验证 + 渲染逐页视觉检查 + 自动修复，交付前必过。

## 八条禁令

1. 禁止在未分析用户 Word 文档的情况下直接生成或重新设计文档。
2. 禁止因为"我觉得这样更好看"而改动用户原始格式。
3. 禁止虚构数据与经历。
4. 禁止用空格、空回车、手打页码、手打编号模拟排版。
5. 禁止把整篇文档设成一种字体（中/西/数字三轨道分开）。
6. 禁止破坏标题体系。
7. 禁止伪造截图证据（模拟界面、跨运行拼接、给无输出的单元格插空图）。
8. 禁止交付前不做验证。

## 运行测试

```bash
pip install pytest pytest-cov
pytest tests/ -v
```

## 作为 Agent Skill 使用

将本目录放置到 Agent 的 skill 根目录下，Agent 会自动读取 `SKILL.md` 作为入口。`SKILL.md` 包含完整的六阶段工作流、铁律、任务路由表和脚本速查。

## 贡献

欢迎贡献！请阅读 [CONTRIBUTING.md](CONTRIBUTING.md) 了解开发流程、代码规范和 PR 检查清单。

## 许可证

[MIT](LICENSE) © word-document-skill contributors

大框模板按 [模板大框与连续分页](references/template-frames.md) 处理：先看原模板视觉效果，再判断底层容器；不因单格表格而擅自按题拆行。截图缩放与裁剪映射见 [证据规则 §6.1](references/screenshot-evidence.md#61-缩放变化后的坐标与证据映射)。
