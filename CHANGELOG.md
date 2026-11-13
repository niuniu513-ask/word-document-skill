## 1.5.0

### 新增：一本可执行的指南书

- `AGENT-GUIDE.md`：任务路由表（自然语言 → task → 命令 → 产出）、五阶段闭环与逐阶段验收标准、
  14 个任务的输入输出契约、11 条真实失败模式台账（现象 / 诊断 / 修法）、质量门与边界。
- `capabilities.json`：同一份知识的机器可读清单（14 个任务，含 `when` 触发语、命令示例、输入产出、下一步）。
- `scripts/docx_agent.py`：唯一入口。`list` / `--capabilities` / `--doctor` / `--guide` / `--dry-run`，
  支持中文别名（体检、审查、合规、生成、修复、渲染…），退出码原样透传。`--doctor` 一次检查依赖、
  脚本、规则包与文档是否齐全。
- `tests/test_docx_agent.py`（12 项）：清单自洽（引用不存在的脚本/工作流/文档即失败）、任务契约完整、
  JSON 输出合法、未知任务退出码 2、`--dry-run` 无副作用、`--doctor` 就绪。
- `package_skill.py` 打包时纳入 `AGENT-GUIDE.md`、`capabilities.json` 与 `rules/`。
- README 顶部改为「指南书」定位，新增「一条命令开始」五阶段表；SKILL.md 增加 §10 入口说明。

# Changelog

All notable changes to **word-document-skill** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- CI 的 `test` 作业自 v1.3.0 起在收集阶段失败（`ModuleNotFoundError: No module named 'nbformat'`）：改装修 `requirements-notebook.txt`，并让 `tests/test_lab_report.py` 在缺少 notebook 依赖时 skip 而不是中断收集。


## [1.4.0] - 2026-11-13

### Added
- `scripts/review.py`：一键合规体检。输出得分/等级/三级问题清单（阻断·严重·提示），每条带位置、详情与修复建议；支持 `-o` Markdown、`--html` 自包含报告、`--json` 机器可读、`--baseline` 附带格式漂移检查；退出码 0/1 可做 CI 卡口。
- `rules/`：三套规则包 `railway-plan`（铁路施工方案报审自查，10 个必备章节）、`lab-report`（实验报告七要素）、`generic`（通用结构体检）。新增文种只需加一份 JSON。
- `workflows/compliance-review.md`：体检工作流的适用场景、读报告方法、与 repair/validate 的衔接顺序、已知局限。
- `references/compliance-rules.md`：规则包字段说明与编写建议。
- `tests/test_review.py`：9 个用例，覆盖通过/缺章节/占位符/层级跳级/页码域/禁用语/规则包切换/报告渲染/退出码。

### Changed
- `SKILL.md` 触发条件新增"文档体检/合规初审/能不能报审"类表述，脚本速查表补 `review.py`。
- 测试总数 73 → 82。

## [1.3.0] - 2026-11-11

### Added
- **离屏渲染采集**（`scripts/render_nb_html.py` + `scripts/shoot_nb_cells.cjs`）：notebook → 独立 HTML → 无头浏览器整页截图，再按 `.input` / `.output` / `.cell` 元素边界逐块裁剪。不占屏幕、无滚动推算、无哨兵行；实测 25 个 cell 一次跑完、零重试、约 40 秒，替代原先卡在大输出帧的 VS Code 路线。
- **实验报告资料区** `references/lab-report/`：把实验报告类做法从通用 Word 规则里分出来，按"通用规则 / 本类方法 / 个案记录"三档标注适用范围，并附 `lab-report-config.md`（配置字段）与 `capture-offscreen.md`（离屏流程与字号换算）。
- `scripts/lab_report.py`：`inventory` / `prepare` / `crop` / `contacts` / `build` / `audit` / `frames` 七个子命令；配置驱动、显式映射，不猜题号、不自动生成分析。
- `image_mode: "joint"`：每个 cell 只插一张"代码+输出"连体图，标签可用 `joint_label` 覆盖；`audit` 的期望标签数随模式自动切换。
- `audit` 支持离屏清单：清单没有 `evidence_notebook` 时，改用 `source` / `sourceFirstLine` 与 notebook 逐字对账，不再强制要求证据副本；指向错误 cell 会直接失败。
- `tests/test_lab_report.py`：joint 模式插图与标签、离屏清单审计、错 cell 必须失败、裁图保真与校准比例。
- `requirements-notebook.txt`（notebook 执行依赖）与 `scripts/smoke_lab_standalone.py`（独立安装自检，不读取任何课程资料）。
- 大框模板识别与分页规则：分别记录视觉用途和 XML 容器，避免按题拆表造成布局偏离；明确用户要求无表格时的处理与验收。
- `crop_evidence.py`：显式坐标变换、越界拒绝、无损裁剪及来源哈希旁车记录；不将几何检查当作 cell 语义验证。
- 裁图像素保真、坐标越界和原屏保护回归测试。

### Changed
- `workflows/lab-report-runbook.md`：第 3 节拆成 **3.1 离屏渲染（默认）** 与 **3.2 VS Code 窗口（仅当明确要求 IDE 界面）**，并写明 VS Code 路线的实测代价与失效点（虚拟化列表、按 URI 恢复滚动位置、输出边界只能推算）。
- `SKILL.md`：任务路由把"带截图的报告"指向 runbook，VS Code 采集降为备选；链接改指 `references/lab-report/`。
- `README.md`：目录树、功能表与实际脚本对齐，测试数更新。
- 新脚本统一通过 `ruff check`（CI 强制），`lab_report.py` / `render_nb_html.py` / 测试文件按同一风格格式化。

### Fixed
- 补充框线坐标、段落边框外扩、封面下划线保留和逐页留白验收,区分压页数与真实排版改进。
- 截图标签识别支持中英文括号的编辑器说明（如“代码截图（VS Code）：”）,避免误报大量孤立图片。
- 修正截图工作流中直接解析 `.doc`、默认原地执行 notebook、强制改用户代码和删除原屏的指导。
- 补充 keep-with-next 长链推空首页、长代码输出标签孤立、行单位段距残留、旧渲染页混入的排查。
- 删除互相矛盾的历史数量基线；区分结构点数与截图内容匹配，明确重新裁图后需更新最终 DOCX。
- 禁止涂抹证据图 UI；允许保留原屏后成对裁出代码与输出，区分实测数字与理论推导值。

## [1.2.0] - 2026-10-18

### Added
- **Screenshot Evidence** (`references/screenshot-evidence.md`): what counts as a real screenshot, what is forbidden (synthetic editor UI, cross-run stitching, blank images for cells without output), Windows capture prerequisites (DPI awareness, detached launch), wheel direction, 2-D block-match stitching, per-cell cropping from single frames, tight-crop thresholds, `width_cm ≈ px × 0.0245` sizing so every screenshot keeps the same on-page character size, and the numeric checks that catch errors the eye misses.
- **Lab report workflow** (`workflows/lab-report-screenshots.md`): eight-step pipeline from notebook execution to delivery, with acceptance criteria per step and a real baseline (37 pages / 74 images / 29 code labels / 27 result labels).
- **Evidence Verifier** (`scripts/verify_evidence.py`): counts labels, pairs them with images, flags missing pairs, orphan images, blank images, wide black bars, over-tall images and aspect distortion; DOCX full check plus PDF count-only mode; `--expect LABEL=N` and `--strict` for gating.
- **Window Capture** (`scripts/capture_notebook.py`): Windows scroll-capture with real wheel events, edge-based 2-D offset estimation, auto-stitch, alignment self-check and crop preview. Verified on VS Code 1.9x + Jupyter at 150% scaling.
- 10 unit tests for the evidence verifier.

### Changed
- `SKILL.md`: routing row for screenshot-backed reports; Phase 4 now requires evidence reconciliation; five new high-frequency pitfalls.
- `references/report-writing.md`: new section on per-question code/result screenshot structure.
- `references/visual-qa.md`: screenshot-specific page checklist.
- `README.md`: feature table, directory tree, test count, and an eighth prohibition (no fabricated screenshot evidence).

## [1.1.0] - 2026-09-12

### Added
- 识别运行与样式中的 Word 主题字体引用（ASCII、东亚、西文、复杂文字四轨）。
- 识别主题字色的 tint、shade 和备用 RGB，避免把主题色误报为普通 RGB。
- 增加显式 RGB、主题色和主题字体元数据的自动回归测试。

### Changed
- 清理全部 Ruff 静态检查问题，并将 GitHub Actions 的 Ruff 检查改为强制通过。
- 修正 README 的 CI 仓库地址与测试数量。

## [1.0.1] - 2026-09-07

### Fixed
- 明确并回归测试模板固定文本保护：`rebuild` 仅重建正文，封面标题、课程名、校名和固定说明默认逐字保留。
- 补充模板填充工作流对“模板标题与源资料标题不一致”的处理规则，默认以模板为准并提示用户。

## [1.0.0] - 2026-09-07

### Added
- **Document Analyzer** (`analyze_docx.py`): full Document Specification output — page setup, styles (three-track font resolution via basedOn chain), outline with native/manual numbering, tables (col widths, merges, header repeat), images (distortion detection), headers/footers with PAGE/TOC fields, cover & TOC detection, formatting inventory (actual run fonts/colors), and an `issues` list.
- **Content Extractor** (`extract_content.py`): structured content-model JSON + Markdown view, preserving heading levels, table rows, captions, and image positions.
- **Style Inspector** (`inspect_styles.py`): detailed style report with basedOn chain and effective formatting.
- **Generator** (`generate_docx.py`): three modes — `insert` (anchor-based, least invasive), `fill` (placeholder-only), `rebuild` (keep cover, rebuild body). Supports `{{key}}`, `【key】`, and wide-spaced label+underscore placeholders; content blocks: cover, heading, paragraph, table, table_fill, image, caption, page_break.
- **Renderer** (`render_docx.py`): Word COM (Windows, most faithful) → LibreOffice fallback; PDF + per-page PNG; `--update-fields` refreshes TOC/PAGE.
- **Validator** (`validate_docx.py`): Format Delta (original vs final, sections/styles/headers/tables/images) + structural sanity check (empty paragraphs, distorted images, orphan headings, manual page numbers, overflow tables).
- **Repair** (`repair_docx.py`): dry-run by default; fixes empty-paragraphs, heading-keepnext, image-aspect, manual-pagenum, fake-headings.
- **JSON Schemas**: `document-spec.json`, `content-model.json`, `task-model.json`.
- **7 Workflows**: analyze, generate, edit, template-fill, format-migration, content-migration, batch-generate.
- **7 Reference docs**: document-analysis, word-style-rules, template-rules, report-writing, document-editing, visual-qa, repair-rules.
- **4 Eval cases**: template-fill (lab report), format-migration (notice), precision-edit, layout-repair.

### Fixed
- `add_caption` duplicate text when caption style exists.
- Removed dead code in `load_numbering`.
- Consolidated duplicate `package_flags_orig/final` helpers.

### Security
- All scripts use `sys.path.insert` for local imports; no arbitrary code execution.
- `render_docx.py` operates on a copy when `--update-fields` is used to avoid mutating user originals unexpectedly.
