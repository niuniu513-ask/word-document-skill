# Changelog

All notable changes to **word-document-skill** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
