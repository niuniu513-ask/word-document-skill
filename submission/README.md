# 参赛材料（天翼智创杯 · TeleAgent）

本目录是「石家庄铁道大学天翼智创杯大学生 TeleAgent 智能体创新竞赛」的投稿材料。
作品本体是本仓库（`word-document-skill`），提交名为 **STDU2026-Word文档智能体指南**。

## 作品是什么

不是「一个帮你写 Word 的工具」，而是一本**可执行的 Word 指南书**：

| 层 | 文件 | 作用 |
|---|---|---|
| 判断 | `AGENT-GUIDE.md` | 任务路由表、五阶段验收标准、14 个任务的输入输出契约、11 条踩坑台账 |
| 清单 | `capabilities.json` | 同一份知识的机器可读版，智能体读它就知道该调哪个任务 |
| 动作 | `scripts/docx_agent.py` | 唯一入口：`list` / `--capabilities` / `--doctor` / `--dry-run` / 中文别名 |
| 能力 | `scripts/*.py` + `rules/*.json` | 分析、生成、修复、渲染、体检、证据对账；合规审查只是其中一项 |

## 文件

| 文件 | 说明 |
|---|---|
| `STDU2026-Word文档智能体指南-【姓名】-【专业班级】.pptx` | 三页方案（每页一屏），按赛事模板写的 |
| `build_ppt.py` | 生成上面这份 PPT；改 `PEOPLE` 后重跑即可 |
| `assets/terminal-review.png` | 第 2 页证据：命令行实录，73 C → 100 A（文本取自真实运行输出） |
| `assets/review-draft.png` | 备用证据：真实样本的 HTML 体检报告（64/100 C） |
| `assets/review-final.png` | 备用证据：修好后的 100/100 A |

## 提交前必须替换

1. `PEOPLE["name"]`、`PEOPLE["class_name"]` → 真实姓名与专业班级（文件名会跟着变）；
2. 第 3 页「验证者」一行里的 `试用【待补充：N 人 / M 次，及反馈原文要点】` → 真实试用记录。
   **没做过试用就不要写数字**，宁可写「已自测 N 次 + 计划在 X 班试用」。
3. 文件名按 `技能名称-姓名-专业班级.pptx` 命名（脚本已按此规则生成）。

重新生成（需要 python-pptx，用仓库自带的运行时即可）：

```bash
python submission/build_ppt.py
```

## 三页与评分点的对应

| 页 | 对应权重 | 内容 |
|---|---|---|
| ① 场景与问题 | 问题价值与场景真实性 30% | 智能体改 Word 的三类翻车（改格式 / 编数据 / 改完没验）+ 真实样本三个硬伤（13 处表格内占位符、缺 10 章、0 个页码域） |
| ② 技能与实现 | 技术实现 25% + 可用性 20% + 创新性 15% | 三件套（指南 / 清单 / 入口）+ 读得懂·写得进·验得住三项核心能力 + 命令行实录 + 五步流程 |
| ③ 效果与复用 | 效果与复用 + 展示 10% | 94 测试 / 0.6 秒一份 / 目标之外零漂移 / 14 个任务一条命令；换单位、换文种、换智能体各改一处 |

## 复现三页里的每个数字

```bash
python -m pytest -q                                              # 94 passed
python scripts/docx_agent.py --doctor                            # 依赖 / 脚本 / 规则包 / 文档一致性
python scripts/review.py examples/compliance-review/方案-初稿.docx --profile railway-plan   # 73 C
python scripts/review.py examples/compliance-review/方案-定稿.docx --profile railway-plan   # 100 A
python scripts/validate_docx.py --original examples/compliance-review/方案-初稿.docx \
                               --final examples/compliance-review/方案-定稿.docx          # 差异 1 项 = 新增页码域
```

## 截图的来源

- `terminal-review.png`：把**真实运行输出**（上面两条 `review.py` 命令的 stdout）排成深色终端样式后离屏渲染，
  不是手写的假截图；原始输出可用上表命令逐字复现。
- `review-draft.png` / `review-final.png`：`review.py --html` 生成的报告页面，由无头浏览器整页截图。
