# 参赛材料（天翼智创杯 · TeleAgent）

本目录是「石家庄铁道大学天翼智创杯大学生 TeleAgent 智能体创新竞赛」的投稿材料，
作品本体是本仓库（`word-document-skill`），提交名为 **STDU2026-方案文档合规初审助手**。

## 文件

| 文件 | 说明 |
|---|---|
| `STDU2026-方案文档合规初审助手-【姓名】-【专业班级】.pptx` | 三页方案（每页一屏），按赛事模板写的 |
| `build_ppt.py` | 生成上面这份 PPT；改 `PEOPLE` 后重跑即可 |
| `assets/review-draft.png` | 场景页证据：真实样本体检报告（64/100 C 级，3 条问题） |
| `assets/review-final.png` | 效果页证据：同一样本修好后 100/100 A 级 |
| `assets/terminal-review.png` | 第 2 页「如何证明能运行」：初稿 73 C（退出码 1）→ 定稿 100 A（退出码 0）的命令行实录，文本取自真实运行输出 |

## 提交前必须替换

1. `PEOPLE["name"]`、`PEOPLE["class_name"]` → 真实姓名与专业班级；
2. 第 3 页「验证者」一行里的 `试用【待补充：N 人 / M 次，及反馈原文要点】` → 真实试用记录。
   **没有做过试用就不要写数字**，宁可写"已自测 N 次 + 计划在 X 班试用"。
3. 文件名按 `技能名称-姓名-专业班级.pptx` 命名（脚本已按此规则生成）。

重新生成：

```bash
python submission/build_ppt.py
```

## 三页与评分点的对应

| 页 | 对应权重 | 内容 |
|---|---|---|
| ① 场景与问题 | 问题价值与场景真实性 30% | 目标用户、使用场景、原来卡在哪、真实样本里的三个硬伤（13 处占位符 / 缺 10 个章节 / 无页码域） |
| ② 技能与实现 | 技术实现 25% + 可用性 20% + 创新性 15% | 一键体检 / 规则包 / 零漂移修复；完整流程；技术路径；差异化（判"能不能交"而不是"帮你写"） |
| ③ 效果与复用 | 效果与复用 + 展示 10% | 73 C → 100 A、0.6 秒/份、格式差异仅 1 项（正是新增页码域）、换规则包 = 改 JSON |

## 截图的复现方式

```bash
python scripts/review.py examples/compliance-review/方案-初稿.docx --profile railway-plan --html /tmp/draft.html
python scripts/review.py examples/compliance-review/方案-定稿.docx --profile railway-plan --html /tmp/final.html
# 浏览器打开 /tmp/*.html，整页截图 → docs/screenshots/ 与 submission/assets/
```
