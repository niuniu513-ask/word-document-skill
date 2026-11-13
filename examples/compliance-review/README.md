# 合规体检演示（可在 1 分钟内跑完）

```bash
python examples/compliance-review/make_samples.py           # 生成 初稿/定稿 两份方案
python scripts/review.py examples/compliance-review/方案-初稿.docx --profile railway-plan
python scripts/review.py examples/compliance-review/方案-定稿.docx --profile railway-plan
python scripts/validate_docx.py --original examples/compliance-review/方案-初稿.docx \
                                --final    examples/compliance-review/方案-定稿.docx
```

实测结果（2026-11-13）：

| 步骤 | 输出 |
|---|---|
| 初稿体检 | 得分 73/100，等级 C，阻断 1（缺「应急预案」一节）、严重 2（含「待定」、页脚无 PAGE 域），退出码 1 |
| 定稿体检 | 得分 100/100，等级 A，无问题，退出码 0 |
| 格式对比 | `format_delta_count = 1`，唯一差异是新增的页码域 `footer: [] -> [('第  页', ())]`；正文 20 → 21 行 |

这说明"修内容"和"改格式"被分开了：初稿到定稿只发生了三处**内容侧**编辑，
版式侧唯一的变化就是本来就要加的页码域。
