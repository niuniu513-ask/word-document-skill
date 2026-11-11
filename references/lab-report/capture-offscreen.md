# 离屏渲染采集 notebook 截图

报告要求"真实代码截图 + 运行结果截图"时，有两条路：

| 路线 | 图长什么样 | 适用 |
|---|---|---|
| **离屏渲染**（本页，首选） | notebook 单元格：上方灰底代码、下方真实输出 | 绝大多数实验报告；要的是代码与结果一一对应 |
| 驱动 VS Code 窗口 | VS Code 编辑器界面，含侧栏、行号、语法高亮 | 明确要求"要 VS Code 界面截图"、或要展示 IDE 操作过程 |

默认走离屏。只有当用户点名要 IDE 界面、或报告本身在讲"如何在 VS Code 里操作"时，才用 [runbook §3.2 的 VS Code 采集](../../workflows/lab-report-runbook.md) 那条路。

## 为什么不是驱动 VS Code

VS Code 的 notebook 是**虚拟化列表**：只渲染视口内的行。这会连锁出一串问题。

- 大输出单元格（比如 6 个子图的 matplotlib 图）高度超过窗口，单元格尾部的行不进 DOM，靠"哨兵行"判断边界会一直等不到。
- VS Code 按文件 URI 记住滚动位置。上一次崩溃/中断留下的滚动偏移会被恢复，打开就停在半屏，首行不在视口里。
- 判断"输出区到哪儿结束"要靠相邻单元格的 `style.top` 推算，滚动一变就要重算；步长滚动容易在过冲与欠冲之间振荡。
- 每次打开新文件都要走 Quick Open、等渲染、校验首行，一帧一帧来，几十个单元格要几分钟，中间还可能被编辑器自动更新打断。

实测代价：用 VS Code 路线做 25 个单元格，卡在第 19、20 个（都是大输出）反复重试，改脚本十余轮仍未跑通；改成离屏渲染后**一次跑完 25 个，零重试，约 40 秒**，而且不占屏幕。

离屏路线没有上述任何问题：页面是完整 DOM，元素边界由浏览器直接给出，不存在"推算"和"滚动"。

## 流程

```bash
# 1) notebook -> 独立 HTML（每个 code cell 一个块，data-idx 为原 notebook 索引）
python -X utf8 scripts/render_nb_html.py EVIDENCE/executed.ipynb -o EVIDENCE/nb.html

# 2) 无头浏览器整页截图 + 逐单元格元素截图
node scripts/shoot_nb_cells.cjs EVIDENCE/nb.html EVIDENCE/shots
```

`shoot_nb_cells.cjs` 需要 `playwright-core` 与一个本机 Chrome/Edge：

```bash
npm install playwright-core          # 不下载浏览器，复用系统已装的 Chrome/Edge
# 找不到时显式指定
node scripts/shoot_nb_cells.cjs ... --browser "C:/Program Files/Google/Chrome/Application/chrome.exe"
# 依赖装在别处时
PLAYWRIGHT_MODULE=/path/to/node_modules/playwright-core node scripts/shoot_nb_cells.cjs ...
```

产物：

```text
shots/full.png                 整页长图，用于总览和留档
shots/012-code.png             第 12 个原始 cell 的代码区
shots/012-output.png           同 cell 的输出区（无输出则不生成）
shots/012-cell.png             同 cell 的"代码+输出"连体图
shots/screenshots-cropped.json 清单，可直接作为 lab_report.py build 的 screenshots
shots/manifest.json            同上，附坐标与首行，便于人工核对
```

文件名用**原始 notebook 的 cell 索引**（三位补零），与 `report-config.json` 里的 `{"cell": N}` 一致。notebook 前面有 markdown 单元格时索引会跳号（0、2、4…），这是正确的，不要改成连续编号。

## 字号：决定成败的参数

插入 Word 后字能不能看清，取决于"渲染页宽 / 字号"这个比值，而不是截图的像素数。截图分辨率只决定清晰度。

```text
插入版心的字高(pt) ≈ 插入宽度(cm) / 2.54 × 72 × 字号(px) / 渲染页宽(px)
```

| 渲染页宽 | 字号 | 插入 15.2cm 时字高 | 结论 |
|---|---|---|---|
| 1080px | 13px | ≈5.2pt | 偏小，表格数字吃力 |
| 880px | 14.5px | ≈7.1pt | 可用（实测个案，按本次版心取） |
| 760px | 15px | ≈8.5pt | 更大，但每行代码会折行 |

`render_nb_html.py` 结束时会打印预估字高；低于 5pt 就调小 `--width` 或调大 `--font-size`。

代价是页面变窄后，matplotlib 的图（`figsize` 固定英寸）在页面上占比变小、两侧留白变多。给 `.out-img` 设 `margin: auto` 居中即可，不要为了填满宽度去拉伸图片——那会让图变形。

## 和构建流程对接

`lab_report.py build` 从 `screenshots-cropped.json` 读 `codeCrop` / `outputCrop` 两个字段，所以离屏产物可以直接喂进去，不需要改构建脚本：

```json
{"screenshots": "D:/course/evidence/shots/screenshots-cropped.json"}
```

清单里每项的 `codeAudit` / `outputAudit` 保留 `raw_sha256`、`output_sha256`、`crop_ltrb_px` 三个字段，`audit` 子命令仍会校验裁剪产物与记录哈希一致。离屏模式没有"原屏 → 裁剪"这一步，`raw` 与 `output` 指向同一文件、`crop_ltrb_px` 记为完整框，这是如实记录而不是伪造链路。

想用更紧凑的"代码+输出连体图"时，取 `cellCrop`（`012-cell.png`）：在配置里加 `"image_mode": "joint"`，`build` 就给每个单元格放一张图、一个标签（默认 `代码与运行结果截图：`，可用 `joint_label` 改），`audit` 的期望数也随之变成"单元格数"而不是"代码+输出两个标签"。见 [配置说明](lab-report-config.md)。

## 体积

2160px 宽的 PNG 一张 300~500KB，几十张就有 10MB 上下，进 docx 后体积跟着涨。纯文字和线条图用 256 色量化可以压掉六成，肉眼无差别：

```python
from PIL import Image
im = Image.open(src).convert('RGB')
im.quantize(colors=256, method=Image.MEDIANCUT, dither=Image.NONE).save(dst, optimize=True)
```

带渐变阴影的截图不要这样压（会出现色带），保持原图。

## 验收

- `render_nb_html.py` 报 `errors=0`。有报错输出就停下修代码，不要带着红色 traceback 截图。
- 清单里 `hasOutput` 为真的项数等于"确实有输出的单元格数"。`import`、函数定义这类无输出单元格不该有输出图。
- 随机抽 3 张放大看：代码首行末尾、最长一行末尾、输出表格最后一行是否完整。
- 逐张核对 `sourceFirstLine` 与报告里该小节的代码是否同一段——数量对得上不代表配对正确。
