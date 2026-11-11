# 工作流:带真实截图的实验报告(Lab Report with Screenshots)

触发:「按这个模板写实验报告,要 **VS Code 的真实代码截图和输出截图**」这一类要求。
与 `template-fill.md` 的区别:多了"证据链"——图必须来自真实运行,且数量、位置、比例都要能对账。

严格原模板+notebook任务先执行 [详细runbook](lab-report-runbook.md)。**截图默认走 runbook §3.1 离屏渲染**（`render_nb_html.py` + `shoot_nb_cells.cjs`）：一次整页、按元素边界裁剪，不需要滚动推算。本文件的滚动采集只在“必须拿到真实 VS Code 窗口像素”且离屏路线不可用时作为后备；不得把经验行高模型当成真实坐标。新版校验器覆盖表格内部，调用时填写本次 `--expect` 数量，避免旧版漏查大框后0张图也通过。

先用 `template-fill.md` 的分析与生成骨架,本文件补的是**两端的活**:开头把证据采出来,结尾把证据核对一遍。

## 0. 开工前先问清三件事(或从资料里确认)

| 事项 | 为什么必须确认 | 默认假设 |
|---|---|---|
| 报告模板 | 决定标题体系、封面字段、是否已有"代码截图"标签 | 无模板则按 `references/report-writing.md` |
| 代码在哪、跑不跑得动 | 截图前必须真跑一遍 | 默认在副本上执行并保留原始文件 |
| 题号与代码的对应 | 一题可能跨多个 cell,一 cell 也可能含两题 | 按模板题号分段,逐段确认 |

题号对应不上就先列一张对照表给用户看,**不要凭猜测把图插到错的题下**——错位比缺图更糟。

## 1. 八步流程

```
Step 1  分析模板    analyze_docx.py + 原模板渲染 → spec,锁定固定内容与容器
Step 2  执行        notebook 真跑,输出落盘,映射到模板小节
Step 3  采集        真实前台 + 滚轮 + 逐帧抓取
Step 4  切图        长图定位 → 单帧裁切 → 紧裁
Step 5  生成        写 content.json / 生成脚本 → docx
Step 6  对账        verify_evidence.py:标签数 / 图片数 / 紧凑度
Step 7  目检        渲染 PDF → 逐页 Read
Step 8  交付        清理中间产物 + 汇报
```

### Step 1 执行:先跑,再看

```bash
python -m nbconvert --to notebook --execute test-1.ipynb --output test-1.executed.ipynb
```

验收:

- 每个非空代码 cell 都有 `execution_count`(无输出也要有计数);markdown/raw cell 不参与此项。
- `outputs` 里 `output_type` 为 `stream` / `execute_result` / `display_data` / `error` 的分布,记下来——
  **它直接决定哪几题有"实验结果截图"**。有 `error` 就定位原因;若用户要求按已有代码写,保留错误证据并说明,未经授权不要改写源码来制造成功结果。
- 执行后 git/备份比对:`source` 与你写的代码一致(执行不会改源码,但有人会在执行时手改)。

### Step 2 分析模板

```bash
python scripts/analyze_docx.py "已转换的模板.docx" -o spec.json
```

`.doc` 先转换为 `.docx` 副本并保留原文件。除封面字段、标题、截图标签外,还要渲染原模板并识别大框用途;单格容器不能自动解释为多题表格。按 `references/template-frames.md` 记录。
模板已有截图标签时在原位置插图；可同段换行，也可相邻段。保持标签文字和格式并以PDF核对归属，不能只靠keep设置认定一定同页。

### Step 3 采集

完整方法见 `references/screenshot-evidence.md`。此处只给执行顺序:

```bash
python scripts/capture_notebook.py --list-windows                 # 先确认窗口标题
python scripts/capture_notebook.py --title "test-1.ipynb" \
       --preview _cap/crop.png                                    # 校准裁切线,必做
python scripts/capture_notebook.py --title "test-1.ipynb" \
       --out _cap/test1 --step 2
```

`--preview` 会在窗口上画两条红线:上红线必须在工具栏**下方**(否则每帧都带一遍标签栏),下红线在状态栏**上方**。红线位置不对就调 `--crop-top/--crop-bottom`,别硬跑。

- 采集期间**不要动鼠标键盘**,也不要让采集脚本以外的东西抢焦点。
- 每抓完一段,立刻拼一张长图并肉眼确认:顶部是第一屏、底部是最后一个 cell、中间无重影。
- 脚本会报重叠率,低于 0.35 就减小 `--step` 重抓(重叠不够,对齐会失手)。
- 采集失败重来时,清空该段目录再来,别混着不同缩放比例的帧。

### Step 4 切图

先算 cell 边界,再回单帧裁,最后紧裁:

1. 优先用同帧DOM边界或原屏人工测量；滚动法仅以 notebook 结构估计 cell 高度辅助定位。经验模型 `h ≈ 56 + 28n`(n = 源码行数)
   与字号、缩放绑定,误差 >3% 就说明换过环境,回单帧量两三个 cell 重新拟合。
2. 在长图上求出每题的 y 区间,选出**完整包住该区间的那一帧**。
3. 在该帧上裁代码块、输出块,各留 8~14px 内边距。
4. 生成紧裁版:`*_t.png`(代码)、`*_f.png`(输出)。
5. **无输出的题只留代码图**,删掉输出标签。

验收:每张图单独打开看一遍,顶部不能是半行字,底部不能带下一个 cell 的头部。缩放变化后按 `screenshot-evidence.md` §6.1 重新测坐标;用题号/cell 映射逐对核对代码与输出,不要只点数图片。

### Step 5 生成

两种做法,优先沿用模板直接填充;题数多本身不构成重建理由:

- **A. content.json + generate_docx.py**:适合模板结构规整、题数不多。
- **B. 专用构建脚本**(`python-docx` 直写):题多、每题要"题目 / 代码图 / 输出图 / 分析"四段式时更省事。
  B 的代价是绕过了 `--template` 的格式继承,所以必须在脚本里显式对齐模板的:页面尺寸与页边距、
  正文字体(宋体 + Times New Roman)、标题字号与加粗、表格样式、封面表格的列宽与合并。
  **只要用了 B,validate_docx 的 delta 对比就不能省**——用它证明你没有把模板格式改坏。

图片插入只给宽度:

```python
doc.add_picture(path, width=Cm(width_cm))
width_cm = max(4.0, min(15.2, px_width * 0.0245))   # 见 screenshot-evidence.md §8
```

### Step 6 对账(本工作流的核心关口)

```bash
python scripts/verify_evidence.py 输出.docx \
       --label "代码截图" --label "实验结果截图" \
       -o evidence.json
# 按本次映射计算数量后,可加 --expect "代码截图=N" --expect "实验结果截图=M";N/M 须替换为整数
```

以下结构与图像特征检查通过后,仍需在 Step 7 核对截图内容与题号匹配:

| 检查 | 通过条件 |
|---|---|
| 标签数 | 与题数/有输出的题数一致 |
| 图片数 | 等于标签总数,且每张图都能定位到标签后的段落 |
| 空图 | 无全黑/全白/纯背景图 |
| 宽黑条 | 内容宽度 / 图片宽度 ≥ 0.35 |
| 超高图 | 高宽比 ≤ 8 |
| 图片变形 | 显示宽高比与原图一致(优先只设宽度;同时设置时按相同比例计算) |

不合格就回到 Step 4 重裁,不要靠"Word 里看起来还行"放过。

### Step 7 目检

```bash
python scripts/render_docx.py 输出.docx -o _render --dpi 100 --update-fields
```

逐页 Read,除通用清单(`references/visual-qa.md`)外,本类文档额外查:

- 每个"代码截图:"标签**下面紧跟着图**,不是上一题的图窜到下一题。
- 代码图里的字与正文对比:目测行高接近正文行高,不出现"图里字比标题还大"。
- 输出图不是一根黑条、不是纯黑方块。
- 标题不孤立页尾(注意:`keep_next=True` 在题目清单连排时会**造成**连锁分页,反而要逐段关掉)。
- 最后一页不是只挂一张图。

### Step 8 交付

保留原屏、裁图及其坐标/哈希清单,与交付文件分目录存放;未经用户要求不要删除证据。渲染缓存可在确认终稿对应的 PDF 与检查记录已保留后清理。截图替换后重新构建 DOCX,不能只更新磁盘 PNG 就视为已修复。

交付说明里写清:页数、图片数、标签数、修了哪几类问题、哪些图是"同一窗口连续滚动拼接"。

## 2. 按当前任务确定验收数量

从本次题号到代码单元格的映射计算代码图、输出图数量;导入单元格、共享代码和长输出拆图应单独说明。历史报告的页数、题数和图片数不能作为新报告的硬编码目标。

## 3. 常见变体

**代码不在 notebook,是 .py 脚本**
截图改为终端:跑完保留终端窗口,采集方式相同,但终端没有 cell 分割,按"命令 + 输出"成对切,
切分靠命令提示符行(`>>> ` 或 `$ `)定位。

**题目要求"整屏截图"**
有些课程要求看到整个 IDE。此时不做紧裁,只裁掉窗口外的任务栏与桌面;但要在图注里写明"整屏截图",
并保证图里字可读——必要时把整屏图放到附录,正文仍用裁切图。

**用户已有一份旧报告,只要换图**
走 `edit.md` 的定位-替换:按题号找到旧图段落,替换 `inline_shapes` 的关系目标,不动其他段落与格式。

**代码跑不起来**
不能伪造输出。如实告知哪几题失败、错误信息是什么,报告里这几题写 `【待补充:运行环境问题】`,
并把代码图照常插入(证明代码确实写了)。**禁止编造输出截图。**

## 4. 与其他 workflow 的关系

```
template-fill.md  ── 提供:分析模板、锚点插入、格式继承
generate.md       ── 提供:无模板时的新建骨架
edit.md           ── 提供:只换图的定点修改
references/screenshot-evidence.md  ── 提供:采集与裁切的全部技术细节
references/visual-qa.md            ── 提供:逐页检查的通用清单
scripts/verify_evidence.py         ── 提供:标签/图片对账
```
