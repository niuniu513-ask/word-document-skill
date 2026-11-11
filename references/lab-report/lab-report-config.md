# 严格模板报告配置

`scripts/lab_report.py build` 的 JSON 是显式映射，不自动猜题号，不生成分析。所有路径填写绝对路径；所有索引从 **0** 开始，cell 索引是原 notebook 中的实际索引，包含前面的 markdown 单元格。以下是最小可运行形状，字段值须由模板分析确定。

```json
{
  "template": "D:/course/evidence/template.docx",
  "notebook": "D:/course/evidence/executed.ipynb",
  "output": "D:/course/report.docx",
  "container": {"kind":"cell", "table":0, "row":0, "column":0},
  "keep_prefix": 53,
  "cover_edits": [
    {"paragraph":2, "run":5, "expected":"   ", "text":"  实际班级   "}
  ],
  "prototypes": {"heading":53,"caption":54,"analysis":57,"section":77,"question":78},
  "clear_instruction_format": ["analysis"],
  "screenshots": "D:/course/evidence/crops/screenshots-cropped.json",
  "max_width_cm":15.1,
  "pixels_per_cm":55,
  "max_height_cm":9,
  "groups":[
    {"title":"1.1加载数据集", "items":[
      {"cell":0,"subtitle":"导入库", "analysis":"从真实输出填写观察、数值和原因。"},
      {"cell":1,"analysis":"本题的实际结果分析。"}
    ]}
  ],
  "questions":[
    {"section":"思考题", "template_paragraph":78, "answer":"对原题的完整回答。"},
    {"text":"指导书中的下一道思考题原文", "answer":"依据实验回答。"}
  ]
}
```

## 字段含义与边界

- `container.kind=cell`：在原单格大框里写，保留其 `tblPr`、原节设置、封面及原样式；无新内横线。**本脚本仅适用纯段落单格正文**。原大框有嵌套数据表、合并布局、锚定图形或复杂文本框时，先适配并验证，不能强套本脚本。
- `container.kind=body`：普通正文模板；脚本将保留 `keep_prefix` 个正文段落并追加内容。正文夹杂表格等非段落对象时应使用原有锚点填充工具，不能用它清空混合正文。
- `keep_prefix`：从容器开头保留多少段。例如结果标题是第52段，保留53段，后面的示例截图和答案占位区域由映射重建。该数字**不通用于其他模板**。
- `prototypes`：从原容器中复制对应段落属性与首个文字运行属性，所有索引需在删除占位区前测定。不同字号的代码标签与输出标签可扩展专用角色，勿改整篇字体。
- `clear_instruction_format`：只有被识别为填写提示的红字/黄色高亮样例，才在**新增回答**里清除 `color/highlight/shd`；不改原模板固定文字。使用正文原型更稳妥。
- `cover_edits.expected`：精确核对原run内容，错位直接失败；填值保留原run的下划线。填值后可清空同字段其他空白run，但不要清空标签run；按原总宽度调长度并检查渲染。
- `groups`：按**模板章节顺序**组织，允许执行顺序与书写顺序不同，例如先划分数据避免泄漏，但模板先写标准化。每个证据cell恰好使用一次；导入等公共cell放在相关第一小节，不能另造顶级章节。
- `questions`：模板已有题用 `template_paragraph` 原样复制，指导书额外必答题使用 `text`。不得拿模板中的第一题当作只需回答一题。不得默认增加模板没有的“实验目的/环境/总结”。
- `pixels_per_cm` 是插图起始尺度，不是画面真伪证明。55是本案例试用值；须按本次VS Code字号、缩放、DPR和PDF可读性调整。长行应在写代码阶段自然换行，不能事后缩小到无法阅读。
- `max_width_cm` 按大框**内可用宽**测得，不能超过它；`max_height_cm` 要给同页标签留空间。长图应重新采集完整分段或调整真实UI，勿硬裁。
- `image_mode`：默认（不写）为 `split`，每个单元格插"代码截图 + 实验结果截图"两张；写 `joint` 则每个单元格只插一张代码+输出连体图（清单里的 `cellCrop`），标签用 `joint_label` 覆盖，默认 `代码与运行结果截图：`。`audit` 的期望标签数会跟着模式变，不需要手改。

## 支持文件

`capture-plan.json`：原执行notebook路径/哈希、cell索引、原source、执行计数、是否有输出、证据副本路径。

`screenshots.json`：在上述记录中增加真实截图路径、当前窗口CSS尺寸/DPR、同帧代码/输出DOM边界。允许同一张原屏分别裁代码和输出，不能拼接不同执行的输出。

`screenshotScale: {"x":1.0416666,"y":1.0416666}` 是明确校准的CSS到原屏像素比例，优先于PNG尺寸换算。Electron截图可能在真实UI之外补白，不能把补白算入比例。此数值必须由当前原屏校准；例值不是通用常量。

`screenshots-cropped.json`：再增加实际像素裁切框、原屏与裁图SHA256。默认保守裁切，不按深色阈值删淡色代码。离屏渲染生成的清单没有 `evidence_notebook`（真实证据就是执行后的 notebook 本身），`audit` 改用 `sourceFirstLine` / `source` 与 notebook 逐字对账；两种清单都能过审计，不需要手工补副本。

`report.build.json`：模板/配置哈希、使用cell、图片数。`audit.json`：固定文字/节设置/原框属性/源码与输出/图片哈希/标签图片配对检查。它们均**不代替逐页视觉检查**。
