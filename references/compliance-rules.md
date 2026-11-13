# 规则包字段说明

规则包是 `rules/*.json`,由 `scripts/review.py --profile` 加载。一份规则包描述
"这类文档交上去必须满足什么",与检查代码解耦:换单位、换文种只改 JSON。

```jsonc
{
  "id": "railway-plan",              // 文件名即 id,--profile 用它
  "name": "铁路施工方案(报审自查)",     // 报告里显示的名字
  "description": "一句话说明适用场景",
  "required_sections": [              // 必备章节;缺一个即阻断
    { "key": "编制依据", "aliases": ["编制依据", "依据", "编制说明"] }
  ],
  "heading": {
    "max_level": 3,                   // 超过该层级的标题算提示项
    "numbering": "auto-or-chinese",   // any | auto | auto-or-chinese
    "forbid_level_jump": true         // 一级直接跳到三级算严重
  },
  "body": {
    "allowed_sizes_pt": [10.5, 12, 14, 16],   // 正文字号白名单,空数组=不检查
    "allowed_east_asia_fonts": ["仿宋", "宋体"],// 中文字体白名单,支持子串
    "max_consecutive_empty": 2,               // 连续空段落上限
    "min_body_paragraphs": 10                 // 正文段数下限
  },
  "page": {
    "require_page_number": true,      // 页脚/页眉必须有 PAGE 域
    "require_header": false,
    "allowed_page_sizes": ["A4"]      // A4 | A3 | B5
  },
  "placeholders": {
    "patterns": ["【[^】]{0,60}】", "_{4,}"],   // 正则,不是字面量
    "forbid": true
  },
  "forbidden_phrases": ["待定", "TBD"],         // 正文里不该出现的字眼
  "tables": { "require_header_bold": false, "max_cols": 12 }
}
```

## 编写建议

- **别名要覆盖真实写法**。标题常写成"编制依据/编制说明/依据",只写一个会误报缺失。
- **白名单宁可宽一点**。字号白名单少写一个,整篇都会被判"字号不符",报告可信度立刻下降。
- **不要用规则包表达审美**。"应该用三号字"属单位模板规定,"看起来更舒服"不是规则。
- **正则要短**。占位符匹配靠 `【…】`、连续下划线两类就够;过宽会把正文里的方括号误判成占位符。

## 校验

改完规则包直接跑:

```bash
python scripts/review.py tests/fixtures/sample.docx --profile 你的id --json /tmp/r.json
pytest tests/test_review.py -q          # 现有用例会覆盖规则加载与基本判定
```
