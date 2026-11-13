/// Tests for the document-specification helpers.

test "placeholder spans pick up Chinese template markers" {
  let spans = @docx.placeholder_spans("【待补充:验收情况】")
  inspect(spans.length(), content="1")
  inspect(spans[0], content="【待补充:验收情况】")
}

test "placeholder spans pick up underscore blanks" {
  let spans = @docx.placeholder_spans("姓名：________ 学号：______")
  inspect(spans.length(), content="2")
}

test "plain text has no placeholders" {
  let spans = @docx.placeholder_spans("这是一段正文")
  inspect(spans.length(), content="0")
}

test "manual numbering detects Chinese ordinals" {
  inspect(@docx.manual_number_prefix("一 方案描述"), content="Some(一)")
  inspect(@docx.manual_number_prefix("二、作品需要的材料"), content="Some(二)")
}

test "manual numbering detects arabic prefixes" {
  inspect(@docx.manual_number_prefix("3. 仿真过程"), content="Some(3.)")
}

test "manual numbering ignores ordinary text" {
  inspect(@docx.manual_number_prefix("基于C51的系统"), content="None")
}

test "clip truncates on character boundaries" {
  inspect(@docx.clip("abcdef", 3), content="abc…")
  inspect(@docx.clip("ab", 3), content="ab")
}
