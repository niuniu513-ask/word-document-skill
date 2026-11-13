fn make_para(xml : String) -> @ooxml.Paragraph raise @xml.XmlError {
  let el = @xml.parse_element(xml)
  @ooxml.Paragraph::new(el, @ooxml.StyleTable::empty())
}

test "replace inside one run" {
  let para = make_para(
    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>姓名：【待补充】</w:t></w:r></w:p>",
  )
  inspect(@docx.replace_text(para, "【待补充】", "张三"), content="1")
  inspect(para.text(), content="姓名：张三")
}

test "replace across split runs keeps one replacement" {
  let para = make_para(
    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:rPr><w:b/></w:rPr><w:t>【待补</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>充】</w:t></w:r></w:p>",
  )
  inspect(@docx.replace_text(para, "【待补充】", "张三"), content="1")
  inspect(para.text(), content="张三")
  inspect(para.runs().length(), content="1")
}

test "replace keeps surrounding text of multi-run match" {
  let para = make_para(
    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>前【待</w:t></w:r><w:r><w:t>补充】后</w:t></w:r></w:p>",
  )
  inspect(@docx.replace_text(para, "【待补充】", "值"), content="1")
  inspect(para.text(), content="前值后")
}

test "replace ignores missing needle" {
  let para = make_para(
    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>正文</w:t></w:r></w:p>",
  )
  inspect(@docx.replace_text(para, "【待补充】", "值"), content="0")
  inspect(para.text(), content="正文")
}

test "set_paragraph_text keeps first run formatting" {
  let para = make_para(
    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:rPr><w:b/></w:rPr><w:t>旧</w:t></w:r><w:r><w:t>文</w:t></w:r></w:p>",
  )
  @docx.set_paragraph_text(para, "新文本")
  inspect(para.text(), content="新文本")
  let runs = para.runs()
  inspect(runs.length(), content="1")
  inspect(runs[0].element().child(@xml.ns_w, "rPr").is_empty(), content="false")
}

test "two placeholders in one paragraph" {
  let para = make_para(
    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>A【x】B【x】C</w:t></w:r></w:p>",
  )
  inspect(@docx.replace_text(para, "【x】", "-"), content="2")
  inspect(para.text(), content="A-B-C")
}
