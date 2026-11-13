/// End-to-end checks against a real .docx package.
///
/// The fixture is produced by `scripts/make_fixture.py` and its numbers are
/// fixed, so a regression in the parser, the analyzer or the writer shows up
/// here instead of in a user's document.

let fixture = "tests/fixtures/mini.docx"

fn read_fixture() -> Bytes {
  try {
    @fs.read_file_to_bytes(fixture)
  } catch {
    _ => b""
  }
}

fn open_fixture() -> @ooxml.Docx raise @ooxml.DocxError {
  @ooxml.Docx::open(read_fixture())
}

test "opens the package and sees the parts" {
  let doc = open_fixture()
  inspect(doc.package().part_count() >= 12, content="true")
  inspect(doc.package().has_part("word/document.xml"), content="true")
  inspect(doc.package().has_part("word/styles.xml"), content="true")
}

test "document counts match the fixture" {
  let doc = open_fixture()
  inspect(doc.body_paragraphs().length(), content="7")
  inspect(doc.blocks().length(), content="8")
  inspect(doc.tables().length(), content="1")
  inspect(doc.paragraphs().length(), content="13")
}

test "outline finds the three numbered headings" {
  let doc = open_fixture()
  let mut headings = 0
  for para in doc.body_paragraphs() {
    match para.heading_level() {
      Some(1) => headings = headings + 1
      _ => ()
    }
  }
  inspect(headings, content="3")
}

test "analyzer reports placeholders in body and in table cells" {
  let doc = open_fixture()
  let spec = @docx.analyze(doc, @docx.AnalyzeOptions::default())
  let placeholders = spec.value("placeholders").unwrap().as_array().unwrap()
  inspect(placeholders.length(), content="2")
  let mut in_table = 0
  for item in placeholders {
    if item.value("where").unwrap().as_string().unwrap() == "table" {
      in_table = in_table + 1
    }
  }
  inspect(in_table, content="1")
}

test "analyzer raises the numbering and placeholder findings" {
  let doc = open_fixture()
  let spec = @docx.analyze(doc, @docx.AnalyzeOptions::default())
  let issues = spec.value("issues").unwrap().as_array().unwrap()
  let kinds : Array[String] = []
  for issue in issues {
    kinds.push(issue.value("kind").unwrap().as_string().unwrap())
  }
  inspect(kinds.contains("manual_heading_numbering"), content="true")
  inspect(kinds.contains("unfilled_placeholders"), content="true")
  inspect(kinds.contains("empty_paragraph_runs"), content="true")
}

test "fill then save produces a readable document with the value in place" {
  let doc = open_fixture()
  let count = @docx.fill(doc, [("【待补充】", "STC89C52")])
  inspect(count, content="1")
  let bytes = try {
    doc.save()
  } catch {
    @ooxml.DocxError(msg) => {
      fail(msg)
    }
  }
  let again = @ooxml.Docx::open(bytes)
  inspect(again.text().contains("STC89C52"), content="true")
  inspect(again.tables()[0].grid()[1][2], content="STC89C52")
}

test "saved output keeps its wordprocessing namespace" {
  let doc = open_fixture()
  ignore(@docx.replace_in_document(doc, "【待补充】", "X"))
  let bytes = try {
    doc.save()
  } catch {
    @ooxml.DocxError(msg) => {
      fail(msg)
    }
  }
  let xml = try {
    @ooxml.WordPackage::open(bytes).part_text("word/document.xml").unwrap()
  } catch {
    _ => ""
  }
  inspect(xml.contains("<w:t"), content="true")
  inspect(xml.contains("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""), content="true")
}

test "format delta is zero for a pure text replacement" {
  let before = open_fixture()
  let after = open_fixture()
  ignore(@docx.replace_in_document(after, "【待补充】", "STC89C52"))
  let delta = @docx.format_delta(before, after)
  inspect(delta.value("drift_count").unwrap().as_number().unwrap().to_int(), content="0")
}
