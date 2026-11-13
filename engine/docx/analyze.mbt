/// Builds the Document Specification: a JSON "radiograph" of a .docx file.
///
/// The specification answers the questions you must answer before you are
/// allowed to touch a document: what page grid is it on, which styles are
/// declared, which paragraphs are headings, where are the tables and the
/// pictures, what is still a placeholder, and what is inconsistent.

/// Analysis knobs.
pub struct AnalyzeOptions {
  max_paragraphs : Int
}

/// Default analysis settings.
pub fn AnalyzeOptions::default() -> AnalyzeOptions {
  AnalyzeOptions::{ max_paragraphs: 120 }
}

/// Shortens text to at most `limit` characters, appending an ellipsis.
pub fn clip(text : String, limit : Int) -> String {
  let chars = text.to_array()
  if chars.length() <= limit {
    return text
  }
  let sb = StringBuilder()
  let mut i = 0
  while i < limit {
    sb.write_char(chars[i])
    i = i + 1
  }
  sb.to_string() + "…"
}

/// One heading found in the document.
priv struct OutlineEntry {
  level : Int
  index : Int
  text : String
  style_id : String?
  manual : String?
  native : Bool
}

fn outline_entries(doc : @ooxml.Docx) -> Array[OutlineEntry] {
  let out : Array[OutlineEntry] = []
  let paras = doc.body_paragraphs()
  for i, para in paras {
    match para.heading_level() {
      Some(level) => {
        let text = para.text()
        out.push(OutlineEntry::{
          level,
          index: i,
          text,
          style_id: para.style_id(),
          manual: manual_number_prefix(text),
          native: para.is_list_item(),
        })
      }
      None => ()
    }
  }
  out
}

fn cm_json(twips : Int?) -> @json.Json {
  match twips {
    Some(v) => jnum(@ooxml.twips_to_cm(v))
    None => jnull()
  }
}

fn document_json(doc : @ooxml.Docx) -> @json.Json {
  jobj([
    ("body_block_count", jint(doc.blocks().length())),
    ("paragraph_count", jint(doc.body_paragraphs().length())),
    ("all_paragraph_count", jint(doc.paragraphs().length())),
    ("table_count", jint(doc.tables().length())),
    ("image_count", jint(image_count(doc))),
    ("section_count", jint(doc.body().find_all_w("sectPr").length())),
    ("part_count", jint(doc.package().part_count())),
    ("character_count", jint(doc.text().length())),
  ])
}

fn image_count(doc : @ooxml.Docx) -> Int {
  let mut count = 0
  for para in doc.paragraphs() {
    count = count + para.element().find_all_ns(@xml.ns_wp, "extent").length()
  }
  count
}

fn doc_defaults_json(doc : @ooxml.Docx) -> @json.Json {
  let run = doc.styles().doc_run()
  let para = doc.styles().doc_para()
  jobj([
    ("ascii", jopt_str(run.font_ascii)),
    ("east_asia", jopt_str(run.font_east_asia)),
    ("hansi", jopt_str(run.font_hansi)),
    ("size_pt", jopt_num(run.size_pt)),
    ("alignment", jopt_str(para.alignment)),
    ("default_style", jopt_str(doc.styles().default_paragraph_style())),
  ])
}

fn sections_json(doc : @ooxml.Docx) -> Array[@json.Json] {
  let out : Array[@json.Json] = []
  let prs = doc.body().find_all_w("sectPr")
  for i, pr in prs {
    let mut width : Int? = None
    let mut height : Int? = None
    let mut orient = "portrait"
    match pr.child(@xml.ns_w, "pgSz") {
      Some(sz) => {
        width = sz.attr_int("w")
        height = sz.attr_int("h")
        match sz.attr("orient") {
          Some(o) =>
            if o == "landscape" {
              orient = "landscape"
            }
          None =>
            match (width, height) {
              (Some(w), Some(h)) =>
                if w > h {
                  orient = "landscape"
                }
              _ => ()
            }
        }
      }
      None => ()
    }
    let margins = match pr.child(@xml.ns_w, "pgMar") {
      Some(m) =>
        jobj([
          ("top", cm_json(m.attr_int("top"))),
          ("bottom", cm_json(m.attr_int("bottom"))),
          ("left", cm_json(m.attr_int("left"))),
          ("right", cm_json(m.attr_int("right"))),
          ("header", cm_json(m.attr_int("header"))),
          ("footer", cm_json(m.attr_int("footer"))),
          ("gutter", cm_json(m.attr_int("gutter"))),
        ])
      None => jnull()
    }
    let page_number = match pr.child(@xml.ns_w, "pgNumType") {
      Some(pn) =>
        jobj([
          ("format", jopt_str(pn.attr("fmt"))),
          ("start", jopt_int(pn.attr_int("start"))),
        ])
      None => jnull()
    }
    out.push(jobj([
      ("index", jint(i)),
      ("width_cm", cm_json(width)),
      ("height_cm", cm_json(height)),
      ("orientation", jstr(orient)),
      ("margins_cm", margins),
      ("page_number", page_number),
      ("columns", jopt_int(match pr.child(@xml.ns_w, "cols") {
        Some(c) => c.attr_int("num")
        None => None
      })),
    ]))
  }
  out
}

fn styles_json(doc : @ooxml.Docx) -> Array[@json.Json] {
  let out : Array[@json.Json] = []
  for style in doc.styles().all() {
    let heading = match style.heading_level() {
      Some(level) => jint(level)
      None => jnull()
    }
    out.push(jobj([
      ("id", jstr(style.id)),
      ("name", jstr(style.name)),
      ("type", jstr(style.kind)),
      ("based_on", jopt_str(style.based_on)),
      ("next", jopt_str(style.next)),
      ("default", jbool(style.is_default)),
      ("heading_level", heading),
      ("font_east_asia", jopt_str(style.run.font_east_asia)),
      ("font_ascii", jopt_str(style.run.font_ascii)),
      ("size_pt", jopt_num(style.run.size_pt)),
      ("bold", match style.run.bold {
        Some(v) => jbool(v)
        None => jnull()
      }),
      ("alignment", jopt_str(style.para.alignment)),
      ("outline_level", jopt_int(style.para.outline_level)),
      ("left_twips", jopt_int(style.para.left_twips)),
    ]))
  }
  out
}

fn outline_json(entries : Array[OutlineEntry]) -> Array[@json.Json] {
  let out : Array[@json.Json] = []
  for entry in entries {
    out.push(jobj([
      ("level", jint(entry.level)),
      ("paragraph_index", jint(entry.index)),
      ("text", jstr(clip(entry.text, 60))),
      ("style", jopt_str(entry.style_id)),
      ("manual_number", jopt_str(entry.manual)),
      ("native_numbering", jbool(entry.native)),
    ]))
  }
  out
}

fn paragraphs_json(doc : @ooxml.Docx, opts : AnalyzeOptions) -> Array[@json.Json] {
  let out : Array[@json.Json] = []
  let paras = doc.body_paragraphs()
  let mut i = 0
  while i < paras.length() && i < opts.max_paragraphs {
    let para = paras[i]
    let props = para.props()
    out.push(jobj([
      ("index", jint(i)),
      ("text", jstr(clip(para.text(), 200))),
      ("style", jopt_str(para.style_id())),
      ("heading_level", match para.heading_level() {
        Some(level) => jint(level)
        None => jnull()
      }),
      ("alignment", jopt_str(props.alignment)),
      ("left_twips", jopt_int(props.left_twips)),
      ("first_line_twips", jopt_int(props.first_line_twips)),
      ("hanging_twips", jopt_int(props.hanging_twips)),
      ("space_before_twips", jopt_int(props.space_before_twips)),
      ("space_after_twips", jopt_int(props.space_after_twips)),
      ("line_twips", jopt_int(props.line_twips)),
      ("num_id", jopt_str(props.num_id)),
      ("empty", jbool(para.is_empty())),
      ("run_count", jint(para.runs().length())),
    ]))
    i = i + 1
  }
  out
}

fn tables_json(doc : @ooxml.Docx) -> Array[@json.Json] {
  let out : Array[@json.Json] = []
  for i, table in doc.tables() {
    let grid = table.grid()
    let header = if grid.is_empty() {
      ""
    } else {
      grid[0].join(" | ")
    }
    out.push(jobj([
      ("index", jint(i)),
      ("rows", jint(table.row_count())),
      ("cols", jint(table.col_count())),
      ("style", jopt_str(table.style_id())),
      ("header", jstr(clip(header, 120))),
    ]))
  }
  out
}

fn images_json(doc : @ooxml.Docx) -> Array[@json.Json] {
  let out : Array[@json.Json] = []
  let paras = doc.paragraphs()
  let mut index = 0
  for i, para in paras {
    let inline = para.element().find_all_ns(@xml.ns_wp, "inline").length()
    for extent in para.element().find_all_ns(@xml.ns_wp, "extent") {
      let kind = if inline > 0 { "inline" } else { "anchored" }
      out.push(jobj([
        ("index", jint(index)),
        ("paragraph_index", jint(i)),
        ("type", jstr(kind)),
        ("width_cm", match extent.attr_int("cx") {
          Some(v) => jnum(@ooxml.emu_to_cm(v))
          None => jnull()
        }),
        ("height_cm", match extent.attr_int("cy") {
          Some(v) => jnum(@ooxml.emu_to_cm(v))
          None => jnull()
        }),
      ]))
      index = index + 1
    }
  }
  out
}

fn placeholders_json(doc : @ooxml.Docx) -> Array[@json.Json] {
  let out : Array[@json.Json] = []
  for i, para in doc.body_paragraphs() {
    let spans = placeholder_spans(para.text())
    if spans.is_empty() {
      continue
    }
    out.push(jobj([
      ("where", jstr("paragraph")),
      ("index", jint(i)),
      ("style", jopt_str(para.style_id())),
      ("text", jstr(clip(para.text(), 80))),
      ("spans", jstr_array(spans)),
    ]))
  }
  // Tables are where templates keep the cover sheet, so they must be scanned too.
  for t, table in doc.tables() {
    let grid = table.grid()
    for r, row in grid {
      for c, cell in row {
        let spans = placeholder_spans(cell)
        if spans.is_empty() {
          continue
        }
        out.push(jobj([
          ("where", jstr("table")),
          ("index", jint(t)),
          ("row", jint(r)),
          ("col", jint(c)),
          ("text", jstr(clip(cell, 80))),
          ("spans", jstr_array(spans)),
        ]))
      }
    }
  }
  out
}

/// Font and size usage across the body.
struct Formatting {
  json : @json.Json
  mixed_fonts : Bool
  run_count : Int
  dominant_font : String?
}

fn formatting_json(doc : @ooxml.Docx) -> Formatting {
  let fonts : Map[String, Int] = Map([])
  let sizes : Map[String, Int] = Map([])
  let colors : Map[String, Int] = Map([])
  let mut run_count = 0
  for para in doc.paragraphs() {
    for run in para.runs() {
      if run.is_textless() {
        continue
      }
      run_count = run_count + 1
      let props = run.props()
      let font = match props.font_east_asia {
        Some(f) => f
        None =>
          match props.font_ascii {
            Some(f) => f
            None => "(inherit)"
          }
      }
      bump(fonts, font)
      match props.size_pt {
        Some(size) => bump(sizes, @ooxml.fmt(size, 1) + "pt")
        None => bump(sizes, "(inherit)")
      }
      match props.color {
        Some(color) => bump(colors, color)
        None => ()
      }
    }
  }
  let font_entries = top_entries(fonts, 6)
  let dominant = if font_entries.is_empty() {
    None
  } else {
    Some(font_entries[0].0)
  }
  let font_json : Array[@json.Json] = []
  for entry in font_entries {
    font_json.push(jobj([("value", jstr(entry.0)), ("run_count", jint(entry.1))]))
  }
  let size_json : Array[@json.Json] = []
  for entry in top_entries(sizes, 6) {
    size_json.push(jobj([("value", jstr(entry.0)), ("run_count", jint(entry.1))]))
  }
  let color_json : Array[@json.Json] = []
  for entry in top_entries(colors, 6) {
    color_json.push(jobj([("value", jstr("#" + entry.0)), ("run_count", jint(entry.1))]))
  }
  Formatting::{
    json: jobj([
      ("run_count", jint(run_count)),
      ("fonts", jarr(font_json)),
      ("sizes", jarr(size_json)),
      ("colors", jarr(color_json)),
      ("dominant_font", jopt_str(dominant)),
    ]),
    mixed_fonts: fonts.size() > 2,
    run_count,
    dominant_font: dominant,
  }
}

fn conventions_json(
  doc : @ooxml.Docx,
  outline : Array[OutlineEntry],
) -> @json.Json {
  let body_counts : Map[String, Int] = Map([])
  for para in doc.body_paragraphs() {
    if para.is_empty() {
      continue
    }
    match para.style_id() {
      Some(id) => bump(body_counts, id)
      None =>
        match doc.styles().default_paragraph_style() {
          Some(id) => bump(body_counts, id)
          None => bump(body_counts, "(default)")
        }
    }
  }
  let body_style = match top_entries(body_counts, 1) {
    [] => jnull()
    [entry, ..] => jstr(entry.0)
  }
  let heading_json : Array[@json.Json] = []
  let mut i = 0
  let mut native = false
  for entry in outline {
    if entry.native {
      native = true
    }
    if i < 12 {
      heading_json.push(jobj([
        ("level", jint(entry.level)),
        ("style", jopt_str(entry.style_id)),
        ("sample", jstr(clip(entry.text, 30))),
      ]))
      i = i + 1
    }
  }
  jobj([
    ("body_style", body_style),
    ("heading_styles", jarr(heading_json)),
    ("uses_native_numbering", jbool(native)),
  ])
}

fn compute_issues(
  doc : @ooxml.Docx,
  outline : Array[OutlineEntry],
  formatting : Formatting,
  section_count : Int,
) -> Array[Issue] {
  let out : Array[Issue] = []
  let paras = doc.body_paragraphs()
  // Runs of two or more empty paragraphs: the usual cause of a page that
  // "jumps" when the document is edited later.
  let mut empty_run = 0
  let mut empty_run_start = 0
  let mut empty_runs = 0
  let mut first_empty : Int? = None
  for i, para in paras {
    if para.is_empty() {
      if empty_run == 0 {
        empty_run_start = i
      }
      empty_run = empty_run + 1
    } else {
      if empty_run >= 2 {
        empty_runs = empty_runs + 1
        match first_empty {
          Some(_) => ()
          None => first_empty = Some(empty_run_start)
        }
      }
      empty_run = 0
    }
  }
  if empty_run >= 2 {
    empty_runs = empty_runs + 1
    match first_empty {
      Some(_) => ()
      None => first_empty = Some(empty_run_start)
    }
  }
  if empty_runs > 0 {
    out.push(Issue::new(
      "empty_paragraph_runs",
      Warn,
      "存在 " + empty_runs.to_string() + " 处连续空段落，删除或改为段前段后间距更稳妥",
      first_empty,
      empty_runs,
    ))
  }
  // Heading levels must not skip a level.
  let mut previous = 0
  for entry in outline {
    if previous > 0 && entry.level > previous + 1 {
      out.push(Issue::new(
        "heading_level_jump",
        Warn,
        "标题层级跳跃：H" +
        previous.to_string() +
        " 之后直接出现 H" +
        entry.level.to_string(),
        Some(entry.index),
        1,
      ))
    }
    previous = entry.level
  }
  // Hand-typed numbering inside headings.
  let mut manual_count = 0
  let mut manual_first : Int? = None
  for entry in outline {
    match entry.manual {
      Some(_) => {
        manual_count = manual_count + 1
        match manual_first {
          Some(_) => ()
          None => manual_first = Some(entry.index)
        }
      }
      None => ()
    }
  }
  if manual_count > 0 {
    out.push(Issue::new(
      "manual_heading_numbering",
      Warn,
      "有 " +
      manual_count.to_string() +
      " 个标题把编号写进了文字（如「一、」），自动编号与目录会因此错乱",
      manual_first,
      manual_count,
    ))
  }
  // Paragraphs that only *look* like headings: centered, bold, short, but
  // carrying a body style. Everything downstream that trusts structure rather
  // than appearance (navigation pane, automatic TOC, outline export) misses them.
  let mut fake_count = 0
  let mut fake_first : Int? = None
  for i, para in paras {
    let text = para.text().trim()
    if text.is_empty() || text.length() > 30 {
      continue
    }
    match para.heading_level() {
      Some(_) => continue
      None => ()
    }
    let centered = match para.props().alignment {
      Some(a) => a == "center"
      None => false
    }
    if !centered {
      continue
    }
    let mut bold = false
    let mut seen = false
    for run in para.runs() {
      if seen || run.is_textless() {
        continue
      }
      seen = true
      match run.props().bold {
        Some(true) => bold = true
        _ => ()
      }
    }
    if !bold || text.length() > 25 {
      continue
    }
    fake_count = fake_count + 1
    match fake_first {
      Some(_) => ()
      None => fake_first = Some(i)
    }
  }
  if fake_count > 0 {
    out.push(Issue::new(
      "fake_heading",
      Warn,
      "有 " +
      fake_count.to_string() +
      " 个段落靠居中加粗伪装成标题（未使用标题样式），目录与导航会漏掉它们",
      fake_first,
      fake_count,
    ))
  }
  // Placeholders that were never filled in.
  let mut placeholder_count = 0
  let mut placeholder_first : Int? = None
  for i, para in paras {
    let spans = placeholder_spans(para.text())
    if !spans.is_empty() {
      placeholder_count = placeholder_count + spans.length()
      match placeholder_first {
        Some(_) => ()
        None => placeholder_first = Some(i)
      }
    }
  }
  for _, table in doc.tables() {
    for _, row in table.grid() {
      for _, cell in row {
        placeholder_count = placeholder_count + placeholder_spans(cell).length()
      }
    }
  }
  if placeholder_count > 0 {
    out.push(Issue::new(
      "unfilled_placeholders",
      Warn,
      "仍有 " + placeholder_count.to_string() + " 处占位符未填写",
      placeholder_first,
      placeholder_count,
    ))
  }
  // Section properties are what make the page grid exist at all.
  if section_count == 0 {
    out.push(Issue::new(
      "missing_section_properties",
      Error,
      "缺少 sectPr，Word 会退回默认页面设置，纸张与页边距不可控",
      None,
      1,
    ))
  }
  // Font churn is the most common reason a report looks assembled.
  if formatting.mixed_fonts {
    out.push(Issue::new(
      "mixed_fonts",
      Info,
      "正文出现多种字体，同一份报告建议统一为一种中西文字体组合",
      None,
      1,
    ))
  }
  if formatting.run_count == 0 {
    out.push(Issue::new("no_body_text", Error, "正文没有任何文字", None, 1))
  }
  out
}

/// Builds the Document Specification for a document.
pub fn analyze(doc : @ooxml.Docx, opts : AnalyzeOptions) -> @json.Json {
  let outline = outline_entries(doc)
  let formatting = formatting_json(doc)
  let sections = sections_json(doc)
  let issues = compute_issues(doc, outline, formatting, sections.length())
  let issue_json : Array[@json.Json] = []
  for issue in issues {
    issue_json.push(issue.to_json())
  }
  jobj([
    ("spec_version", jstr("1.0")),
    ("engine", jstr("docx-mbt")),
    ("document", document_json(doc)),
    ("doc_defaults", doc_defaults_json(doc)),
    ("sections", jarr(sections)),
    ("styles", jarr(styles_json(doc))),
    ("outline", jarr(outline_json(outline))),
    ("paragraphs_preview", jarr(paragraphs_json(doc, opts))),
    ("tables", jarr(tables_json(doc))),
    ("images", jarr(images_json(doc))),
    ("placeholders", jarr(placeholders_json(doc))),
    ("formatting_inventory", formatting.json),
    ("conventions", conventions_json(doc, outline)),
    ("issues", jarr(issue_json)),
  ])
}
