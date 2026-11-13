/// `docx` — command line front end for the docx-mbt engine.

let commands : Array[String] = [
  "info",
  "text",
  "outline",
  "parts",
  "styles",
  "paragraphs",
  "tables",
  "analyze",
  "issues",
  "fill",
  "validate",
  "help",
]

fn usage() -> String {
  [
    "docx-mbt - a pure-MoonBit Word (.docx) engine",
    "",
    "usage: docx <command> <file.docx>",
    "",
    "commands:",
    "  info        package and document summary",
    "  text        plain text of the document",
    "  outline     heading structure with levels",
    "  parts       list package parts",
    "  styles      list the styles defined in the document",
    "  paragraphs  numbered paragraph dump with style ids",
    "  tables      table dimensions and first row",
    "  analyze     full Document Specification as JSON (--out FILE)",
    "  issues      only the findings, human readable",
    "  fill        replace placeholders and write a new file",
    "              fill in.docx --set \"【待补充】=value\" --out out.docx",
    "  validate    compare formatting of two files",
    "              validate before.docx after.docx [--json]",
    "  help        this message",
  ].join("\n")
}

fn main {
  let argv = @env.args()
  let mut command = ""
  let rest : Array[String] = []
  let mut i = 0
  let mut found = false
  while i < argv.length() {
    let arg = argv[i]
    if !found && commands.contains(arg) {
      command = arg
      found = true
    } else if found {
      rest.push(arg)
    }
    i = i + 1
  }
  if !found || command == "help" || rest.is_empty() {
    println(usage())
    return
  }
  if command == "validate" {
    run_validate(rest)
    return
  }
  let path = rest[0]
  let bytes = try {
    @fs.read_file_to_bytes(path)
  } catch {
    _ => {
      println("error: cannot read file " + path)
      return
    }
  }
  let doc = try {
    @ooxml.Docx::open(bytes)
  } catch {
    @ooxml.DocxError(msg) => {
      println("error: " + msg)
      return
    }
  }
  match command {
    "info" => print_info(doc, path, bytes.length())
    "text" => println(doc.text())
    "outline" => print_outline(doc)
    "parts" => {
      for part in doc.package().parts() {
        println("  " + part)
      }
    }
    "styles" => print_styles(doc)
    "paragraphs" => print_paragraphs(doc)
    "tables" => print_tables(doc)
    "analyze" => print_analyze(doc, rest)
    "issues" => print_issues(doc)
    "fill" => print_fill(doc, rest)
    _ => println(usage())
  }
}

fn print_info(doc : @ooxml.Docx, path : String, size : Int) -> Unit {
  println("file          " + path)
  println("size          " + size.to_string() + " bytes")
  println("parts         " + doc.package().part_count().to_string())
  println("paragraphs    " + doc.paragraphs().length().to_string())
  println("body blocks   " + doc.blocks().length().to_string())
  println("tables        " + doc.tables().length().to_string())
  println("styles        " + doc.styles().count().to_string())
  println("characters    " + doc.text().length().to_string())
}

fn print_outline(doc : @ooxml.Docx) -> Unit {
  for block in doc.blocks() {
    match block {
      @ooxml.Block::Para(para) =>
        match para.heading_level() {
          Some(level) => {
            let pad = StringBuilder()
            let mut i = 1
            while i < level {
              pad.write_string("  ")
              i = i + 1
            }
            println(pad.to_string() + "H" + level.to_string() + " " + para.text())
          }
          None => ()
        }
      _ => ()
    }
  }
}

fn print_styles(doc : @ooxml.Docx) -> Unit {
  for style in doc.styles().all() {
    let mut note = ""
    match style.heading_level() {
      Some(level) => note = "  (heading level " + level.to_string() + ")"
      None => ()
    }
    println(
      "  " +
      style.kind +
      "  " +
      style.id +
      "  =  " +
      style.name +
      note,
    )
  }
}

fn print_paragraphs(doc : @ooxml.Docx) -> Unit {
  for i, para in doc.paragraphs() {
    let style = match para.style_id() {
      Some(id) => id
      None => "-"
    }
    let indent = if para.left_twips() != 0 {
      " indent=" + para.left_twips().to_string()
    } else {
      ""
    }
    let bullet = if para.is_list_item() { " list" } else { "" }
    println(
      "[" +
      i.to_string() +
      "] {" +
      style +
      indent +
      bullet +
      "} " +
      para.text(),
    )
  }
}

fn print_tables(doc : @ooxml.Docx) -> Unit {
  for i, table in doc.tables() {
    println(
      "table " +
      i.to_string() +
      ": " +
      table.row_count().to_string() +
      " rows x " +
      table.col_count().to_string() +
      " cols",
    )
    let grid = table.grid()
    if !grid.is_empty() {
      println("  header: " + grid[0].join(" | "))
    }
  }
}

fn print_analyze(doc : @ooxml.Docx, rest : Array[String]) -> Unit {
  let mut out_path = ""
  let mut i = 1
  while i < rest.length() {
    if rest[i] == "--out" && i + 1 < rest.length() {
      out_path = rest[i + 1]
    }
    i = i + 1
  }
  let spec = @docx.analyze(doc, @docx.AnalyzeOptions::default())
  let text = spec.stringify(indent=2)
  if out_path == "" {
    println(text)
  } else {
    try {
      @fs.write_string_to_file(out_path, text)
      println("wrote " + out_path)
    } catch {
      _ => println("error: cannot write " + out_path)
    }
  }
}

fn print_issues(doc : @ooxml.Docx) -> Unit {
  let spec = @docx.analyze(doc, @docx.AnalyzeOptions::default())
  match spec.value("issues") {
    Some(issues) =>
      match issues.as_array() {
        Some(items) =>
          if items.is_empty() {
            println("no issues found")
          } else {
            for item in items {
              let severity = match item.value("severity") {
                Some(v) => v.as_string().unwrap_or("info")
                None => "info"
              }
              let message = match item.value("message") {
                Some(v) => v.as_string().unwrap_or("")
                None => ""
              }
              println("[" + severity + "] " + message)
            }
          }
        None => println("no issues found")
      }
    None => println("no issues found")
  }
}

fn split_pair(arg : String) -> (String, String)? {
  let chars = arg.to_array()
  let mut i = 0
  while i < chars.length() {
    if chars[i] == '=' {
      let key = StringBuilder()
      let mut k = 0
      while k < i {
        key.write_char(chars[k])
        k = k + 1
      }
      let value = StringBuilder()
      let mut v = i + 1
      while v < chars.length() {
        value.write_char(chars[v])
        v = v + 1
      }
      return Some((key.to_string(), value.to_string()))
    }
    i = i + 1
  }
  None
}

fn json_str(v : @json.Json?, fallback : String) -> String {
  match v {
    Some(value) =>
      match value.as_string() {
        Some(s) => s
        None => fallback
      }
    None => fallback
  }
}

fn json_int(v : @json.Json?, fallback : Int) -> Int {
  match v {
    Some(value) =>
      match value.as_number() {
        Some(n) => n.to_int()
        None => fallback
      }
    None => fallback
  }
}

fn print_fill(doc : @ooxml.Docx, rest : Array[String]) -> Unit {
  let pairs : Array[(String, String)] = []
  let mut out_path = ""
  let mut i = 1
  while i < rest.length() {
    if rest[i] == "--set" && i + 1 < rest.length() {
      match split_pair(rest[i + 1]) {
        Some(pair) => pairs.push(pair)
        None => println("error: --set expects KEY=VALUE")
      }
      i = i + 2
    } else if rest[i] == "--out" && i + 1 < rest.length() {
      out_path = rest[i + 1]
      i = i + 2
    } else {
      i = i + 1
    }
  }
  if pairs.is_empty() {
    println("nothing to do: pass at least one --set \"KEY=VALUE\"")
    return
  }
  let count = @docx.fill(doc, pairs)
  println("replaced " + count.to_string() + " occurrence(s)")
  if out_path == "" {
    println("(no --out given, nothing written)")
    return
  }
  let bytes = try {
    doc.save()
  } catch {
    @ooxml.DocxError(msg) => {
      println("error: " + msg)
      return
    }
  }
  try {
    @fs.write_bytes_to_file(out_path, bytes)
    println("wrote " + out_path)
  } catch {
    _ => println("error: cannot write " + out_path)
  }
}

fn run_validate(rest : Array[String]) -> Unit {
  let paths : Array[String] = []
  let mut as_json = false
  for arg in rest {
    if arg == "--json" {
      as_json = true
    } else if !arg.starts_with("--") {
      paths.push(arg)
    }
  }
  if paths.length() < 2 {
    println("usage: docx validate before.docx after.docx [--json]")
    return
  }
  let before = try {
    @ooxml.Docx::open(@fs.read_file_to_bytes(paths[0]))
  } catch {
    _ => {
      println("error: cannot open " + paths[0])
      return
    }
  }
  let after = try {
    @ooxml.Docx::open(@fs.read_file_to_bytes(paths[1]))
  } catch {
    _ => {
      println("error: cannot open " + paths[1])
      return
    }
  }
  let delta = @docx.format_delta(before, after)
  if as_json {
    println(delta.stringify(indent=2))
    return
  }
  println("before  " + paths[0])
  println("after   " + paths[1])
  println("summary " + json_str(delta.value("summary"), ""))
  let drift_count = json_int(delta.value("drift_count"), 0)
  if drift_count == 0 {
    println("formatting delta 0 - the edit did not move any formatting")
    return
  }
  println("drift   " + drift_count.to_string() + " change(s)")
  match delta.value("drift") {
    Some(items) =>
      match items.as_array() {
        Some(list) =>
          for item in list {
            println(
              "  " +
              json_str(item.value("field"), "?") +
              ": " +
              json_str(item.value("before"), "?") +
              " -> " +
              json_str(item.value("after"), "?") +
              "   (" +
              json_str(item.value("where"), "") +
              ")",
            )
          }
        None => ()
      }
    None => ()
  }
}
