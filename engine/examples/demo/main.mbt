/// End-to-end demo: analyze a document, fill a placeholder, prove the
/// formatting did not move.
///
/// Run it with
///     moon run examples/demo --target js
/// or point it at your own file
///     moon run examples/demo --target js -- report.docx

fn main {
  // argv is [runtime, compiled script, user arguments...] on the js target.
  let args = @env.args()
  let mut path = "tests/fixtures/mini.docx"
  if args.length() > 2 {
    path = args[2]
  }
  let bytes = try {
    @fs.read_file_to_bytes(path)
  } catch {
    _ => {
      println("cannot read " + path)
      return
    }
  }
  let before = try {
    @ooxml.Docx::open(bytes)
  } catch {
    @ooxml.DocxError(msg) => {
      println("error: " + msg)
      return
    }
  }
  println("== " + path)
  println("parts      " + before.package().part_count().to_string())
  println("paragraphs " + before.body_paragraphs().length().to_string())
  println("tables     " + before.tables().length().to_string())
  println("")
  println("== outline")
  for block in before.blocks() {
    match block {
      @ooxml.Block::Para(para) =>
        match para.heading_level() {
          Some(level) => println("  H" + level.to_string() + "  " + para.text())
          None => ()
        }
      _ => ()
    }
  }
  let spec = @docx.analyze(before, @docx.AnalyzeOptions::default())
  println("")
  println("== issues")
  match spec.value("issues") {
    Some(issues) =>
      match issues.as_array() {
        Some(items) =>
          for item in items {
            println(
              "  [" +
              item.value("severity").unwrap().as_string().unwrap() +
              "] " +
              item.value("message").unwrap().as_string().unwrap(),
            )
          }
        None => ()
      }
    None => ()
  }
  // A second copy, so the before/after diff has something to compare.
  let after = try {
    @ooxml.Docx::open(bytes)
  } catch {
    @ooxml.DocxError(msg) => {
      println("error: " + msg)
      return
    }
  }
  let filled = @docx.fill(after, [("【待补充】", "STC89C52")])
  println("")
  println("== fill")
  println("  replaced " + filled.to_string() + " placeholder(s)")
  let delta = @docx.format_delta(before, after)
  match delta.value("summary") {
    Some(summary) => println("  " + summary.as_string().unwrap())
    None => ()
  }
  match delta.value("drift_count") {
    Some(count) =>
      if count.as_number().unwrap().to_int() == 0 {
        println("  formatting untouched: safe to ship")
      }
    None => ()
  }
  let out = try {
    after.save()
  } catch {
    @ooxml.DocxError(msg) => {
      println("error: " + msg)
      return
    }
  }
  try {
    @fs.write_bytes_to_file("out/demo-filled.docx", out)
    println("  wrote out/demo-filled.docx")
  } catch {
    _ => println("  (cannot write out/demo-filled.docx)")
  }
}
