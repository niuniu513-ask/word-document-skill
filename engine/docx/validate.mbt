/// Formatting drift detection.
///
/// The whole point of editing inside a document's own formatting system is that
/// the *format* must not move: after filling placeholders, every paragraph must
/// still carry the same style, indents, spacing and alignment. This module turns
/// that promise into a checkable diff, so an edit either proves itself clean or
/// reports exactly which property drifted and where.

pub struct Drift {
  kind : String
  location : String
  field : String
  before : String
  after : String
}

fn opt_int(value : Int?) -> String {
  match value {
    Some(v) => v.to_string()
    None => "-"
  }
}

fn opt_str(value : String?) -> String {
  match value {
    Some(v) => v
    None => "-"
  }
}

fn opt_bool(value : Bool?) -> String {
  match value {
    Some(v) => (if v { "on" } else { "off" })
    None => "-"
  }
}

fn attr_or_dash(el : @xml.Element, name : String) -> String {
  match el.attr(name) {
    Some(v) => v
    None => "-"
  }
}

fn para_pairs(para : @ooxml.Paragraph) -> Array[(String, String)] {
  let p = para.props()
  [
    ("style", opt_str(para.style_id())),
    ("heading_level", opt_int(para.heading_level())),
    ("alignment", opt_str(p.alignment)),
    ("outline_level", opt_int(p.outline_level)),
    ("left_twips", opt_int(p.left_twips)),
    ("right_twips", opt_int(p.right_twips)),
    ("first_line_twips", opt_int(p.first_line_twips)),
    ("hanging_twips", opt_int(p.hanging_twips)),
    ("space_before_twips", opt_int(p.space_before_twips)),
    ("space_after_twips", opt_int(p.space_after_twips)),
    ("line_twips", opt_int(p.line_twips)),
    ("line_rule", opt_str(p.line_rule)),
    ("num_id", opt_str(p.num_id)),
    ("ilvl", opt_int(p.ilvl)),
    ("keep_next", opt_bool(p.keep_next)),
    ("page_break_before", opt_bool(p.page_break_before)),
  ]
}

fn page_pairs(doc : @ooxml.Docx) -> Array[(String, String)] {
  let out : Array[(String, String)] = []
  match doc.body().child(@xml.ns_w, "sectPr") {
    None => ()
    Some(sec) =>
      for child in sec.children_in(@xml.ns_w) {
        match child.local() {
          "pgSz" => {
            out.push(("page.width_twips", attr_or_dash(child, "w")))
            out.push(("page.height_twips", attr_or_dash(child, "h")))
            out.push(("page.orientation", attr_or_dash(child, "orient")))
          }
          "pgMar" => {
            out.push(("page.margin_top", attr_or_dash(child, "top")))
            out.push(("page.margin_bottom", attr_or_dash(child, "bottom")))
            out.push(("page.margin_left", attr_or_dash(child, "left")))
            out.push(("page.margin_right", attr_or_dash(child, "right")))
            out.push(("page.header", attr_or_dash(child, "header")))
            out.push(("page.footer", attr_or_dash(child, "footer")))
            out.push(("page.gutter", attr_or_dash(child, "gutter")))
          }
          "cols" => out.push(("page.columns", attr_or_dash(child, "num")))
          "pgNumType" => {
            out.push(("page.number_format", attr_or_dash(child, "fmt")))
            out.push(("page.number_start", attr_or_dash(child, "start")))
          }
          _ => ()
        }
      }
  }
  out
}

fn defaults_pairs(doc : @ooxml.Docx) -> Array[(String, String)] {
  let styles = doc.styles()
  let run = styles.doc_run()
  let para = styles.doc_para()
  [
    ("defaults.font_ascii", opt_str(run.font_ascii)),
    ("defaults.font_east_asia", opt_str(run.font_east_asia)),
    ("defaults.size_pt", opt_str(run.size_pt.map(fn(v : Double) { @ooxml.fmt(v, 2) }))),
    ("defaults.alignment", opt_str(para.alignment)),
    ("defaults.style", opt_str(styles.default_paragraph_style())),
  ]
}

fn cmp_pairs(
  kind : String,
  location : String,
  before : Array[(String, String)],
  after : Array[(String, String)],
  drift : Array[Drift],
) -> Unit {
  let keys : Array[String] = []
  for pair in before {
    if !keys.contains(pair.0) {
      keys.push(pair.0)
    }
  }
  for pair in after {
    if !keys.contains(pair.0) {
      keys.push(pair.0)
    }
  }
  for key in keys {
    let mut b = ""
    let mut a = ""
    for pair in before {
      if pair.0 == key {
        b = pair.1
      }
    }
    for pair in after {
      if pair.0 == key {
        a = pair.1
      }
    }
    if b != a {
      drift.push(Drift::{ kind, location, field: key, before: b, after: a })
    }
  }
}

fn table_pairs(table : @ooxml.Table) -> Array[(String, String)] {
  [
    ("rows", table.row_count().to_string()),
    ("cols", table.col_count().to_string()),
    ("style", opt_str(table.style_id())),
  ]
}

fn drift_json(drift : Array[Drift], limit : Int) -> @json.Json {
  let items : Array[@json.Json] = []
  let mut shown = 0
  for d in drift {
    if shown >= limit {
      break
    }
    items.push(
      jobj([
        ("kind", jstr(d.kind)),
        ("where", jstr(d.location)),
        ("field", jstr(d.field)),
        ("before", jstr(d.before)),
        ("after", jstr(d.after)),
      ]),
    )
    shown = shown + 1
  }
  jarr(items)
}

/// Compares two documents and reports every formatting property that moved.
///
/// Text differences are deliberately ignored: replacing a placeholder is the
/// intended change. Everything that gives the document its look — styles,
/// indents, spacing, page geometry, table shape, document defaults — is compared.
pub fn format_delta(before : @ooxml.Docx, after : @ooxml.Docx) -> @json.Json {
  let drift : Array[Drift] = []
  cmp_pairs("page", "document", page_pairs(before), page_pairs(after), drift)
  cmp_pairs(
    "defaults",
    "document",
    defaults_pairs(before),
    defaults_pairs(after),
    drift,
  )
  let paras_before = before.body_paragraphs()
  let paras_after = after.body_paragraphs()
  let mut i = 0
  while i < paras_before.length() && i < paras_after.length() {
    cmp_pairs(
      "paragraph",
      "body paragraph " + (i + 1).to_string(),
      para_pairs(paras_before[i]),
      para_pairs(paras_after[i]),
      drift,
    )
    i = i + 1
  }
  let tables_before = before.tables()
  let tables_after = after.tables()
  let mut t = 0
  while t < tables_before.length() && t < tables_after.length() {
    cmp_pairs(
      "table",
      "table " + (t + 1).to_string(),
      table_pairs(tables_before[t]),
      table_pairs(tables_after[t]),
      drift,
    )
    t = t + 1
  }
  let added = paras_after.length() - paras_before.length()
  let table_delta = tables_after.length() - tables_before.length()
  let summary = if drift.is_empty() {
    "formatting delta 0: no drift detected"
  } else {
    drift.length().to_string() + " formatting properties drifted"
  }
  jobj([
    ("ok", jbool(drift.is_empty())),
    ("summary", jstr(summary)),
    (
      "checked",
      jobj([
        ("paragraphs", jint(paras_before.length())),
        ("tables", jint(tables_before.length())),
        ("sections", jint(page_count(before))),
      ]),
    ),
    (
      "structure",
      jobj([
        ("paragraph_delta", jint(added)),
        ("table_delta", jint(table_delta)),
      ]),
    ),
    ("drift_count", jint(drift.length())),
    ("drift", drift_json(drift, 40)),
  ])
}

fn page_count(doc : @ooxml.Docx) -> Int {
  let sections = doc.root().find_all_w("sectPr")
  if sections.is_empty() {
    0
  } else {
    sections.length()
  }
}
