/// Style table: the document's named formatting recipes.

/// One named style from `word/styles.xml`.
pub struct Style {
  id : String
  name : String
  kind : String
  based_on : String?
  next : String?
  is_default : Bool
  para : ParaProps
  run : RunProps
}

/// Heading level implied by a style, 1-based, or `None` for body text.
///
/// Word stores the true signal in `w:outlineLvl`, but older and hand-built
/// documents frequently only name the style "heading 2" / "标题 2".
pub fn Style::heading_level(self : Style) -> Int? {
  match self.para.outline_level {
    Some(level) =>
      if level >= 0 && level < 9 {
        return Some(level + 1)
      }
    None => ()
  }
  match heading_from_name(self.name) {
    Some(level) => Some(level)
    None => heading_from_name(self.id)
  }
}

fn heading_from_name(name : String) -> Int? {
  let lower = name.to_lower()
  let digits = ["1", "2", "3", "4", "5", "6", "7", "8", "9"]
  if lower.has_prefix("heading") || lower.has_prefix("title") || name.has_prefix("标题") {
    for i, d in digits {
      if lower.has_suffix(d) || name.has_suffix(d) {
        return Some(i + 1)
      }
    }
  }
  None
}

/// The document style table plus its document defaults.
pub struct StyleTable {
  styles : Map[String, Style]
  by_name : Map[String, String]
  default_para : String?
  default_char : String?
  default_table : String?
  doc_run : RunProps
  doc_para : ParaProps
}

/// An empty table, used for documents without a styles part.
pub fn StyleTable::empty() -> StyleTable {
  StyleTable::{
    styles: Map([]),
    by_name: Map([]),
    default_para: None,
    default_char: None,
    default_table: None,
    doc_run: RunProps::none(),
    doc_para: ParaProps::none(),
  }
}

/// Parses `word/styles.xml`.
pub fn StyleTable::parse(source : String) -> StyleTable raise @xml.XmlError {
  let root = @xml.parse(source).root()
  let mut table = StyleTable::empty()
  // Document defaults.
  match root.child(@xml.ns_w, "docDefaults") {
    Some(defaults) => {
      match defaults.child(@xml.ns_w, "rPrDefault") {
        Some(rd) =>
          match rd.child(@xml.ns_w, "rPr") {
            Some(rpr) => table = { ..table, doc_run: read_run_props(Some(rpr)) }
            None => ()
          }
        None => ()
      }
      match defaults.child(@xml.ns_w, "pPrDefault") {
        Some(pd) =>
          match pd.child(@xml.ns_w, "pPr") {
            Some(ppr) => table = { ..table, doc_para: read_para_props(Some(ppr)) }
            None => ()
          }
        None => ()
      }
    }
    None => ()
  }
  for el in root.w_children("style") {
    let id = el.attr_or("styleId", "")
    if id == "" {
      continue
    }
    let name = match el.child(@xml.ns_w, "name") {
      Some(n) => n.attr_or("val", id)
      None => id
    }
    let kind = el.attr_or("type", "paragraph")
    let is_default = match el.attr("default") {
      Some(v) => v == "1" || v == "true"
      None => false
    }
    let style = Style::{
      id,
      name,
      kind,
      based_on: match el.child(@xml.ns_w, "basedOn") {
        Some(b) => b.attr("val")
        None => None
      },
      next: match el.child(@xml.ns_w, "next") {
        Some(b) => b.attr("val")
        None => None
      },
      is_default,
      para: read_para_props(el.child(@xml.ns_w, "pPr")),
      run: read_run_props(el.child(@xml.ns_w, "rPr")),
    }
    table.styles.set(id, style)
    table.by_name.set(name, id)
    if is_default {
      match kind {
        "paragraph" => table = { ..table, default_para: Some(id) }
        "character" => table = { ..table, default_char: Some(id) }
        "table" => table = { ..table, default_table: Some(id) }
        _ => ()
      }
    }
  }
  table
}

/// Looks up a style by id.
pub fn StyleTable::get(self : StyleTable, id : String) -> Style? {
  self.styles.get(id)
}

/// Looks up a style by its display name.
pub fn StyleTable::get_by_name(self : StyleTable, name : String) -> Style? {
  match self.by_name.get(name) {
    Some(id) => self.styles.get(id)
    None => None
  }
}

/// Number of styles.
pub fn StyleTable::count(self : StyleTable) -> Int {
  self.styles.size()
}

/// All styles, sorted by kind then display name.
pub fn StyleTable::all(self : StyleTable) -> Array[Style] {
  let out = []
  for _, style in self.styles {
    out.push(style)
  }
  out.sort_by(fn(a, b) {
    let ka = a.kind + "\u0000" + a.name
    let kb = b.kind + "\u0000" + b.name
    ka.compare(kb)
  })
  out
}

/// Effective character formatting for a style id (defaults + inheritance chain).
pub fn StyleTable::resolve_run(self : StyleTable, id : String?) -> RunProps {
  let chain = self.chain(id)
  let mut props = self.doc_run
  let mut i = chain.length() - 1
  while i >= 0 {
    props = props.over(chain[i].run)
    i = i - 1
  }
  props
}

/// Effective paragraph layout for a style id (defaults + inheritance chain).
pub fn StyleTable::resolve_para(self : StyleTable, id : String?) -> ParaProps {
  let chain = self.chain(id)
  let mut props = self.doc_para
  let mut i = chain.length() - 1
  while i >= 0 {
    props = props.over(chain[i].para)
    i = i - 1
  }
  props
}

/// The inheritance chain, most derived first, following `w:basedOn`.
fn StyleTable::chain(self : StyleTable, id : String?) -> Array[Style] {
  let out : Array[Style] = []
  let mut current = id
  let mut depth = 0
  while depth < 32 {
    match current {
      Some(cid) =>
        match self.styles.get(cid) {
          Some(style) => {
            for seen in out {
              if seen.id == cid {
                return out
              }
            }
            out.push(style)
            current = style.based_on
          }
          None => return out
        }
      None => return out
    }
    depth = depth + 1
  }
  out
}

/// Document-level default character properties.
pub fn StyleTable::doc_run(self : StyleTable) -> RunProps {
  self.doc_run
}

/// Document-level default paragraph properties.
pub fn StyleTable::doc_para(self : StyleTable) -> ParaProps {
  self.doc_para
}

/// The default paragraph style id, if the document declares one.
pub fn StyleTable::default_paragraph_style(self : StyleTable) -> String? {
  self.default_para
}
