/// The WordprocessingML document model: body, paragraphs, runs, tables.

/// Package name of the main document part.
pub let part_document : String = "word/document.xml"

/// Package name of the style table.
pub let part_styles : String = "word/styles.xml"

/// Package name of the numbering definitions.
pub let part_numbering : String = "word/numbering.xml"

/// A node in the document body.
pub(all) enum Block {
  Para(Paragraph)
  Tbl(Table)
}

/// Appends the visible text of an element, honouring Word's text primitives.
pub fn gather_text(el : @xml.Element, sb : StringBuilder) -> Unit {
  for node in el.nodes() {
    match node {
      @xml.XmlNode::Elem(child) => {
        if child.namespace() == @xml.ns_w {
          match child.local() {
            "t" => sb.write_string(child.text())
            "tab" => sb.write_char('\t')
            "br" => sb.write_char('\n')
            "cr" => sb.write_char('\n')
            "noBreakHyphen" => sb.write_char('-')
            "softHyphen" => ()
            "pPr" => ()
            "rPr" => ()
            "del" => ()
            "moveFrom" => ()
            "delText" => ()
            "instrText" => ()
            "drawing" => ()
            "pict" => ()
            "object" => ()
            _ => gather_text(child, sb)
          }
        } else if child.namespace() != @xml.ns_mc {
          gather_text(child, sb)
        }
      }
      _ => ()
    }
  }
}

/// One paragraph.
pub struct Paragraph {
  el : @xml.Element
  styles : StyleTable
}

/// Wraps an existing `w:p` element.
pub fn Paragraph::new(el : @xml.Element, styles : StyleTable) -> Paragraph {
  Paragraph::{ el, styles }
}

/// The paragraph's plain text, one line per visual line.
pub fn Paragraph::text(self : Paragraph) -> String {
  let sb = StringBuilder()
  gather_text(self.el, sb)
  sb.to_string()
}

/// Whether the paragraph carries no text at all.
pub fn Paragraph::is_empty(self : Paragraph) -> Bool {
  self.text().trim() == ""
}

/// Style id referenced directly by `w:pStyle`, if any.
pub fn Paragraph::style_id(self : Paragraph) -> String? {
  match self.el.child(@xml.ns_w, "pPr") {
    Some(ppr) =>
      match ppr.child(@xml.ns_w, "pStyle") {
        Some(pstyle) => pstyle.attr("val")
        None => None
      }
    None => None
  }
}

/// The style object behind `style_id`.
pub fn Paragraph::style(self : Paragraph) -> Style? {
  match self.style_id() {
    Some(id) => self.styles.get(id)
    None => None
  }
}

/// Properties declared directly on this paragraph.
pub fn Paragraph::direct_props(self : Paragraph) -> ParaProps {
  read_para_props(self.el.child(@xml.ns_w, "pPr"))
}

/// Properties after folding in the style chain.
pub fn Paragraph::props(self : Paragraph) -> ParaProps {
  let style = match self.style_id() {
    Some(id) => Some(id)
    None => self.styles.default_para
  }
  self.styles.resolve_para(style).over(self.direct_props())
}

/// Heading level (1..9) or `None` for body text.
pub fn Paragraph::heading_level(self : Paragraph) -> Int? {
  let props = self.props()
  match props.outline_level {
    Some(level) =>
      if level >= 0 && level < 9 {
        return Some(level + 1)
      }
    None => ()
  }
  match self.style() {
    Some(style) => style.heading_level()
    None => None
  }
}

/// Left indent in twips.
pub fn Paragraph::left_twips(self : Paragraph) -> Int {
  match self.props().left_twips {
    Some(v) => v
    None => 0
  }
}

/// Whether the paragraph is part of a numbered or bulleted list.
pub fn Paragraph::is_list_item(self : Paragraph) -> Bool {
  match self.props().num_id {
    Some(num) => num != "0"
    None => false
  }
}

/// The runs of this paragraph, including runs nested in hyperlinks and fields.
pub fn Paragraph::runs(self : Paragraph) -> Array[Run] {
  let out : Array[Run] = []
  let para_props = self.props()
  collect_runs(self.el, self.styles, para_props, out)
  out
}

fn collect_runs(
  el : @xml.Element,
  styles : StyleTable,
  para_props : ParaProps,
  out : Array[Run],
) -> Unit {
  for node in el.nodes() {
    match node {
      @xml.XmlNode::Elem(child) => {
        if child.is_w("r") {
          out.push(Run::new(child, styles, para_props))
        } else if child.is_w("pPr") || child.is_w("del") || child.is_w("moveFrom") {
          continue
        } else {
          collect_runs(child, styles, para_props, out)
        }
      }
      _ => ()
    }
  }
}

/// Deletes `el` from its parent.
pub fn detach(parent : @xml.Element, el : @xml.Element) -> Bool {
  let index = parent.index_of(el)
  if index < 0 {
    return false
  }
  let nodes = parent.nodes()
  let kept : Array[@xml.XmlNode] = []
  for j, node in nodes {
    if j != index {
      kept.push(node)
    }
  }
  parent.set_children(kept)
  true
}

/// One text run.
pub struct Run {
  el : @xml.Element
  styles : StyleTable
  para_props : ParaProps
}

pub fn Run::new(
  el : @xml.Element,
  styles : StyleTable,
  para_props : ParaProps,
) -> Run {
  Run::{ el, styles, para_props }
}

/// The run's visible text.
pub fn Run::text(self : Run) -> String {
  let sb = StringBuilder()
  gather_text(self.el, sb)
  sb.to_string()
}

/// Whether the run carries no visible text (a break, a field marker, ...).
pub fn Run::is_textless(self : Run) -> Bool {
  self.text() == ""
}

/// Character style id referenced by `w:rStyle`.
pub fn Run::char_style_id(self : Run) -> String? {
  match self.el.child(@xml.ns_w, "rPr") {
    Some(rpr) =>
      match rpr.child(@xml.ns_w, "rStyle") {
        Some(rs) => rs.attr("val")
        None => None
      }
    None => None
  }
}

/// Effective character properties: character style, then run overrides.
pub fn Run::props(self : Run) -> RunProps {
  let char_style = match self.char_style_id() {
    Some(id) => Some(id)
    None => self.styles.default_char
  }
  self.styles
    .resolve_run(char_style)
    .over(read_run_props(self.el.child(@xml.ns_w, "rPr")))
}

/// Sets the text of this run, keeping its formatting.
///
/// The run keeps a single `w:t` child; any additional text primitives (tabs,
/// breaks) are dropped, which is what "replace the placeholder" means.
pub fn Run::set_text(self : Run, text : String) -> Unit {
  let kept : Array[@xml.XmlNode] = []
  let t = @xml.Element::make(@xml.ns_w, "t")
  t.set_attr("xml:space", "preserve")
  t.push(@xml.XmlNode::Text(text))
  for node in self.el.nodes() {
    match node {
      @xml.XmlNode::Elem(child) =>
        if child.is_w("rPr") {
          kept.push(node)
        }
      _ => ()
    }
  }
  kept.push(@xml.XmlNode::Elem(t))
  self.el.set_children(kept)
}

/// One table.
pub struct Table {
  el : @xml.Element
  styles : StyleTable
}

/// Number of rows.
pub fn Table::row_count(self : Table) -> Int {
  self.el.w_children("tr").length()
}

/// Cell texts in row-major order.
pub fn Table::grid(self : Table) -> Array[Array[String]] {
  let out : Array[Array[String]] = []
  for row in self.el.w_children("tr") {
    let cells : Array[String] = []
    for cell in row.w_children("tc") {
      let sb = StringBuilder()
      let paras = cell.w_children("p")
      for i, para in paras {
        if i > 0 {
          sb.write_string("\n")
        }
        gather_text(para, sb)
      }
      cells.push(sb.to_string())
    }
    out.push(cells)
  }
  out
}

/// Number of columns, counting the widest row.
pub fn Table::col_count(self : Table) -> Int {
  let mut widest = 0
  for row in self.grid() {
    if row.length() > widest {
      widest = row.length()
    }
  }
  widest
}

/// Table style id, if any.
pub fn Table::style_id(self : Table) -> String? {
  match self.el.child(@xml.ns_w, "tblPr") {
    Some(pr) =>
      match pr.child(@xml.ns_w, "tblStyle") {
        Some(s) => s.attr("val")
        None => None
      }
    None => None
  }
}

/// A `.docx` document opened for reading and editing.
pub struct Docx {
  pkg : WordPackage
  doc : @xml.Document
  styles : StyleTable
}

/// Opens a document from `.docx` bytes.
pub fn Docx::open(bytes : Bytes) -> Docx raise DocxError {
  let pkg = WordPackage::open(bytes)
  let source = match pkg.part_text(part_document) {
    Some(text) => text
    None => raise DocxError("not a Word document: word/document.xml is missing")
  }
  let doc = try {
    @xml.parse(source)
  } catch {
    _ => raise DocxError("word/document.xml is not well-formed XML")
  }
  let styles = match pkg.part_text(part_styles) {
    Some(text) =>
      try {
        StyleTable::parse(text)
      } catch {
        _ => StyleTable::empty()
      }
    None => StyleTable::empty()
  }
  Docx::{ pkg, doc, styles }
}

/// The underlying zip package.
pub fn Docx::package(self : Docx) -> WordPackage {
  self.pkg
}

/// The document's style table.
pub fn Docx::styles(self : Docx) -> StyleTable {
  self.styles
}

/// The root `w:document` element.
pub fn Docx::root(self : Docx) -> @xml.Element {
  self.doc.root
}

/// The `w:body` element.
pub fn Docx::body(self : Docx) -> @xml.Element {
  match self.root().child(@xml.ns_w, "body") {
    Some(body) => body
    None => self.root()
  }
}

/// Top-level body blocks in document order.
pub fn Docx::blocks(self : Docx) -> Array[Block] {
  blocks_of(self.body(), self.styles)
}

fn blocks_of(el : @xml.Element, styles : StyleTable) -> Array[Block] {
  let out : Array[Block] = []
  for child in el.children_in(@xml.ns_w) {
    match child.local() {
      "p" => out.push(Block::Para(Paragraph::{ el: child, styles }))
      "tbl" => out.push(Block::Tbl(Table::{ el: child, styles }))
      "sdt" =>
        match child.child(@xml.ns_w, "sdtContent") {
          Some(content) => for block in blocks_of(content, styles) {
            out.push(block)
          }
          None => ()
        }
      _ => ()
    }
  }
  out
}

/// All paragraphs, including those inside table cells, in document order.
pub fn Docx::paragraphs(self : Docx) -> Array[Paragraph] {
  let out : Array[Paragraph] = []
  for el in self.body().find_all_w("p") {
    out.push(Paragraph::{ el, styles: self.styles })
  }
  out
}

/// Top-level body paragraphs only.
pub fn Docx::body_paragraphs(self : Docx) -> Array[Paragraph] {
  let out : Array[Paragraph] = []
  for block in self.blocks() {
    match block {
      Block::Para(p) => out.push(p)
      _ => ()
    }
  }
  out
}

/// All tables in the document.
pub fn Docx::tables(self : Docx) -> Array[Table] {
  let out : Array[Table] = []
  for el in self.body().find_all_w("tbl") {
    out.push(Table::{ el, styles: self.styles })
  }
  out
}

/// Serialises the document back to `.docx` bytes.
pub fn Docx::save(self : Docx) -> Bytes raise DocxError {
  self.pkg.set_part_text(part_document, @xml.document_to_string(self.doc))
  self.pkg.save()
}

/// The full text of the document, one line per paragraph.
pub fn Docx::text(self : Docx) -> String {
  let sb = StringBuilder()
  let paras = self.paragraphs()
  for i, para in paras {
    if i > 0 {
      sb.write_char('\n')
    }
    sb.write_string(para.text())
  }
  sb.to_string()
}

/// The element behind a paragraph, for low-level queries.
pub fn Paragraph::element(self : Paragraph) -> @xml.Element {
  self.el
}

/// The element behind a table.
pub fn Table::element(self : Table) -> @xml.Element {
  self.el
}

/// The element behind a run.
pub fn Run::element(self : Run) -> @xml.Element {
  self.el
}
