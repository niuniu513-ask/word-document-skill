/// Character and paragraph property reading.

/// Reads an OOXML on/off property: a bare element means "on", an explicit
/// `w:val` of 0/false/off means "off".
pub fn ooxml_flag(el : @xml.Element) -> Bool {
  match el.attr("val") {
    Some(v) => !(v == "0" || v == "false" || v == "off")
    None => true
  }
}

fn[T] pick(base : T?, over : T?) -> T? {
  match over {
    Some(_) => over
    None => base
  }
}

/// Character formatting. Every field is optional so that unresolved properties
/// inherit from the style chain and finally from document defaults.
pub struct RunProps {
  bold : Bool?
  italic : Bool?
  underline : Bool?
  strike : Bool?
  font_ascii : String?
  font_east_asia : String?
  font_hansi : String?
  font_cs : String?
  size_pt : Double?
  color : String?
  highlight : String?
  vert_align : String?
  char_spacing_pt : Double?
}

/// The empty property set.
pub fn RunProps::none() -> RunProps {
  RunProps::{
    bold: None,
    italic: None,
    underline: None,
    strike: None,
    font_ascii: None,
    font_east_asia: None,
    font_hansi: None,
    font_cs: None,
    size_pt: None,
    color: None,
    highlight: None,
    vert_align: None,
    char_spacing_pt: None,
  }
}

/// Layers `over` on top of `self`, letting defined values win.
pub fn RunProps::over(self : RunProps, over : RunProps) -> RunProps {
  RunProps::{
    bold: pick(self.bold, over.bold),
    italic: pick(self.italic, over.italic),
    underline: pick(self.underline, over.underline),
    strike: pick(self.strike, over.strike),
    font_ascii: pick(self.font_ascii, over.font_ascii),
    font_east_asia: pick(self.font_east_asia, over.font_east_asia),
    font_hansi: pick(self.font_hansi, over.font_hansi),
    font_cs: pick(self.font_cs, over.font_cs),
    size_pt: pick(self.size_pt, over.size_pt),
    color: pick(self.color, over.color),
    highlight: pick(self.highlight, over.highlight),
    vert_align: pick(self.vert_align, over.vert_align),
    char_spacing_pt: pick(self.char_spacing_pt, over.char_spacing_pt),
  }
}

/// Reads a `w:rPr` element into a property set.
pub fn read_run_props(rpr : @xml.Element?) -> RunProps {
  let props = RunProps::none()
  match rpr {
    None => props
    Some(rpr) => {
      let mut out = props
      for child in rpr.children_in(@xml.ns_w) {
        match child.local() {
          "b" => out = { ..out, bold: Some(ooxml_flag(child)) }
          "i" => out = { ..out, italic: Some(ooxml_flag(child)) }
          "u" =>
            match child.attr("val") {
              Some("none") => out = { ..out, underline: Some(false) }
              _ => out = { ..out, underline: Some(true) }
            }
          "strike" => out = { ..out, strike: Some(ooxml_flag(child)) }
          "rFonts" => {
            out = {
              ..out,
              font_ascii: pick(out.font_ascii, child.attr("ascii")),
              font_east_asia: pick(out.font_east_asia, child.attr("eastAsia")),
              font_hansi: pick(out.font_hansi, child.attr("hAnsi")),
              font_cs: pick(out.font_cs, child.attr("cs")),
            }
          }
          "sz" =>
            match child.attr_int("val") {
              Some(v) => out = { ..out, size_pt: Some(half_points_to_pt(v)) }
              None => ()
            }
          "color" =>
            match child.attr("val") {
              Some(v) =>
                if v != "auto" {
                  out = { ..out, color: Some(v) }
                }
              None => ()
            }
          "highlight" => out = { ..out, highlight: child.attr("val") }
          "vertAlign" => out = { ..out, vert_align: child.attr("val") }
          "spacing" =>
            match child.attr_int("val") {
              Some(v) => out = { ..out, char_spacing_pt: Some(twips_to_pt(v)) }
              None => ()
            }
          _ => ()
        }
      }
      out
    }
  }
}

/// Paragraph layout properties.
pub struct ParaProps {
  style_id : String?
  alignment : String?
  outline_level : Int?
  left_twips : Int?
  right_twips : Int?
  first_line_twips : Int?
  hanging_twips : Int?
  space_before_twips : Int?
  space_after_twips : Int?
  line_twips : Int?
  line_rule : String?
  num_id : String?
  ilvl : Int?
  keep_next : Bool?
  page_break_before : Bool?
  contextual_spacing : Bool?
}

/// The empty paragraph property set.
pub fn ParaProps::none() -> ParaProps {
  ParaProps::{
    style_id: None,
    alignment: None,
    outline_level: None,
    left_twips: None,
    right_twips: None,
    first_line_twips: None,
    hanging_twips: None,
    space_before_twips: None,
    space_after_twips: None,
    line_twips: None,
    line_rule: None,
    num_id: None,
    ilvl: None,
    keep_next: None,
    page_break_before: None,
    contextual_spacing: None,
  }
}

/// Layer `over` on top of `self`.
pub fn ParaProps::over(self : ParaProps, over : ParaProps) -> ParaProps {
  ParaProps::{
    style_id: pick(self.style_id, over.style_id),
    alignment: pick(self.alignment, over.alignment),
    outline_level: pick(self.outline_level, over.outline_level),
    left_twips: pick(self.left_twips, over.left_twips),
    right_twips: pick(self.right_twips, over.right_twips),
    first_line_twips: pick(self.first_line_twips, over.first_line_twips),
    hanging_twips: pick(self.hanging_twips, over.hanging_twips),
    space_before_twips: pick(self.space_before_twips, over.space_before_twips),
    space_after_twips: pick(self.space_after_twips, over.space_after_twips),
    line_twips: pick(self.line_twips, over.line_twips),
    line_rule: pick(self.line_rule, over.line_rule),
    num_id: pick(self.num_id, over.num_id),
    ilvl: pick(self.ilvl, over.ilvl),
    keep_next: pick(self.keep_next, over.keep_next),
    page_break_before: pick(self.page_break_before, over.page_break_before),
    contextual_spacing: pick(self.contextual_spacing, over.contextual_spacing),
  }
}

/// Reads a `w:pPr` element into a paragraph property set.
pub fn read_para_props(ppr : @xml.Element?) -> ParaProps {
  let base = ParaProps::none()
  match ppr {
    None => base
    Some(ppr) => {
      let mut out = base
      for child in ppr.children_in(@xml.ns_w) {
        match child.local() {
          "pStyle" => out = { ..out, style_id: child.attr("val") }
          "jc" => out = { ..out, alignment: child.attr("val") }
          "outlineLvl" => out = { ..out, outline_level: child.attr_int("val") }
          "ind" => {
            out = {
              ..out,
              left_twips: pick(out.left_twips, child.attr_int("left")),
              right_twips: pick(out.right_twips, child.attr_int("right")),
              first_line_twips: pick(out.first_line_twips, child.attr_int("firstLine")),
              hanging_twips: pick(out.hanging_twips, child.attr_int("hanging")),
            }
          }
          "spacing" => {
            out = {
              ..out,
              space_before_twips: pick(out.space_before_twips, child.attr_int("before")),
              space_after_twips: pick(out.space_after_twips, child.attr_int("after")),
              line_twips: pick(out.line_twips, child.attr_int("line")),
              line_rule: pick(out.line_rule, child.attr("lineRule")),
            }
          }
          "keepNext" => out = { ..out, keep_next: Some(ooxml_flag(child)) }
          "pageBreakBefore" => out = { ..out, page_break_before: Some(ooxml_flag(child)) }
          "contextualSpacing" => out = { ..out, contextual_spacing: Some(ooxml_flag(child)) }
          "numPr" => {
            for item in child.children_in(@xml.ns_w) {
              match item.local() {
                "numId" => out = { ..out, num_id: item.attr("val") }
                "ilvl" => out = { ..out, ilvl: item.attr_int("val") }
                _ => ()
              }
            }
          }
          _ => ()
        }
      }
      out
    }
  }
}
