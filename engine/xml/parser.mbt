/// A dependency-free XML 1.0 parser that keeps namespaces, entities, comments
/// and processing instructions. It is designed for OOXML: one pass, an explicit
/// element stack (so hostile nesting cannot blow the MoonBit stack), and a tree
/// that re-serializes without loss.

/// Raised for malformed input. The message carries a 1-based byte offset.
pub suberror XmlError {
  XmlError(String)
}

/// A parsed XML document: the original declaration plus the root element.
pub struct Document {
  decl : String
  root : Element
}

/// The XML declaration as it appeared in the source (`""` when absent).
pub fn Document::declaration(self : Document) -> String {
  self.decl
}

/// The root element.
pub fn Document::root(self : Document) -> Element {
  self.root
}

/// Parses a whole document and returns it.
pub fn parse(source : StringView) -> Document raise XmlError {
  let text = source.to_string()
  let cur = Cursor::at(text, 0)
  cur.skip_bom()
  let decl = cur.read_decl()
  cur.skip_misc()
  if cur.eof() {
    raise XmlError("empty document: no root element")
  }
  let root = cur.read_tree()
  { decl, root }
}

/// Parses a document and returns only its root element.
pub fn parse_element(source : StringView) -> Element raise XmlError {
  (parse(source)).root
}

/// Parses a detached fragment by wrapping it in a synthetic root element.
pub fn parse_fragment(source : StringView) -> Element raise XmlError {
  let wrapped = "<fragment>" + source.to_string() + "</fragment>"
  let root = parse_element(wrapped)
  let kids = root.elements()
  if kids.is_empty() {
    raise XmlError("empty fragment")
  }
  kids[0]
}

// ---------------------------------------------------------------- cursor

struct RawAttr {
  qname : String
  value : String
}

struct Cursor {
  src : String
  mut pos : Int32
  len : Int32
  scope : Map[String, String]
  history : Array[Array[(String, String?)]]
}

fn Cursor::at(src : String, pos : Int32) -> Cursor {
  let scope : Map[String, String] = Map([])
  scope["xml"] = ns_xml
  { src, pos, len: src.length(), scope, history: [] }
}

fn Cursor::eof(self : Cursor) -> Bool {
  self.pos >= self.len
}

fn Cursor::peek(self : Cursor) -> Char? {
  self.src.get_char(self.pos)
}

fn Cursor::error(self : Cursor, message : String) -> XmlError {
  XmlError("offset " + self.pos.to_string() + ": " + message)
}

fn Cursor::starts_with(self : Cursor, s : String) -> Bool {
  let n = s.length()
  if self.pos + n > self.len {
    return false
  }
  let mut i = 0
  while i < n {
    if self.src.get_char(self.pos + i) != s.get_char(i) {
      return false
    }
    i = i + 1
  }
  true
}

fn Cursor::skip_bom(self : Cursor) -> Unit {
  if self.peek() == Some('\u{FEFF}') {
    self.pos = self.pos + 1
  }
}

fn Cursor::skip_ws(self : Cursor) -> Unit {
  while true {
    match self.peek() {
      Some(c) => {
        if c == ' ' || c == '\t' || c == '\n' || c == '\r' {
          self.pos = self.pos + 1
        } else {
          break
        }
      }
      None => break
    }
  }
}

fn Cursor::read_decl(self : Cursor) -> String {
  if self.starts_with("<?xml") {
    let start = self.pos + 2
    while not(self.eof()) && not(self.starts_with("?>")) {
      self.pos = self.pos + 1
    }
    let text = self.src.substring(start~, end=self.pos)
    if self.starts_with("?>") {
      self.pos = self.pos + 2
    }
    text
  } else {
    ""
  }
}

fn Cursor::skip_misc(self : Cursor) -> Unit {
  while true {
    self.skip_ws()
    if self.starts_with("<!--") {
      self.skip_comment()
    } else if self.starts_with("<?") {
      self.skip_pi()
    } else if self.starts_with("<!") {
      self.skip_doctype()
    } else {
      break
    }
  }
}

fn Cursor::skip_comment(self : Cursor) -> Unit {
  self.pos = self.pos + 4
  while not(self.eof()) && not(self.starts_with("-->")) {
    self.pos = self.pos + 1
  }
  if self.starts_with("-->") {
    self.pos = self.pos + 3
  }
}

fn Cursor::skip_pi(self : Cursor) -> Unit {
  self.pos = self.pos + 2
  while not(self.eof()) && not(self.starts_with("?>")) {
    self.pos = self.pos + 1
  }
  if self.starts_with("?>") {
    self.pos = self.pos + 2
  }
}

fn Cursor::skip_doctype(self : Cursor) -> Unit {
  let mut depth = 0
  while not(self.eof()) {
    match self.peek() {
      Some('[') => depth = depth + 1
      Some(']') => if depth > 0 { depth = depth - 1 }
      Some('>') =>
        if depth == 0 {
          self.pos = self.pos + 1
          return
        }
      _ => ()
    }
    self.pos = self.pos + 1
  }
}

// ---------------------------------------------------------------- names

fn split_qname(qname : String) -> (String, String) {
  match qname.find(":") {
    Some(i) => (qname.substring(end=i), qname.substring(start=i + 1))
    None => ("", qname)
  }
}

fn Cursor::resolve(self : Cursor, prefix : String) -> String {
  match self.scope.get(prefix) {
    Some(uri) => uri
    None => ""
  }
}

fn Cursor::enter_scope(self : Cursor, raw_attrs : Array[RawAttr]) -> Unit {
  let undo : Array[(String, String?)] = []
  for a in raw_attrs {
    if a.qname == "xmlns" {
      undo.push(("", self.scope.get("")))
      self.scope[""] = a.value
    } else if a.qname.find("xmlns:") == Some(0) && a.qname.length() > 6 {
      let prefix = a.qname.substring(start=6)
      undo.push((prefix, self.scope.get(prefix)))
      self.scope[prefix] = a.value
    }
  }
  self.history.push(undo)
}

fn Cursor::leave_scope(self : Cursor) -> Unit {
  if self.history.is_empty() {
    return
  }
  let undo = self.history.pop().unwrap()
  let mut i = undo.length() - 1
  while i >= 0 {
    let (prefix, previous) = undo[i]
    match previous {
      Some(v) => self.scope[prefix] = v
      None => self.scope[prefix] = ""
    }
    i = i - 1
  }
}

// ---------------------------------------------------------------- reader

fn Cursor::read_name(self : Cursor) -> String raise XmlError {
  let start = self.pos
  while not(self.eof()) {
    match self.peek() {
      Some(c) => {
        if is_name_char(c) {
          self.pos = self.pos + 1
        } else {
          break
        }
      }
      None => break
    }
  }
  if self.pos == start {
    raise self.error("expected a name")
  }
  self.src.substring(start~, end=self.pos)
}

fn is_name_char(c : Char) -> Bool {
  let v = c.to_int()
  (v >= 97 && v <= 122) ||
  (v >= 65 && v <= 90) ||
  (v >= 48 && v <= 57) ||
  v == 95 || v == 58 || v == 45 || v == 46
}

fn Cursor::read_attrs(self : Cursor) -> Array[RawAttr] raise XmlError {
  let out : Array[RawAttr] = []
  while true {
    self.skip_ws()
    match self.peek() {
      Some('>') => break
      Some('/') => break
      None => raise self.error("unterminated start tag")
      _ => ()
    }
    let name = self.read_name()
    self.skip_ws()
    if self.peek() != Some('=') {
      raise self.error("expected '=' after attribute " + name)
    }
    self.pos = self.pos + 1
    self.skip_ws()
    let quote = match self.peek() {
      Some(q) =>
        if q == '"' || q == '\'' {
          q
        } else {
          raise self.error("expected a quoted attribute value")
        }
      None => raise self.error("expected a quoted attribute value")
    }
    self.pos = self.pos + 1
    let start = self.pos
    while not(self.eof()) && self.peek() != Some(quote) {
      self.pos = self.pos + 1
    }
    if self.eof() {
      raise self.error("unterminated attribute value")
    }
    let raw = self.src.substring(start~, end=self.pos)
    self.pos = self.pos + 1
    out.push({ qname: name, value: unescape(raw) })
  }
  out
}

fn Cursor::read_tree(self : Cursor) -> Element raise XmlError {
  let stack : Array[Element] = []
  let mut root : Element? = None
  while not(self.eof()) {
    if self.starts_with("</") {
      self.pos = self.pos + 2
      let name = self.read_name()
      self.skip_ws()
      if self.peek() != Some('>') {
        raise self.error("expected '>' closing </" + name + ">")
      }
      self.pos = self.pos + 1
      match stack.pop() {
        Some(el) =>
          if el.qname != name {
            raise self.error(
              "mismatched closing tag </" + name + ">, expected </" + el.qname + ">",
            )
          }
        None => raise self.error("unexpected closing tag </" + name + ">")
      }
      self.leave_scope()
      if stack.is_empty() {
        break
      }
    } else if self.starts_with("<!--") {
      let start = self.pos + 4
      self.skip_comment()
      let text = self.src.substring(start~, end=self.pos - 3)
      attach(stack, Comment(text))
    } else if self.starts_with("<![CDATA[") {
      self.pos = self.pos + 9
      let start = self.pos
      while not(self.eof()) && not(self.starts_with("]]>")) {
        self.pos = self.pos + 1
      }
      let text = self.src.substring(start~, end=self.pos)
      if self.starts_with("]]>") {
        self.pos = self.pos + 3
      }
      attach(stack, Text(text))
    } else if self.starts_with("<?") {
      let start = self.pos + 2
      self.skip_pi()
      let text = self.src.substring(start~, end=self.pos - 2)
      attach(stack, Pi(text))
    } else if self.starts_with("<!") {
      self.skip_doctype()
    } else if self.peek() == Some('<') {
      self.pos = self.pos + 1
      let qname = self.read_name()
      let raw_attrs = self.read_attrs()
      self.enter_scope(raw_attrs)
      let (prefix, local_name) = split_qname(qname)
      let ns = self.resolve(prefix)
      let attrs : Array[Attr] = []
      for a in raw_attrs {
        let (aprefix, alocal) = split_qname(a.qname)
        let is_decl = a.qname == "xmlns" || a.qname.find("xmlns:") == Some(0)
        let ans = if is_decl || aprefix == "" {
          ""
        } else {
          self.resolve(aprefix)
        }
        attrs.push({ qname: a.qname, ns: ans, local_name: alocal, value: a.value })
      }
      let el : Element = { uid: fresh_element_id(), qname, ns, local_name, attrs, children: [] }
      let self_closing = if self.peek() == Some('/') {
        self.pos = self.pos + 1
        if self.peek() != Some('>') {
          raise self.error("expected '>' in self-closing tag")
        }
        self.pos = self.pos + 1
        true
      } else {
        if self.peek() != Some('>') {
          raise self.error("expected '>' in start tag <" + qname + ">")
        }
        self.pos = self.pos + 1
        false
      }
      if stack.is_empty() {
        root = Some(el)
      } else {
        stack[stack.length() - 1].children.push(Elem(el))
      }
      if self_closing {
        self.leave_scope()
        if stack.is_empty() {
          break
        }
      } else {
        stack.push(el)
      }
    } else {
      let start = self.pos
      while not(self.eof()) && self.peek() != Some('<') {
        self.pos = self.pos + 1
      }
      let text = self.src.substring(start~, end=self.pos)
      if stack.is_empty() {
        if text.trim().is_empty() {
          continue
        }
        raise self.error("text outside the root element")
      }
      stack[stack.length() - 1].children.push(Text(unescape(text)))
    }
  }
  match root {
    Some(el) => el
    None => raise self.error("no root element found")
  }
}

fn attach(stack : Array[Element], node : XmlNode) -> Unit {
  if stack.is_empty() {
    return
  }
  stack[stack.length() - 1].children.push(node)
}

// ---------------------------------------------------------------- entities

/// Decodes the five predefined entities plus numeric character references.
pub fn unescape(s : String) -> String {
  match s.find("&") {
    None => return s
    _ => ()
  }
  let sb = StringBuilder()
  let mut i = 0
  let n = s.length()
  while i < n {
    let c = s.get_char(i)
    if c == Some('&') {
      let mut stop = -1
      let mut j = i + 1
      while j < n && j < i + 12 {
        if s.get_char(j) == Some(';') {
          stop = j
          break
        }
        j = j + 1
      }
      if stop < 0 {
        sb.write_char('&')
        i = i + 1
      } else {
        let entity = s.substring(start=i + 1, end=stop)
        if entity == "amp" {
          sb.write_char('&')
        } else if entity == "lt" {
          sb.write_char('<')
        } else if entity == "gt" {
          sb.write_char('>')
        } else if entity == "quot" {
          sb.write_char('"')
        } else if entity == "apos" {
          sb.write_char('\'')
        } else if entity == "nbsp" {
          sb.write_char('\u{00A0}')
        } else if entity.find("#") == Some(0) {
          match decode_char_ref(entity) {
            Some(ch) => sb.write_char(ch)
            None => sb.write_string(s.substring(start=i, end=stop + 1))
          }
        } else {
          sb.write_string(s.substring(start=i, end=stop + 1))
        }
        i = stop + 1
      }
    } else {
      match c {
        Some(ch) => sb.write_char(ch)
        None => ()
      }
      i = i + 1
    }
  }
  sb.to_string()
}

fn decode_char_ref(entity : String) -> Char? {
  let body = entity.substring(start=1)
  if body.length() == 0 {
    return None
  }
  let mut radix = 10
  let mut digits = body
  if body.get_char(0) == Some('x') || body.get_char(0) == Some('X') {
    radix = 16
    digits = body.substring(start=1)
  }
  if digits.length() == 0 {
    return None
  }
  let mut acc = 0
  let mut i = 0
  while i < digits.length() {
    let d = match digits.get_char(i) {
      Some(ch) => digit_value(ch)
      None => return None
    }
    if d < 0 || d >= radix {
      return None
    }
    acc = acc * radix + d
    i = i + 1
  }
  if acc <= 0 || acc > 0x10FFFF {
    return None
  }
  acc.to_char()
}

fn digit_value(c : Char) -> Int {
  let v = c.to_int()
  if v >= 48 && v <= 57 {
    v - 48
  } else if v >= 97 && v <= 102 {
    v - 87
  } else if v >= 65 && v <= 70 {
    v - 55
  } else {
    -1
  }
}
