/// XML tree model and query helpers.
///
/// Every element keeps three names: the qualified name exactly as written
/// (`w:p`), the resolved namespace URI, and the local name (`p`).
/// Word documents are namespace-heavy, so all DOCX code queries by URI,
/// never by prefix.

/// Namespace of the `xml:` prefix, defined by the XML spec itself.
pub let ns_xml : String = "http://www.w3.org/XML/1998/namespace"

/// An attribute, resolved against the in-scope namespace declarations.
pub struct Attr {
  qname : String
  ns : String
  local_name : String
  value : String
} derive(Debug)

/// One node of the tree. Text, comments and processing instructions are
/// preserved so that a document can be re-serialized without losing anything.
pub(all) enum XmlNode {
  Elem(Element)
  Text(String)
  Comment(String)
  Pi(String)
} derive(Debug)

/// Monotonic counter used to give every element a stable identity.
let element_seq : Array[Int] = [0]

fn fresh_element_id() -> Int {
  element_seq[0] = element_seq[0] + 1
  element_seq[0]
}

/// An element node.
///
/// `uid` is a process-local identity. Two elements parsed from identical XML
/// compare unequal by id, which is what edit operations need when they have to
/// locate "this very paragraph" inside a parent.
pub struct Element {
  uid : Int
  qname : String
  ns : String
  local_name : String
  attrs : Array[Attr]
  children : Array[XmlNode]
} derive(Debug)

/// Creates an element with an explicit namespace.
pub fn Element::make(ns : String, local_name : String) -> Element {
  let prefix = prefix_for(ns)
  let qname = if prefix == "" { local_name } else { prefix + ":" + local_name }
  { uid: fresh_element_id(), qname, ns, local_name, attrs: [], children: [] }
}

/// The process-local identity of this element.
pub fn Element::id(self : Element) -> Int {
  self.uid
}

/// The qualified name as written in the source document.
pub fn Element::qname(self : Element) -> String {
  self.qname
}

/// The namespace URI, or "" when the element has no namespace.
pub fn Element::namespace(self : Element) -> String {
  self.ns
}

/// The local name (prefix stripped).
pub fn Element::local(self : Element) -> String {
  self.local_name
}

/// All child nodes, including text.
pub fn Element::nodes(self : Element) -> Array[XmlNode] {
  self.children
}

/// All attributes.
pub fn Element::attributes(self : Element) -> Array[Attr] {
  self.attrs
}

/// Direct element children, text nodes filtered out.
pub fn Element::elements(self : Element) -> Array[Element] {
  let out : Array[Element] = []
  for node in self.children {
    match node {
      Elem(e) => out.push(e)
      _ => ()
    }
  }
  out
}

/// Number of direct element children.
pub fn Element::child_count(self : Element) -> Int {
  let mut n = 0
  for node in self.children {
    match node {
      Elem(_) => n = n + 1
      _ => ()
    }
  }
  n
}

/// True when the element matches the given namespace and local name.
pub fn Element::is_ns(self : Element, ns : StringView, local_name : StringView) -> Bool {
  self.local_name == local_name.to_string() && self.ns == ns.to_string()
}

/// True when this is a WordprocessingML element (`w:` namespace).
pub fn Element::is_w(self : Element, local_name : StringView) -> Bool {
  self.local_name == local_name.to_string() && self.ns == ns_w
}

/// First direct child with the given namespace and local name.
pub fn Element::child(self : Element, ns : String, local_name : String) -> Element? {
  for node in self.children {
    match node {
      Elem(e) => if e.local_name == local_name && e.ns == ns { return Some(e) }
      _ => ()
    }
  }
  None
}

/// First direct WordprocessingML child with the given local name.
pub fn Element::w_child(self : Element, local_name : String) -> Element? {
  self.child(ns_w, local_name)
}

/// All direct WordprocessingML children with the given local name.
pub fn Element::w_children(self : Element, local_name : String) -> Array[Element] {
  let out : Array[Element] = []
  for node in self.children {
    match node {
      Elem(e) => if e.is_w(local_name) { out.push(e) }
      _ => ()
    }
  }
  out
}

/// All direct children in the given namespace.
pub fn Element::children_in(self : Element, ns : String) -> Array[Element] {
  let out : Array[Element] = []
  for node in self.children {
    match node {
      Elem(e) => if e.ns == ns { out.push(e) }
      _ => ()
    }
  }
  out
}

/// Attribute value by local name, whatever its namespace.
pub fn Element::attr(self : Element, local_name : StringView) -> String? {
  let want = local_name.to_string()
  for a in self.attrs {
    if a.local_name == want { return Some(a.value) }
  }
  None
}

/// Attribute value by local name with a fallback.
pub fn Element::attr_or(self : Element, local_name : StringView, fallback : String) -> String {
  match self.attr(local_name) {
    Some(v) => v
    None => fallback
  }
}

/// Attribute value parsed as a decimal integer.
pub fn Element::attr_int(self : Element, local_name : StringView) -> Int? {
  match self.attr(local_name) {
    Some(v) => {
      let t = v.trim()
      if t.is_empty() { return None }
      let mut sign = 1
      let mut i = 0
      let s = t.to_string()
      if s.get_char(0) == Some('-') {
        sign = -1
        i = 1
      }
      let mut acc = 0
      let mut any = false
      while i < s.length() {
        match s.get_char(i) {
          Some(c) => {
            let d = c.to_int() - 48
            if d < 0 || d > 9 { return None }
            acc = acc * 10 + d
            any = true
          }
          None => break
        }
        i = i + 1
      }
      if any { Some(acc * sign) } else { None }
    }
    None => None
  }
}

/// True when the attribute is present with value "1" or "true" or "on".
pub fn Element::attr_bool(self : Element, local_name : StringView) -> Bool {
  match self.attr(local_name) {
    Some(v) => v == "1" || v == "true" || v == "on"
    None => false
  }
}

/// Sets (or appends) an attribute by local name.
pub fn Element::set_attr(self : Element, local_name : String, value : String) -> Unit {
  for i in 0..<self.attrs.length() {
    if self.attrs[i].local_name == local_name {
      self.attrs[i] = { qname: self.attrs[i].qname, ns: self.attrs[i].ns, local_name, value }
      return
    }
  }
  self.attrs.push({ qname: local_name, ns: "", local_name, value })
}

/// Sets an attribute in a namespace, reusing the canonical prefix given.
pub fn Element::set_attr_ns(
  self : Element,
  prefix : String,
  ns : String,
  local_name : String,
  value : String,
) -> Unit {
  let qname = if prefix == "" { local_name } else { prefix + ":" + local_name }
  for i in 0..<self.attrs.length() {
    if self.attrs[i].local_name == local_name && self.attrs[i].ns == ns {
      self.attrs[i] = { qname, ns, local_name, value }
      return
    }
  }
  self.attrs.push({ qname, ns, local_name, value })
}

/// Removes an attribute by local name; returns true when something was removed.
pub fn Element::remove_attr(self : Element, local_name : String) -> Bool {
  let mut idx = -1
  for i in 0..<self.attrs.length() {
    if self.attrs[i].local_name == local_name {
      idx = i
      break
    }
  }
  if idx < 0 { return false }
  let kept : Array[Attr] = []
  for i in 0..<self.attrs.length() {
    if i != idx { kept.push(self.attrs[i]) }
  }
  self.attrs.clear()
  for a in kept { self.attrs.push(a) }
  true
}

/// Appends a child node.
pub fn Element::push(self : Element, node : XmlNode) -> Unit {
  self.children.push(node)
}

/// Appends an element child and returns it.
pub fn Element::append(self : Element, child : Element) -> Unit {
  self.children.push(Elem(child))
}

/// Concatenated text of this element and all descendants.
pub fn Element::text(self : Element) -> String {
  let sb = StringBuilder()
  collect_text(self, sb)
  sb.to_string()
}

fn collect_text(el : Element, sb : StringBuilder) -> Unit {
  for node in el.children {
    match node {
      Text(t) => sb.write_string(t)
      Elem(e) => collect_text(e, sb)
      _ => ()
    }
  }
}

/// Text of the direct text children only.
pub fn Element::own_text(self : Element) -> String {
  let sb = StringBuilder()
  for node in self.children {
    match node {
      Text(t) => sb.write_string(t)
      _ => ()
    }
  }
  sb.to_string()
}

/// True when the element has neither children nor text.
pub fn Element::is_empty(self : Element) -> Bool {
  self.children.is_empty()
}

/// Depth-first search for every descendant (excluding self) that satisfies
/// `pred`, in document order.
pub fn Element::find_all(self : Element, pred : (Element) -> Bool) -> Array[Element] {
  let out : Array[Element] = []
  let stack : Array[Element] = []
  push_children_reversed(self, stack)
  while not(stack.is_empty()) {
    let el = stack.pop().unwrap()
    if pred(el) { out.push(el) }
    push_children_reversed(el, stack)
  }
  out
}

fn push_children_reversed(el : Element, stack : Array[Element]) -> Unit {
  let kids = el.elements()
  let mut i = kids.length() - 1
  while i >= 0 {
    stack.push(kids[i])
    i = i - 1
  }
}

/// Every descendant in a namespace with a given local name, document order.
pub fn Element::find_all_ns(self : Element, ns : String, local_name : String) -> Array[Element] {
  self.find_all(fn(e) { e.ns == ns && e.local_name == local_name })
}

/// Every descendant in the WordprocessingML namespace with a local name.
pub fn Element::find_all_w(self : Element, local_name : String) -> Array[Element] {
  self.find_all_ns(ns_w, local_name)
}

/// First descendant (excluding self) matching namespace and local name.
pub fn Element::find_ns(self : Element, ns : String, local_name : String) -> Element? {
  let hits = self.find_all_ns(ns, local_name)
  if hits.is_empty() { None } else { Some(hits[0]) }
}

/// First descendant in the WordprocessingML namespace with a local name.
pub fn Element::find_w(self : Element, local_name : String) -> Element? {
  self.find_ns(ns_w, local_name)
}

/// Replaces the children of this element.
pub fn Element::set_children(self : Element, nodes : Array[XmlNode]) -> Unit {
  while self.children.length() > 0 {
    ignore(self.children.pop())
  }
  for node in nodes {
    self.children.push(node)
  }
}

/// Inserts a child at `index`, clamping out-of-range indices.
pub fn Element::insert_at(self : Element, index : Int, node : XmlNode) -> Unit {
  let len = self.children.length()
  let at = if index < 0 {
    0
  } else if index > len {
    len
  } else {
    index
  }
  let tail : Array[XmlNode] = []
  let mut i = at
  while i < len {
    tail.push(self.children[i])
    i = i + 1
  }
  while self.children.length() > at {
    ignore(self.children.pop())
  }
  self.children.push(node)
  for item in tail {
    self.children.push(item)
  }
}

/// Removes and returns the child at `index`.
pub fn Element::remove_at(self : Element, index : Int) -> Bool {
  let len = self.children.length()
  if index < 0 || index >= len {
    return false
  }
  let tail : Array[XmlNode] = []
  let mut i = index + 1
  while i < len {
    tail.push(self.children[i])
    i = i + 1
  }
  while self.children.length() > index {
    ignore(self.children.pop())
  }
  for item in tail {
    self.children.push(item)
  }
  true
}

/// Position of a child element inside this element, or -1.
pub fn Element::index_of(self : Element, target : Element) -> Int {
  let mut index = 0
  for node in self.children {
    match node {
      Elem(child) => {
        if child.uid == target.uid {
          return index
        }
        index = index + 1
      }
      _ => ()
    }
  }
  -1
}

/// Deep-copies an element, giving the copy a fresh identity.
pub fn Element::copy(self : Element) -> Element {
  let el : Element = {
    uid: fresh_element_id(),
    qname: self.qname,
    ns: self.ns,
    local_name: self.local_name,
    attrs: [],
    children: [],
  }
  for attr in self.attrs {
    el.attrs.push(attr)
  }
  for node in self.children {
    match node {
      Elem(child) => el.children.push(Elem(child.copy()))
      Text(text) => el.children.push(Text(text))
      Comment(text) => el.children.push(Comment(text))
      Pi(text) => el.children.push(Pi(text))
    }
  }
  el
}
