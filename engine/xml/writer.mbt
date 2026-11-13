/// Serialization back to XML text.
///
/// The writer is namespace aware: Word documents bind the main namespace to
/// whatever prefix the producer liked, and elements we create ourselves may
/// carry a prefix that is not declared at the root. Every element therefore
/// resolves its prefix against the declarations in scope and emits an inline
/// `xmlns:` declaration when needed, so generated markup stays well formed and
/// keeps its namespaces.

/// Escapes character data.
pub fn escape_text(s : String) -> String {
  let sb = StringBuilder()
  for c in s {
    if c == '&' {
      sb.write_string("&amp;")
    } else if c == '<' {
      sb.write_string("&lt;")
    } else if c == '>' {
      sb.write_string("&gt;")
    } else {
      sb.write_char(c)
    }
  }
  sb.to_string()
}

/// Escapes an attribute value (adds quotes, newlines and tabs).
pub fn escape_attr(s : String) -> String {
  let sb = StringBuilder()
  for c in s {
    if c == '&' {
      sb.write_string("&amp;")
    } else if c == '<' {
      sb.write_string("&lt;")
    } else if c == '>' {
      sb.write_string("&gt;")
    } else if c == '"' {
      sb.write_string("&quot;")
    } else if c == '\n' {
      sb.write_string("&#10;")
    } else if c == '\r' {
      sb.write_string("&#13;")
    } else if c == '\t' {
      sb.write_string("&#9;")
    } else {
      sb.write_char(c)
    }
  }
  sb.to_string()
}

fn map_clone(m : Map[String, String]) -> Map[String, String] {
  let out : Map[String, String] = Map([])
  for k, v in m {
    out[k] = v
  }
  out
}

fn split_prefix(qname : String) -> (String, String) {
  let chars = qname.to_array()
  let mut i = 0
  while i < chars.length() {
    if chars[i] == ':' {
      let p = StringBuilder()
      let mut j = 0
      while j < i {
        p.write_char(chars[j])
        j = j + 1
      }
      let l = StringBuilder()
      let mut k = i + 1
      while k < chars.length() {
        l.write_char(chars[k])
        k = k + 1
      }
      return (p.to_string(), l.to_string())
    }
    i = i + 1
  }
  ("", qname)
}

fn scope_of(el : Element, parent : Map[String, String]) -> Map[String, String] {
  let out = map_clone(parent)
  for a in el.attrs {
    if a.qname == "xmlns" {
      out[""] = a.value
    } else if a.qname.has_prefix("xmlns:") {
      let (_, prefix) = split_prefix(a.qname)
      out[prefix] = a.value
    }
  }
  out
}

/// The declarations an element has to add for its own name and attributes.
fn declarations(
  el : Element,
  scope : Map[String, String],
) -> (Array[(String, String)], Array[String]) {
  let extra : Array[(String, String)] = []
  let (prefix, _) = split_prefix(el.qname)
  if el.ns != "" && prefix != "xml" {
    if scope.get(prefix).unwrap_or("") != el.ns {
      extra.push((prefix, el.ns))
    }
  }
  let names : Array[String] = []
  for a in el.attrs {
    let (aprefix, _) = split_prefix(a.qname)
    if a.ns != "" && aprefix == "" {
      let want = prefix_for(a.ns)
      let use = if want == "" { "ns1" } else { want }
      names.push(use + ":" + a.local_name)
      if scope.get(use).unwrap_or("") != a.ns {
        extra.push((use, a.ns))
      }
    } else {
      names.push(a.qname)
    }
  }
  (extra, names)
}

fn write_open_tag(
  el : Element,
  sb : StringBuilder,
  close : Bool,
  extra : Array[(String, String)],
  names : Array[String],
) -> Unit {
  sb.write_char('<')
  sb.write_string(el.qname)
  for decl in extra {
    if decl.0 == "" {
      sb.write_string(" xmlns=\"")
    } else {
      sb.write_string(" xmlns:")
      sb.write_string(decl.0)
      sb.write_string("=\"")
    }
    sb.write_string(escape_attr(decl.1))
    sb.write_char('"')
  }
  for i, a in el.attrs {
    sb.write_char(' ')
    sb.write_string(names[i])
    sb.write_string("=\"")
    sb.write_string(escape_attr(a.value))
    sb.write_char('"')
  }
  if close {
    sb.write_string("/>")
  } else {
    sb.write_char('>')
  }
}

fn write_node(node : XmlNode, sb : StringBuilder, scope : Map[String, String]) -> Unit {
  match node {
    Elem(el) => write_element_in(el, sb, scope)
    Text(t) => sb.write_string(escape_text(t))
    Comment(c) => {
      sb.write_string("<!--")
      sb.write_string(c)
      sb.write_string("-->")
    }
    Pi(p) => {
      sb.write_string("<?")
      sb.write_string(p)
      sb.write_string("?>")
    }
  }
}

fn write_element_in(el : Element, sb : StringBuilder, parent : Map[String, String]) -> Unit {
  let scope = scope_of(el, parent)
  let (extra, names) = declarations(el, scope)
  if el.children.is_empty() {
    write_open_tag(el, sb, true, extra, names)
    return
  }
  write_open_tag(el, sb, false, extra, names)
  let child_scope = map_clone(scope)
  for decl in extra {
    child_scope[decl.0] = decl.1
  }
  for node in el.children {
    write_node(node, sb, child_scope)
  }
  sb.write_string("</")
  sb.write_string(el.qname)
  sb.write_char('>')
}

/// Writes an element (and its subtree) into a builder.
pub fn write_element(el : Element, sb : StringBuilder) -> Unit {
  write_element_in(el, sb, Map([]))
}

/// Serializes an element to a string.
pub fn to_string(el : Element) -> String {
  let sb = StringBuilder()
  write_element(el, sb)
  sb.to_string()
}

/// Serializes a document, keeping its original declaration when present.
pub fn document_to_string(doc : Document) -> String {
  let sb = StringBuilder()
  let decl = doc.decl
  if decl.length() > 0 {
    sb.write_string("<?")
    sb.write_string(decl)
    sb.write_string("?>")
  } else {
    sb.write_string("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
  }
  write_element(doc.root, sb)
  sb.to_string()
}

/// Serializes with two-space indentation, for humans and diffing.
pub fn to_string_pretty(el : Element) -> String {
  let sb = StringBuilder()
  write_pretty(el, sb, 0, Map([]))
  sb.to_string()
}

fn indent(sb : StringBuilder, depth : Int) -> Unit {
  let mut i = 0
  while i < depth {
    sb.write_string("  ")
    i = i + 1
  }
}

fn write_pretty(
  el : Element,
  sb : StringBuilder,
  depth : Int,
  parent : Map[String, String],
) -> Unit {
  let scope = scope_of(el, parent)
  let (extra, names) = declarations(el, scope)
  let child_scope = map_clone(scope)
  for decl in extra {
    child_scope[decl.0] = decl.1
  }
  let kids = el.children
  if kids.is_empty() {
    indent(sb, depth)
    write_open_tag(el, sb, true, extra, names)
    sb.write_char('\n')
    return
  }
  indent(sb, depth)
  write_open_tag(el, sb, false, extra, names)
  let has_element = el.child_count() > 0
  if has_element {
    sb.write_char('\n')
  }
  for node in kids {
    match node {
      Elem(child) => write_pretty(child, sb, depth + 1, child_scope)
      Text(t) =>
        if t.trim().is_empty() {
          () 
        } else {
          indent(sb, depth + 1)
          sb.write_string(escape_text(t))
          sb.write_char('\n')
        }
      _ => ()
    }
  }
  if has_element {
    indent(sb, depth)
  }
  sb.write_string("</")
  sb.write_string(el.qname)
  sb.write_char('>')
  sb.write_char('\n')
}
