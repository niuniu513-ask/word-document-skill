/// Tests for the XML reader, writers and query helpers.

test "parses a document declaration and root" {
  let doc = @xml.parse(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"" +
    @xml.ns_w +
    "\"><w:body/></w:document>",
  )
  assert_eq(doc.declaration(), "xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"")
  let root = doc.root()
  assert_eq(root.local(), "document")
  assert_eq(root.namespace(), @xml.ns_w)
  assert_eq(root.child_count(), 1)
}

test "resolves nested namespace prefixes and default namespace" {
  let doc = @xml.parse("<r xmlns=\"urn:a\"><p xmlns:q=\"urn:b\"><q:c/></p><d/></r>")
  let root = doc.root()
  assert_eq(root.namespace(), "urn:a")
  let kids = root.elements()
  assert_eq(kids[0].namespace(), "urn:a")
  let inner = kids[0].elements()[0]
  assert_eq(inner.namespace(), "urn:b")
  assert_eq(inner.local(), "c")
  assert_eq(kids[1].namespace(), "urn:a")
}

test "keeps attributes in their namespace" {
  let doc = @xml.parse("<w:p xmlns:w=\"" + @xml.ns_w + "\" w:val=\"x\" plain=\"y\"/>")
  let p = doc.root()
  assert_eq(p.attr("val"), Some("x"))
  assert_eq(p.attr("plain"), Some("y"))
  let attrs = p.attributes()
  assert_eq(attrs[1].qname, "w:val")
  assert_eq(attrs[1].ns, @xml.ns_w)
}

test "decodes entities and numeric references" {
  let doc = @xml.parse("<t xmlns=\"" + @xml.ns_w + "\">a &amp; b &lt;c&gt; &#65;&#x4E2D;</t>")
  assert_eq(doc.root().text(), "a & b <c> A中")
}

test "unescapes and re-escapes without loss" {
  let raw = "a &amp; b &lt; c &quot;q&quot;"
  let decoded = @xml.unescape(raw)
  assert_eq(decoded, "a & b < c \"q\"")
  assert_eq(@xml.escape_text(decoded), "a &amp; b &lt; c \"q\"")
}

test "serializes a tree back to xml" {
  let el = @xml.parse_fragment(
    "<w:p xmlns:w=\"" + @xml.ns_w + "\"><w:r><w:t>hi</w:t></w:r></w:p>",
  )
  assert_eq(
    @xml.to_string(el),
    "<w:p xmlns:w=\"" + @xml.ns_w + "\"><w:r><w:t>hi</w:t></w:r></w:p>",
  )
}

test "keeps comments, cdata and processing instructions" {
  let doc = @xml.parse("<r><!--note--><![CDATA[x<y]]><?pi data?></r>")
  let nodes = doc.root().nodes()
  assert_eq(nodes.length(), 3)
  assert_eq(@xml.to_string(doc.root()), "<r><!--note-->x&lt;y<?pi data?></r>")
}

test "reports mismatched tags" {
  let result : Result[Unit, @xml.XmlError] = try {
    let _ = @xml.parse("<a><b></a></b>")
    Ok(())
  } catch {
    e => Err(e)
  }
  assert_true(result is Err(_))
  match result {
    Err(@xml.XmlError(msg)) => assert_true(msg.contains("mismatched"))
    _ => fail("expected XmlError")
  }
}

test "self closing root" {
  let doc = @xml.parse("<s:svg xmlns:s=\"urn:s\" />")
  assert_eq(doc.root().qname(), "s:svg")
  assert_eq(doc.root().child_count(), 0)
}

test "finds descendants in document order" {
  let doc = @xml.parse(
    "<w:body xmlns:w=\"" +
    @xml.ns_w +
    "\"><w:p><w:r><w:t>1</w:t></w:r></w:p><w:p><w:r><w:t>2</w:t></w:r></w:p></w:body>",
  )
  let ts = doc.root().find_all_w("t")
  assert_eq(ts.length(), 2)
  assert_eq(ts[0].text(), "1")
  assert_eq(ts[1].text(), "2")
}

test "integer attributes" {
  let doc = @xml.parse("<w:ind xmlns:w=\"" + @xml.ns_w + "\" w:left=\"720\" w:hanging=\"-280\"/>")
  let el = doc.root()
  assert_eq(el.attr_int("left"), Some(720))
  assert_eq(el.attr_int("hanging"), Some(-280))
  assert_eq(el.attr_int("missing"), None)
}

test "created elements keep their namespace prefix" {
  let root = @xml.parse_element(
    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>",
  )
  let body = root.child(@xml.ns_w, "body").unwrap()
  let para = @xml.Element::make(@xml.ns_w, "p")
  let run = @xml.Element::make(@xml.ns_w, "r")
  let text = @xml.Element::make(@xml.ns_w, "t")
  text.push(@xml.XmlNode::Text("TBD"))
  run.push(@xml.XmlNode::Elem(text))
  para.push(@xml.XmlNode::Elem(run))
  body.push(@xml.XmlNode::Elem(para))
  let out = @xml.to_string(root)
  inspect(out.contains("<w:t>TBD</w:t>"), content="true")
}

test "a foreign namespace prefix is declared inline" {
  let root = @xml.parse_element("<a:root xmlns:a=\"urn:a\"><a:child/></a:root>")
  let child = @xml.Element::make("urn:b", "leaf")
  root.push(@xml.XmlNode::Elem(child))
  let out = @xml.to_string(root)
  inspect(out.contains("xmlns:ns1=\"urn:b\""), content="false")
}
