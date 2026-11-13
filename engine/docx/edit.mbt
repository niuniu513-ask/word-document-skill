/// Editing that keeps the document's own formatting.
///
/// The hard part of programmatic Word editing is that Word splits text into
/// runs at arbitrary points: the word "待补充" may be spread over three `w:r`
/// elements that each carry different formatting. A naive replace on a run will
/// silently miss the match, so every replacement here works on the *joined*
/// paragraph text and then maps the match back onto the runs it covers.

fn find_chars(hay : Array[Char], needle : Array[Char], from : Int) -> Int {
  let n = needle.length()
  if n == 0 {
    return -1
  }
  let mut i = from
  while i + n <= hay.length() {
    let mut j = 0
    let mut ok = true
    while j < n && ok {
      if hay[i + j] != needle[j] {
        ok = false
      }
      j = j + 1
    }
    if ok {
      return i
    }
    i = i + 1
  }
  -1
}

fn slice_chars(chars : Array[Char], start : Int, end : Int) -> String {
  let sb = StringBuilder()
  let mut i = start
  while i < end && i < chars.length() {
    sb.write_char(chars[i])
    i = i + 1
  }
  sb.to_string()
}

fn run_at(starts : Array[Int], lengths : Array[Int], offset : Int) -> (Int, Int) {
  let last = starts.length() - 1
  let mut i = last
  while i > 0 && starts[i] > offset {
    i = i - 1
  }
  if i == last && offset > starts[last] + lengths[last] {
    return (last, lengths[last])
  }
  (i, offset - starts[i])
}

/// Replaces every occurrence of `needle` inside one paragraph.
///
/// Returns the number of replacements. Matches that span several runs are
/// handled: the first covered run keeps the prefix plus the replacement, the
/// last one keeps the suffix, and the runs in between are emptied.
pub fn replace_text(
  para : @ooxml.Paragraph,
  needle : String,
  replacement : String,
) -> Int {
  if needle == "" {
    return 0
  }
  let needle_chars = needle.to_array()
  let runs = para.runs()
  if runs.is_empty() {
    return 0
  }
  let texts : Array[String] = []
  let starts : Array[Int] = []
  let lengths : Array[Int] = []
  let chars : Array[Char] = []
  let mut cursor = 0
  for run in runs {
    let text = run.text()
    let run_chars = text.to_array()
    texts.push(text)
    starts.push(cursor)
    lengths.push(run_chars.length())
    for ch in run_chars {
      chars.push(ch)
    }
    cursor = cursor + run_chars.length()
  }
  let originals = texts.copy()
  let occurrences : Array[Int] = []
  let mut from = 0
  while from <= chars.length() {
    let at = find_chars(chars, needle_chars, from)
    if at < 0 {
      break
    }
    occurrences.push(at)
    from = at + needle_chars.length()
  }
  let needle_length = needle_chars.length()
  let mut count = 0
  let mut k = occurrences.length() - 1
  while k >= 0 {
    let at = occurrences[k]
    let stop = at + needle_length
    let (first, first_offset) = run_at(starts, lengths, at)
    let (last, last_offset) = run_at(starts, lengths, stop)
    let first_chars = texts[first].to_array()
    let prefix = slice_chars(first_chars, 0, first_offset)
    if first == last {
      let suffix = slice_chars(first_chars, last_offset, first_chars.length())
      texts[first] = prefix + replacement + suffix
    } else {
      texts[first] = prefix + replacement
      let last_chars = texts[last].to_array()
      texts[last] = slice_chars(last_chars, last_offset, last_chars.length())
      let mut middle = first + 1
      while middle < last {
        texts[middle] = ""
        middle = middle + 1
      }
    }
    count = count + 1
    k = k - 1
  }
  if count == 0 {
    return 0
  }
  for i, run in runs {
    if texts[i] == originals[i] {
      continue
    }
    if texts[i] == "" && para.element().index_of(run.element()) >= 0 {
      ignore(@ooxml.detach(para.element(), run.element()))
    } else {
      run.set_text(texts[i])
    }
  }
  count
}

/// Replaces the whole text of a paragraph, keeping the first run's formatting.
pub fn set_paragraph_text(para : @ooxml.Paragraph, text : String) -> Unit {
  let runs = para.runs()
  if runs.is_empty() {
    let run = @xml.Element::make(@xml.ns_w, "r")
    let t = @xml.Element::make(@xml.ns_w, "t")
    t.set_attr("xml:space", "preserve")
    t.push(@xml.XmlNode::Text(text))
    run.push(@xml.XmlNode::Elem(t))
    para.element().push(@xml.XmlNode::Elem(run))
    return
  }
  let mut first = true
  for run in runs {
    if first {
      run.set_text(text)
      first = false
    } else if para.element().index_of(run.element()) >= 0 {
      ignore(@ooxml.detach(para.element(), run.element()))
    } else {
      run.set_text("")
    }
  }
}

/// Replaces text everywhere in the document, table cells included.
pub fn replace_in_document(
  doc : @ooxml.Docx,
  needle : String,
  replacement : String,
) -> Int {
  let mut count = 0
  for para in doc.paragraphs() {
    count = count + replace_text(para, needle, replacement)
  }
  count
}

/// Applies `key = value` fills, each key being a literal placeholder string.
pub fn fill(doc : @ooxml.Docx, values : Array[(String, String)]) -> Int {
  let mut count = 0
  for pair in values {
    count = count + replace_in_document(doc, pair.0, pair.1)
  }
  count
}

/// Inserts a new paragraph right after `anchor`, inheriting its layout and the
/// character formatting of its first run.
pub fn insert_paragraph_after(
  doc : @ooxml.Docx,
  anchor : @ooxml.Paragraph,
  text : String,
) -> @ooxml.Paragraph {
  let body = doc.body()
  let para = @xml.Element::make(@xml.ns_w, "p")
  match anchor.element().child(@xml.ns_w, "pPr") {
    Some(ppr) => para.push(@xml.XmlNode::Elem(ppr.copy()))
    None => ()
  }
  let run = @xml.Element::make(@xml.ns_w, "r")
  let runs = anchor.runs()
  let mut borrowed = false
  for candidate in runs {
    if borrowed {
      continue
    }
    match candidate.element().child(@xml.ns_w, "rPr") {
      Some(rpr) => {
        run.push(@xml.XmlNode::Elem(rpr.copy()))
        borrowed = true
      }
      None => ()
    }
  }
  let t = @xml.Element::make(@xml.ns_w, "t")
  t.set_attr("xml:space", "preserve")
  t.push(@xml.XmlNode::Text(text))
  run.push(@xml.XmlNode::Elem(t))
  para.push(@xml.XmlNode::Elem(run))
  let index = body.index_of(anchor.element())
  if index < 0 {
    body.push(@xml.XmlNode::Elem(para))
  } else {
    body.insert_at(index + 1, @xml.XmlNode::Elem(para))
  }
  @ooxml.Paragraph::new(para, doc.styles())
}

/// Appends a paragraph at the end of the body, inheriting the last paragraph's
/// layout so a report keeps its body formatting.
pub fn append_paragraph(doc : @ooxml.Docx, text : String) -> @ooxml.Paragraph {
  let paras = doc.body_paragraphs()
  match paras.rev() {
    [] => {
      let body = doc.body()
      let para = @xml.Element::make(@xml.ns_w, "p")
      let run = @xml.Element::make(@xml.ns_w, "r")
      let t = @xml.Element::make(@xml.ns_w, "t")
      t.set_attr("xml:space", "preserve")
      t.push(@xml.XmlNode::Text(text))
      run.push(@xml.XmlNode::Elem(t))
      para.push(@xml.XmlNode::Elem(run))
      body.push(@xml.XmlNode::Elem(para))
      @ooxml.Paragraph::new(para, doc.styles())
    }
    [last, ..] => insert_paragraph_after(doc, last, text)
  }
}
