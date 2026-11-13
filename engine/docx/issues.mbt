/// The rule engine: turns raw document facts into actionable findings.
///
/// Every issue carries a stable `kind`, a severity, a human message in Chinese
/// (the working language of the reports this engine targets) and, when it can be
/// attributed, the paragraph index a fix would touch.

/// Severity of a finding.
pub(all) enum Severity {
  Info
  Warn
  Error
} derive(Debug)

/// What a severity is called in JSON output.
pub fn Severity::label(self : Severity) -> String {
  match self {
    Info => "info"
    Warn => "warn"
    Error => "error"
  }
}

/// One finding.
pub struct Issue {
  kind : String
  severity : Severity
  message : String
  paragraph_index : Int?
  count : Int
}

/// Creates an issue.
pub fn Issue::new(
  kind : String,
  severity : Severity,
  message : String,
  paragraph_index : Int?,
  count : Int,
) -> Issue {
  Issue::{ kind, severity, message, paragraph_index, count }
}

/// JSON form of an issue.
pub fn Issue::to_json(self : Issue) -> @json.Json {
  jobj([
    ("kind", jstr(self.kind)),
    ("severity", jstr(self.severity.label())),
    ("message", jstr(self.message)),
    ("paragraph_index", jopt_int(self.paragraph_index)),
    ("count", jint(self.count)),
  ])
}

/// Characters that begin a Chinese ordinal list item.
let cn_ordinals : String = "一二三四五六七八九十"

/// Detects hand-typed numbering such as "一、", "1.", "(二)" or "第3章".
///
/// Word's own numbering is a property of the paragraph; a number that lives in
/// the text is invisible to the outline and cannot be renumbered, so it is a
/// reliable smell of a document that was formatted by hand.
pub fn manual_number_prefix(text : String) -> String? {
  let trimmed = text.trim().to_owned()
  if trimmed.is_empty() {
    return None
  }
  let chars = trimmed.to_array()
  let first = chars[0].to_string()
  if cn_ordinals.contains(first) && chars.length() > 1 {
    return Some(first)
  }
  // "1." / "1、" / "1．" / "1)"
  let mut digits = 0
  while digits < chars.length() && chars[digits] >= '0' && chars[digits] <= '9' {
    digits = digits + 1
  }
  if digits > 0 && digits < chars.length() {
    let sep = chars[digits].to_string()
    if sep == "." || sep == "、" || sep == "．" || sep == ")" || sep == "）" {
      return Some(trimmed.substring(start=0, end=digits + 1))
    }
  }
  None
}

/// Whether a text contains a placeholder the author still has to fill in.
pub fn placeholder_spans(text : String) -> Array[String] {
  let out : Array[String] = []
  let mut buffer = StringBuilder()
  let mut inside = false
  let mut underscores = 0
  for ch in text {
    if ch == '【' {
      inside = true
      ignore(buffer)
      buffer = StringBuilder()
    } else if ch == '】' && inside {
      out.push("【" + buffer.to_string() + "】")
      inside = false
      buffer = StringBuilder()
    } else if inside {
      buffer.write_char(ch)
    } else if ch == '_' || ch == '＿' {
      underscores = underscores + 1
    } else {
      if underscores >= 4 {
        out.push("____")
      }
      underscores = 0
    }
  }
  if underscores >= 4 {
    out.push("____")
  }
  out
}
