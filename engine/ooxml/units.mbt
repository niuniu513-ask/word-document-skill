/// Unit conversions and small formatting helpers.
///
/// OOXML mixes five unit systems: twentieths of a point (twips) for spacing and
/// indents, half-points for font size, eighths of a point for borders, EMUs for
/// drawings, and fiftieths of a centimetre for measurement attributes.

/// Twips (1/20 pt) to points.
pub fn twips_to_pt(v : Int) -> Double {
  v.to_double() / 20.0
}

/// Twips to centimetres.
pub fn twips_to_cm(v : Int) -> Double {
  v.to_double() / 566.929
}

/// Half-points to points.
pub fn half_points_to_pt(v : Int) -> Double {
  v.to_double() / 2.0
}

/// EMUs (English Metric Units, 914400 per inch) to centimetres.
pub fn emu_to_cm(v : Int) -> Double {
  v.to_double() / 360000.0
}

/// EMUs to points.
pub fn emu_to_pt(v : Int) -> Double {
  v.to_double() / 12700.0
}

/// Fiftieths of a centimetre to centimetres.
pub fn fiftieth_cm_to_cm(v : Int) -> Double {
  v.to_double() / 50.0
}

/// Rounds to `digits` decimal places and drops trailing zeros.
pub fn fmt(value : Double, digits : Int) -> String {
  if value.is_nan() || value.is_inf() {
    return "null"
  }
  let mut scale = 1.0
  let mut i = 0
  while i < digits {
    scale = scale * 10.0
    i = i + 1
  }
  let rounded = (value * scale).round() / scale
  let text = rounded.to_string()
  if text.contains(".") {
    trim_number(text)
  } else {
    text
  }
}

fn trim_number(text : String) -> String {
  let mut end = text.length()
  let mut stop = false
  while end > 0 && !stop {
    match text.get_char(end - 1) {
      Some('0') => end = end - 1
      Some('.') => {
        end = end - 1
        stop = true
      }
      _ => stop = true
    }
  }
  text.substring(start=0, end=end)
}
