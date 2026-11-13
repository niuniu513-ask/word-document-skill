/// The OOXML package: a zip container holding parts addressed by name.

/// Errors raised while reading or writing a .docx package.
pub suberror DocxError {
  DocxError(String)
}

/// A Word package backed by an in-memory zip archive.
///
/// Parts are addressed by their package name (for example `word/document.xml`).
/// The archive is kept intact so that anything this library does not understand
/// survives a read/write round trip untouched.
pub struct WordPackage {
  archive : @zip.Archive
}

/// Opens a package from the raw bytes of a .docx file.
pub fn WordPackage::open(bytes : Bytes) -> WordPackage raise DocxError {
  if bytes.length() < 4 {
    raise DocxError("not a .docx file: data is too short")
  }
  let archive = try {
    @zip.read_all(bytes)
  } catch {
    _ => raise DocxError("not a valid .docx container (zip decoding failed)")
  }
  WordPackage::{ archive }
}

/// All part names in the package, in archive order.
pub fn WordPackage::parts(self : WordPackage) -> Array[String] {
  let out = []
  for entry in self.archive.entries() {
    let name = entry.name()
    if !name.has_suffix("/") {
      out.push(name)
    }
  }
  out
}

/// Number of parts in the package.
pub fn WordPackage::part_count(self : WordPackage) -> Int {
  self.parts().length()
}

/// Whether a part exists.
pub fn WordPackage::has_part(self : WordPackage, name : String) -> Bool {
  match self.archive.get(name) {
    Some(_) => true
    None => false
  }
}

/// Raw bytes of a part.
pub fn WordPackage::part_bytes(self : WordPackage, name : String) -> Bytes? {
  match self.archive.get(name) {
    Some(view) => Some(view.to_owned())
    None => None
  }
}

/// UTF-8 text of a part.
pub fn WordPackage::part_text(self : WordPackage, name : String) -> String? raise DocxError {
  match self.archive.get(name) {
    Some(view) =>
      Some(
        try {
          @utf8.decode(view)
        } catch {
          _ => raise DocxError("part " + name + " is not valid UTF-8")
        },
      )
    None => None
  }
}

/// Replaces the contents of an existing part, or adds it when missing.
pub fn WordPackage::set_part(self : WordPackage, name : String, data : Bytes) -> Unit {
  if self.archive.get(name) is Some(_) {
    ignore(self.archive.replace(name, data))
  } else {
    self.archive.add(name, data)
  }
}

/// Replaces the contents of an existing part with UTF-8 text.
pub fn WordPackage::set_part_text(self : WordPackage, name : String, text : String) -> Unit {
  self.set_part(name, @utf8.encode(text))
}

/// Serialises the whole package back to .docx bytes.
pub fn WordPackage::save(self : WordPackage) -> Bytes raise DocxError {
  try {
    @zip.write(self.archive)
  } catch {
    _ => raise DocxError("failed to write .docx container")
  }
}
