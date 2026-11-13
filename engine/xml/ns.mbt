/// Well-known namespace URIs used by OOXML and the XML spec.

pub let ns_w : String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
pub let ns_r : String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
pub let ns_wp : String = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
pub let ns_a : String = "http://schemas.openxmlformats.org/drawingml/2006/main"
pub let ns_pic : String = "http://schemas.openxmlformats.org/drawingml/2006/picture"
pub let ns_mc : String = "http://schemas.openxmlformats.org/markup-compatibility/2006"
pub let ns_w14 : String = "http://schemas.microsoft.com/office/word/2010/wordml"
pub let ns_w15 : String = "http://schemas.microsoft.com/office/word/2012/wordml"
pub let ns_v : String = "urn:schemas-microsoft-com:vml"
pub let ns_o : String = "urn:schemas-microsoft-com:office:office"
pub let ns_m : String = "http://schemas.openxmlformats.org/officeDocument/2006/math"
pub let ns_rel_pkg : String = "http://schemas.openxmlformats.org/package/2006/relationships"
pub let ns_ct : String = "http://schemas.openxmlformats.org/package/2006/content-types"
pub let ns_cp : String = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
pub let ns_dc : String = "http://purl.org/dc/elements/1.1/"
pub let ns_dcterms : String = "http://purl.org/dc/terms/"

/// Canonical prefix for a namespace URI, used when creating new markup.
pub fn prefix_for(ns : String) -> String {
  match ns {
    _ if ns == ns_w => "w"
    _ if ns == ns_r => "r"
    _ if ns == ns_wp => "wp"
    _ if ns == ns_a => "a"
    _ if ns == ns_pic => "pic"
    _ if ns == ns_mc => "mc"
    _ if ns == ns_m => "m"
    _ if ns == ns_v => "v"
    _ if ns == ns_o => "o"
    _ if ns == ns_xml => "xml"
    _ => ""
  }
}
