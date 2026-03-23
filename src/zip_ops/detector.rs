//! Auto-detect linked Excel file from PPTX XML.
//!
//! Parses the PPTX ZIP to find the first external `file:///` link target
//! in slide relationship files.

use std::io::Read;
use std::path::{Path, PathBuf};

/// Detect the linked Excel file from a PPTX's slide relationship files.
///
/// Searches `slides/_rels/*.rels` for external `file:///` targets and returns
/// the first one found. Does NOT require the file to exist on disk (unlike Python).
///
/// Returns `None` if no external Excel link is found or the PPTX is invalid.
pub fn detect_linked_excel(pptx_path: &Path) -> Option<PathBuf> {
    let file = std::fs::File::open(pptx_path).ok()?;
    let mut archive = zip::ZipArchive::new(file).ok()?;

    for i in 0..archive.len() {
        let entry = archive.by_index(i).ok()?;
        let name = entry.name().to_string();

        // Only look in slides/_rels/ (not charts/_rels/)
        if !name.ends_with(".rels") {
            continue;
        }
        let normalized = name.replace('\\', "/");
        let parts: Vec<&str> = normalized.split('/').collect();
        if parts.len() < 3 {
            continue;
        }
        let rels_dir = parts[parts.len() - 2];
        let parent_dir = parts[parts.len() - 3];
        if rels_dir != "_rels" || parent_dir != "slides" {
            continue;
        }

        // Drop the entry and re-read to avoid borrow issues
        drop(entry);
        let mut entry = archive.by_index(i).ok()?;
        let mut xml_data = Vec::new();
        entry.read_to_end(&mut xml_data).ok()?;

        if let Some(path) = extract_excel_path_from_rels(&xml_data) {
            return Some(path);
        }
    }

    None
}

/// Parse a .rels XML to find the first external file:/// target pointing to an Excel file.
fn extract_excel_path_from_rels(data: &[u8]) -> Option<PathBuf> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(data);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() != b"Relationship" {
                    continue;
                }

                let target_mode = e.try_get_attribute("TargetMode")
                    .ok()
                    .flatten()
                    .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string());

                if target_mode.as_deref() != Some("External") {
                    continue;
                }

                let target = e.try_get_attribute("Target")
                    .ok()
                    .flatten()
                    .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string())?;

                if !target.starts_with("file:///") {
                    continue;
                }

                // Extract file path: strip file:/// prefix and !sheet!range suffix
                let path_str = &target["file:///".len()..];

                // Find the extension end (handle ! in filenames)
                let path_str = strip_suffix_after_extension(path_str);

                // Decode percent-encoded characters (e.g., %20 → space)
                let path_str = percent_decode(&path_str);

                // Normalize slashes
                let path_str = path_str.replace('/', std::path::MAIN_SEPARATOR_STR);

                return Some(PathBuf::from(path_str));
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => continue,
        }
    }

    None
}

/// Decode percent-encoded characters in a URI path (e.g., `%20` → space).
///
/// Only decodes common characters found in file URIs. This is simpler and
/// more predictable than a full URI decoder.
fn percent_decode(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let Ok(byte) = u8::from_str_radix(
                std::str::from_utf8(&bytes[i + 1..i + 3]).unwrap_or(""),
                16,
            ) {
                result.push(byte as char);
                i += 3;
                continue;
            }
        }
        result.push(bytes[i] as char);
        i += 1;
    }
    result
}

/// Strip the !Sheet!Range suffix from a path, handling ! in filenames.
fn strip_suffix_after_extension(path: &str) -> &str {
    let lower = path.to_lowercase();
    for ext in &[".xlsx", ".xls", ".xlsm", ".xlsb"] {
        if let Some(pos) = lower.find(ext) {
            let end = pos + ext.len();
            return &path[..end];
        }
    }
    // No known extension — fall back to first ! as separator
    if let Some(pos) = path.find('!') {
        &path[..pos]
    } else {
        path
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_strip_suffix_xlsx() {
        assert_eq!(
            strip_suffix_after_extension("C:/Data/report.xlsx!Sheet1!R1C1"),
            "C:/Data/report.xlsx"
        );
    }

    #[test]
    fn test_strip_suffix_no_suffix() {
        assert_eq!(
            strip_suffix_after_extension("C:/Data/report.xlsx"),
            "C:/Data/report.xlsx"
        );
    }

    #[test]
    fn test_strip_suffix_xls() {
        assert_eq!(
            strip_suffix_after_extension("C:/Data/old.xls!Sheet1!A1"),
            "C:/Data/old.xls"
        );
    }

    #[test]
    fn test_strip_suffix_bang_in_name() {
        assert_eq!(
            strip_suffix_after_extension("C:/Data/report!v2.xlsx!Sheet1"),
            "C:/Data/report!v2.xlsx"
        );
    }

    #[test]
    fn test_strip_suffix_no_extension() {
        assert_eq!(
            strip_suffix_after_extension("C:/Data/report!Sheet1"),
            "C:/Data/report"
        );
    }

    #[test]
    fn test_extract_from_rels_xml() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="file:///C:/Users/data/report.xlsx!Tables!R1C1:R5C5" TargetMode="External"/>
</Relationships>"#;

        let result = extract_excel_path_from_rels(xml);
        assert!(result.is_some());
        let path = result.unwrap();
        assert!(path.to_string_lossy().contains("report.xlsx"));
    }

    #[test]
    fn test_extract_no_external_links() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="../media/image.png"/>
</Relationships>"#;

        assert!(extract_excel_path_from_rels(xml).is_none());
    }

    #[test]
    fn test_percent_decode() {
        assert_eq!(percent_decode("hello%20world"), "hello world");
        assert_eq!(percent_decode("no%20change%21"), "no change!");
        assert_eq!(percent_decode("plain"), "plain");
        assert_eq!(percent_decode("100%25"), "100%");
        assert_eq!(percent_decode("~%20PMI"), "~ PMI");
    }

    #[test]
    fn test_detect_linked_excel_on_real_pptx() {
        let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("quick_test_files")
            .join("test_template.pptx");
        if !path.exists() {
            return;
        }

        let result = detect_linked_excel(&path);
        // test_template.pptx has OLE links, so we should find something
        assert!(result.is_some(), "Expected to find linked Excel in test_template.pptx");
        let excel_path = result.unwrap();
        println!("Detected Excel: {}", excel_path.display());
        assert!(excel_path.to_string_lossy().contains(".xlsx"));
    }
}
