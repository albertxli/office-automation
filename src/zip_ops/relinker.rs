//! Rewrite OLE/chart link paths in PPTX .rels XML files.
//!
//! This is the critical performance optimization (GOTCHA #22):
//! - Python COM: ~90-100s for 186 links
//! - ZIP relink: ~0.12s for the same
//!
//! Fixes over Python version:
//! - Normalizes slashes in both directions (Python bug #1)
//! - Proper path-based .rels detection (Python bug #7)
//! - Better !sheet!range split handling (Python bug #8)
//! - Input validation (Python bug #9)
//! - UNC path support (Python bug #10)
//! - Per-file error handling with context (Python bug #5)

use std::io::{Read, Write, Seek};
use std::path::Path;

use crate::zip_ops::xml_stream::rewrite_xml_attributes;

/// Rewrite OLE and chart link paths in a PPTX file.
///
/// Modifies the PPTX in-place (via temp file + rename) to repoint all
/// external `file:///` references to `new_excel_path`.
///
/// Returns the number of links rewritten.
pub fn relink_pptx_zip(pptx_path: &Path, new_excel_path: &Path) -> Result<usize, String> {
    // Validate inputs
    if !pptx_path.exists() {
        return Err(format!("PPTX file not found: {}", pptx_path.display()));
    }
    if !new_excel_path.exists() {
        return Err(format!("Excel file not found: {}", new_excel_path.display()));
    }

    let new_excel_abs = new_excel_path.canonicalize()
        .map_err(|e| format!("Cannot resolve Excel path: {e}"))?;

    let new_file_uri = path_to_file_uri(&new_excel_abs);

    // Read the ZIP, rewrite .rels files, write to temp, then replace
    let pptx_data = std::fs::read(pptx_path)
        .map_err(|e| format!("Failed to read PPTX: {e}"))?;

    let mut reader = zip::ZipArchive::new(std::io::Cursor::new(&pptx_data))
        .map_err(|e| format!("Failed to open PPTX as ZIP: {e}"))?;

    let tmp_path = pptx_path.with_extension("pptx.tmp");
    let tmp_file = std::fs::File::create(&tmp_path)
        .map_err(|e| format!("Failed to create temp file: {e}"))?;
    let mut writer = zip::ZipWriter::new(tmp_file);

    let mut total_rewritten = 0;

    for i in 0..reader.len() {
        let mut entry = reader.by_index(i)
            .map_err(|e| format!("Failed to read ZIP entry {i}: {e}"))?;

        let name = entry.name().to_string();
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(entry.compression());

        if is_rels_file(&name) {
            // Read .rels XML and rewrite links
            let mut xml_data = Vec::new();
            entry.read_to_end(&mut xml_data)
                .map_err(|e| format!("Failed to read {name}: {e}"))?;

            match rewrite_rels_xml(&xml_data, &new_file_uri) {
                Ok((modified_xml, count)) => {
                    if count > 0 {
                        writer.start_file(&name, options)
                            .map_err(|e| format!("Failed to write {name}: {e}"))?;
                        writer.write_all(&modified_xml)
                            .map_err(|e| format!("Failed to write {name}: {e}"))?;
                        total_rewritten += count;
                    } else {
                        // No changes — write original
                        writer.start_file(&name, options)
                            .map_err(|e| format!("Failed to write {name}: {e}"))?;
                        writer.write_all(&xml_data)
                            .map_err(|e| format!("Failed to write {name}: {e}"))?;
                    }
                }
                Err(e) => {
                    // Per-file error handling — log and skip, don't fail entire operation
                    eprintln!("Warning: skipping {name}: {e}");
                    writer.start_file(&name, options)
                        .map_err(|e| format!("Failed to write {name}: {e}"))?;
                    writer.write_all(&xml_data)
                        .map_err(|e| format!("Failed to write {name}: {e}"))?;
                }
            }
        } else {
            // Copy non-.rels files unchanged
            writer.raw_copy_file(entry)
                .map_err(|e| format!("Failed to copy {name}: {e}"))?;
        }
    }

    writer.finish().map_err(|e| format!("Failed to finalize ZIP: {e}"))?;

    // Replace original with temp
    std::fs::rename(&tmp_path, pptx_path)
        .map_err(|e| {
            // Try to clean up temp file on failure
            let _ = std::fs::remove_file(&tmp_path);
            format!("Failed to replace PPTX: {e}")
        })?;

    Ok(total_rewritten)
}

/// Check if a ZIP entry is a .rels file in slides/ or charts/.
///
/// Proper path-based detection instead of fragile substring matching (fixes Python bug #7).
fn is_rels_file(name: &str) -> bool {
    if !name.ends_with(".rels") {
        return false;
    }

    // Normalize to forward slashes for consistent parsing
    let normalized = name.replace('\\', "/");
    let parts: Vec<&str> = normalized.split('/').collect();

    // Expected: ppt/slides/_rels/slide1.xml.rels or ppt/charts/_rels/chart1.xml.rels
    // So we look for: _rels as second-to-last, and slides or charts as third-to-last
    if parts.len() >= 3 {
        let rels_dir = parts[parts.len() - 2];
        let parent_dir = parts[parts.len() - 3];
        return rels_dir == "_rels" && (parent_dir == "slides" || parent_dir == "charts");
    }

    false
}

/// Rewrite external link targets in a .rels XML file.
///
/// Finds all `<Relationship>` elements with `TargetMode="External"` and `Target`
/// starting with `file:///`, then replaces the file path portion while preserving
/// the `!sheet!range` suffix.
fn rewrite_rels_xml(data: &[u8], new_file_uri: &str) -> Result<(Vec<u8>, usize), String> {
    // Normalize the new URI for comparison
    let new_file_uri_normalized = new_file_uri.replace('\\', "/").to_lowercase();

    rewrite_xml_attributes(data, |elem| {
        // Only process Relationship elements
        let local_name = elem.local_name();
        if local_name.as_ref() != b"Relationship" {
            return false;
        }

        let target_mode = elem.try_get_attribute("TargetMode")
            .ok()
            .flatten()
            .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string())
            .unwrap_or_default();

        if target_mode != "External" {
            return false;
        }

        let target = match elem.try_get_attribute("Target").ok().flatten() {
            Some(a) => String::from_utf8_lossy(a.value.as_ref()).to_string(),
            None => return false,
        };

        if !target.starts_with("file:///") {
            return false;
        }

        // Normalize the existing target for comparison (fixes Python bug #1: slash mismatch)
        let target_normalized = target.replace('\\', "/");

        // Split into file path and !sheet!range suffix
        // We need to be careful with '!' in filenames (Python bug #8)
        let (existing_file_part, suffix) = split_file_uri_and_suffix(&target_normalized);
        let existing_file_normalized = existing_file_part.to_lowercase();

        // Only rewrite if the file part actually differs
        if existing_file_normalized == new_file_uri_normalized {
            return false;
        }

        // Build new target: new file URI + original suffix
        let new_target = if suffix.is_empty() {
            new_file_uri.to_string()
        } else {
            format!("{new_file_uri}{suffix}")
        };

        // Rebuild all attributes, replacing Target
        let attrs: Vec<(String, String)> = elem.attributes()
            .filter_map(|a| a.ok())
            .map(|a| {
                let key = String::from_utf8_lossy(a.key.as_ref()).to_string();
                let val = if key == "Target" {
                    new_target.clone()
                } else {
                    String::from_utf8_lossy(a.value.as_ref()).to_string()
                };
                (key, val)
            })
            .collect();

        elem.clear_attributes();
        for (k, v) in attrs {
            elem.push_attribute((k.as_str(), v.as_str()));
        }

        true
    })
}

/// Split a `file:///path!Sheet!Range` URI into the file part and the suffix.
///
/// Handles filenames with `!` by looking for the pattern after the file extension.
/// The suffix always starts with `!` if present.
fn split_file_uri_and_suffix(uri: &str) -> (&str, &str) {
    // Find the last occurrence of a file extension (.xlsx, .xls, .xlsm) followed by !
    let lower = uri.to_lowercase();
    for ext in &[".xlsx!", ".xls!", ".xlsm!", ".xlsb!"] {
        if let Some(pos) = lower.find(ext) {
            let split_at = pos + ext.len() - 1; // Position of the '!' after extension
            return (&uri[..split_at], &uri[split_at..]);
        }
    }

    // Fallback: no known extension found, return entire URI as file part
    (uri, "")
}

/// Convert a Windows path to a file:/// URI.
///
/// Handles both local paths (`C:\...`) and UNC paths (`\\server\...`).
fn path_to_file_uri(path: &Path) -> String {
    let path_str = path.to_string_lossy().replace('\\', "/");

    // Strip \\?\ extended-length prefix first (before UNC check)
    let path_str = path_str.strip_prefix("//?/").unwrap_or(&path_str);

    // UNC paths: \\server\share → file://server/share
    if path_str.starts_with("//") {
        format!("file:{path_str}")
    } else {
        format!("file:///{path_str}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // --- is_rels_file tests ---

    #[test]
    fn test_is_rels_slide() {
        assert!(is_rels_file("ppt/slides/_rels/slide1.xml.rels"));
    }

    #[test]
    fn test_is_rels_chart() {
        assert!(is_rels_file("ppt/charts/_rels/chart1.xml.rels"));
    }

    #[test]
    fn test_is_rels_not_rels() {
        assert!(!is_rels_file("ppt/slides/slide1.xml"));
    }

    #[test]
    fn test_is_rels_document_level() {
        // ppt/_rels/presentation.xml.rels should NOT be processed
        assert!(!is_rels_file("ppt/_rels/presentation.xml.rels"));
    }

    #[test]
    fn test_is_rels_top_level() {
        assert!(!is_rels_file("_rels/.rels"));
    }

    #[test]
    fn test_is_rels_arbitrary_path() {
        assert!(!is_rels_file("ppt/myslides/_rels/fake.xml.rels"));
    }

    // --- split_file_uri_and_suffix tests ---

    #[test]
    fn test_split_with_suffix() {
        let (file, suffix) = split_file_uri_and_suffix(
            "file:///C:/Data/report.xlsx!Sheet1!R1C1:R5C5",
        );
        assert_eq!(file, "file:///C:/Data/report.xlsx");
        assert_eq!(suffix, "!Sheet1!R1C1:R5C5");
    }

    #[test]
    fn test_split_no_suffix() {
        let (file, suffix) = split_file_uri_and_suffix("file:///C:/Data/report.xlsx");
        assert_eq!(file, "file:///C:/Data/report.xlsx");
        assert_eq!(suffix, "");
    }

    #[test]
    fn test_split_xls_extension() {
        let (file, suffix) = split_file_uri_and_suffix(
            "file:///C:/Data/old.xls!Sheet1!A1",
        );
        assert_eq!(file, "file:///C:/Data/old.xls");
        assert_eq!(suffix, "!Sheet1!A1");
    }

    #[test]
    fn test_split_bang_in_filename() {
        // Python bug #8: filename with ! should not confuse the parser
        let (file, suffix) = split_file_uri_and_suffix(
            "file:///C:/Data/report!v2.xlsx!Sheet1!A1",
        );
        assert_eq!(file, "file:///C:/Data/report!v2.xlsx");
        assert_eq!(suffix, "!Sheet1!A1");
    }

    // --- path_to_file_uri tests ---

    #[test]
    fn test_local_path_to_uri() {
        let path = Path::new(r"C:\Users\data\report.xlsx");
        let uri = path_to_file_uri(path);
        assert_eq!(uri, "file:///C:/Users/data/report.xlsx");
    }

    #[test]
    fn test_extended_prefix_stripped() {
        let path = Path::new(r"\\?\C:\Users\data\report.xlsx");
        let uri = path_to_file_uri(path);
        assert_eq!(uri, "file:///C:/Users/data/report.xlsx");
    }

    // --- rewrite_rels_xml tests ---

    #[test]
    fn test_rewrite_rels_basic() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="file:///C:/old/data.xlsx!Tables!R1C1:R5C5" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let (output, count) = rewrite_rels_xml(xml, "file:///C:/new/report.xlsx").unwrap();
        assert_eq!(count, 1);

        let output_str = String::from_utf8(output).unwrap();
        assert!(output_str.contains("file:///C:/new/report.xlsx!Tables!R1C1:R5C5"));
        assert!(output_str.contains("../media/image1.png")); // unchanged
    }

    #[test]
    fn test_rewrite_no_external_links() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com" Target="../media/image1.png"/>
</Relationships>"#;

        let (_, count) = rewrite_rels_xml(xml, "file:///C:/new/report.xlsx").unwrap();
        assert_eq!(count, 0);
    }

    #[test]
    fn test_rewrite_preserves_suffix() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="file:///C:/old/data.xlsx!Revenue!R1C1:R100C30" TargetMode="External"/>
</Relationships>"#;

        let (output, count) = rewrite_rels_xml(xml, "file:///C:/new/report.xlsx").unwrap();
        assert_eq!(count, 1);
        let output_str = String::from_utf8(output).unwrap();
        assert!(output_str.contains("file:///C:/new/report.xlsx!Revenue!R1C1:R100C30"));
    }

    #[test]
    fn test_rewrite_same_target_no_change() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="file:///C:/same/data.xlsx!Sheet1!A1" TargetMode="External"/>
</Relationships>"#;

        let (_, count) = rewrite_rels_xml(xml, "file:///C:/same/data.xlsx").unwrap();
        assert_eq!(count, 0); // Same file, no change needed
    }

    #[test]
    fn test_rewrite_normalizes_backslashes() {
        // Python bug #1: existing target has backslashes, new URI has forward slashes
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="file:///C:\old\data.xlsx!Sheet1!R1C1" TargetMode="External"/>
</Relationships>"#;

        let (output, count) = rewrite_rels_xml(xml, "file:///C:/new/report.xlsx").unwrap();
        assert_eq!(count, 1);
        let output_str = String::from_utf8(output).unwrap();
        assert!(output_str.contains("file:///C:/new/report.xlsx!Sheet1!R1C1"));
    }

    #[test]
    fn test_rewrite_multiple_links() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="file:///C:/old/data.xlsx!Sheet1!R1C1" TargetMode="External"/>
  <Relationship Id="rId2" Target="file:///C:/old/data.xlsx!Sheet2!R1C1" TargetMode="External"/>
  <Relationship Id="rId3" Target="../media/image.png"/>
</Relationships>"#;

        let (_, count) = rewrite_rels_xml(xml, "file:///C:/new/report.xlsx").unwrap();
        assert_eq!(count, 2);
    }
}
