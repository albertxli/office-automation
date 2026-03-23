//! Parse OLE SourceFullName strings into components.

use crate::utils::cell_ref::r1c1_to_a1;

/// Parsed OLE link components.
#[derive(Debug, Clone, PartialEq)]
pub struct LinkParts {
    pub file_path: String,
    pub sheet_name: String,
    pub range_address: String,
}

/// Parse an OLE `SourceFullName` into its components.
///
/// Format: `filepath!sheetname!R1C1:R5C5`
///
/// Returns "Not Specified" for missing parts. Converts the range from R1C1 to A1 notation.
pub fn parse_source_full_name(source_full_name: &str) -> LinkParts {
    let parts: Vec<&str> = source_full_name.split('!').collect();

    let file_path = if !parts.is_empty() && !parts[0].is_empty() {
        parts[0].to_string()
    } else {
        "Not Specified".to_string()
    };

    let sheet_name = if parts.len() >= 2 && !parts[1].is_empty() {
        parts[1].to_string()
    } else {
        "Not Specified".to_string()
    };

    let range_address = if parts.len() >= 3 && !parts[2].is_empty() {
        r1c1_to_a1(parts[2]).unwrap_or_else(|_| parts[2].to_string())
    } else {
        "Not Specified".to_string()
    };

    LinkParts {
        file_path,
        sheet_name,
        range_address,
    }
}

/// Extract just the file path from a SourceFullName (everything before first `!`).
pub fn extract_file_path(source_full_name: &str) -> String {
    source_full_name
        .split('!')
        .next()
        .unwrap_or("Not Specified")
        .to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_full_link() {
        let parts = parse_source_full_name(
            r"C:\Data\report.xlsx!Sheet1!R1C1:R5C5",
        );
        assert_eq!(parts.file_path, r"C:\Data\report.xlsx");
        assert_eq!(parts.sheet_name, "Sheet1");
        assert_eq!(parts.range_address, "A1:E5");
    }

    #[test]
    fn test_file_and_sheet_only() {
        let parts = parse_source_full_name(r"C:\Data\report.xlsx!Tables");
        assert_eq!(parts.file_path, r"C:\Data\report.xlsx");
        assert_eq!(parts.sheet_name, "Tables");
        assert_eq!(parts.range_address, "Not Specified");
    }

    #[test]
    fn test_file_only() {
        let parts = parse_source_full_name(r"C:\Data\report.xlsx");
        assert_eq!(parts.file_path, r"C:\Data\report.xlsx");
        assert_eq!(parts.sheet_name, "Not Specified");
        assert_eq!(parts.range_address, "Not Specified");
    }

    #[test]
    fn test_empty_string() {
        let parts = parse_source_full_name("");
        assert_eq!(parts.file_path, "Not Specified");
        assert_eq!(parts.sheet_name, "Not Specified");
        assert_eq!(parts.range_address, "Not Specified");
    }

    #[test]
    fn test_range_conversion() {
        let parts = parse_source_full_name("file.xlsx!Sheet!R1C27:R10C30");
        assert_eq!(parts.range_address, "AA1:AD10");
    }

    #[test]
    fn test_extract_file_path() {
        assert_eq!(
            extract_file_path(r"C:\Data\report.xlsx!Sheet1!R1C1"),
            r"C:\Data\report.xlsx"
        );
    }

    #[test]
    fn test_extract_file_path_no_exclamation() {
        assert_eq!(
            extract_file_path(r"C:\Data\report.xlsx"),
            r"C:\Data\report.xlsx"
        );
    }

    #[test]
    fn test_network_path() {
        let parts = parse_source_full_name(
            r"\\server\share\data.xlsx!Revenue!R1C1:R50C10",
        );
        assert_eq!(parts.file_path, r"\\server\share\data.xlsx");
        assert_eq!(parts.sheet_name, "Revenue");
        assert_eq!(parts.range_address, "A1:J50");
    }
}
