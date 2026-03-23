//! Shape name classification and token matching.
//!
//! Classifies shapes by their name prefix (ntbl_, htmp_, trns_, delt_, _ccst)
//! and provides exact token matching for associating shapes with OLE objects.

/// The type of special shape, determined by name prefix.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapePrefix {
    /// Normal table (ntbl_) — preserves formatting across runs.
    NormalTable,
    /// Heatmap table (htmp_) — recalculates 3-color scale on each run.
    Heatmap,
    /// Transposed table (trns_) — swaps rows and columns.
    Transposed,
    /// Delta indicator (delt_) — arrow shape indicating value sign.
    Delta,
    /// Color-coded table (_ccst) — sign-based cell coloring.
    ColorCoded,
}

/// Table type for priority ordering (ntbl > htmp > trns).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum TableType {
    Normal,     // ntbl_ — highest priority
    Heatmap,    // htmp_
    Transposed, // trns_ — lowest priority
}

/// Classify a shape name by its prefix. Returns None for unrecognized names.
pub fn classify_shape_name(name: &str) -> Option<ShapePrefix> {
    if name.contains("ntbl_") {
        Some(ShapePrefix::NormalTable)
    } else if name.contains("htmp_") {
        Some(ShapePrefix::Heatmap)
    } else if name.contains("trns_") {
        Some(ShapePrefix::Transposed)
    } else if name.contains("delt_") {
        Some(ShapePrefix::Delta)
    } else if name.contains("_ccst") {
        Some(ShapePrefix::ColorCoded)
    } else {
        None
    }
}

/// Get the table type from a prefix (only for table-type prefixes).
pub fn prefix_to_table_type(prefix: ShapePrefix) -> Option<TableType> {
    match prefix {
        ShapePrefix::NormalTable => Some(TableType::Normal),
        ShapePrefix::Heatmap => Some(TableType::Heatmap),
        ShapePrefix::Transposed => Some(TableType::Transposed),
        _ => None,
    }
}

/// Check if `linked_name` appears as a complete token in `shape_name`.
///
/// A token boundary is: start/end of string, or any non-alphanumeric character
/// (underscore, space, hyphen, etc.).
///
/// # Examples
/// ```
/// use office_automation::shapes::matcher::is_exact_token_match;
///
/// assert!(is_exact_token_match("ntbl_Revenue", "Revenue"));
/// assert!(is_exact_token_match("ntbl_Revenue_Q4", "Revenue"));
/// assert!(!is_exact_token_match("ntbl_RevenueTotal", "Revenue"));
/// ```
pub fn is_exact_token_match(shape_name: &str, linked_name: &str) -> bool {
    if linked_name.is_empty() {
        return false;
    }

    let linked_len = linked_name.len();
    let name_len = shape_name.len();
    let mut pos = 0;

    while let Some(found) = shape_name[pos..].find(linked_name) {
        let abs_pos = pos + found;

        // Check character before match
        let before_ok = abs_pos == 0 || !shape_name.as_bytes()[abs_pos - 1].is_ascii_alphanumeric();

        // Check character after match
        let end_pos = abs_pos + linked_len;
        let after_ok = end_pos >= name_len || !shape_name.as_bytes()[end_pos].is_ascii_alphanumeric();

        if before_ok && after_ok {
            return true;
        }

        pos = abs_pos + 1;
        if pos >= name_len {
            break;
        }
    }

    false
}

/// Strip the sign suffix (_pos, _neg, _none) from a delta shape name.
pub fn strip_sign_suffix(name: &str) -> &str {
    for suffix in &["_pos", "_neg", "_none"] {
        if let Some(stripped) = name.strip_suffix(suffix) {
            return stripped;
        }
    }
    name
}

#[cfg(test)]
mod tests {
    use super::*;

    // --- classify_shape_name tests ---

    #[test]
    fn test_classify_ntbl() {
        assert_eq!(classify_shape_name("ntbl_Revenue"), Some(ShapePrefix::NormalTable));
    }

    #[test]
    fn test_classify_htmp() {
        assert_eq!(classify_shape_name("htmp_Scores"), Some(ShapePrefix::Heatmap));
    }

    #[test]
    fn test_classify_trns() {
        assert_eq!(classify_shape_name("trns_Matrix"), Some(ShapePrefix::Transposed));
    }

    #[test]
    fn test_classify_delt() {
        assert_eq!(classify_shape_name("delt_Growth_pos"), Some(ShapePrefix::Delta));
    }

    #[test]
    fn test_classify_ccst() {
        assert_eq!(classify_shape_name("table_ccst"), Some(ShapePrefix::ColorCoded));
    }

    #[test]
    fn test_classify_unknown() {
        assert_eq!(classify_shape_name("regular_shape"), None);
    }

    #[test]
    fn test_classify_priority_ntbl_over_htmp() {
        // If a name somehow contains both, ntbl_ wins (checked first)
        assert_eq!(classify_shape_name("ntbl_htmp_test"), Some(ShapePrefix::NormalTable));
    }

    // --- is_exact_token_match tests ---

    #[test]
    fn test_token_match_exact() {
        assert!(is_exact_token_match("ntbl_Revenue", "Revenue"));
    }

    #[test]
    fn test_token_match_middle() {
        assert!(is_exact_token_match("ntbl_Revenue_Q4", "Revenue"));
    }

    #[test]
    fn test_token_no_match_partial() {
        // "Revenue" should NOT match "RevenueTotal" (no boundary after)
        assert!(!is_exact_token_match("ntbl_RevenueTotal", "Revenue"));
    }

    #[test]
    fn test_token_no_match_partial_before() {
        // "Revenue" should NOT match "TotalRevenue" (no boundary before)
        assert!(!is_exact_token_match("ntbl_TotalRevenue", "Revenue"));
    }

    #[test]
    fn test_token_match_at_start() {
        assert!(is_exact_token_match("Revenue_table", "Revenue"));
    }

    #[test]
    fn test_token_match_at_end() {
        assert!(is_exact_token_match("ntbl_Revenue", "Revenue"));
    }

    #[test]
    fn test_token_match_whole_string() {
        assert!(is_exact_token_match("Revenue", "Revenue"));
    }

    #[test]
    fn test_token_match_underscore_boundary() {
        assert!(is_exact_token_match("data_Revenue_2024", "Revenue"));
    }

    #[test]
    fn test_token_match_hyphen_boundary() {
        assert!(is_exact_token_match("data-Revenue-2024", "Revenue"));
    }

    #[test]
    fn test_token_match_space_boundary() {
        assert!(is_exact_token_match("data Revenue 2024", "Revenue"));
    }

    #[test]
    fn test_token_no_match_empty_linked() {
        assert!(!is_exact_token_match("ntbl_Revenue", ""));
    }

    #[test]
    fn test_token_no_match_not_found() {
        assert!(!is_exact_token_match("ntbl_Revenue", "Costs"));
    }

    // --- strip_sign_suffix tests ---

    #[test]
    fn test_strip_pos() {
        assert_eq!(strip_sign_suffix("delt_Growth_pos"), "delt_Growth");
    }

    #[test]
    fn test_strip_neg() {
        assert_eq!(strip_sign_suffix("delt_Growth_neg"), "delt_Growth");
    }

    #[test]
    fn test_strip_none() {
        assert_eq!(strip_sign_suffix("delt_Growth_none"), "delt_Growth");
    }

    #[test]
    fn test_strip_no_suffix() {
        assert_eq!(strip_sign_suffix("delt_Growth"), "delt_Growth");
    }

    // --- prefix_to_table_type tests ---

    #[test]
    fn test_table_type_priority() {
        // Normal < Heatmap < Transposed (lower = higher priority)
        assert!(TableType::Normal < TableType::Heatmap);
        assert!(TableType::Heatmap < TableType::Transposed);
    }
}
