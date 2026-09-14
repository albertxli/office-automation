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
    } else if delta_set(name).is_some() {
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

/// Extract the delta template-set number from a shape name.
///
/// Recognises `delt` followed by an optional run of ASCII digits and an underscore,
/// anywhere in the name (mirrors the historical `contains("delt_")` check):
/// - `delt_Rev`   → `Some(1)` (default set)
/// - `delt1_Rev`  → `Some(1)` (explicit alias of set 1)
/// - `delt2_Rev`  → `Some(2)`, `delt12_Rev` → `Some(12)`
/// - `delt0_Rev`, `delta_Rev`, `delt2Rev` → `None`
///
/// Set N pairs with templates named by [`template_name_for_set`].
pub fn delta_set(name: &str) -> Option<u32> {
    let bytes = name.as_bytes();
    let mut search = 0;

    while let Some(found) = name[search..].find("delt") {
        let start = search + found + 4; // byte index just past "delt"
        let digits_end = bytes[start..]
            .iter()
            .position(|b| !b.is_ascii_digit())
            .map_or(bytes.len(), |p| start + p);

        if bytes.get(digits_end) == Some(&b'_') {
            let digits = &name[start..digits_end];
            if digits.is_empty() {
                return Some(1);
            }
            if let Ok(n) = digits.parse::<u32>()
                && n >= 1
            {
                return Some(n);
            }
        }

        search = search + found + 1;
        if search >= bytes.len() {
            break;
        }
    }

    None
}

/// The OLE-facing token of a special shape name: type prefix and sign suffix removed.
/// `delt2_globalnet_f_pos` → `globalnet_f`, `ntbl_Object_edu` → `Object_edu`, `htmp_x` → `x`.
pub fn special_token(name: &str) -> &str {
    let mut rest = name;
    if let Some(pos) = rest.find("delt") {
        let after = &rest[pos + 4..];
        let digits = after.bytes().take_while(|b| b.is_ascii_digit()).count();
        if after[digits..].starts_with('_') {
            rest = &after[digits + 1..];
        }
    } else {
        for prefix in ["ntbl_", "htmp_", "trns_"] {
            if let Some(r) = rest.strip_prefix(prefix) {
                rest = r;
                break;
            }
        }
    }
    strip_sign_suffix(rest)
}

/// Levenshtein distance over chars (small inputs only — shape names).
fn levenshtein(a: &str, b: &str) -> usize {
    let a: Vec<char> = a.chars().collect();
    let b: Vec<char> = b.chars().collect();
    let mut prev: Vec<usize> = (0..=b.len()).collect();
    let mut cur = vec![0; b.len() + 1];
    for (i, ca) in a.iter().enumerate() {
        cur[0] = i + 1;
        for (j, cb) in b.iter().enumerate() {
            let cost = usize::from(ca != cb);
            cur[j + 1] = (prev[j + 1] + 1).min(cur[j] + 1).min(prev[j] + cost);
        }
        std::mem::swap(&mut prev, &mut cur);
    }
    prev[b.len()]
}

/// The candidate closest to `target` when it is plausibly a typo of it: edit distance ≤ 2
/// and no more than half the target's length. Ties → longest shared prefix, then first.
/// Used for hints like `delt2_globalnet_f … (closest: globalnet_g)` (GOTCHA #49).
pub fn closest_name<'a>(target: &str, candidates: &[&'a str]) -> Option<&'a str> {
    let target_len = target.chars().count();
    if target.is_empty() || target_len == 0 {
        return None;
    }
    let max_d = 2usize.min((target_len / 2).max(1));
    let common_prefix = |c: &str| target.chars().zip(c.chars()).take_while(|(x, y)| x == y).count();
    let mut best: Option<(&'a str, usize, usize)> = None; // (name, distance, shared prefix)
    for &c in candidates {
        let d = levenshtein(target, c);
        if d > max_d {
            continue;
        }
        let p = common_prefix(c);
        let better = match best {
            None => true,
            Some((_, bd, bp)) => d < bd || (d == bd && p > bp),
        };
        if better {
            best = Some((c, d, p));
        }
    }
    best.map(|(c, _, _)| c)
}

/// Derive the template shape name for a given delta set.
///
/// - Set 1 returns `base` unchanged (the configured `tmpl_delta_*` name).
/// - Set N ≥ 2 replaces a leading `tmpl_` with `tmpl<N>_`, e.g.
///   `tmpl_delta_pos` → `tmpl2_delta_pos`. If `base` has no `tmpl_` prefix
///   (custom override), falls back to `tmpl<N>_delta_<sign>`.
pub fn template_name_for_set(base: &str, set: u32, sign: &str) -> String {
    if set <= 1 {
        return base.to_string();
    }
    match base.strip_prefix("tmpl_") {
        Some(rest) => format!("tmpl{set}_{rest}"),
        None => format!("tmpl{set}_delta_{sign}"),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── GOTCHA #49: unpaired special shapes ────────────────

    #[test]
    fn test_special_token() {
        assert_eq!(special_token("delt2_globalnet_f"), "globalnet_f");
        assert_eq!(special_token("delt2_globalnet_f_pos"), "globalnet_f");
        assert_eq!(special_token("delt_marketnet_none"), "marketnet");
        assert_eq!(special_token("delt12_Rev_DE_neg"), "Rev_DE");
        assert_eq!(special_token("ntbl_Object_edu"), "Object_edu");
        assert_eq!(special_token("htmp_Heat_1"), "Heat_1");
        assert_eq!(special_token("trns_T"), "T");
        assert_eq!(special_token("Object_plain"), "Object_plain");
    }

    #[test]
    fn test_levenshtein() {
        assert_eq!(levenshtein("globalnet_f", "globalnet_g"), 1);
        assert_eq!(levenshtein("abc", "abc"), 0);
        assert_eq!(levenshtein("", "abc"), 3);
        assert_eq!(levenshtein("kitten", "sitting"), 3);
    }

    #[test]
    fn test_closest_name_typo_hint() {
        let oles = ["globalnet_a", "globalnet_b", "globalnet_c", "globalnet_d", "globalnet_e", "globalnet_g"];
        // all are distance 1 from globalnet_f — tie broken by shared prefix, then first
        assert_eq!(closest_name("globalnet_f", &oles), Some("globalnet_a"));
        assert_eq!(closest_name("globalnet_f", &["marketnet_f", "globalnet_g"]), Some("globalnet_g"));
        assert_eq!(closest_name("globalnet_f", &["Object_edu", "Object_age"]), None, "nothing near");
        assert_eq!(closest_name("globalnet_f", &[]), None);
        assert_eq!(closest_name("ab", &["xy"]), None, "short names need distance ≤ 1");
        assert_eq!(closest_name("ab", &["ac"]), Some("ac"));
    }

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

    #[test]
    fn test_strip_numbered_set() {
        assert_eq!(strip_sign_suffix("delt2_Growth_neg"), "delt2_Growth");
        assert_eq!(strip_sign_suffix("delt12_Growth_none"), "delt12_Growth");
    }

    // --- delta_set tests ---

    #[test]
    fn test_delta_set_default() {
        assert_eq!(delta_set("delt_Growth_pos"), Some(1));
        assert_eq!(delta_set("delt_"), Some(1));
    }

    #[test]
    fn test_delta_set_explicit_one_is_alias() {
        assert_eq!(delta_set("delt1_Growth_pos"), Some(1));
    }

    #[test]
    fn test_delta_set_numbered() {
        assert_eq!(delta_set("delt2_Growth_pos"), Some(2));
        assert_eq!(delta_set("delt9_Growth"), Some(9));
        assert_eq!(delta_set("delt12_Growth_none"), Some(12));
        assert_eq!(delta_set("delt250_X"), Some(250));
    }

    #[test]
    fn test_delta_set_anywhere_in_name() {
        // Mirrors the old contains("delt_") semantics
        assert_eq!(delta_set("Group delt_Growth"), Some(1));
        assert_eq!(delta_set("Xdelt2_Growth"), Some(2));
    }

    #[test]
    fn test_delta_set_zero_rejected() {
        assert_eq!(delta_set("delt0_Growth"), None);
    }

    #[test]
    fn test_delta_set_not_a_delta() {
        assert_eq!(delta_set("delta_Growth"), None);
        assert_eq!(delta_set("delt2Growth"), None);
        assert_eq!(delta_set("delt"), None);
        assert_eq!(delta_set("ntbl_Growth"), None);
        assert_eq!(delta_set(""), None);
    }

    #[test]
    fn test_delta_set_skips_bad_then_finds_good() {
        // First "delt" has no underscore after digits; second one is valid
        assert_eq!(delta_set("deltX_delt2_Growth"), Some(2));
    }

    #[test]
    fn test_classify_numbered_delt() {
        assert_eq!(classify_shape_name("delt2_Growth_pos"), Some(ShapePrefix::Delta));
        assert_eq!(classify_shape_name("delt1_Growth_pos"), Some(ShapePrefix::Delta));
        assert_eq!(classify_shape_name("delt0_Growth_pos"), None);
    }

    // --- template_name_for_set tests ---

    #[test]
    fn test_template_name_set_one_unchanged() {
        assert_eq!(template_name_for_set("tmpl_delta_pos", 1, "pos"), "tmpl_delta_pos");
        assert_eq!(template_name_for_set("custom_up", 1, "pos"), "custom_up");
    }

    #[test]
    fn test_template_name_numbered() {
        assert_eq!(template_name_for_set("tmpl_delta_pos", 2, "pos"), "tmpl2_delta_pos");
        assert_eq!(template_name_for_set("tmpl_delta_neg", 2, "neg"), "tmpl2_delta_neg");
        assert_eq!(template_name_for_set("tmpl_delta_none", 12, "none"), "tmpl12_delta_none");
    }

    #[test]
    fn test_template_name_custom_base_falls_back() {
        assert_eq!(template_name_for_set("custom_up", 2, "pos"), "tmpl2_delta_pos");
    }

    // --- prefix_to_table_type tests ---

    #[test]
    fn test_table_type_priority() {
        // Normal < Heatmap < Transposed (lower = higher priority)
        assert!(TableType::Normal < TableType::Heatmap);
        assert!(TableType::Heatmap < TableType::Transposed);
    }
}
