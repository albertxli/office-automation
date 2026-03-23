//! `oa check` — validate PPT values against Excel source data.
//!
//! Cell-by-cell comparison of tables, delta sign verification.
//! Returns exit code 0 if all match, 1 if mismatches found.

use crate::com::dispatch::Dispatch;
use crate::com::session::{create_instance, init_com_sta, spawn_dialog_dismisser, stop_dialog_dismisser};
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::office::constants::MsoTriState;
use crate::pipeline::color_coder::parse_numeric;
use crate::pipeline::delta_updater::determine_sign;
use crate::pipeline::table_updater::open_or_get_workbook;
use crate::shapes::inventory::{build_inventory, SlideInventory};
use crate::shapes::matcher::{strip_sign_suffix, TableType};
use crate::utils::link_parser::parse_source_full_name;

/// A single mismatch found during validation.
#[derive(Debug)]
pub struct Mismatch {
    pub slide: i32,
    pub shape: String,
    pub category: String,
    pub detail: String,
}

/// Results from running the check.
#[derive(Debug, Default)]
pub struct CheckResult {
    pub tbl_checked: usize,
    pub tbl_mismatches: Vec<Mismatch>,
    pub delt_checked: usize,
    pub delt_mismatches: Vec<Mismatch>,
}

impl CheckResult {
    pub fn total_checked(&self) -> usize {
        self.tbl_checked + self.delt_checked
    }

    pub fn all_mismatches(&self) -> Vec<&Mismatch> {
        self.tbl_mismatches.iter().chain(self.delt_mismatches.iter()).collect()
    }

    pub fn passed(&self) -> bool {
        self.tbl_mismatches.is_empty() && self.delt_mismatches.is_empty()
    }
}

fn strip_unc(path: &std::path::Path) -> String {
    let s = path.to_string_lossy().to_string();
    s.strip_prefix(r"\\?\").unwrap_or(&s).to_string()
}

/// Apply the same transformation that color_coder does, so we can compare
/// the expected PPT text against the actual PPT text for _ccst tables.
pub fn apply_ccst_transform(text: &str, config: &Config) -> String {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return trimmed.to_string();
    }

    let mut s = trimmed.to_string();
    let had_percent = s.ends_with('%');

    // Strip % for numeric test
    let test_val = if had_percent {
        s[..s.len() - 1].trim().to_string()
    } else {
        s.clone()
    };

    let parsed = parse_numeric(&s);

    match parsed {
        Some(value) if value > 0.0 => {
            // Add positive prefix
            let prefix = &config.ccst.positive_prefix;
            if !prefix.is_empty() && !s.starts_with(prefix.as_str()) {
                if had_percent {
                    s = format!("{prefix}{}", test_val.trim());
                    s.push('%');
                } else {
                    s = format!("{prefix}{}", test_val.trim());
                }
            }
        }
        _ => {}
    }

    // Symbol removal (mirrors color_coder)
    let removal = &config.ccst.symbol_removal;
    if !removal.is_empty() {
        if removal.contains('%') && s.ends_with('%') {
            s.pop();
        }
        if removal.contains('+') && s.starts_with('+') {
            s.remove(0);
        }
        if removal.contains('-') && s.starts_with('-') {
            s.remove(0);
        }
    }

    s
}

/// Compare two cell texts with whitespace tolerance.
fn compare_cell_text(ppt: &str, excel: &str) -> bool {
    ppt.trim() == excel.trim()
}

/// Run the `oa check` command.
pub fn run_check(pptx_path: &str, excel_path: Option<&str>, config: &Config) -> OaResult<CheckResult> {
    let pptx = std::path::Path::new(pptx_path);
    if !pptx.exists() {
        return Err(crate::error::OaError::Other(format!("File not found: {pptx_path}")));
    }
    let pptx_str = strip_unc(&pptx.canonicalize()?);

    let _com = init_com_sta()?;
    let (stop, handle) = spawn_dialog_dismisser();

    let mut excel_app = create_instance("Excel.Application")?;
    excel_app.put("Visible", Variant::from(false))?;
    excel_app.put("DisplayAlerts", Variant::from(false))?;

    let mut ppt_app = create_instance("PowerPoint.Application")?;
    ppt_app.put("DisplayAlerts", Variant::from(0i32))?;

    let mut presentations = Dispatch::new(ppt_app.get("Presentations")?.as_dispatch()?);
    let pres_v = presentations.call("Open", &[
        Variant::from(pptx_str.as_str()),
        Variant::from(MsoTriState::True as i32),  // ReadOnly
        Variant::from(MsoTriState::True as i32),  // Untitled
        Variant::from(MsoTriState::False as i32), // WithWindow
    ])?;
    let mut presentation = Dispatch::new(pres_v.as_dispatch()?);
    let inventory = build_inventory(&mut presentation);

    // Determine Excel path
    let excel_str = if let Some(ep) = excel_path {
        strip_unc(&std::path::Path::new(ep).canonicalize()?)
    } else {
        crate::zip_ops::detector::detect_linked_excel(pptx)
            .map(|p| strip_unc(&p))
            .ok_or_else(|| crate::error::OaError::Other("Cannot auto-detect Excel file. Use -e.".into()))?
    };

    let mut result = CheckResult::default();

    // Check tables (cell-by-cell)
    check_tables(&inventory, &mut excel_app, &excel_str, config, &mut result);

    // Check deltas (sign verification)
    check_deltas(&inventory, &mut excel_app, &excel_str, &mut result);

    // Cleanup (GOTCHA #21)
    drop(inventory);
    presentation.call("Close", &[])?;
    drop(presentation);
    drop(presentations);
    excel_app.call0("Quit")?;
    drop(excel_app);
    ppt_app.call0("Quit")?;
    drop(ppt_app);
    stop_dialog_dismisser(stop, handle);

    // Print results
    println!();
    if !result.tbl_mismatches.is_empty() {
        println!("TABLE MISMATCHES ({}):", result.tbl_mismatches.len());
        for m in &result.tbl_mismatches {
            println!("  Slide {}, {}: {}", m.slide, m.shape, m.detail);
        }
    }
    if !result.delt_mismatches.is_empty() {
        println!("DELTA MISMATCHES ({}):", result.delt_mismatches.len());
        for m in &result.delt_mismatches {
            println!("  Slide {}, {}: {}", m.slide, m.shape, m.detail);
        }
    }

    println!();
    if result.passed() {
        println!("CHECK PASSED: {} tables, {} deltas verified",
            result.tbl_checked, result.delt_checked);
    } else {
        println!("CHECK FAILED: {} table mismatches, {} delta mismatches (of {} checked)",
            result.tbl_mismatches.len(),
            result.delt_mismatches.len(),
            result.total_checked());
    }

    Ok(result)
}

/// Check table cell values against Excel — cell by cell.
fn check_tables(
    inventory: &SlideInventory,
    excel_app: &mut Dispatch,
    excel_path: &str,
    config: &Config,
    result: &mut CheckResult,
) {
    let mut workbooks = match excel_app.get("Workbooks")
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .map(Dispatch::new)
    {
        Ok(wb) => wb,
        Err(_) => return,
    };

    for ole_ref in &inventory.ole_shapes {
        // Skip template slide
        if ole_ref.slide_index <= 1 {
            continue;
        }

        let key = (ole_ref.slide_index, ole_ref.name.clone());
        let table_info = match inventory.tables.get(&key) {
            Some(ti) => ti,
            None => continue,
        };

        // Get link parts for sheet/range
        let mut ole_shape = ole_ref.dispatch.clone();
        let source_full = match ole_shape.nav("LinkFormat")
            .and_then(|mut lf| lf.get("SourceFullName"))
            .and_then(|v| v.as_string().map_err(|e| e))
        {
            Ok(s) => s,
            Err(_) => continue,
        };

        let parts = parse_source_full_name(&source_full);
        if parts.range_address == "Not Specified" || parts.sheet_name == "Not Specified" {
            continue;
        }

        // Open Excel workbook and get range (use CLI excel_path, GOTCHA #29)
        let mut wb = match open_or_get_workbook(&mut workbooks, excel_path) {
            Ok(wb) => wb,
            Err(_) => continue,
        };

        let excel_range = match wb.get("Worksheets")
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(parts.sheet_name.as_str())]))
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(parts.range_address.as_str())]))
            .and_then(|v| v.as_dispatch().map_err(|e| e))
        {
            Ok(r) => r,
            Err(_) => continue,
        };

        let mut range = Dispatch::new(excel_range);
        let excel_rows = range.get("Rows")
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).get("Count"))
            .and_then(|v| v.as_i32().map_err(|e| e))
            .unwrap_or(0);
        let excel_cols = range.get("Columns")
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).get("Count"))
            .and_then(|v| v.as_i32().map_err(|e| e))
            .unwrap_or(0);

        let mut cells = match range.get("Cells")
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .map(Dispatch::new)
        {
            Ok(c) => c,
            Err(_) => continue,
        };

        // Get PPT table
        let mut tbl_shape = table_info.dispatch.clone();
        let mut tbl = match tbl_shape.get("Table")
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .map(Dispatch::new)
        {
            Ok(t) => t,
            Err(_) => continue,
        };

        let do_transpose = table_info.table_type == TableType::Transposed;
        let is_ccst = table_info.name.contains("_ccst");

        // Cell-by-cell comparison
        for r in 1..=excel_rows {
            for c in 1..=excel_cols {
                // Read Excel cell text
                let excel_text = cells.call("Item", &[Variant::from(r), Variant::from(c)])
                    .and_then(|v| v.as_dispatch().map_err(|e| e))
                    .and_then(|d| Dispatch::new(d).get("Text"))
                    .and_then(|v| v.as_string().map_err(|e| e))
                    .unwrap_or_default();

                // Apply transpose for trns_ tables
                let (ppt_r, ppt_c) = if do_transpose { (c, r) } else { (r, c) };

                // Read PPT cell text
                let ppt_text = tbl.call("Cell", &[Variant::from(ppt_r), Variant::from(ppt_c)])
                    .and_then(|v| v.as_dispatch().map_err(|e| e))
                    .and_then(|d| Dispatch::new(d).nav("Shape.TextFrame.TextRange"))
                    .and_then(|mut tr| tr.get("Text"))
                    .and_then(|v| v.as_string().map_err(|e| e))
                    .unwrap_or_default();

                // For _ccst tables: apply the same transform color_coder does
                let expected = if is_ccst {
                    apply_ccst_transform(&excel_text, config)
                } else {
                    excel_text.trim().to_string()
                };

                result.tbl_checked += 1;

                if !compare_cell_text(&ppt_text, &expected) {
                    result.tbl_mismatches.push(Mismatch {
                        slide: ole_ref.slide_index,
                        shape: ole_ref.name.clone(),
                        category: "table".into(),
                        detail: format!("Cell ({ppt_r},{ppt_c}): PPT={:?} vs Expected={:?}",
                            ppt_text.trim(), expected),
                    });
                }
            }
        }
    }
}

/// Check delta indicator signs against Excel values.
fn check_deltas(
    inventory: &SlideInventory,
    excel_app: &mut Dispatch,
    excel_path: &str,
    result: &mut CheckResult,
) {
    let mut workbooks = match excel_app.get("Workbooks")
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .map(Dispatch::new)
    {
        Ok(wb) => wb,
        Err(_) => return,
    };

    for ole_ref in &inventory.ole_shapes {
        if ole_ref.slide_index <= 1 {
            continue;
        }

        let key = (ole_ref.slide_index, ole_ref.name.clone());
        let delt_ref = match inventory.delts.get(&key) {
            Some(d) => d,
            None => continue,
        };

        // Extract actual sign from shape name suffix
        let actual_sign = if delt_ref.name.ends_with("_pos") {
            "pos"
        } else if delt_ref.name.ends_with("_neg") {
            "neg"
        } else if delt_ref.name.ends_with("_none") {
            "none"
        } else {
            continue; // No sign suffix — can't verify
        };

        // Get expected sign from Excel
        let mut ole_shape = ole_ref.dispatch.clone();
        let source_full = match ole_shape.nav("LinkFormat")
            .and_then(|mut lf| lf.get("SourceFullName"))
            .and_then(|v| v.as_string().map_err(|e| e))
        {
            Ok(s) => s,
            Err(_) => continue,
        };

        let parts = parse_source_full_name(&source_full);
        if parts.range_address == "Not Specified" || parts.sheet_name == "Not Specified" {
            continue;
        }

        let excel_text = {
            let mut wb = match open_or_get_workbook(&mut workbooks, excel_path) {
                Ok(wb) => wb,
                Err(_) => continue,
            };

            wb.get("Worksheets")
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(parts.sheet_name.as_str())]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(parts.range_address.as_str())]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| {
                    let mut range = Dispatch::new(d);
                    range.call("Cells", &[Variant::from(1i32), Variant::from(1i32)])
                })
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).get("Text"))
                .and_then(|v| v.as_string().map_err(|e| e))
                .unwrap_or_default()
        };

        let expected_sign = determine_sign(&excel_text);

        result.delt_checked += 1;

        if actual_sign != expected_sign {
            result.delt_mismatches.push(Mismatch {
                slide: ole_ref.slide_index,
                shape: delt_ref.name.clone(),
                category: "delta".into(),
                detail: format!("Actual={actual_sign}, Expected={expected_sign} (Excel value: {excel_text:?})"),
            });
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // --- compare_cell_text ---

    #[test]
    fn test_compare_exact() {
        assert!(compare_cell_text("12.5%", "12.5%"));
    }

    #[test]
    fn test_compare_mismatch() {
        assert!(!compare_cell_text("12.5%", "12.3%"));
    }

    #[test]
    fn test_compare_whitespace() {
        assert!(compare_cell_text("  12.5%  ", "12.5%"));
    }

    #[test]
    fn test_compare_empty() {
        assert!(compare_cell_text("", "  "));
    }

    // --- apply_ccst_transform ---

    #[test]
    fn test_ccst_positive_default_config() {
        let config = Config::default(); // prefix="+", removal="%" (only % removed)
        // "5.2%" → add prefix → "+5.2%" → remove % → "+5.2"
        assert_eq!(apply_ccst_transform("5.2%", &config), "+5.2");
    }

    #[test]
    fn test_ccst_negative_default_config() {
        let config = Config::default();
        // "-3.1%" → no prefix added → remove % → "-3.1"
        assert_eq!(apply_ccst_transform("-3.1%", &config), "-3.1");
    }

    #[test]
    fn test_ccst_zero() {
        let config = Config::default();
        // "0%" → no prefix (not >0) → remove % → "0"
        assert_eq!(apply_ccst_transform("0%", &config), "0");
    }

    #[test]
    fn test_ccst_non_numeric() {
        let config = Config::default();
        assert_eq!(apply_ccst_transform("N/A", &config), "N/A");
    }

    #[test]
    fn test_ccst_empty() {
        let config = Config::default();
        assert_eq!(apply_ccst_transform("", &config), "");
    }

    #[test]
    fn test_ccst_no_removal() {
        let mut config = Config::default();
        config.ccst.symbol_removal = String::new();
        config.ccst.positive_prefix = String::new();
        assert_eq!(apply_ccst_transform("5.2%", &config), "5.2%");
    }

    #[test]
    fn test_ccst_already_has_prefix() {
        let config = Config::default();
        // "+5.2%" already has prefix → don't double → remove % → "+5.2"
        assert_eq!(apply_ccst_transform("+5.2%", &config), "+5.2");
    }

    #[test]
    fn test_ccst_full_removal() {
        // When symbol_removal includes all three: %, +, -
        let mut config = Config::default();
        config.ccst.symbol_removal = "%+-".into();
        // "5.2%" → "+5.2%" → remove % → "+5.2" → remove + → "5.2"
        assert_eq!(apply_ccst_transform("5.2%", &config), "5.2");
        // "-3.1%" → remove % → "-3.1" → remove - → "3.1"
        assert_eq!(apply_ccst_transform("-3.1%", &config), "3.1");
    }

    #[test]
    fn test_ccst_no_prefix_no_percent() {
        let config = Config::default();
        // "5.2" (no %) → add prefix → "+5.2" → no % to remove → "+5.2"
        assert_eq!(apply_ccst_transform("5.2", &config), "+5.2");
    }
}
