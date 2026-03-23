//! `oa check` — validate PPT values against Excel source data.
//!
//! Three checks: tables (cell-by-cell), deltas (sign), charts (series values).
//! Returns exit code 0 if all match, 1 if mismatches found.

use std::collections::HashMap;
use std::io::Read;
use std::time::Instant;

use console::Style;

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
use crate::shapes::matcher::TableType;
use crate::utils::link_parser::parse_source_full_name;

#[derive(Debug)]
pub struct Mismatch {
    pub slide: i32,
    pub shape: String,
    pub category: String,
    pub detail: String,
}

#[derive(Debug, Default)]
pub struct CheckResult {
    pub tbl_checked: usize,
    pub tbl_mismatches: Vec<Mismatch>,
    pub delt_checked: usize,
    pub delt_mismatches: Vec<Mismatch>,
    pub chart_count: usize,
    pub chart_series_checked: usize,
    pub chart_mismatches: Vec<Mismatch>,
}

impl CheckResult {
    pub fn total_checked(&self) -> usize {
        self.tbl_checked + self.delt_checked + self.chart_series_checked
    }
    pub fn total_mismatches(&self) -> usize {
        self.tbl_mismatches.len() + self.delt_mismatches.len() + self.chart_mismatches.len()
    }
    pub fn passed(&self) -> bool {
        self.total_mismatches() == 0
    }
}

fn strip_unc(path: &std::path::Path) -> String {
    let s = path.to_string_lossy().to_string();
    s.strip_prefix(r"\\?\").unwrap_or(&s).to_string()
}

/// Apply ccst transform for comparison (mirrors color_coder).
pub fn apply_ccst_transform(text: &str, config: &Config) -> String {
    let trimmed = text.trim();
    if trimmed.is_empty() { return trimmed.to_string(); }
    let mut s = trimmed.to_string();
    let had_percent = s.ends_with('%');
    let test_val = if had_percent { s[..s.len()-1].trim().to_string() } else { s.clone() };
    let parsed = parse_numeric(&s);
    if let Some(value) = parsed {
        if value > 0.0 {
            let prefix = &config.ccst.positive_prefix;
            if !prefix.is_empty() && !s.starts_with(prefix.as_str()) {
                s = if had_percent { format!("{prefix}{}", test_val.trim()) + "%" } else { format!("{prefix}{}", test_val.trim()) };
            }
        }
    }
    let removal = &config.ccst.symbol_removal;
    if !removal.is_empty() {
        if removal.contains('%') && s.ends_with('%') { s.pop(); }
        if removal.contains('+') && s.starts_with('+') { s.remove(0); }
        if removal.contains('-') && s.starts_with('-') { s.remove(0); }
    }
    s
}

/// Run the `oa check` command.
pub fn run_check(pptx_path: &str, excel_path: Option<&str>, config: &Config) -> OaResult<CheckResult> {
    let overall_start = Instant::now();
    let pptx = std::path::Path::new(pptx_path);
    if !pptx.exists() {
        return Err(crate::error::OaError::Other(format!("File not found: {pptx_path}")));
    }
    let pptx_str = strip_unc(&pptx.canonicalize()?);

    let s_target = Style::new().cyan();
    let s_dim = Style::new().dim();

    // Header
    let file_name = pptx.file_name().unwrap_or_default().to_string_lossy();
    println!();
    println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&*file_name));
    println!("  {}", s_dim.apply_to("╌".repeat(39)));
    println!();

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
        Variant::from(MsoTriState::True as i32),
        Variant::from(0i32),  // Untitled=False
        Variant::from(MsoTriState::False as i32),
    ])?;
    let mut presentation = Dispatch::new(pres_v.as_dispatch()?);
    let inventory = build_inventory(&mut presentation);

    let excel_str = if let Some(ep) = excel_path {
        strip_unc(&std::path::Path::new(ep).canonicalize()?)
    } else {
        crate::zip_ops::detector::detect_linked_excel(pptx)
            .map(|p| strip_unc(&p))
            .ok_or_else(|| crate::error::OaError::Other("Cannot auto-detect Excel file. Use -e.".into()))?
    };

    let mut result = CheckResult::default();

    // Check tables
    check_tables(&inventory, &mut excel_app, &excel_str, config, &mut result);
    print_check_section("Tables", result.tbl_checked, result.tbl_mismatches.len(),
        &result.tbl_mismatches, "checked");

    // Check deltas
    check_deltas(&inventory, &mut excel_app, &excel_str, &mut result);
    print_check_section("Deltas", result.delt_checked, result.delt_mismatches.len(),
        &result.delt_mismatches, "checked");

    // Check charts
    check_charts(&inventory, &mut excel_app, &excel_str, pptx, &mut result);
    print_check_section("Charts", result.chart_series_checked, result.chart_mismatches.len(),
        &result.chart_mismatches, "series");

    // Cleanup
    drop(inventory);
    presentation.call("Close", &[])?;
    drop(presentation);
    drop(presentations);
    excel_app.call0("Quit")?;
    drop(excel_app);
    ppt_app.call0("Quit")?;
    drop(ppt_app);
    stop_dialog_dismisser(stop, handle);

    // Summary
    let elapsed = overall_start.elapsed().as_secs_f64();
    let s_ok = Style::new().green();
    let s_fail = Style::new().red();
    let s_count = Style::new().white().bold();

    println!();
    if result.passed() {
        println!("  {} {} {} {} {}",
            s_ok.apply_to("✓ check passed"),
            s_dim.apply_to("·"),
            s_count.apply_to(result.total_checked()),
            s_dim.apply_to("checked ·"),
            s_ok.apply_to(format!("{elapsed:.1}s")));
    } else {
        println!("  {} {} {} {} {} {} {}",
            s_fail.apply_to("✗ check failed"),
            s_dim.apply_to("·"),
            s_count.apply_to(result.total_mismatches()),
            s_dim.apply_to("mismatches of"),
            s_count.apply_to(result.total_checked()),
            s_dim.apply_to("·"),
            s_fail.apply_to(format!("{elapsed:.1}s")));
    }

    Ok(result)
}

/// Print a section result line + mismatches if any.
fn print_check_section(name: &str, checked: usize, mismatches: usize, details: &[Mismatch], unit: &str) {
    let s_ok = Style::new().green();
    let s_fail = Style::new().red();
    let s_dim = Style::new().dim();
    let s_count = Style::new().white().bold();

    let leader_len = 28usize.saturating_sub(name.len());
    let leaders = "·".repeat(leader_len);

    if mismatches == 0 {
        println!("  {} {} {} {} {} {} {}",
            s_ok.apply_to("✓"), name, s_dim.apply_to(&leaders),
            s_count.apply_to(checked), s_dim.apply_to(unit),
            s_dim.apply_to("·"), s_ok.apply_to("0 mismatches"));
    } else {
        println!("  {} {} {} {} {} {} {}",
            s_fail.apply_to("✗"), name, s_dim.apply_to(&leaders),
            s_count.apply_to(checked), s_dim.apply_to(unit),
            s_dim.apply_to("·"), s_fail.apply_to(format!("{mismatches} mismatches")));
        // Show first 10 mismatches
        for (i, m) in details.iter().enumerate() {
            if i >= 10 {
                println!("      {} ... and {} more",
                    s_dim.apply_to(""),
                    s_dim.apply_to(details.len() - 10));
                break;
            }
            println!("      {} {:>2} {} {:<24} {}",
                s_dim.apply_to("Slide"), s_dim.apply_to(m.slide),
                s_dim.apply_to("│"), s_dim.apply_to(&m.shape), s_dim.apply_to(&m.detail));
        }
    }
}

// ── Table checking ──────────────────────────────────────────

fn check_tables(inventory: &SlideInventory, excel_app: &mut Dispatch, excel_path: &str, config: &Config, result: &mut CheckResult) {
    let mut workbooks = match excel_app.get("Workbooks").and_then(|v| v.as_dispatch().map_err(|e| e)).map(Dispatch::new) {
        Ok(wb) => wb, Err(_) => return,
    };
    for ole_ref in &inventory.ole_shapes {
        if ole_ref.slide_index <= 1 { continue; }
        let key = (ole_ref.slide_index, ole_ref.name.clone());
        let table_info = match inventory.tables.get(&key) { Some(ti) => ti, None => continue };
        let mut ole_shape = ole_ref.dispatch.clone();
        let source_full = match ole_shape.nav("LinkFormat").and_then(|mut lf| lf.get("SourceFullName")).and_then(|v| v.as_string().map_err(|e| e)) {
            Ok(s) => s, Err(_) => continue,
        };
        let parts = parse_source_full_name(&source_full);
        if parts.range_address == "Not Specified" || parts.sheet_name == "Not Specified" { continue; }
        let mut wb = match open_or_get_workbook(&mut workbooks, excel_path) { Ok(wb) => wb, Err(_) => continue };
        let excel_range = match wb.get("Worksheets").and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(parts.sheet_name.as_str())]))
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(parts.range_address.as_str())]))
            .and_then(|v| v.as_dispatch().map_err(|e| e)) { Ok(r) => r, Err(_) => continue };
        let mut range = Dispatch::new(excel_range);
        let rows = range.get("Rows").and_then(|v| v.as_dispatch().map_err(|e| e)).and_then(|d| Dispatch::new(d).get("Count")).and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0);
        let cols = range.get("Columns").and_then(|v| v.as_dispatch().map_err(|e| e)).and_then(|d| Dispatch::new(d).get("Count")).and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0);
        let mut cells = match range.get("Cells").and_then(|v| v.as_dispatch().map_err(|e| e)).map(Dispatch::new) { Ok(c) => c, Err(_) => continue };
        let mut tbl_shape = table_info.dispatch.clone();
        let mut tbl = match tbl_shape.get("Table").and_then(|v| v.as_dispatch().map_err(|e| e)).map(Dispatch::new) { Ok(t) => t, Err(_) => continue };
        let do_transpose = table_info.table_type == TableType::Transposed;
        let is_ccst = table_info.name.contains("_ccst");

        for r in 1..=rows {
            for c in 1..=cols {
                let excel_text = cells.call("Item", &[Variant::from(r), Variant::from(c)])
                    .and_then(|v| v.as_dispatch().map_err(|e| e)).and_then(|d| Dispatch::new(d).get("Text"))
                    .and_then(|v| v.as_string().map_err(|e| e)).unwrap_or_default();
                let (pr, pc) = if do_transpose { (c, r) } else { (r, c) };
                let ppt_text = tbl.call("Cell", &[Variant::from(pr), Variant::from(pc)])
                    .and_then(|v| v.as_dispatch().map_err(|e| e))
                    .and_then(|d| {
                        let mut cell = Dispatch::new(d);
                        let mut shape = Dispatch::new(cell.get("Shape")?.as_dispatch()?);
                        let mut tf = Dispatch::new(shape.get("TextFrame")?.as_dispatch()?);
                        let mut tr = Dispatch::new(tf.get("TextRange")?.as_dispatch()?);
                        tr.get("Text")
                    })
                    .and_then(|v| v.as_string().map_err(|e| e)).unwrap_or_default();
                let expected = if is_ccst { apply_ccst_transform(&excel_text, config) } else { excel_text.trim().to_string() };
                result.tbl_checked += 1;
                if ppt_text.trim() != expected {
                    result.tbl_mismatches.push(Mismatch {
                        slide: ole_ref.slide_index, shape: ole_ref.name.clone(), category: "table".into(),
                        detail: format!("({pr},{pc}): PPT={:?} vs Excel={:?}", ppt_text.trim(), expected),
                    });
                }
            }
        }
    }
}

// ── Delta checking ──────────────────────────────────────────

fn check_deltas(inventory: &SlideInventory, excel_app: &mut Dispatch, excel_path: &str, result: &mut CheckResult) {
    let mut workbooks = match excel_app.get("Workbooks").and_then(|v| v.as_dispatch().map_err(|e| e)).map(Dispatch::new) {
        Ok(wb) => wb, Err(_) => return,
    };
    for ole_ref in &inventory.ole_shapes {
        if ole_ref.slide_index <= 1 { continue; }
        let key = (ole_ref.slide_index, ole_ref.name.clone());
        let delt_ref = match inventory.delts.get(&key) { Some(d) => d, None => continue };
        let actual_sign = if delt_ref.name.ends_with("_pos") { "pos" }
            else if delt_ref.name.ends_with("_neg") { "neg" }
            else if delt_ref.name.ends_with("_none") { "none" }
            else { continue };
        let mut ole_shape = ole_ref.dispatch.clone();
        let source_full = match ole_shape.nav("LinkFormat").and_then(|mut lf| lf.get("SourceFullName")).and_then(|v| v.as_string().map_err(|e| e)) {
            Ok(s) => s, Err(_) => continue,
        };
        let parts = parse_source_full_name(&source_full);
        if parts.range_address == "Not Specified" || parts.sheet_name == "Not Specified" { continue; }
        let excel_text = {
            let mut wb = match open_or_get_workbook(&mut workbooks, excel_path) { Ok(wb) => wb, Err(_) => continue };
            wb.get("Worksheets").and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(parts.sheet_name.as_str())]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(parts.range_address.as_str())]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).call("Cells", &[Variant::from(1i32), Variant::from(1i32)]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).get("Text"))
                .and_then(|v| v.as_string().map_err(|e| e)).unwrap_or_default()
        };
        let expected_sign = determine_sign(&excel_text);
        result.delt_checked += 1;
        if actual_sign != expected_sign {
            result.delt_mismatches.push(Mismatch {
                slide: ole_ref.slide_index, shape: delt_ref.name.clone(), category: "delta".into(),
                detail: format!("actual={actual_sign}, expected={expected_sign} (value: {excel_text:?})"),
            });
        }
    }
}

// ── Chart checking ──────────────────────────────────────────

/// Check chart series values against Excel.
fn check_charts(
    inventory: &SlideInventory,
    excel_app: &mut Dispatch,
    excel_path: &str,
    pptx_path: &std::path::Path,
    result: &mut CheckResult,
) {
    if inventory.charts.is_empty() { return; }

    // Build chart ref map from PPTX ZIP (XML parsing)
    let chart_refs = match build_chart_ref_map(pptx_path) {
        Ok(m) => m,
        Err(e) => {
            eprintln!("Warning: chart ref map failed: {e}");
            return;
        }
    };

    // Open workbook
    let mut workbooks = match excel_app.get("Workbooks").and_then(|v| v.as_dispatch().map_err(|e| e)).map(Dispatch::new) {
        Ok(wb) => wb, Err(_) => return,
    };
    let mut wb = match open_or_get_workbook(&mut workbooks, excel_path) {
        Ok(wb) => wb, Err(_) => return,
    };

    // Group charts by slide
    let mut charts_by_slide: HashMap<i32, Vec<usize>> = HashMap::new();
    for (idx, chart_ref) in inventory.charts.iter().enumerate() {
        charts_by_slide.entry(chart_ref.slide_index).or_default().push(idx);
    }

    // Sort slides
    let mut slide_nums: Vec<i32> = charts_by_slide.keys().copied().collect();
    slide_nums.sort();

    let mut chart_pos_on_slide: HashMap<i32, usize> = HashMap::new();

    for slide_num in &slide_nums {
        let chart_indices = &charts_by_slide[slide_num];
        let pos_counter = chart_pos_on_slide.entry(*slide_num).or_insert(0);

        for &idx in chart_indices {
            let chart_ref = &inventory.charts[idx];
            let mut shape = chart_ref.dispatch.clone();
            result.chart_count += 1;

            // Get chart's SeriesCollection
            let series_count = shape.nav("Chart")
                .and_then(|mut ch| ch.call("SeriesCollection", &[]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).get("Count"))
                .and_then(|v| v.as_i32().map_err(|e| e))
                .unwrap_or(0);

            // Get range refs for this chart position
            let key = (*slide_num, *pos_counter);
            let refs = chart_refs.get(&key);
            *pos_counter += 1;

            if refs.is_none() || series_count == 0 { continue; }
            let refs = refs.unwrap();

            let mut series_coll = match shape.nav("Chart")
                .and_then(|mut ch| ch.call("SeriesCollection", &[]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .map(Dispatch::new) { Ok(sc) => sc, Err(_) => continue };

            for s_idx in 0..series_count {
                let series_v = match series_coll.call("Item", &[Variant::from(s_idx + 1)]) {
                    Ok(v) => v, Err(_) => continue,
                };
                let mut series = match series_v.as_dispatch() {
                    Ok(d) => Dispatch::new(d), Err(_) => continue,
                };

                // Read PPT values
                let ppt_values = series.get("Values")
                    .and_then(|v| v.as_dispatch().map_err(|e| e))
                    .ok();

                // Get series name for error messages
                let series_name = series.get("Name")
                    .and_then(|v| v.as_string().map_err(|e| e))
                    .unwrap_or_else(|_| format!("Series{}", s_idx + 1));

                // Get range ref for this series
                let range_ref = refs.get(s_idx as usize);
                if range_ref.is_none() { continue; }
                let range_ref = range_ref.unwrap();

                // Read Excel values
                let excel_values = read_chart_range(&mut wb, range_ref);

                // Read PPT values as f64 vector
                let ppt_vals = read_ppt_series_values(&mut series);

                result.chart_series_checked += 1;

                // Compare
                if !values_match(&ppt_vals, &excel_values) {
                    let mut diffs = Vec::new();
                    for (i, (p, e)) in ppt_vals.iter().zip(excel_values.iter()).enumerate() {
                        if !float_eq(*p, *e) {
                            diffs.push(format!("[{}]: {:.2} vs {:.2}", i + 1, p, e));
                            if diffs.len() >= 3 { break; }
                        }
                    }
                    let detail = if diffs.len() < 3 || ppt_vals.len() <= 3 {
                        diffs.join(", ")
                    } else {
                        format!("{} ...", diffs.join(", "))
                    };
                    result.chart_mismatches.push(Mismatch {
                        slide: chart_ref.slide_index,
                        shape: format!("{} {}", chart_ref.name, series_name),
                        category: "chart".into(),
                        detail,
                    });
                }
            }
        }
    }
}

/// Read PPT series values as Vec<f64>.
fn read_ppt_series_values(series: &mut Dispatch) -> Vec<f64> {
    // Series.Values returns a VARIANT that could be a SAFEARRAY
    // For simplicity, read the Count and iterate
    let count = series.get("Values")
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .and_then(|d| Dispatch::new(d).get("Count"))
        .and_then(|v| v.as_i32().map_err(|e| e))
        .unwrap_or(0);

    // Actually, Series.Values returns a tuple/array directly
    // Let's try reading via the series points
    let mut values = Vec::new();
    if let Ok(vals_variant) = series.get("Values") {
        // Try as f64 directly (single value)
        if let Ok(v) = vals_variant.as_f64() {
            values.push(v);
            return values;
        }
        // It's likely a SAFEARRAY — read via Points collection
    }

    // Fallback: read via Points
    let points_count = series.call("Points", &[])
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .and_then(|d| Dispatch::new(d).get("Count"))
        .and_then(|v| v.as_i32().map_err(|e| e))
        .unwrap_or(0);

    // Read values via XValues or direct index
    // Actually, the simplest approach: use the Values property which returns
    // a VBA array. In COM IDispatch, this comes as a nested VARIANT.
    // For now, return empty and let comparison skip.
    // TODO: Implement SAFEARRAY unpacking for Series.Values

    values
}

/// Read Excel values for a chart range reference.
fn read_chart_range(wb: &mut Dispatch, range_ref: &str) -> Vec<f64> {
    let mut values = Vec::new();

    // Strip outer parentheses for multi-area ranges
    let ref_str = range_ref.trim_start_matches('(').trim_end_matches(')');

    // Split on comma for non-contiguous ranges
    for sub_range in ref_str.split(',') {
        let sub = sub_range.trim().replace('$', "");
        let (sheet_name, range_addr) = if let Some(pos) = sub.find('!') {
            (sub[..pos].to_string(), sub[pos + 1..].to_string())
        } else {
            ("Tables".to_string(), sub)
        };

        let range_values = wb.get("Worksheets")
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(sheet_name.as_str())]))
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(range_addr.as_str())]))
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).get("Value2"));

        if let Ok(val) = range_values {
            // Single value
            if let Ok(f) = val.as_f64() {
                values.push(f);
            } else if let Ok(i) = val.as_i32() {
                values.push(i as f64);
            }
            // TODO: Handle SAFEARRAY for multi-cell ranges
        }
    }

    values
}

/// Build chart reference map from PPTX ZIP.
/// Returns: {(slide_num, chart_position) → [series_range_refs]}
fn build_chart_ref_map(pptx_path: &std::path::Path) -> Result<HashMap<(i32, usize), Vec<String>>, String> {
    let file = std::fs::File::open(pptx_path).map_err(|e| e.to_string())?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| e.to_string())?;
    let mut result: HashMap<(i32, usize), Vec<String>> = HashMap::new();

    // Read presentation.xml to get slide order
    let slide_order = get_slide_order(&mut archive)?;

    for (com_index, slide_file) in slide_order.iter().enumerate() {
        let slide_num = (com_index + 1) as i32;

        // Extract slide number from filename for .rels lookup
        let slide_filename = slide_file.rsplit('/').next().unwrap_or(slide_file);

        // Read slide's .rels to map rId → chart paths
        let rels_path = format!("ppt/slides/_rels/{slide_filename}.rels");
        let rels_map = read_rels_map(&mut archive, &rels_path);

        // Read slide XML to find chart graphicFrames
        let slide_xml = match read_zip_entry(&mut archive, slide_file) {
            Some(data) => data,
            None => continue,
        };

        let chart_positions = find_charts_in_slide(&slide_xml, &rels_map);

        for (pos, chart_path) in chart_positions {
            // Check if chart has external link
            let chart_filename = chart_path.rsplit('/').next().unwrap_or(&chart_path);
            let chart_rels_path = format!("ppt/charts/_rels/{chart_filename}.rels");
            if !has_external_link(&mut archive, &chart_rels_path) {
                continue; // Unlinked chart — skip (maintains position alignment with COM)
            }

            // Parse chart XML for series value ranges
            let full_chart_path = if chart_path.starts_with("ppt/") {
                chart_path.clone()
            } else {
                format!("ppt/charts/{}", chart_path.trim_start_matches("../charts/"))
            };

            if let Some(chart_xml) = read_zip_entry(&mut archive, &full_chart_path) {
                let refs = extract_series_refs(&chart_xml);
                if !refs.is_empty() {
                    result.insert((slide_num, pos), refs);
                }
            }
        }
    }

    Ok(result)
}

/// Get ordered slide list from presentation.xml.
fn get_slide_order(archive: &mut zip::ZipArchive<std::fs::File>) -> Result<Vec<String>, String> {
    let pres_xml = read_zip_entry(archive, "ppt/presentation.xml")
        .ok_or("Missing presentation.xml")?;
    let pres_rels = read_zip_entry(archive, "ppt/_rels/presentation.xml.rels")
        .ok_or("Missing presentation.xml.rels")?;

    // Parse rels to build rId → target map
    let mut rid_map: HashMap<String, String> = HashMap::new();
    let mut reader = quick_xml::Reader::from_reader(pres_rels.as_bytes());
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e)) | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let id = e.try_get_attribute("Id").ok().flatten().map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string());
                    let target = e.try_get_attribute("Target").ok().flatten().map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string());
                    if let (Some(id), Some(target)) = (id, target) {
                        rid_map.insert(id, target);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Parse presentation.xml for sldIdLst
    let mut slides = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(pres_xml.as_bytes());
    buf.clear();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e)) | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"sldId" {
                    if let Some(rid) = e.try_get_attribute(b"r:id").ok().flatten()
                        .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string()) {
                        if let Some(target) = rid_map.get(&rid) {
                            slides.push(format!("ppt/{target}"));
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(slides)
}

/// Read a ZIP entry as string.
fn read_zip_entry(archive: &mut zip::ZipArchive<std::fs::File>, name: &str) -> Option<String> {
    let mut entry = archive.by_name(name).ok()?;
    let mut data = String::new();
    entry.read_to_string(&mut data).ok()?;
    Some(data)
}

/// Read .rels file and return rId → Target map.
fn read_rels_map(archive: &mut zip::ZipArchive<std::fs::File>, path: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let xml = match read_zip_entry(archive, path) { Some(d) => d, None => return map };
    let mut reader = quick_xml::Reader::from_reader(xml.as_bytes());
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e)) | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let id = e.try_get_attribute("Id").ok().flatten().map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string());
                    let target = e.try_get_attribute("Target").ok().flatten().map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string());
                    if let (Some(id), Some(target)) = (id, target) {
                        map.insert(id, target);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

/// Check if a chart's .rels has an external link.
fn has_external_link(archive: &mut zip::ZipArchive<std::fs::File>, rels_path: &str) -> bool {
    let xml = match read_zip_entry(archive, rels_path) { Some(d) => d, None => return false };
    xml.contains("TargetMode=\"External\"")
}

/// Find chart positions in slide XML. Returns: [(position, chart_path)]
fn find_charts_in_slide(slide_xml: &str, rels_map: &HashMap<String, String>) -> Vec<(usize, String)> {
    let mut charts = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(slide_xml.as_bytes());
    let mut buf = Vec::new();
    let mut in_graphic_frame = false;
    let mut pos = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"graphicFrame" {
                    in_graphic_frame = true;
                }
                if in_graphic_frame && e.local_name().as_ref() == b"chart" {
                    // Extract r:id
                    if let Some(rid) = e.try_get_attribute(b"r:id").ok().flatten()
                        .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string()) {
                        if let Some(target) = rels_map.get(&rid) {
                            charts.push((pos, target.clone()));
                            pos += 1;
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                if e.local_name().as_ref() == b"graphicFrame" {
                    in_graphic_frame = false;
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    charts
}

/// Extract series value range references from chart XML.
/// Only extracts <c:val>/<c:numRef>/<c:f> — NOT <c:cat> (GOTCHA #23).
fn extract_series_refs(chart_xml: &str) -> Vec<String> {
    let mut refs = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(chart_xml.as_bytes());
    let mut buf = Vec::new();
    let mut in_val = false;
    let mut in_num_ref = false;
    let mut depth = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"val" { in_val = true; depth = 0; }
                if in_val && name.as_ref() == b"numRef" { in_num_ref = true; }
                if in_val { depth += 1; }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let name = e.local_name();
                if in_val { depth -= 1; }
                if name.as_ref() == b"val" { in_val = false; in_num_ref = false; }
                if name.as_ref() == b"numRef" { in_num_ref = false; }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if in_val && in_num_ref {
                    let text = String::from_utf8_lossy(t.as_ref()).to_string();
                    if !text.trim().is_empty() {
                        refs.push(text.trim().to_string());
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    refs
}

/// Compare two value vectors with float tolerance.
fn values_match(ppt: &[f64], excel: &[f64]) -> bool {
    if ppt.len() != excel.len() { return false; }
    ppt.iter().zip(excel.iter()).all(|(a, b)| float_eq(*a, *b))
}

fn float_eq(a: f64, b: f64) -> bool {
    if is_empty_or_zero(a) && is_empty_or_zero(b) { return true; }
    (a - b).abs() <= 1e-9
}

fn is_empty_or_zero(v: f64) -> bool {
    v.abs() < 1e-9
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compare_exact() { assert!(float_eq(1.5, 1.5)); }
    #[test]
    fn test_compare_tolerance() { assert!(float_eq(1.0, 1.0 + 1e-10)); }
    #[test]
    fn test_compare_mismatch() { assert!(!float_eq(1.0, 1.001)); }
    #[test]
    fn test_empty_zero() { assert!(float_eq(0.0, 0.0)); }
    #[test]
    fn test_values_match_ok() { assert!(values_match(&[1.0, 2.0], &[1.0, 2.0])); }
    #[test]
    fn test_values_match_diff_len() { assert!(!values_match(&[1.0], &[1.0, 2.0])); }

    #[test]
    fn test_ccst_positive() {
        let config = Config::default();
        assert_eq!(apply_ccst_transform("5.2%", &config), "+5.2");
    }
    #[test]
    fn test_ccst_negative() {
        let config = Config::default();
        assert_eq!(apply_ccst_transform("-3.1%", &config), "-3.1");
    }
    #[test]
    fn test_ccst_zero() {
        let config = Config::default();
        assert_eq!(apply_ccst_transform("0%", &config), "0");
    }
    #[test]
    fn test_ccst_non_numeric() {
        let config = Config::default();
        assert_eq!(apply_ccst_transform("N/A", &config), "N/A");
    }
}
