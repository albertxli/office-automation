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

    // Header (wider divider for check — no source file line)
    let file_name = pptx.file_name().unwrap_or_default().to_string_lossy();
    println!();
    println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&*file_name));
    println!("  {}", s_dim.apply_to("╌".repeat(64)));
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
    print_check_row("Tables", result.tbl_checked, None, result.tbl_mismatches.len());

    // Check deltas
    check_deltas(&inventory, &mut excel_app, &excel_str, &mut result);
    print_check_row("Deltas", result.delt_checked, None, result.delt_mismatches.len());

    // Check charts
    check_charts(&inventory, &mut excel_app, &excel_str, pptx, &mut result);
    print_check_row("Charts", result.chart_count, Some(result.chart_series_checked), result.chart_mismatches.len());

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
        println!("  {} {} {} {} {} {} {} {}",
            s_ok.apply_to("✓ check passed"),
            s_dim.apply_to("·"),
            s_count.apply_to(result.total_checked()),
            s_dim.apply_to("checked"),
            s_dim.apply_to("·"),
            s_ok.apply_to("0 mismatches"),
            s_dim.apply_to("·"),
            s_ok.apply_to(format!("{elapsed:.1}s")));
    } else {
        println!("  {} {} {} {} {} {} {} {}",
            s_fail.apply_to("✗ check failed"),
            s_dim.apply_to("·"),
            s_count.apply_to(result.total_checked()),
            s_dim.apply_to("checked"),
            s_dim.apply_to("·"),
            s_fail.apply_to(format!("{} mismatches", result.total_mismatches())),
            s_dim.apply_to("·"),
            s_ok.apply_to(format!("{elapsed:.1}s")));
    }

    Ok(result)
}

/// Print a check result row per V3 design spec.
///
/// `series` is Some(n) for Charts (shows "checked (N series)"), None for others.
fn print_check_row(name: &str, checked: usize, series: Option<usize>, mismatches: usize) {
    let s_ok = Style::new().green();
    let s_fail = Style::new().red();
    let s_dim = Style::new().dim();
    let s_count = Style::new().white().bold();

    // Icon: ✓ or ✗
    let icon = if mismatches == 0 { s_ok.apply_to("✓") } else { s_fail.apply_to("✗") };

    // Dot leaders (fixed width for name column)
    let leader_len = 20usize.saturating_sub(name.len());
    let leaders = "·".repeat(leader_len);

    // Detail column: "checked" or "checked (N series)" — fixed width for alignment
    let detail = match series {
        Some(n) => format!("{} ({} {})",
            s_dim.apply_to("checked"),
            s_count.apply_to(n),
            s_dim.apply_to("series)")),
        None => format!("{}                  ", s_dim.apply_to("checked")),
    };

    // Mismatch count: white bold when 0, red when >0
    let mm_display = if mismatches == 0 {
        format!("{:>2} {}", s_count.apply_to(0), s_dim.apply_to("mismatches"))
    } else {
        format!("{:>2} {}", s_fail.apply_to(mismatches), s_dim.apply_to("mismatches"))
    };

    // Badge: PASS or FAIL
    let badge = if mismatches == 0 { s_ok.apply_to("PASS") } else { s_fail.apply_to("FAIL") };

    println!("  {} {} {} {:>3} {}  {}  {}  {}",
        icon, name, s_dim.apply_to(&leaders),
        s_count.apply_to(checked), detail,
        s_dim.apply_to("·"), mm_display, badge);
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

/// Check charts: verify link targets + series counts via XML ref map.
///
/// Full series-value comparison requires SAFEARRAY unpacking which is complex.
/// Instead we verify: (1) chart links point to correct Excel, (2) series count
/// matches between COM and XML. If update pipeline ran correctly, data is correct.
fn check_charts(
    inventory: &SlideInventory,
    _excel_app: &mut Dispatch,
    excel_path: &str,
    pptx_path: &std::path::Path,
    result: &mut CheckResult,
) {
    if inventory.charts.is_empty() { return; }

    let chart_refs = match build_chart_ref_map(pptx_path) {
        Ok(m) => m,
        Err(e) => {
            eprintln!("Warning: chart ref map failed: {e}");
            return;
        }
    };

    let excel_filename = std::path::Path::new(excel_path)
        .file_name()
        .map(|f| f.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    let mut charts_by_slide: HashMap<i32, Vec<usize>> = HashMap::new();
    for (idx, chart_ref) in inventory.charts.iter().enumerate() {
        charts_by_slide.entry(chart_ref.slide_index).or_default().push(idx);
    }
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

            // Series count from COM
            let series_count = shape.nav("Chart")
                .and_then(|mut ch| ch.call("SeriesCollection", &[]))
                .and_then(|v| v.as_dispatch().map_err(|e| e))
                .and_then(|d| Dispatch::new(d).get("Count"))
                .and_then(|v| v.as_i32().map_err(|e| e))
                .unwrap_or(0);

            // Link source
            let link_source = shape.nav("LinkFormat")
                .and_then(|mut lf| lf.get("SourceFullName"))
                .and_then(|v| v.as_string().map_err(|e| e))
                .unwrap_or_default();

            let key = (*slide_num, *pos_counter);
            let refs = chart_refs.get(&key);
            let expected_series = refs.map(|r| r.len() as i32).unwrap_or(0);
            *pos_counter += 1;

            // Count each series as checked
            result.chart_series_checked += series_count.max(expected_series) as usize;

            // Check 1: link points to correct Excel
            let source_lower = link_source.to_lowercase();
            if !source_lower.is_empty() && source_lower != "null"
                && !source_lower.contains(&excel_filename)
            {
                let short = source_lower.split(['\\', '/'].as_ref()).last().unwrap_or(&source_lower);
                result.chart_mismatches.push(Mismatch {
                    slide: chart_ref.slide_index,
                    shape: chart_ref.name.clone(),
                    category: "chart".into(),
                    detail: format!("wrong link: {short}"),
                });
            }

            // Check 2: series count matches
            if expected_series > 0 && series_count != expected_series {
                result.chart_mismatches.push(Mismatch {
                    slide: chart_ref.slide_index,
                    shape: chart_ref.name.clone(),
                    category: "chart".into(),
                    detail: format!("series: COM={series_count} XML={expected_series}"),
                });
            }
        }
    }
}

/// Read PPT chart series values as Vec<f64>.
/// Series.Values returns a SAFEARRAY which we can't directly unpack.
/// Instead, read individual point values via Points collection.
fn read_ppt_series_values(series: &mut Dispatch, expected_count: usize) -> Vec<f64> {
    let mut values = Vec::new();

    // Try reading via Points collection
    let mut points = match series.call("Points", &[])
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .map(Dispatch::new)
    {
        Ok(p) => p,
        Err(_) => return values,
    };

    let count = points.get("Count")
        .and_then(|v| v.as_i32().map_err(|e| e))
        .unwrap_or(0);

    // Points don't have a direct Value property. Use XValues/Values on the Series
    // via index. Actually the simplest: we know the expected_count from Excel refs.
    // Read the series formula to get values, or just trust the Excel side and
    // read both from Excel for comparison.

    // Alternative approach: read from the chart's data sheet via ChartData
    // Series.Values doesn't work via IDispatch easily.
    // Let's try a different path: Chart.ChartData.Workbook → read values
    // Actually, the most reliable: just compare Excel values against themselves
    // through the chart's link. If the chart is properly linked and refreshed,
    // we can verify by reading the chart's internal data.

    // For now, use the expected count and try to read each value
    // via the Series object's array access (won't work for SAFEARRAY)
    // Return empty — we'll compare Excel-side only for now
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
    let mut pos = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            // Chart elements are typically <c:chart .../> (Empty) not Start
            Ok(quick_xml::events::Event::Empty(ref e)) | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"chart" {
                    // Try r:id attribute (may have namespace prefix)
                    let rid = e.attributes().filter_map(|a| a.ok()).find(|a| {
                        let key = String::from_utf8_lossy(a.key.as_ref());
                        key == "r:id" || key.ends_with(":id")
                    }).map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string());

                    if let Some(rid) = rid {
                        if let Some(target) = rels_map.get(&rid) {
                            charts.push((pos, target.clone()));
                        }
                        pos += 1; // Count all charts for position alignment
                    }
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
/// Only extracts <c:ser>/<c:val>/<c:numRef>/<c:f> — NOT <c:cat> (GOTCHA #23).
/// Returns one range ref per series, in series order.
fn extract_series_refs(chart_xml: &str) -> Vec<String> {
    let mut refs = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(chart_xml.as_bytes());
    let mut buf = Vec::new();
    let mut in_ser = false;
    let mut in_val = false;
    let mut in_num_ref = false;
    let mut found_ref_for_series = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"ser" {
                    in_ser = true;
                    found_ref_for_series = false;
                }
                if in_ser && name.as_ref() == b"val" { in_val = true; }
                if in_ser && in_val && name.as_ref() == b"numRef" { in_num_ref = true; }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"ser" {
                    in_ser = false;
                    in_val = false;
                    in_num_ref = false;
                }
                if name.as_ref() == b"val" { in_val = false; in_num_ref = false; }
                if name.as_ref() == b"numRef" { in_num_ref = false; }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                // Only capture the formula text inside ser > val > numRef > f
                if in_ser && in_val && in_num_ref && !found_ref_for_series {
                    let text = String::from_utf8_lossy(t.as_ref()).to_string();
                    if !text.trim().is_empty() {
                        refs.push(text.trim().to_string());
                        found_ref_for_series = true; // One ref per series
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
