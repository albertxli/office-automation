//! `oa check` — validate PPT values against Excel source data.
//!
//! Three checks: tables (cell-by-cell), deltas (sign), charts (series values).
//! Returns exit code 0 if all match, 1 if mismatches found.

use std::collections::HashMap;
use std::io::Read;
use std::time::Instant;

use console::Style;

use crate::com::dispatch::Dispatch;
use crate::com::session::{create_instance, init_com_sta, spawn_dialog_dismisser, stop_dialog_dismisser, ComSession};
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::{OaError, OaResult};
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
pub fn run_check(pptx_path: &str, excel_path: Option<&str>, config: &Config, verbose: bool) -> OaResult<CheckResult> {
    let overall_start = Instant::now();
    crate::pipeline::verbose::set_verbose(verbose);
    let pptx = std::path::Path::new(pptx_path);
    if !pptx.exists() {
        return Err(crate::error::OaError::Other(format!("File not found: {pptx_path}")));
    }
    let pptx_str = strip_unc(&pptx.canonicalize()?);

    let s_target = Style::new().cyan();
    let s_dim = Style::new().dim();

    // Header with data file path
    let file_name = pptx.file_name().unwrap_or_default().to_string_lossy();
    let s_source = Style::new().yellow();
    println!();
    println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&*file_name));

    // Resolve Excel path BEFORE COM init (fast fail on multi-link or missing file)
    let excel_str = if let Some(ep) = excel_path {
        strip_unc(&std::path::Path::new(ep).canonicalize()?)
    } else {
        let all_excels = crate::zip_ops::detector::detect_all_linked_excels(pptx);
        if all_excels.is_empty() {
            return Err(crate::error::OaError::Other("Cannot auto-detect Excel file. Use -e.".into()));
        }
        if all_excels.len() > 1 {
            let names: Vec<String> = all_excels.iter()
                .filter_map(|p| p.file_name().map(|f| f.to_string_lossy().to_string()))
                .collect();
            return Err(crate::error::OaError::Other(format!(
                "Multiple Excel links found ({}). Use -e to specify which file to check against.",
                names.join(", ")
            )));
        }
        let detected = &all_excels[0];
        if !detected.exists() {
            return Err(crate::error::OaError::Other(format!(
                "Auto-detected Excel file not found: {}. Use -e to specify.",
                detected.display()
            )));
        }
        strip_unc(&detected.canonicalize()?)
    };

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

    // Finish header with data path + divider
    let excel_filename = std::path::Path::new(&excel_str)
        .file_name()
        .map(|f| f.to_string_lossy().to_string())
        .unwrap_or_else(|| excel_str.clone());
    println!("    {} {}", s_source.apply_to("←"), s_source.apply_to(&excel_filename));
    println!("  {}", s_dim.apply_to("╌".repeat(64)));
    println!();

    let mut result = CheckResult::default();

    // Check tables
    let sp = if !verbose { Some(make_check_spinner("Tables")) } else { None };
    check_tables(&inventory, &mut excel_app, &excel_str, config, &mut result);
    if let Some(sp) = sp { sp.finish_and_clear(); }

    let sp = if !verbose { Some(make_check_spinner("Deltas")) } else { None };
    check_deltas(&inventory, &mut excel_app, &excel_str, &mut result);
    if let Some(sp) = sp { sp.finish_and_clear(); }

    let sp = if !verbose { Some(make_check_spinner("Charts")) } else { None };
    check_charts(&inventory, &mut excel_app, &excel_str, pptx, &mut result);
    if let Some(sp) = sp { sp.finish_and_clear(); }

    // Summary table (always shown — in verbose mode it's the only summary)
    println!();
    print_check_row("Tables", result.tbl_checked, None, result.tbl_mismatches.len());
    print_check_row("Deltas", result.delt_checked, None, result.delt_mismatches.len());
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

/// Run `oa check` using an existing COM session (for batch reuse).
pub fn run_check_with_session(
    session: &mut ComSession,
    pptx_path: &str,
    excel_path: &str,
    config: &Config,
    verbose: bool,
) -> OaResult<CheckResult> {
    let overall_start = Instant::now();
    crate::pipeline::verbose::set_verbose(verbose);
    let pptx = std::path::Path::new(pptx_path);
    if !pptx.exists() {
        return Err(OaError::Other(format!("File not found: {pptx_path}")));
    }
    let pptx_str = strip_unc(&pptx.canonicalize()?);
    let excel_str = strip_unc(&std::path::Path::new(excel_path).canonicalize()?);

    let s_target = Style::new().cyan();
    let s_source = Style::new().yellow();
    let s_dim = Style::new().dim();

    let file_name = pptx.file_name().unwrap_or_default().to_string_lossy();
    let excel_filename = std::path::Path::new(&excel_str)
        .file_name()
        .map(|f| f.to_string_lossy().to_string())
        .unwrap_or_else(|| excel_str.clone());

    println!();
    println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&*file_name));
    println!("    {} {}", s_source.apply_to("←"), s_source.apply_to(&excel_filename));
    println!("  {}", s_dim.apply_to("╌".repeat(64)));
    println!();

    // Open presentation using session's PPT app
    let mut presentations = Dispatch::new(session.ppt_app.get("Presentations")?.as_dispatch()?);
    let pres_v = presentations.call("Open", &[
        Variant::from(pptx_str.as_str()),
        Variant::from(MsoTriState::True as i32),
        Variant::from(0i32),
        Variant::from(MsoTriState::False as i32),
    ])?;
    let mut presentation = Dispatch::new(pres_v.as_dispatch()?);
    let inventory = build_inventory(&mut presentation);

    let mut result = CheckResult::default();

    // Check tables
    let sp = if !verbose { Some(make_check_spinner("Tables")) } else { None };
    check_tables(&inventory, &mut session.excel_app, &excel_str, config, &mut result);
    if let Some(sp) = sp { sp.finish_and_clear(); }

    let sp = if !verbose { Some(make_check_spinner("Deltas")) } else { None };
    check_deltas(&inventory, &mut session.excel_app, &excel_str, &mut result);
    if let Some(sp) = sp { sp.finish_and_clear(); }

    let sp = if !verbose { Some(make_check_spinner("Charts")) } else { None };
    check_charts(&inventory, &mut session.excel_app, &excel_str, pptx, &mut result);
    if let Some(sp) = sp { sp.finish_and_clear(); }

    // Per-file summary table
    println!();
    print_check_row("Tables", result.tbl_checked, None, result.tbl_mismatches.len());
    print_check_row("Deltas", result.delt_checked, None, result.delt_mismatches.len());
    print_check_row("Charts", result.chart_count, Some(result.chart_series_checked), result.chart_mismatches.len());

    // Per-file completion line
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

    // Cleanup: drop inventory + close pres + close workbooks (keep apps alive)
    drop(inventory);
    let _ = presentation.call("Close", &[]);
    drop(presentation);
    drop(presentations);
    session.close_all_workbooks();

    Ok(result)
}

/// Run `oa check` against all jobs in a runfile. Returns true if all passed.
pub fn run_check_runfile(
    runfile_path: &str,
    config: &Config,
    verbose: bool,
) -> OaResult<bool> {
    use crate::commands::run::{parse_runfile, fmt_time};

    let runfile_path = std::path::Path::new(runfile_path);
    let (jobs, config_overrides, _steps) = parse_runfile(runfile_path)?;
    if jobs.is_empty() {
        println!("No jobs found in runfile.");
        return Ok(true);
    }

    // Apply runfile config overrides
    let mut config = config.clone();
    config.apply_overrides(&config_overrides)?;

    println!("Check runfile: {} ({} jobs)", runfile_path.display(), jobs.len());

    let mut session = ComSession::new()?;
    let total_start = Instant::now();

    enum CheckOutcome {
        Passed { checked: usize, secs: f64 },
        Failed { checked: usize, mismatches: usize, secs: f64 },
        Error(String),
    }
    struct CheckJobResult {
        name: String,
        outcome: CheckOutcome,
    }

    let mut results: Vec<CheckJobResult> = Vec::with_capacity(jobs.len());

    for (i, job) in jobs.iter().enumerate() {
        println!("\n--- Check {}/{}: {} ---", i + 1, jobs.len(), job.name);

        let pptx_path = job.output.to_string_lossy().to_string();
        let excel_path = job.excel.to_string_lossy().to_string();

        let start = Instant::now();
        match run_check_with_session(&mut session, &pptx_path, &excel_path, &config, verbose) {
            Ok(result) => {
                let secs = start.elapsed().as_secs_f64();
                if result.passed() {
                    results.push(CheckJobResult {
                        name: job.name.clone(),
                        outcome: CheckOutcome::Passed { checked: result.total_checked(), secs },
                    });
                } else {
                    results.push(CheckJobResult {
                        name: job.name.clone(),
                        outcome: CheckOutcome::Failed {
                            checked: result.total_checked(),
                            mismatches: result.total_mismatches(),
                            secs,
                        },
                    });
                }
            }
            Err(e) => {
                eprintln!("  Check '{}' error: {e}", job.name);
                results.push(CheckJobResult {
                    name: job.name.clone(),
                    outcome: CheckOutcome::Error(e.to_string()),
                });
            }
        }
    }

    drop(session);

    // Batch summary
    let total_elapsed = total_start.elapsed().as_secs_f64();
    let s_dim = Style::new().dim();
    let s_ok = Style::new().green();
    let s_err = Style::new().red();
    let s_count = Style::new().white().bold();

    println!();
    println!("  {}", s_dim.apply_to("═".repeat(39)));
    println!("  {}", s_dim.apply_to("Check summary"));
    println!();

    let col: usize = 48;
    for r in &results {
        let prefix_len = 4; // "  ✓ " or "  ✗ "
        let name_len = r.name.chars().count();
        let leader_len = col.saturating_sub(prefix_len + name_len + 1);
        let leaders = "·".repeat(leader_len);

        match &r.outcome {
            CheckOutcome::Passed { checked, secs } => {
                println!("  {} {} {} {} {} {}",
                    s_ok.apply_to("✓"),
                    r.name,
                    s_dim.apply_to(&leaders),
                    s_count.apply_to(format!("{checked:>4}")),
                    s_dim.apply_to("checked"),
                    s_dim.apply_to(format!("{:>6}", fmt_time(*secs))),
                );
            }
            CheckOutcome::Failed { mismatches, secs, .. } => {
                println!("  {} {} {} {} {}",
                    s_err.apply_to("✗"),
                    r.name,
                    s_dim.apply_to(&leaders),
                    s_err.apply_to(format!("{mismatches} mismatches")),
                    s_dim.apply_to(format!("{:>6}", fmt_time(*secs))),
                );
            }
            CheckOutcome::Error(msg) => {
                println!("  {} {} {} {}",
                    s_err.apply_to("✗"),
                    r.name,
                    s_dim.apply_to(&leaders),
                    s_err.apply_to(msg),
                );
            }
        }
    }

    let pass_count = results.iter()
        .filter(|r| matches!(&r.outcome, CheckOutcome::Passed { .. }))
        .count();
    let fail_count = results.len() - pass_count;
    let total_checked: usize = results.iter()
        .map(|r| match &r.outcome {
            CheckOutcome::Passed { checked, .. } | CheckOutcome::Failed { checked, .. } => *checked,
            CheckOutcome::Error(_) => 0,
        })
        .sum();

    let status = if fail_count == 0 {
        format!("{} {}", s_ok.apply_to("✓"), s_ok.apply_to("all checks passed"))
    } else {
        format!("{} {}", s_err.apply_to("✗"),
            s_err.apply_to(format!("{fail_count} check{} failed",
                if fail_count == 1 { "" } else { "s" })))
    };

    println!();
    println!("  {} {} {}{} {} {} {} {}",
        status,
        s_dim.apply_to("·"),
        s_count.apply_to(pass_count),
        s_dim.apply_to(format!("/{}", results.len())),
        s_dim.apply_to("files ·"),
        s_count.apply_to(total_checked),
        s_dim.apply_to("checked ·"),
        s_ok.apply_to(fmt_time(total_elapsed)),
    );

    Ok(fail_count == 0)
}

/// Create a spinner for a check step.
fn make_check_spinner(name: &str) -> indicatif::ProgressBar {
    let leader_len = 20usize.saturating_sub(name.len());
    let leaders = "·".repeat(leader_len);
    let pb = indicatif::ProgressBar::new_spinner();
    pb.set_style(
        indicatif::ProgressStyle::default_spinner()
            .tick_strings(&["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"])
            .template("  {spinner:.cyan} {msg}")
            .unwrap()
    );
    pb.set_message(format!("{name} {leaders}"));
    pb.enable_steady_tick(std::time::Duration::from_millis(80));
    pb
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

    // Detail column: "checked" or "checked (N series)" — use plain string for width,
    // then style it. ANSI escapes break width calculation so we pad manually.
    let (detail_plain, detail_styled) = match series {
        Some(n) => {
            let plain = format!("checked ({n} series)");
            let styled = format!("{} ({} {}",
                s_dim.apply_to("checked"),
                s_count.apply_to(n),
                s_dim.apply_to("series)"));
            (plain, styled)
        }
        None => {
            let plain = "checked".to_string();
            let styled = format!("{}", s_dim.apply_to("checked"));
            (plain, styled)
        }
    };
    // Pad to fixed width (25 chars) for alignment
    let pad = 25usize.saturating_sub(detail_plain.len());
    let detail = format!("{}{}", detail_styled, " ".repeat(pad));

    // Mismatch column: fixed width (16 chars plain) so PASS/FAIL aligns
    let mm_plain = format!("{:>3} mismatches", mismatches);
    let mm_styled = if mismatches == 0 {
        format!("{:>3} {}", s_count.apply_to(0), s_dim.apply_to("mismatches"))
    } else {
        format!("{:>3} {}", s_fail.apply_to(mismatches), s_dim.apply_to("mismatches"))
    };
    let mm_pad = 16usize.saturating_sub(mm_plain.len());
    let mm_display = format!("{}{}", " ".repeat(mm_pad), mm_styled);

    // Badge: PASS or FAIL
    let badge = if mismatches == 0 { s_ok.apply_to("PASS") } else { s_fail.apply_to("FAIL") };

    println!("  {} {} {} {:>3} {}  {} {}  {}",
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

        // Shared DISPID caches — same-class COM objects reuse resolved DISPIDs
        let excel_cell_cache = cells.cache();
        let ppt_cell_cache = tbl.cache();
        // GOTCHA #31: DISPIDs are per-COM-class — each class needs its own cache
        let shape_cache: std::rc::Rc<std::cell::RefCell<std::collections::HashMap<String, i32>>> =
            std::rc::Rc::new(std::cell::RefCell::new(std::collections::HashMap::new()));
        let tf_cache: std::rc::Rc<std::cell::RefCell<std::collections::HashMap<String, i32>>> =
            std::rc::Rc::new(std::cell::RefCell::new(std::collections::HashMap::new()));
        let tr_cache: std::rc::Rc<std::cell::RefCell<std::collections::HashMap<String, i32>>> =
            std::rc::Rc::new(std::cell::RefCell::new(std::collections::HashMap::new()));

        for r in 1..=rows {
            for c in 1..=cols {
                let excel_text = cells.call("Item", &[Variant::from(r), Variant::from(c)])
                    .and_then(|v| v.as_dispatch()).and_then(|d| Dispatch::new_with_cache(d, excel_cell_cache.clone()).get("Text"))
                    .and_then(|v| v.as_string()).unwrap_or_default();
                let (pr, pc) = if do_transpose { (c, r) } else { (r, c) };
                let ppt_text = tbl.call("Cell", &[Variant::from(pr), Variant::from(pc)])
                    .and_then(|v| v.as_dispatch())
                    .and_then(|d| {
                        let mut cell = Dispatch::new_with_cache(d, ppt_cell_cache.clone());
                        let mut shape = Dispatch::new_with_cache(cell.get("Shape")?.as_dispatch()?, shape_cache.clone());
                        let mut tf = Dispatch::new_with_cache(shape.get("TextFrame")?.as_dispatch()?, tf_cache.clone());
                        let mut tr = Dispatch::new_with_cache(tf.get("TextRange")?.as_dispatch()?, tr_cache.clone());
                        tr.get("Text")
                    })
                    .and_then(|v| v.as_string()).unwrap_or_default();
                let expected = if is_ccst { apply_ccst_transform(&excel_text, config) } else { excel_text.trim().to_string() };
                result.tbl_checked += 1;
                if ppt_text.trim() != expected {
                    result.tbl_mismatches.push(Mismatch {
                        slide: ole_ref.slide_index, shape: ole_ref.name.clone(), category: "table".into(),
                        detail: format!("({pr},{pc}): PPT={:?} vs Excel={:?}", ppt_text.trim(), expected),
                    });
                    crate::pipeline::verbose::check_detail(
                        ole_ref.slide_index, "table", &table_info.name, false,
                        &format!("({pr},{pc}) PPT='{}' Excel='{}'", ppt_text.trim(), expected));
                } else {
                    crate::pipeline::verbose::check_detail(
                        ole_ref.slide_index, "table", &table_info.name, true,
                        &format!("({pr},{pc}) '{}'", ppt_text.trim()));
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
            crate::pipeline::verbose::check_detail(
                ole_ref.slide_index, "delta", &delt_ref.name, false,
                &format!("sign={actual_sign} expected={expected_sign} ({excel_text})"));
        } else {
            crate::pipeline::verbose::check_detail(
                ole_ref.slide_index, "delta", &delt_ref.name, true,
                &format!("sign={actual_sign} ({excel_text})"));
        }
    }
}

// ── Chart checking ──────────────────────────────────────────

/// Check charts: verify link targets, series counts, and series values.
///
/// Three checks per chart:
/// 1. Link target points to correct Excel file
/// 2. Series count matches between COM and XML
/// 3. Series values match between PPT (Series.Values) and Excel (Range.Value2)
fn check_charts(
    inventory: &SlideInventory,
    excel_app: &mut Dispatch,
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

    // Pre-read chart cached values from ZIP — avoids COM Series.Values calls
    let chart_cache_map = match build_chart_cache_map(pptx_path) {
        Ok(m) => m,
        Err(_) => HashMap::new(), // Fall back to empty — will use COM
    };

    let excel_filename = std::path::Path::new(excel_path)
        .file_name()
        .map(|f| f.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    // Get workbooks collection for Excel reads
    let mut workbooks = match excel_app.get("Workbooks")
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .map(Dispatch::new)
    {
        Ok(wb) => wb,
        Err(_) => return,
    };

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

            // Check 1: link points to correct Excel
            let source_lower = link_source.to_lowercase();
            let mut chart_ok = true;
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
                chart_ok = false;
                let s_red = console::Style::new().red();
                let s_dim = console::Style::new().dim();
                crate::pipeline::verbose::check_detail(
                    chart_ref.slide_index, "chart", &chart_ref.name, false,
                    &format!("{} {}", s_red.apply_to("wrong link:"), s_dim.apply_to(short)));
            }

            // Check 2: series count matches
            if expected_series > 0 && series_count != expected_series {
                result.chart_mismatches.push(Mismatch {
                    slide: chart_ref.slide_index,
                    shape: chart_ref.name.clone(),
                    category: "chart".into(),
                    detail: format!("series: COM={series_count} XML={expected_series}"),
                });
                chart_ok = false;
                crate::pipeline::verbose::check_detail(
                    chart_ref.slide_index, "chart", &chart_ref.name, false,
                    &format!("series: COM={series_count} XML={expected_series}"));
            }

            // Check 3: series values match (always attempt if we have refs)
            if let Some(series_refs) = refs {
                let cached = chart_cache_map.get(&key).cloned().unwrap_or_default();
                let value_ok = check_chart_series_values(
                    &cached, &mut workbooks, excel_path, series_refs,
                    chart_ref, series_count, result,
                );
                if value_ok && chart_ok {
                    crate::pipeline::verbose::check_detail(
                        chart_ref.slide_index, "chart", &chart_ref.name, true,
                        &format!("verified ({series_count} series)"));
                }
                // Note: when !value_ok, the ✗ parent line + ╰ diffs are printed
                // inside check_chart_series_values before it returns.
            } else {
                // No XML refs — count-only
                result.chart_series_checked += series_count as usize;
                if chart_ok {
                    crate::pipeline::verbose::check_detail(
                        chart_ref.slide_index, "chart", &chart_ref.name, true,
                        &format!("linked ({series_count} series)"));
                }
            }
        }
    }
}

/// Compare series values between PPT and Excel for one chart.
///
/// Returns true if all series match, false if any mismatch found.
fn check_chart_series_values(
    ppt_cached: &[(String, Vec<f64>)],  // ZIP-cached values (ref, values) per series
    workbooks: &mut Dispatch,
    excel_path: &str,
    series_refs: &[String],
    chart_ref: &crate::shapes::inventory::ShapeRef,
    series_count: i32,
    result: &mut CheckResult,
) -> bool {
    let mut wb = match open_or_get_workbook(workbooks, excel_path) {
        Ok(wb) => wb,
        Err(_) => {
            result.chart_series_checked += series_count as usize;
            return true;
        }
    };

    struct SeriesMismatch {
        name: String,
        diff_count: usize,
        total: usize,
        pairs: Vec<(f64, f64)>,
        has_more: bool,
    }
    let mut mismatches: Vec<SeriesMismatch> = Vec::new();

    for (i, range_ref) in series_refs.iter().enumerate() {
        let series_idx = (i + 1) as i32;
        result.chart_series_checked += 1;

        // PPT side: read from ZIP cache instead of COM Series.Values
        let ppt_values = if i < ppt_cached.len() {
            let vals = &ppt_cached[i].1;
            if vals.is_empty() {
                continue; // Cache exists but no data points (stale/empty cache) — skip
            }
            vals.clone()
        } else {
            continue; // No cached data for this series
        };

        let excel_values = match read_chart_range(&mut wb, range_ref) {
            Ok(v) => v,
            Err(_) => continue,
        };

        if !values_match(&ppt_values, &excel_values) {
            // Use series name from the ZIP cache ref, or fallback
            let series_name_raw = format!("Series {series_idx}");
            let series_name = crate::pipeline::verbose::truncate_middle(&series_name_raw);
            let (diff_count, pairs, has_more) = collect_mismatch_pairs(&ppt_values, &excel_values, 6);
            let total = ppt_values.len().max(excel_values.len());

            result.chart_mismatches.push(Mismatch {
                slide: chart_ref.slide_index,
                shape: chart_ref.name.clone(),
                category: "chart".into(),
                detail: format!("'{series_name}' {diff_count}/{total} differ"),
            });

            mismatches.push(SeriesMismatch { name: series_name, diff_count, total, pairs, has_more });
        }
    }

    if !mismatches.is_empty() {
        // Print ✗ parent line before ╰ continuation lines
        let total_diffs: usize = mismatches.iter().map(|m| m.diff_count).sum();
        crate::pipeline::verbose::check_detail(
            chart_ref.slide_index, "chart", &chart_ref.name, false,
            &format!("{total_diffs} values differ ({series_count} series)"));

        let max_name_len = mismatches.iter().map(|m| m.name.len()).max().unwrap_or(0);
        for m in &mismatches {
            crate::pipeline::verbose::check_chart_series_diff(
                &m.name, max_name_len, m.diff_count, m.total, &m.pairs, m.has_more);
        }
    }

    mismatches.is_empty()
}

/// Collect first N mismatched (ppt, excel) value pairs for compact display.
///
/// Returns: (total_diff_count, pairs_vec, has_more_beyond_max)
fn collect_mismatch_pairs(ppt: &[f64], excel: &[f64], max: usize) -> (usize, Vec<(f64, f64)>, bool) {
    let mut pairs = Vec::new();
    let mut diff_count = 0;

    for (p, e) in ppt.iter().zip(excel.iter()) {
        if !float_eq(*p, *e) {
            diff_count += 1;
            if pairs.len() < max {
                pairs.push((*p, *e));
            }
        }
    }

    // Length mismatch: count extra elements as diffs too
    if ppt.len() != excel.len() {
        diff_count += ppt.len().abs_diff(excel.len());
    }

    let has_more = diff_count > pairs.len();
    (diff_count, pairs, has_more)
}

/// Read PPT chart series values via Series.Values (SAFEARRAY of doubles).
fn read_ppt_series_values(series: &mut Dispatch) -> OaResult<Vec<f64>> {
    let values_variant = series.get("Values")?;
    if values_variant.is_empty() {
        return Ok(vec![]);
    }
    values_variant.as_flat_f64_vec()
}

/// Read Excel values for a chart range reference (supports multi-cell SAFEARRAY).
///
/// Handles non-contiguous ranges like "(Tables!$C$10,Tables!$F$10)" — GOTCHA #20.
fn read_chart_range(wb: &mut Dispatch, range_ref: &str) -> OaResult<Vec<f64>> {
    let mut values = Vec::new();

    // Strip outer parentheses for multi-area ranges
    let ref_str = range_ref.trim_start_matches('(').trim_end_matches(')');

    // Split on comma for non-contiguous ranges (GOTCHA #20)
    for sub_range in ref_str.split(',') {
        let sub = sub_range.trim().replace('$', "");
        let (sheet_name, range_addr) = if let Some(pos) = sub.find('!') {
            (sub[..pos].to_string(), sub[pos + 1..].to_string())
        } else {
            ("Tables".to_string(), sub)
        };

        let val = wb.get("Worksheets")
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(sheet_name.as_str())]))
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(range_addr.as_str())]))
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).get("Value2"))?;

        if val.is_array() {
            values.extend(val.as_flat_f64_vec()?);
        } else if val.is_empty() {
            values.push(0.0);
        } else {
            values.push(val.as_numeric()?);
        }
    }

    Ok(values)
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

/// Build chart CACHED VALUES map from PPTX ZIP.
/// Returns: {(slide_num, chart_position) → Vec<(range_ref, cached_values)>}
/// This reads <c:numCache> values directly from the ZIP — no COM needed.
fn build_chart_cache_map(pptx_path: &std::path::Path) -> Result<HashMap<(i32, usize), Vec<(String, Vec<f64>)>>, String> {
    let file = std::fs::File::open(pptx_path).map_err(|e| e.to_string())?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| e.to_string())?;
    let mut result: HashMap<(i32, usize), Vec<(String, Vec<f64>)>> = HashMap::new();

    let slide_order = get_slide_order(&mut archive)?;

    for (com_index, slide_file) in slide_order.iter().enumerate() {
        let slide_num = (com_index + 1) as i32;
        let slide_filename = slide_file.rsplit('/').next().unwrap_or(slide_file);
        let rels_path = format!("ppt/slides/_rels/{slide_filename}.rels");
        let rels_map = read_rels_map(&mut archive, &rels_path);
        let slide_xml = match read_zip_entry(&mut archive, slide_file) {
            Some(data) => data,
            None => continue,
        };
        let chart_positions = find_charts_in_slide(&slide_xml, &rels_map);

        for (pos, chart_path) in chart_positions {
            let chart_filename = chart_path.rsplit('/').next().unwrap_or(&chart_path);
            let chart_rels_path = format!("ppt/charts/_rels/{chart_filename}.rels");
            if !has_external_link(&mut archive, &chart_rels_path) {
                continue;
            }
            let full_chart_path = if chart_path.starts_with("ppt/") {
                chart_path.clone()
            } else {
                format!("ppt/charts/{}", chart_path.trim_start_matches("../charts/"))
            };
            if let Some(chart_xml) = read_zip_entry(&mut archive, &full_chart_path) {
                let cached = crate::zip_ops::chart_data::extract_cached_values(&chart_xml);
                if !cached.is_empty() {
                    result.insert((slide_num, pos), cached);
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

    // --- collect_mismatch_pairs tests ---

    #[test]
    fn test_collect_pairs_basic() {
        let (count, pairs, more) = collect_mismatch_pairs(&[1.0, 2.0, 3.0], &[1.0, 2.5, 3.0], 5);
        assert_eq!(count, 1);
        assert_eq!(pairs.len(), 1);
        assert!((pairs[0].0 - 2.0).abs() < f64::EPSILON);
        assert!((pairs[0].1 - 2.5).abs() < f64::EPSILON);
        assert!(!more);
    }

    #[test]
    fn test_collect_pairs_truncated() {
        let (count, pairs, more) = collect_mismatch_pairs(
            &[1.0, 2.0, 3.0, 4.0, 5.0],
            &[0.0, 0.0, 0.0, 0.0, 0.0], 3);
        assert_eq!(count, 5);
        assert_eq!(pairs.len(), 3);
        assert!(more);
    }

    #[test]
    fn test_collect_pairs_length_mismatch() {
        let (count, _pairs, _more) = collect_mismatch_pairs(&[1.0], &[1.0, 2.0], 5);
        assert_eq!(count, 1); // 1 extra element
    }

    #[test]
    fn test_collect_pairs_all_match() {
        let (count, pairs, more) = collect_mismatch_pairs(&[1.0, 2.0], &[1.0, 2.0], 5);
        assert_eq!(count, 0);
        assert!(pairs.is_empty());
        assert!(!more);
    }

    #[test]
    fn test_values_match_zero_vs_zero() {
        assert!(values_match(&[0.0, 0.0], &[0.0, 0.0]));
    }
}
