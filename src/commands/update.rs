//! `oa update` — the main pipeline command.
//!
//! Processes one or more PPTX files: ZIP pre-relink → COM session → pipeline → save.
//! Supports single file, batch via --pair, and glob patterns.

use std::path::{Path, PathBuf};
use std::time::Instant;

use console::Style;

use crate::cli::{parse_pair, UpdateArgs};
use crate::com::dispatch::Dispatch;
use crate::com::session::ComSession;
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::{OaError, OaResult};
use crate::pipeline::{self, PipelineResults};
use crate::shapes::inventory::build_inventory;
use crate::zip_ops::relinker::relink_pptx_zip;

/// A resolved (pptx, excel) pair ready for processing.
struct FilePair {
    pptx: PathBuf,
    excel: PathBuf,
    output: Option<PathBuf>,
}

/// Strip the \\?\ UNC prefix from a canonicalized Windows path.
fn strip_unc(path: &Path) -> String {
    let s = path.to_string_lossy().to_string();
    s.strip_prefix(r"\\?\").unwrap_or(&s).to_string()
}

/// Run the `oa update` command (creates its own COM session).
pub fn run_update(args: &UpdateArgs) -> OaResult<()> {
    let mut session = ComSession::new()?;
    run_update_with_session(args, &mut session)?;
    Ok(())
}

/// Run the `oa update` command using an existing COM session.
///
/// Returns `(total_objects, total_elapsed_secs)` for batch summary use.
/// Called directly by `oa run` to reuse one session across all jobs,
/// avoiding the 0x80010001 error from rapid COM churn (GOTCHA #39).
pub fn run_update_with_session(args: &UpdateArgs, session: &mut ComSession) -> OaResult<(usize, f64)> {
    // Build config with overrides
    let mut config = Config::default();
    config.apply_overrides(&args.set)?;

    // Resolve file pairs
    let pairs = resolve_file_pairs(args)?;
    if pairs.is_empty() {
        return Err(OaError::Other("No files to process. Provide PPTX files, --pair, or --pick.".into()));
    }

    let total_start = Instant::now();
    let mut all_results: Vec<(String, PipelineResults, f64)> = Vec::new();

    for pair in &pairs {
        let file_name = pair.pptx.file_name()
            .map(|f| f.to_string_lossy().to_string())
            .unwrap_or_default();
        let excel_name = pair.excel.file_name()
            .map(|f| f.to_string_lossy().to_string())
            .unwrap_or_default();

        let s_target = Style::new().cyan();
        let s_source = Style::new().yellow();
        let s_dim = Style::new().dim();

        if !args.quiet {
            println!();
            println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&file_name));
            println!("    {} {}", s_source.apply_to("←"), s_source.apply_to(&excel_name));
            println!("  {}", s_dim.apply_to("╌".repeat(39)));
            println!();
        }

        // Determine output path
        let work_path = if let Some(output) = &pair.output {
            if output.is_dir() {
                let out = output.join(pair.pptx.file_name().unwrap_or_default());
                std::fs::copy(&pair.pptx, &out)?;
                out
            } else {
                std::fs::copy(&pair.pptx, output)?;
                output.clone()
            }
        } else {
            pair.pptx.clone() // In-place
        };

        let start = Instant::now();

        let result = process_single_file(
            session,
            &work_path,
            &pair.excel,
            &config,
            args,
        );

        match result {
            Ok(results) => {
                let elapsed = start.elapsed().as_secs_f64();
                if !args.quiet {
                    print_completion(&results, elapsed, args.dry_run);
                }
                all_results.push((file_name, results, elapsed));
            }
            Err(e) => {
                let s_err = Style::new().red().bold();
                eprintln!("  {} {e}", s_err.apply_to("Error:"));
            }
        }
    }

    let total_elapsed = total_start.elapsed().as_secs_f64();

    // Batch summary (for standalone oa update with multiple files)
    if pairs.len() > 1 && !args.quiet {
        print_batch_completion(&all_results, total_elapsed);
    }

    let total_objects: usize = all_results.iter().map(|(_, r, _)| r.total_objects()).sum();
    Ok((total_objects, total_elapsed))
}

/// Process a single PPTX file using an existing COM session.
///
/// Performs: ZIP pre-relink -> chart pre-update -> open presentation ->
/// build inventory -> run pipeline -> save -> cleanup (close pres + workbooks).
/// The COM session (Excel/PPT apps) stays alive for reuse by the next file.
fn process_single_file(
    session: &mut ComSession,
    pptx_path: &Path,
    excel_path: &Path,
    config: &Config,
    args: &UpdateArgs,
) -> OaResult<PipelineResults> {
    let pptx_str = strip_unc(&pptx_path.canonicalize()?);
    let excel_str = strip_unc(&excel_path.canonicalize()?);
    let dry_run = args.dry_run;
    let quiet = args.quiet;
    let verbose = args.verbose;

    use crate::pipeline::verbose as pverbose;
    pverbose::set_verbose(verbose);

    // ZIP pre-relink (before Open, so PowerPoint reads corrected paths)
    if !dry_run {
        let use_spinner = !quiet && !verbose;
        let relink_spinner = if use_spinner { Some(pipeline::make_spinner_pub("Relink")) } else { None };
        let relink_t = std::time::Instant::now();

        let relink_result = relink_pptx_zip(pptx_path, excel_path)
            .unwrap_or_else(|e| {
                eprintln!("  ZIP pre-relink warning: {e}");
                crate::zip_ops::relinker::RelinkResult { total: 0, ole: 0, charts: 0 }
            });

        let relink_elapsed = relink_t.elapsed().as_secs_f64();
        if let Some(pb) = relink_spinner { pb.finish_and_clear(); }
        if !quiet {
            if verbose {
                pverbose::note(&format!(
                    "Relink ····················· {} links ({} OLE + {} charts) · {:.1}s",
                    relink_result.total, relink_result.ole, relink_result.charts, relink_elapsed
                ));
            } else {
                println!("{}", pipeline::format_step_line_pub("Relink", relink_result.total, relink_elapsed));
            }
        }
    }

    // ZIP chart data pre-update: rewrite numCache values directly in chart XML.
    // This bypasses the slow LinkFormat.Update() COM call.
    let mut chart_data_ok = false;
    if !dry_run {
        let chart_t = std::time::Instant::now();
        match zip_chart_preupdate(pptx_path, excel_path, &mut session.excel_app, quiet, verbose) {
            Ok(result) => {
                chart_data_ok = result.charts_updated > 0 || result.series_updated == 0;
                let chart_elapsed = chart_t.elapsed().as_secs_f64();
                if verbose {
                    pverbose::note(&format!(
                        "Chart pre-update ··········· {} charts ({} series) · {:.1}s",
                        result.charts_updated, result.series_updated, chart_elapsed
                    ));
                }
            }
            Err(e) => {
                if verbose {
                    pverbose::note(&format!("Chart pre-update skipped: {e}"));
                }
            }
        }
    }

    // Open presentation
    let t_open = std::time::Instant::now();
    let mut presentations = Dispatch::new(session.ppt_app.get("Presentations")?.as_dispatch()?);
    let pres_variant = presentations.call("Open", &[
        Variant::from(pptx_str.as_str()),
        Variant::from(0i32),  // ReadOnly = False
        Variant::from(0i32),  // Untitled = False (MUST be False so Save() works)
        Variant::from(0i32),  // WithWindow = False
    ])?;
    let mut presentation = Dispatch::new(pres_variant.as_dispatch()?);
    pverbose::note(&format!("Open PPTX ·················· {:.1}s", t_open.elapsed().as_secs_f64()));

    // Build inventory
    let t_inv = std::time::Instant::now();
    let inventory = build_inventory(&mut presentation);
    pverbose::note(&format!("Build inventory ············ {:.1}s", t_inv.elapsed().as_secs_f64()));

    // Run pipeline
    let pipeline_result = pipeline::run_pipeline(
        &inventory,
        config,
        &mut presentation,
        &mut session.excel_app,
        &excel_str,
        &args.steps,
        &args.skip,
        quiet,
        verbose,
        chart_data_ok,
    );

    // Save (unless dry-run or pipeline failed)
    if let Ok(ref _results) = pipeline_result {
        let t_save = std::time::Instant::now();
        if !dry_run {
            presentation.call0("Save")?;
        }
        pverbose::note(&format!("Save ······················· {:.1}s", t_save.elapsed().as_secs_f64()));
    }

    // GOTCHA #21: Explicit drop ordering to prevent 60s hang.
    // Drop all shape/inventory refs before closing the presentation.
    // Apps stay alive for the next file.
    let t_cleanup = std::time::Instant::now();
    drop(inventory);
    let _ = presentation.call("Close", &[]);
    drop(presentation);
    drop(presentations);

    // Close workbooks between files to free memory
    session.close_all_workbooks();
    pverbose::note(&format!("Cleanup ···················· {:.1}s", t_cleanup.elapsed().as_secs_f64()));

    pipeline_result
}

/// Resolve CLI arguments into (pptx, excel) file pairs.
fn resolve_file_pairs(args: &UpdateArgs) -> OaResult<Vec<FilePair>> {
    let mut pairs = Vec::new();

    // Mode 1: --pair PPT=XLSX (explicit pairs)
    if !args.pair.is_empty() {
        for pair_str in &args.pair {
            let (pptx, xlsx) = parse_pair(pair_str)
                .map_err(OaError::Config)?;
            pairs.push(FilePair {
                pptx: PathBuf::from(&pptx),
                excel: PathBuf::from(&xlsx),
                output: args.output.as_ref().map(PathBuf::from),
            });
        }
        return Ok(pairs);
    }

    // Mode 2: FILE(s) + --excel
    if !args.files.is_empty() {
        // Resolve globs
        let mut pptx_files: Vec<PathBuf> = Vec::new();
        for pattern in &args.files {
            let matches: Vec<_> = glob::glob(pattern)
                .map_err(|e| OaError::Config(format!("Invalid glob pattern: {e}")))?
                .filter_map(|r| r.ok())
                .collect();
            if matches.is_empty() {
                return Err(OaError::Config(format!("No files match: {pattern}")));
            }
            pptx_files.extend(matches);
        }

        // Determine Excel file
        let excel_path = if let Some(ref excel) = args.excel {
            PathBuf::from(excel)
        } else if args.pick {
            // File picker
            pick_excel_file()?
        } else {
            // Auto-detect from first PPTX — check for multiple links
            let all_excels = crate::zip_ops::detector::detect_all_linked_excels(&pptx_files[0]);
            if all_excels.is_empty() {
                return Err(OaError::Config("Cannot auto-detect Excel file. Use -e to specify.".into()));
            }
            if all_excels.len() > 1 {
                let names: Vec<String> = all_excels.iter()
                    .filter_map(|p| p.file_name().map(|f| f.to_string_lossy().to_string()))
                    .collect();
                return Err(OaError::Config(format!(
                    "Multiple Excel links found ({}). Use -e to specify which file to update with.",
                    names.join(", ")
                )));
            }
            let detected = &all_excels[0];
            if !detected.exists() {
                return Err(OaError::Config(format!(
                    "Auto-detected Excel file not found: {}. Use -e to specify.",
                    detected.display()
                )));
            }
            detected.clone()
        };

        for pptx in pptx_files {
            pairs.push(FilePair {
                pptx,
                excel: excel_path.clone(),
                output: args.output.as_ref().map(PathBuf::from),
            });
        }

        return Ok(pairs);
    }

    Ok(pairs)
}

fn pick_excel_file() -> OaResult<PathBuf> {
    let result = rfd::FileDialog::new()
        .set_title("Select Excel Data File")
        .add_filter("Excel Files", &["xlsx", "xls", "xlsm"])
        .pick_file();

    match result {
        Some(path) => Ok(path),
        None => Err(OaError::Other("File selection cancelled".into())),
    }
}

/// ZIP chart data pre-update: scan chart ranges, read values from Excel, rewrite ZIP.
///
/// Opens the Excel workbook (via already-open Excel app), reads all chart range values
/// in batch using Range.Value2 (SAFEARRAY), then rewrites the PPTX chart XML cache.
fn zip_chart_preupdate(
    pptx_path: &Path,
    excel_path: &Path,
    excel_app: &mut Dispatch,
    _quiet: bool,
    _verbose: bool,
) -> Result<crate::zip_ops::chart_data::ChartDataResult, String> {
    use crate::zip_ops::chart_data;
    use crate::pipeline::table_updater::open_or_get_workbook;

    // Step 1: Scan chart XML for range references
    let chart_ranges = chart_data::scan_chart_ranges(pptx_path)?;
    if chart_ranges.is_empty() {
        return Ok(chart_data::ChartDataResult { charts_updated: 0, series_updated: 0 });
    }

    // Step 2: Collect unique range refs and read values from Excel via COM
    let unique_ranges = chart_data::collect_unique_ranges(&chart_ranges);
    if unique_ranges.is_empty() {
        return Ok(chart_data::ChartDataResult { charts_updated: 0, series_updated: 0 });
    }

    let excel_str = strip_unc(&excel_path.canonicalize().map_err(|e| e.to_string())?);
    let mut workbooks = Dispatch::new(
        excel_app.get("Workbooks").map_err(|e| e.to_string())?
            .as_dispatch().map_err(|e| e.to_string())?
    );
    let mut wb = open_or_get_workbook(&mut workbooks, &excel_str).map_err(|e| e.to_string())?;

    let mut range_values: std::collections::HashMap<String, Vec<f64>> = std::collections::HashMap::new();

    for range_ref in &unique_ranges {
        // Parse "Sheet!Range" format
        let (sheet_name, range_addr) = if let Some(pos) = range_ref.find('!') {
            (range_ref[..pos].to_string(), range_ref[pos + 1..].to_string())
        } else {
            ("Tables".to_string(), range_ref.clone())
        };

        // Read via Range.Value2 (SAFEARRAY batch — one COM call per range)
        let val = wb.get("Worksheets")
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(sheet_name.as_str())]))
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(range_addr.as_str())]))
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).get("Value2"))
            .map_err(|e| format!("Failed to read range {range_ref}: {e}"))?;

        let values = val.as_flat_f64_vec().map_err(|e| format!("Failed to unpack {range_ref}: {e}"))?;
        range_values.insert(range_ref.clone(), values);
    }

    // Step 3: Rewrite chart XML cache in the PPTX ZIP
    chart_data::update_chart_data(pptx_path, &range_values)
}

/// Build a comfy-table with our standard style.
/// Print the completion line after all steps finish.
fn print_completion(results: &PipelineResults, total_secs: f64, dry_run: bool) {
    let s_ok = Style::new().green();
    let s_count = Style::new().white().bold();
    let s_dim = Style::new().dim();

    // In verbose mode, add a divider + per-step summary since detail lines are long
    if crate::pipeline::verbose::is_verbose() {
        println!();
        println!("  {}", s_dim.apply_to("╌".repeat(39)));
        for step in &results.steps {
            if step.count > 0 {
                println!("  {:<14} {:>4}   {}",
                    step.name,
                    s_count.apply_to(step.count),
                    s_dim.apply_to(format!("{:.1}s", step.elapsed_secs)));
            }
        }
    }

    println!();
    if dry_run {
        let s_warn = Style::new().yellow();
        println!("  {} {} {} {} {} {}",
            s_warn.apply_to("⚠ dry run"),
            s_dim.apply_to("·"),
            s_count.apply_to(results.total_objects()),
            s_dim.apply_to("objects"),
            s_dim.apply_to("·"),
            s_warn.apply_to("not saved"));
    } else {
        println!("  {} {} {} {} {}",
            s_ok.apply_to("✓ completed"),
            s_dim.apply_to("·"),
            s_count.apply_to(results.total_objects()),
            s_dim.apply_to("objects ·"),
            s_ok.apply_to(format!("{total_secs:.1}s")));
    }
}

/// Print batch summary across multiple files.
fn print_batch_completion(all_results: &[(String, PipelineResults, f64)], total_secs: f64) {
    let s_ok = Style::new().green();
    let s_count = Style::new().white().bold();
    let s_dim = Style::new().dim();

    let total_objects: usize = all_results.iter().map(|(_, r, _)| r.total_objects()).sum();

    println!();
    println!("  {} {} {} {} {} {} {}",
        s_ok.apply_to("✓ batch complete"),
        s_dim.apply_to("·"),
        s_count.apply_to(all_results.len()),
        s_dim.apply_to("files ·"),
        s_count.apply_to(total_objects),
        s_dim.apply_to("objects ·"),
        s_ok.apply_to(format!("{total_secs:.1}s")));
}
