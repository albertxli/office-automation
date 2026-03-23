//! `oa update` — the main pipeline command.
//!
//! Processes one or more PPTX files: ZIP pre-relink → COM session → pipeline → save.
//! Supports single file, batch via --pair, and glob patterns.

use std::path::{Path, PathBuf};
use std::time::Instant;

use console::Style;

use crate::cli::{parse_pair, UpdateArgs};
use crate::com::dispatch::Dispatch;
use crate::com::session::{create_instance, init_com_sta, spawn_dialog_dismisser, stop_dialog_dismisser};
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::{OaError, OaResult};
use crate::office::constants::{MsoTriState, XlCalculation};
use crate::pipeline::{self, PipelineResults};
use crate::shapes::inventory::build_inventory;
use crate::zip_ops::detector::detect_linked_excel;
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

/// Run the `oa update` command.
pub fn run_update(args: &UpdateArgs) -> OaResult<()> {
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

        // Step 0: ZIP pre-relink (before COM, much faster)
        // Skip for dry-run since it modifies the file
        if !args.dry_run {
            let relink_spinner = if !args.quiet {
                Some(pipeline::make_spinner_pub("Relink"))
            } else {
                None
            };
            let relink_t = Instant::now();

            let relinked = relink_pptx_zip(&work_path, &pair.excel)
                .unwrap_or_else(|e| {
                    eprintln!("  ZIP pre-relink warning: {e}");
                    0
                });

            let relink_elapsed = relink_t.elapsed().as_secs_f64();
            if let Some(pb) = relink_spinner {
                pb.finish_and_clear();
                println!("{}", pipeline::format_step_line_pub("Relink", relinked, relink_elapsed));
            }
        }

        // COM session
        let result = run_com_pipeline(
            &work_path,
            &pair.excel,
            &config,
            &args.steps,
            &args.skip,
            args.dry_run,
            args.quiet,
            args.verbose,
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

    // Batch summary
    if pairs.len() > 1 && !args.quiet {
        let total_elapsed = total_start.elapsed().as_secs_f64();
        print_batch_completion(&all_results, total_elapsed);
    }

    Ok(())
}

/// Run the COM-based pipeline on a single PPTX file.
fn run_com_pipeline(
    pptx_path: &Path,
    excel_path: &Path,
    config: &Config,
    steps_include: &[String],
    steps_skip: &[String],
    dry_run: bool,
    quiet: bool,
    verbose: bool,
) -> OaResult<PipelineResults> {
    let pptx_str = strip_unc(&pptx_path.canonicalize()?);
    let excel_str = strip_unc(&excel_path.canonicalize()?);

    // Initialize COM
    let _com = init_com_sta()?;
    let (stop, handle) = spawn_dialog_dismisser();

    // Create Excel (GOTCHA #28: don't set Calculation mode until workbook is open)
    let mut excel_app = create_instance("Excel.Application")?;
    excel_app.put("Visible", Variant::from(false))?;
    excel_app.put("DisplayAlerts", Variant::from(false))?;
    excel_app.put("ScreenUpdating", Variant::from(false))?;
    excel_app.put("EnableEvents", Variant::from(false))?;

    // Create PowerPoint
    let mut ppt_app = create_instance("PowerPoint.Application")?;
    ppt_app.put("DisplayAlerts", Variant::from(0i32))?;

    // Open presentation
    let mut presentations = Dispatch::new(ppt_app.get("Presentations")?.as_dispatch()?);
    let pres_variant = presentations.call("Open", &[
        Variant::from(pptx_str.as_str()),
        Variant::from(0i32),  // ReadOnly = False
        Variant::from(-1i32), // Untitled = True
        Variant::from(0i32),  // WithWindow = False
    ])?;
    let mut presentation = Dispatch::new(pres_variant.as_dispatch()?);

    // Build inventory
    let inventory = build_inventory(&mut presentation);

    // Run pipeline
    let results = pipeline::run_pipeline(
        &inventory,
        config,
        &mut presentation,
        &mut excel_app,
        &excel_str,
        steps_include,
        steps_skip,
        quiet,
        verbose,
    )?;

    // Save (unless dry-run)
    if !dry_run {
        presentation.call0("Save")?;
    }

    // GOTCHA #21: Explicit drop ordering to prevent 60s hang
    // Drop inventory refs (they hold IDispatch pointers into the presentation)
    drop(inventory);

    // Close presentation
    presentation.call("Close", &[])?;
    drop(presentation);
    drop(presentations);

    // Restore Excel calculation mode and quit
    let _ = excel_app.put("Calculation", Variant::from(XlCalculation::Automatic as i32));
    excel_app.call0("Quit")?;
    drop(excel_app);

    // Quit PowerPoint
    ppt_app.call0("Quit")?;
    drop(ppt_app);

    // Stop dialog dismisser
    stop_dialog_dismisser(stop, handle);

    Ok(results)
}

/// Resolve CLI arguments into (pptx, excel) file pairs.
fn resolve_file_pairs(args: &UpdateArgs) -> OaResult<Vec<FilePair>> {
    let mut pairs = Vec::new();

    // Mode 1: --pair PPT=XLSX (explicit pairs)
    if !args.pair.is_empty() {
        for pair_str in &args.pair {
            let (pptx, xlsx) = parse_pair(pair_str)
                .map_err(|e| OaError::Config(e))?;
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
            // Auto-detect from first PPTX
            detect_linked_excel(&pptx_files[0])
                .ok_or_else(|| OaError::Config(
                    "Cannot auto-detect Excel file. Use -e to specify.".into()
                ))?
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

/// Build a comfy-table with our standard style.
/// Print the completion line after all steps finish.
fn print_completion(results: &PipelineResults, total_secs: f64, dry_run: bool) {
    let s_ok = Style::new().green();
    let s_count = Style::new().white().bold();
    let s_dim = Style::new().dim();

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
