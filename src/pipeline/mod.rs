//! Pipeline orchestration — step sequencing, progress reporting.
//!
//! Pipeline order (fixed, cannot be reordered):
//! 1. links    — Re-point OLE links to new Excel file
//! 2. tables   — Populate PPT tables from Excel ranges
//! 3. deltas   — Swap delta indicator arrows based on sign
//! 4. coloring — Apply sign-based color coding (_ccst shapes)
//! 5. charts   — Update chart data links

pub mod chart_updater;
pub mod color_coder;
pub mod delta_updater;
pub mod linker;
pub mod table_updater;

use std::time::Instant;

use indicatif::{ProgressBar, ProgressStyle};

use crate::cli::resolve_steps;
use crate::com::dispatch::Dispatch;
use crate::config::Config;
use crate::error::OaResult;
use crate::shapes::inventory::SlideInventory;

/// Per-step result with timing.
#[derive(Debug, Clone)]
pub struct StepResult {
    pub name: &'static str,
    pub count: usize,
    pub elapsed_secs: f64,
    pub ok: bool,
}

/// Results from running the pipeline on a single presentation.
#[derive(Debug, Default)]
pub struct PipelineResults {
    pub steps: Vec<StepResult>,
    pub links_updated: usize,
    pub tables_updated: usize,
    pub deltas_updated: usize,
    pub tables_colored: usize,
    pub charts_updated: usize,
}

/// Create a spinner for a pipeline step.
fn make_spinner(step_name: &str) -> ProgressBar {
    let pb = ProgressBar::new_spinner();
    pb.set_style(
        ProgressStyle::default_spinner()
            .template("  {spinner:.cyan} {msg}")
            .unwrap()
    );
    pb.set_message(format!("{step_name}..."));
    pb.enable_steady_tick(std::time::Duration::from_millis(80));
    pb
}

/// Run the update pipeline on a presentation with the given steps.
pub fn run_pipeline(
    inventory: &SlideInventory,
    config: &Config,
    presentation: &mut Dispatch,
    excel_app: &mut Dispatch,
    excel_path: &str,
    steps_include: &[String],
    steps_skip: &[String],
    quiet: bool,
) -> OaResult<PipelineResults> {
    let active_steps = resolve_steps(steps_include, steps_skip)
        .map_err(|e| crate::error::OaError::Config(e))?;

    let mut results = PipelineResults::default();

    // Step 1: Links
    if active_steps.iter().any(|s| s == "links") {
        let spinner = if !quiet { Some(make_spinner("Links")) } else { None };
        let t = Instant::now();

        let count = linker::update_links(inventory, excel_path, config)?;
        results.links_updated = count;

        let elapsed = t.elapsed().as_secs_f64();
        if let Some(pb) = spinner { pb.finish_and_clear(); }
        results.steps.push(StepResult { name: "Links", count, elapsed_secs: elapsed, ok: true });
    }

    // Step 2: Tables
    if active_steps.iter().any(|s| s == "tables") {
        let spinner = if !quiet { Some(make_spinner("Tables")) } else { None };
        let t = Instant::now();

        let count = table_updater::update_tables(inventory, config, excel_app, excel_path)?;
        results.tables_updated = count;

        let elapsed = t.elapsed().as_secs_f64();
        if let Some(pb) = spinner { pb.finish_and_clear(); }
        results.steps.push(StepResult { name: "Tables", count, elapsed_secs: elapsed, ok: true });
    }

    // Step 3: Deltas
    if active_steps.iter().any(|s| s == "deltas") {
        let spinner = if !quiet { Some(make_spinner("Deltas")) } else { None };
        let t = Instant::now();

        let count = delta_updater::update_deltas(
            inventory, config, presentation, excel_path, excel_app,
        )?;
        results.deltas_updated = count;

        let elapsed = t.elapsed().as_secs_f64();
        if let Some(pb) = spinner { pb.finish_and_clear(); }
        results.steps.push(StepResult { name: "Deltas", count, elapsed_secs: elapsed, ok: true });
    }

    // Step 4: Coloring
    if active_steps.iter().any(|s| s == "coloring") {
        let spinner = if !quiet { Some(make_spinner("Coloring")) } else { None };
        let t = Instant::now();

        let count = color_coder::apply_color_coding(inventory, config)?;
        results.tables_colored = count;

        let elapsed = t.elapsed().as_secs_f64();
        if let Some(pb) = spinner { pb.finish_and_clear(); }
        results.steps.push(StepResult { name: "Coloring", count, elapsed_secs: elapsed, ok: true });
    }

    // Step 5: Charts
    if active_steps.iter().any(|s| s == "charts") {
        let spinner = if !quiet { Some(make_spinner("Charts")) } else { None };
        let t = Instant::now();

        let count = chart_updater::update_charts(inventory, excel_path)?;
        results.charts_updated = count;

        let elapsed = t.elapsed().as_secs_f64();
        if let Some(pb) = spinner { pb.finish_and_clear(); }
        results.steps.push(StepResult { name: "Charts", count, elapsed_secs: elapsed, ok: true });
    }

    Ok(results)
}
