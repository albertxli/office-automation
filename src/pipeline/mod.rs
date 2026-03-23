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

use crate::cli::resolve_steps;
use crate::com::dispatch::Dispatch;
use crate::config::Config;
use crate::error::OaResult;
use crate::shapes::inventory::SlideInventory;

/// Results from running the pipeline on a single presentation.
#[derive(Debug, Default)]
pub struct PipelineResults {
    pub links_updated: usize,
    pub tables_updated: usize,
    pub deltas_updated: usize,
    pub tables_colored: usize,
    pub charts_updated: usize,
}

/// Run the update pipeline on a presentation with the given steps.
///
/// `steps_include` and `steps_skip` control which steps run (from --steps/--skip flags).
/// Returns per-step counts.
pub fn run_pipeline(
    inventory: &SlideInventory,
    config: &Config,
    presentation: &mut Dispatch,
    excel_app: &mut Dispatch,
    excel_path: &str,
    steps_include: &[String],
    steps_skip: &[String],
) -> OaResult<PipelineResults> {
    let active_steps = resolve_steps(steps_include, steps_skip)
        .map_err(|e| crate::error::OaError::Config(e))?;

    let mut results = PipelineResults::default();

    // Step 1: Links (AutoUpdate=Manual only — paths handled by ZIP pre-relink)
    if active_steps.iter().any(|s| s == "links") {
        results.links_updated = linker::update_links(inventory, excel_path, config)?;
    }

    // Step 2: Tables
    if active_steps.iter().any(|s| s == "tables") {
        results.tables_updated = table_updater::update_tables(inventory, config, excel_app, excel_path)?;
    }

    // Step 3: Deltas
    if active_steps.iter().any(|s| s == "deltas") {
        results.deltas_updated = delta_updater::update_deltas(
            inventory,
            config,
            presentation,
            excel_path,
            excel_app,
        )?;
    }

    // Step 4: Coloring
    if active_steps.iter().any(|s| s == "coloring") {
        results.tables_colored = color_coder::apply_color_coding(inventory, config)?;
    }

    // Step 5: Charts
    if active_steps.iter().any(|s| s == "charts") {
        results.charts_updated = chart_updater::update_charts(inventory, excel_path)?;
    }

    Ok(results)
}
