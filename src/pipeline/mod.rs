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
pub mod verbose;

use std::time::{Duration, Instant};

use console::Style;
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
    #[allow(dead_code)]
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

impl PipelineResults {
    pub fn total_objects(&self) -> usize {
        self.steps.iter().map(|s| s.count).sum()
    }
}

// ── Styles ──────────────────────────────────────────────────
fn s_ok() -> Style { Style::new().green() }
fn s_count() -> Style { Style::new().white().bold() }
fn s_dim() -> Style { Style::new().dim() }

/// Fixed width for step name + dot leaders (ensures count column aligns).
const STEP_WIDTH: usize = 30;

/// Format a completed step line:  `• Links ··················· 86   0.0s`
pub fn format_step_line_pub(name: &str, count: usize, secs: f64) -> String {
    let leader_len = STEP_WIDTH.saturating_sub(name.len() + 1);
    let leaders = "·".repeat(leader_len);
    format!(
        "  {} {} {} {:>4}   {}",
        s_ok().apply_to("•"),
        name,
        s_dim().apply_to(&leaders),
        s_count().apply_to(count),
        s_dim().apply_to(format!("{secs:.1}s")),
    )
}

/// Public wrapper for update.rs to use for the Relink step.
pub fn make_spinner_pub(step_name: &str) -> ProgressBar { make_spinner(step_name) }

/// Create a spinner for a pipeline step.
fn make_spinner(step_name: &str) -> ProgressBar {
    let leader_len = STEP_WIDTH.saturating_sub(step_name.len() + 1);
    let leaders = "·".repeat(leader_len);
    let msg = format!("{step_name} {leaders}");

    let pb = ProgressBar::new_spinner();
    pb.set_style(
        ProgressStyle::default_spinner()
            .tick_strings(&["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"])
            .template("  {spinner:.cyan} {msg}")
            .unwrap()
    );
    pb.set_message(msg);
    pb.enable_steady_tick(Duration::from_millis(80));
    pb
}

/// Run a single pipeline step with spinner.
macro_rules! run_step {
    ($results:expr, $quiet:expr, $name:expr, $field:ident, $body:expr) => {{
        // Only show spinner in non-quiet, non-verbose mode.
        // In verbose mode, detail lines would conflict with the spinner redraw.
        let use_spinner = !$quiet && !verbose::is_verbose();
        let spinner = if use_spinner { Some(make_spinner($name)) } else { None };
        let t = Instant::now();

        let count = $body;
        $results.$field = count;

        let elapsed = t.elapsed().as_secs_f64();
        if let Some(pb) = spinner {
            pb.finish_and_clear();
        }
        // In verbose mode, detail lines already printed — skip the summary dot line.
        // In normal mode, the dot line IS the output.
        if !$quiet && !verbose::is_verbose() {
            println!("{}", format_step_line_pub($name, count, elapsed));
        }
        $results.steps.push(StepResult { name: $name, count, elapsed_secs: elapsed, ok: true });
    }};
}

/// Run the update pipeline on a presentation with the given steps.
#[allow(clippy::too_many_arguments)]
pub fn run_pipeline(
    inventory: &SlideInventory,
    config: &Config,
    presentation: &mut Dispatch,
    excel_app: &mut Dispatch,
    excel_path: &str,
    steps_include: &[String],
    steps_skip: &[String],
    quiet: bool,
    verbose: bool,
    skip_chart_refresh: bool,
) -> OaResult<PipelineResults> {
    verbose::set_verbose(verbose);

    let active_steps = resolve_steps(steps_include, steps_skip)
        .map_err(crate::error::OaError::Config)?;

    let mut results = PipelineResults::default();

    if active_steps.iter().any(|s| s == "links") {
        run_step!(results, quiet, "Links", links_updated,
            linker::update_links(inventory, excel_path, config)?);
    }

    if active_steps.iter().any(|s| s == "tables") {
        run_step!(results, quiet, "Tables", tables_updated,
            table_updater::update_tables(inventory, config, excel_app, excel_path)?);
    }

    if active_steps.iter().any(|s| s == "deltas") {
        run_step!(results, quiet, "Deltas", deltas_updated,
            delta_updater::update_deltas(inventory, config, presentation, excel_path, excel_app)?);
    }

    if active_steps.iter().any(|s| s == "coloring") {
        run_step!(results, quiet, "Coloring", tables_colored,
            color_coder::apply_color_coding(inventory, config)?);
    }

    if active_steps.iter().any(|s| s == "charts") {
        run_step!(results, quiet, "Charts", charts_updated,
            chart_updater::update_charts(inventory, excel_path, skip_chart_refresh)?);
    }

    Ok(results)
}
