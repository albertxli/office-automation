//! `oa run` — execute a TOML or Python runfile for batch processing.
//!
//! Supports both .toml and .py runfiles (auto-detected by extension).
//! Python parser is in `py_parser.rs`.
//!
//! TOML format (v2):
//! ```toml
//! [templates]
//! t1 = "template_1.pptx"
//!
//! data_path = "../excel_data"          # optional — prepended to job data paths
//! default_output = "output/{name}.pptx" # optional — {name} replaced with job name
//! # steps = ["links", "tables"]        # optional — omit to run all
//! # [config]                           # optional — config overrides
//!
//! [[job]]
//! name = "Argentina"
//! template = "t1"                       # alias from [templates] or direct path
//! data = "rpm_tracking_Argentina.xlsx"  # filename (data_path prepended) or full path
//! # output = "custom/argentina.pptx"   # optional — overrides default_output
//! ```

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use serde::Deserialize;

use console::Style;

use crate::cli::UpdateArgs;
use crate::com::session::ComSession;
use crate::commands::update::run_update_with_session;
use crate::error::{OaError, OaResult};

/// Result of a single job in a batch run.
enum JobOutcome {
    Ok { objects: usize, secs: f64 },
    Err(String),
}

struct JobResult {
    name: String,
    outcome: JobOutcome,
}

/// Parsed runfile (TOML v2 format).
#[derive(Debug, Deserialize)]
pub struct RunFile {
    /// Template aliases: `[templates]` section mapping short names to paths.
    #[serde(default)]
    pub templates: HashMap<String, String>,
    /// Prepended to relative job `data` paths. Optional.
    #[serde(default)]
    pub data_path: Option<String>,
    /// Output path template. `{name}` replaced with job name. Optional.
    #[serde(default)]
    pub default_output: Option<String>,
    /// Pipeline steps to run. Optional — omit to run all.
    #[serde(default)]
    pub steps: Option<Vec<String>>,
    /// Config overrides (`[config]` section). Optional.
    #[serde(default)]
    pub config: HashMap<String, toml::Value>,
    /// Job list (`[[job]]` array). At least one required.
    #[serde(default)]
    pub job: Vec<Job>,

    // --- Legacy format support (old [jobs."template"] style) ---
    #[serde(default)]
    pub jobs: Option<HashMap<String, HashMap<String, JobValue>>>,
}

/// A single job entry from `[[job]]`.
#[derive(Debug, Deserialize)]
pub struct Job {
    pub name: String,
    pub template: String,
    pub data: String,
    #[serde(default)]
    pub output: Option<String>,
}

/// Legacy job value (old format): either a plain string or `{data, output}`.
#[derive(Debug, Deserialize)]
#[serde(untagged)]
pub enum JobValue {
    Simple(String),
    Detailed { data: String, output: Option<String> },
}

impl JobValue {
    pub fn excel_path(&self) -> &str {
        match self {
            JobValue::Simple(s) => s,
            JobValue::Detailed { data, .. } => data,
        }
    }

    pub fn output_path(&self) -> Option<&str> {
        match self {
            JobValue::Simple(_) => None,
            JobValue::Detailed { output, .. } => output.as_deref(),
        }
    }
}

#[derive(Debug)]
pub struct ResolvedJob {
    pub name: String,
    pub template: PathBuf,
    pub excel: PathBuf,
    pub output: PathBuf,
}

/// Parse a runfile and resolve all jobs. Returns (jobs, config_overrides, steps).
pub fn parse_runfile(runfile_path: &Path) -> OaResult<(Vec<ResolvedJob>, Vec<String>, Vec<String>)> {
    if !runfile_path.exists() {
        return Err(OaError::Other(format!("Runfile not found: {}", runfile_path.display())));
    }

    let content = std::fs::read_to_string(runfile_path)?;
    let ext = runfile_path.extension().and_then(|e| e.to_str()).unwrap_or("");
    let runfile: RunFile = match ext {
        "py" => super::py_parser::parse_py_runfile(&content)?,
        "toml" => toml::from_str(&content)
            .map_err(|e| OaError::Config(format!("Failed to parse TOML runfile: {e}")))?,
        _ => return Err(OaError::Config(format!("Unsupported runfile format: .{ext} (use .toml or .py)"))),
    };

    let base_dir = runfile_path.parent().unwrap_or(Path::new("."));
    let config_overrides: Vec<String> = runfile.config.iter()
        .map(|(k, v)| format!("{k}={}", toml_value_to_string(v)))
        .collect();

    let steps = runfile.steps.clone().unwrap_or_default();
    let jobs = resolve_jobs(&runfile, base_dir)?;
    Ok((jobs, config_overrides, steps))
}

/// Run the `oa run` command.
pub fn run_runfile(
    runfile_path: &str,
    check_after: bool,
    dry_run: bool,
    verbose: bool,
    quiet: bool,
) -> OaResult<()> {
    let runfile_path = Path::new(runfile_path);
    let (jobs, config_overrides, steps) = parse_runfile(runfile_path)?;
    if jobs.is_empty() {
        println!("No jobs found in runfile.");
        return Ok(());
    }

    println!("Runfile: {} ({} jobs)", runfile_path.display(), jobs.len());

    // Create one COM session for all jobs (GOTCHA #39: avoids 0x80010001
    // from rapid COM create/destroy, saves ~1.1s setup per job).
    let mut session = ComSession::new()?;
    let total_start = std::time::Instant::now();
    let mut results: Vec<JobResult> = Vec::with_capacity(jobs.len());

    for (i, job) in jobs.iter().enumerate() {
        println!("\n--- Job {}/{}: {} ---", i + 1, jobs.len(), job.name);
        if let Some(parent) = job.output.parent()
            && !parent.exists()
        {
            std::fs::create_dir_all(parent)?;
        }

        let steps = steps.clone();

        let args = UpdateArgs {
            files: vec![job.template.to_string_lossy().to_string()],
            excel: Some(job.excel.to_string_lossy().to_string()),
            pick: false,
            pair: Vec::new(),
            output: Some(job.output.to_string_lossy().to_string()),
            steps,
            skip: Vec::new(),
            set: config_overrides.clone(),
            check: check_after,
            dry_run,
            verbose,
            quiet,
        };

        match run_update_with_session(&args, &mut session) {
            Ok((objects, secs)) => {
                results.push(JobResult {
                    name: job.name.clone(),
                    outcome: JobOutcome::Ok { objects, secs },
                });
            }
            Err(e) => {
                eprintln!("  Job '{}' failed: {e}", job.name);
                results.push(JobResult {
                    name: job.name.clone(),
                    outcome: JobOutcome::Err(e.to_string()),
                });
            }
        }
    }

    // Session drops here: Quit Excel → Quit PPT → CoUninitialize
    drop(session);

    let total_elapsed = total_start.elapsed().as_secs_f64();
    print_run_summary(&results, total_elapsed);
    Ok(())
}

pub fn resolve_jobs(runfile: &RunFile, base_dir: &Path) -> OaResult<Vec<ResolvedJob>> {
    // New v2 format: [[job]] array
    if !runfile.job.is_empty() {
        return resolve_jobs_v2(runfile, base_dir);
    }

    // Legacy format: [jobs."template_path"]
    if let Some(ref jobs_map) = runfile.jobs {
        return resolve_jobs_legacy(runfile, jobs_map, base_dir);
    }

    Ok(vec![])
}

/// Resolve jobs from new v2 `[[job]]` format.
fn resolve_jobs_v2(runfile: &RunFile, base_dir: &Path) -> OaResult<Vec<ResolvedJob>> {
    let mut jobs = Vec::new();

    for job in &runfile.job {
        // Resolve template: check aliases first, then use as path
        let template_path = runfile.templates.get(&job.template)
            .map(|s| s.as_str())
            .unwrap_or(&job.template);
        let template = resolve_path(base_dir, template_path);
        if !template.exists() {
            eprintln!("Warning: template not found for job '{}': {}", job.name, template.display());
            continue;
        }

        // Resolve data: if absolute use as-is, else prepend data_path (if set) or base_dir
        let data_path_str = &job.data;
        let excel = if Path::new(data_path_str).is_absolute() {
            PathBuf::from(data_path_str)
        } else if let Some(ref dp) = runfile.data_path {
            resolve_path(base_dir, dp).join(data_path_str)
        } else {
            resolve_path(base_dir, data_path_str)
        };
        if !excel.exists() {
            eprintln!("Warning: Excel not found for job '{}': {}", job.name, excel.display());
            continue;
        }

        // Resolve output: per-job output > default_output > fallback
        let output = if let Some(ref custom) = job.output {
            resolve_path(base_dir, custom)
        } else if let Some(ref default) = runfile.default_output {
            resolve_path(base_dir, &default.replace("{name}", &job.name))
        } else {
            base_dir.join(format!("{}.pptx", job.name))
        };

        jobs.push(ResolvedJob { name: job.name.clone(), template, excel, output });
    }

    Ok(jobs)
}

/// Resolve jobs from legacy `[jobs."template"]` format.
fn resolve_jobs_legacy(
    runfile: &RunFile,
    jobs_map: &HashMap<String, HashMap<String, JobValue>>,
    base_dir: &Path,
) -> OaResult<Vec<ResolvedJob>> {
    let mut jobs = Vec::new();
    for (template_path, job_map) in jobs_map {
        let template = resolve_path(base_dir, template_path);
        if !template.exists() {
            eprintln!("Warning: template not found: {}", template.display());
            continue;
        }
        for (name, value) in job_map {
            let excel = resolve_path(base_dir, value.excel_path());
            if !excel.exists() {
                eprintln!("Warning: Excel not found for job '{}': {}", name, excel.display());
                continue;
            }
            let output = if let Some(custom_output) = value.output_path() {
                resolve_path(base_dir, custom_output)
            } else if let Some(ref default_output) = runfile.default_output {
                resolve_path(base_dir, &default_output.replace("{name}", name))
            } else {
                base_dir.join(format!("{name}.pptx"))
            };
            jobs.push(ResolvedJob { name: name.clone(), template: template.clone(), excel, output });
        }
    }
    Ok(jobs)
}

fn resolve_path(base_dir: &Path, path: &str) -> PathBuf {
    let p = Path::new(path);
    if p.is_absolute() { p.to_path_buf() } else { base_dir.join(p) }
}

fn toml_value_to_string(v: &toml::Value) -> String {
    match v {
        toml::Value::String(s) => s.clone(),
        toml::Value::Integer(i) => i.to_string(),
        toml::Value::Float(f) => f.to_string(),
        toml::Value::Boolean(b) => b.to_string(),
        other => other.to_string(),
    }
}

// ── Batch summary ───────────────────────────────────────────

/// Target column for dot-leader alignment (matches info.rs).
const SUMMARY_COL: usize = 48;

/// Format elapsed time: under 60s → "42.2s", over → "1m 12s".
pub fn fmt_time(secs: f64) -> String {
    if secs < 60.0 {
        format!("{secs:.1}s")
    } else {
        let mins = secs as u64 / 60;
        let rem = secs % 60.0;
        format!("{mins}m {rem:04.1}s")
    }
}

/// Print the batch summary table after all jobs complete.
fn print_run_summary(results: &[JobResult], total_elapsed: f64) {
    let s_dim = Style::new().dim();
    let s_ok = Style::new().green();
    let s_err = Style::new().red();
    let s_count = Style::new().white().bold();

    // Separator + header
    println!();
    println!("  {}", s_dim.apply_to("═".repeat(39)));
    println!("  {}", s_dim.apply_to("Job summary"));
    println!();

    // Per-job recap rows
    for result in results {
        let (icon, icon_style) = match &result.outcome {
            JobOutcome::Ok { .. } => ("✓", &s_ok),
            JobOutcome::Err(_) => ("✗", &s_err),
        };

        // "  ✓ " = 4 chars prefix
        let prefix_len = 4;
        let name_len = result.name.chars().count();
        let leader_len = SUMMARY_COL.saturating_sub(prefix_len + name_len + 1);
        let leaders = "·".repeat(leader_len);

        match &result.outcome {
            JobOutcome::Ok { objects, secs } => {
                println!("  {} {} {} {} {} {}",
                    icon_style.apply_to(icon),
                    result.name,
                    s_dim.apply_to(&leaders),
                    s_count.apply_to(format!("{objects:>4}")),
                    s_dim.apply_to("objects"),
                    s_dim.apply_to(format!("{:>6}", fmt_time(*secs))),
                );
            }
            JobOutcome::Err(msg) => {
                println!("  {} {} {} {}",
                    icon_style.apply_to(icon),
                    result.name,
                    s_dim.apply_to(&leaders),
                    s_err.apply_to(msg),
                );
            }
        }
    }

    // Totals
    let passed: Vec<&JobResult> = results.iter()
        .filter(|r| matches!(&r.outcome, JobOutcome::Ok { .. }))
        .collect();
    let fail_count = results.len() - passed.len();
    let total_objects: usize = passed.iter()
        .map(|r| match &r.outcome { JobOutcome::Ok { objects, .. } => *objects, _ => 0 })
        .sum();
    let avg_secs = if passed.is_empty() {
        0.0
    } else {
        let sum: f64 = passed.iter()
            .map(|r| match &r.outcome { JobOutcome::Ok { secs, .. } => *secs, _ => 0.0 })
            .sum();
        sum / passed.len() as f64
    };

    // Totals line
    let status = if fail_count == 0 {
        format!("{} {}", s_ok.apply_to("✓"), s_ok.apply_to("all jobs complete"))
    } else {
        format!("{} {}", s_err.apply_to("✗"),
            s_err.apply_to(format!("{fail_count} job{} failed",
                if fail_count == 1 { "" } else { "s" })))
    };

    println!();
    println!("  {} {} {}{} {} {} {} {} {} {}{}",
        status,
        s_dim.apply_to("·"),
        s_count.apply_to(passed.len()),
        s_dim.apply_to(format!("/{}", results.len())),
        s_dim.apply_to("files ·"),
        s_count.apply_to(total_objects),
        s_dim.apply_to("objects ·"),
        s_ok.apply_to(fmt_time(total_elapsed)),
        s_dim.apply_to("· avg"),
        s_count.apply_to(fmt_time(avg_secs)),
        s_dim.apply_to("/file"),
    );
}

#[cfg(test)]
mod tests {
    use super::*;

    // --- New v2 format tests ---

    #[test]
    fn test_parse_v2_simple() {
        let toml_str = r#"
[[job]]
name = "Argentina"
template = "template.pptx"
data = "data/argentina.xlsx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert_eq!(rf.job.len(), 1);
        assert_eq!(rf.job[0].name, "Argentina");
        assert_eq!(rf.job[0].template, "template.pptx");
        assert_eq!(rf.job[0].data, "data/argentina.xlsx");
        assert!(rf.job[0].output.is_none());
    }

    #[test]
    fn test_parse_v2_with_templates_and_data_path() {
        let toml_str = r#"
data_path = "../excel_data"
default_output = "output/{name}.pptx"

[templates]
t1 = "templates/market_report.pptx"

[[job]]
name = "Argentina"
template = "t1"
data = "rpm_tracking_Argentina.xlsx"

[[job]]
name = "Brazil"
template = "t1"
data = "rpm_tracking_Brazil.xlsx"
output = "custom/brazil.pptx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert_eq!(rf.templates.get("t1").unwrap(), "templates/market_report.pptx");
        assert_eq!(rf.data_path.as_deref(), Some("../excel_data"));
        assert_eq!(rf.default_output.as_deref(), Some("output/{name}.pptx"));
        assert_eq!(rf.job.len(), 2);
        assert_eq!(rf.job[1].output.as_deref(), Some("custom/brazil.pptx"));
    }

    #[test]
    fn test_parse_v2_minimal() {
        let toml_str = r#"
[[job]]
name = "test"
template = "t.pptx"
data = "C:/full/path/data.xlsx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert!(rf.templates.is_empty());
        assert!(rf.data_path.is_none());
        assert!(rf.default_output.is_none());
        assert!(rf.steps.is_none());
        assert!(rf.config.is_empty());
        assert_eq!(rf.job.len(), 1);
    }

    #[test]
    fn test_parse_v2_with_steps_and_config() {
        let toml_str = r#"
steps = ["links", "tables"]

[config]
"ccst.positive_prefix" = ""

[[job]]
name = "test"
template = "t.pptx"
data = "data.xlsx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert_eq!(rf.steps.as_ref().unwrap(), &["links", "tables"]);
        assert_eq!(rf.config.len(), 1);
    }

    // --- Legacy format tests ---

    #[test]
    fn test_parse_legacy_format() {
        let toml_str = r#"
default_output = "output/{name}.pptx"
[jobs."template.pptx"]
us = "data/us.xlsx"
mx = "data/mx.xlsx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert!(rf.job.is_empty());
        assert!(rf.jobs.is_some());
        let jobs = rf.jobs.as_ref().unwrap();
        assert_eq!(jobs.len(), 1);
        assert_eq!(jobs.get("template.pptx").unwrap().len(), 2);
    }

    // --- Path resolution tests ---

    #[test]
    fn test_resolve_path_absolute() {
        let base = Path::new("C:/project");
        assert_eq!(resolve_path(base, "C:/data/file.xlsx"), PathBuf::from("C:/data/file.xlsx"));
    }

    #[test]
    fn test_resolve_path_relative() {
        let base = Path::new("C:/project");
        assert_eq!(resolve_path(base, "data/file.xlsx"), PathBuf::from("C:/project/data/file.xlsx"));
    }

    #[test]
    fn test_output_expansion() {
        let expanded = "output/rpm_2024_{name}.pptx".replace("{name}", "australia");
        assert_eq!(expanded, "output/rpm_2024_australia.pptx");
    }

    // --- Time formatting tests ---

    #[test]
    fn test_fmt_time_under_60() {
        assert_eq!(fmt_time(0.0), "0.0s");
        assert_eq!(fmt_time(5.4), "5.4s");
        assert_eq!(fmt_time(42.24), "42.2s");
        assert_eq!(fmt_time(59.9), "59.9s");
    }

    #[test]
    fn test_fmt_time_over_60() {
        assert_eq!(fmt_time(60.0), "1m 00.0s");
        assert_eq!(fmt_time(72.3), "1m 12.3s");
        assert_eq!(fmt_time(272.0), "4m 32.0s");
    }
}
