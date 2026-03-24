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

use crate::cli::UpdateArgs;
use crate::commands::update::run_update;
use crate::error::{OaError, OaResult};

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
struct ResolvedJob {
    name: String,
    template: PathBuf,
    excel: PathBuf,
    output: PathBuf,
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

    let jobs = resolve_jobs(&runfile, base_dir)?;
    if jobs.is_empty() {
        println!("No jobs found in runfile.");
        return Ok(());
    }

    println!("Runfile: {} ({} jobs)", runfile_path.display(), jobs.len());

    for job in &jobs {
        println!("\n--- Job: {} ---", job.name);
        if let Some(parent) = job.output.parent() {
            if !parent.exists() {
                std::fs::create_dir_all(parent)?;
            }
        }

        let mut steps = Vec::new();
        if let Some(ref s) = runfile.steps {
            steps = s.clone();
        }

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

        if let Err(e) = run_update(&args) {
            eprintln!("  Job '{}' failed: {e}", job.name);
        }
    }

    println!("\nRunfile complete.");
    Ok(())
}

fn resolve_jobs(runfile: &RunFile, base_dir: &Path) -> OaResult<Vec<ResolvedJob>> {
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
}
