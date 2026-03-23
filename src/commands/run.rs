//! `oa run` — execute a TOML or Python runfile for batch processing.
//!
//! Supports both .toml and .py runfiles (auto-detected by extension).
//! Python parser is in `py_parser.rs`.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use serde::Deserialize;

use crate::cli::UpdateArgs;
use crate::commands::update::run_update;
use crate::error::{OaError, OaResult};

/// Parsed runfile (from TOML or Python).
#[derive(Debug, Deserialize)]
pub struct RunFile {
    #[serde(default)]
    pub default_output: Option<String>,
    #[serde(default)]
    pub steps: Option<Vec<String>>,
    #[serde(default)]
    pub config: HashMap<String, toml::Value>,
    pub jobs: HashMap<String, HashMap<String, JobValue>>,
}

/// A job value: either a plain Excel path string, or a table with `data` + optional `output`.
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
    let mut jobs = Vec::new();
    for (template_path, job_map) in &runfile.jobs {
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

    #[test]
    fn test_parse_simple_runfile() {
        let toml_str = r#"
default_output = "output/{name}.pptx"
[jobs."template.pptx"]
us = "data/us.xlsx"
mx = "data/mx.xlsx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert_eq!(rf.default_output.as_deref(), Some("output/{name}.pptx"));
        assert_eq!(rf.jobs.len(), 1);
        assert_eq!(rf.jobs.get("template.pptx").unwrap().len(), 2);
    }

    #[test]
    fn test_parse_with_detailed_job() {
        let toml_str = r#"
[jobs."template.pptx"]
us = "data/us.xlsx"
mx = { data = "data/mx.xlsx", output = "special/mx.pptx" }
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        let mx = rf.jobs.get("template.pptx").unwrap().get("mx").unwrap();
        assert_eq!(mx.excel_path(), "data/mx.xlsx");
        assert_eq!(mx.output_path(), Some("special/mx.pptx"));
    }

    #[test]
    fn test_parse_with_steps_and_config() {
        let toml_str = r#"
steps = ["links", "tables"]
[config]
"ccst.positive_prefix" = ""
"links.set_manual" = true
[jobs."template.pptx"]
us = "data/us.xlsx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert_eq!(rf.steps.as_ref().unwrap(), &["links", "tables"]);
        assert_eq!(rf.config.len(), 2);
    }

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
