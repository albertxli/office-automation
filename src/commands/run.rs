//! `oa run` — execute a TOML runfile for batch processing.
//!
//! TOML format replaces Python runfiles. Structure:
//! ```toml
//! default_output = "output/rpm_2024_{name}.pptx"
//! steps = ["links", "tables", "deltas", "coloring", "charts"]
//!
//! [config]
//! ccst.positive_prefix = ""
//!
//! [jobs."templates/template.pptx"]
//! australia = "data/australia.xlsx"
//! japan = { data = "data/japan.xlsx", output = "special/japan.pptx" }
//! ```

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use serde::Deserialize;

use crate::cli::UpdateArgs;
use crate::commands::update::run_update;
use crate::error::{OaError, OaResult};

/// Parsed TOML runfile.
#[derive(Debug, Deserialize)]
pub struct RunFile {
    /// Default output path template. `{name}` is replaced with the job name.
    /// Can be a directory (ends with `/`) or a file pattern.
    #[serde(default)]
    pub default_output: Option<String>,

    /// Pipeline steps to run (defaults to all).
    #[serde(default)]
    pub steps: Option<Vec<String>>,

    /// Config overrides (dot-notation keys).
    #[serde(default)]
    pub config: HashMap<String, toml::Value>,

    /// Jobs: template path → { job_name → excel_path_or_spec }
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
    fn excel_path(&self) -> &str {
        match self {
            JobValue::Simple(s) => s,
            JobValue::Detailed { data, .. } => data,
        }
    }

    fn output_path(&self) -> Option<&str> {
        match self {
            JobValue::Simple(_) => None,
            JobValue::Detailed { output, .. } => output.as_deref(),
        }
    }
}

/// A resolved job ready for processing.
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
        "py" => parse_py_runfile(&content)?,
        "toml" => toml::from_str(&content)
            .map_err(|e| OaError::Config(format!("Failed to parse TOML runfile: {e}")))?,
        _ => return Err(OaError::Config(format!("Unsupported runfile format: .{ext} (use .toml or .py)"))),
    };

    // Resolve paths relative to runfile directory
    let base_dir = runfile_path.parent().unwrap_or(Path::new("."));

    // Resolve config overrides to --set format
    let config_overrides: Vec<String> = runfile.config.iter()
        .map(|(k, v)| format!("{k}={}", toml_value_to_string(v)))
        .collect();

    // Resolve all jobs
    let jobs = resolve_jobs(&runfile, base_dir)?;

    if jobs.is_empty() {
        println!("No jobs found in runfile.");
        return Ok(());
    }

    println!("Runfile: {} ({} jobs)", runfile_path.display(), jobs.len());

    // Process each job by building UpdateArgs and calling run_update
    for job in &jobs {
        println!("\n--- Job: {} ---", job.name);

        // Ensure output directory exists
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

/// Resolve all jobs from the runfile into concrete file paths.
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
                let expanded = default_output.replace("{name}", name);
                resolve_path(base_dir, &expanded)
            } else {
                // Default: same directory as template, named {name}.pptx
                base_dir.join(format!("{name}.pptx"))
            };

            jobs.push(ResolvedJob {
                name: name.clone(),
                template: template.clone(),
                excel,
                output,
            });
        }
    }

    Ok(jobs)
}

fn resolve_path(base_dir: &Path, path: &str) -> PathBuf {
    let p = Path::new(path);
    if p.is_absolute() {
        p.to_path_buf()
    } else {
        base_dir.join(p)
    }
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

// ── Python runfile parser ─────────────────────────────────────

/// Parse a Python runfile (.py) into a RunFile struct.
///
/// Supports the decx runfile format:
/// - `jobs = { "template.pptx": { "name": "data.xlsx" } }`
/// - `default_output = "output/{name}.pptx"`
/// - `steps = ["links", "tables"]`
/// - `config = { "key": "value" }`
/// - `VARIABLE = "value"` for f-string substitution
/// - `#` comments, `"""` docstrings
fn parse_py_runfile(content: &str) -> OaResult<RunFile> {
    use regex::Regex;

    // Step 1: Strip docstrings (""" ... """)
    let re_docstring = Regex::new(r#""""[\s\S]*?""""#).unwrap();
    let stripped = re_docstring.replace_all(content, "");

    // Step 2: Collect user variables (UPPERCASE = "value") for f-string substitution
    let mut variables: HashMap<String, String> = HashMap::new();
    let re_var = Regex::new(r#"^([A-Z_][A-Z0-9_]*)\s*=\s*"([^"]*)"#).unwrap();
    for line in stripped.lines() {
        let line = line.trim();
        if let Some(caps) = re_var.captures(line) {
            variables.insert(caps[1].to_string(), caps[2].to_string());
        }
    }

    // Step 3: Extract top-level assignments by finding `varname = ...` and collecting
    // everything until the next top-level assignment or EOF
    let assignments = extract_assignments(&stripped);

    // Step 4: Parse each assignment
    let default_output = assignments.get("default_output")
        .map(|s| substitute_fstrings(extract_string_value(s), &variables));

    let steps = assignments.get("steps")
        .map(|s| extract_string_list(s));

    let config_raw = assignments.get("config")
        .map(|s| extract_string_dict(s))
        .unwrap_or_default();
    let config: HashMap<String, toml::Value> = config_raw.into_iter()
        .map(|(k, v)| {
            let val = if v == "True" || v == "true" {
                toml::Value::Boolean(true)
            } else if v == "False" || v == "false" {
                toml::Value::Boolean(false)
            } else if let Ok(i) = v.parse::<i64>() {
                toml::Value::Integer(i)
            } else {
                toml::Value::String(v)
            };
            (k, val)
        })
        .collect();

    let jobs_raw = assignments.get("jobs")
        .map(|s| extract_nested_dict(s, &variables))
        .unwrap_or_default();

    // Convert to RunFile format
    let mut jobs: HashMap<String, HashMap<String, JobValue>> = HashMap::new();
    for (template, job_map) in jobs_raw {
        let mut inner: HashMap<String, JobValue> = HashMap::new();
        for (name, value) in job_map {
            if let Some(data_output) = parse_job_dict_value(&value) {
                inner.insert(name, data_output);
            } else {
                inner.insert(name, JobValue::Simple(value));
            }
        }
        jobs.insert(template, inner);
    }

    Ok(RunFile {
        default_output,
        steps,
        config,
        jobs,
    })
}

/// Extract top-level assignments from Python source.
/// Returns map of variable_name → full expression text.
fn extract_assignments(source: &str) -> HashMap<String, String> {
    let mut result: HashMap<String, String> = HashMap::new();
    let mut current_var: Option<String> = None;
    let mut current_expr = String::new();
    let mut brace_depth = 0i32;
    let mut bracket_depth = 0i32;

    for line in source.lines() {
        let trimmed = line.trim();

        // Skip comments and empty lines at top level
        if trimmed.starts_with('#') || trimmed.is_empty() {
            if brace_depth > 0 || bracket_depth > 0 {
                // Inside a dict/list — keep going but skip comment content
                continue;
            }
            continue;
        }

        // Check for new top-level assignment
        if brace_depth == 0 && bracket_depth == 0 {
            if let Some(eq_pos) = trimmed.find('=') {
                let var_name = trimmed[..eq_pos].trim();
                // Must be a simple identifier (not inside a dict)
                if var_name.chars().all(|c| c.is_alphanumeric() || c == '_') && !var_name.is_empty() {
                    // Save previous assignment
                    if let Some(prev_var) = current_var.take() {
                        result.insert(prev_var, current_expr.trim().to_string());
                    }
                    current_var = Some(var_name.to_string());
                    current_expr = trimmed[eq_pos + 1..].to_string();

                    // Count braces in this line
                    for ch in current_expr.chars() {
                        match ch {
                            '{' => brace_depth += 1,
                            '}' => brace_depth -= 1,
                            '[' => bracket_depth += 1,
                            ']' => bracket_depth -= 1,
                            _ => {}
                        }
                    }
                    continue;
                }
            }
        }

        // Continue collecting multi-line expression
        if current_var.is_some() {
            current_expr.push('\n');
            current_expr.push_str(trimmed);
            for ch in trimmed.chars() {
                match ch {
                    '{' => brace_depth += 1,
                    '}' => brace_depth -= 1,
                    '[' => bracket_depth += 1,
                    ']' => bracket_depth -= 1,
                    _ => {}
                }
            }
        }
    }

    // Save last assignment
    if let Some(var) = current_var {
        result.insert(var, current_expr.trim().to_string());
    }

    result
}

/// Extract a quoted string value from a Python expression.
fn extract_string_value(expr: &str) -> String {
    let trimmed = expr.trim().trim_matches('"').trim_matches('\'');
    // Handle f-strings: f"..."
    let trimmed = trimmed.strip_prefix("f\"").unwrap_or(trimmed);
    let trimmed = trimmed.strip_suffix('"').unwrap_or(trimmed);
    trimmed.to_string()
}

/// Extract a list of strings from `["a", "b", "c"]`.
fn extract_string_list(expr: &str) -> Vec<String> {
    let re = regex::Regex::new(r#""([^"]+)""#).unwrap();
    re.captures_iter(expr)
        .map(|c| c[1].to_string())
        .collect()
}

/// Extract a flat dict of strings from `{"key": "value"}`.
fn extract_string_dict(expr: &str) -> HashMap<String, String> {
    let re = regex::Regex::new(r#""([^"]+)"\s*:\s*"([^"]*)"#).unwrap();
    let mut map = HashMap::new();
    for caps in re.captures_iter(expr) {
        map.insert(caps[1].to_string(), caps[2].to_string());
    }
    // Also handle non-string values like True, False, integers
    let re_bool = regex::Regex::new(r#""([^"]+)"\s*:\s*(True|False|\d+)"#).unwrap();
    for caps in re_bool.captures_iter(expr) {
        map.entry(caps[1].to_string()).or_insert_with(|| caps[2].to_string());
    }
    map
}

/// Extract a nested dict: { "template": { "name": "value" | {"data": ..., "output": ...} } }
fn extract_nested_dict(expr: &str, variables: &HashMap<String, String>) -> HashMap<String, HashMap<String, String>> {
    let mut result: HashMap<String, HashMap<String, String>> = HashMap::new();

    // Find outer dict entries: "template_path": { ... }
    // Strategy: find each quoted key followed by : {, then collect until matching }
    let chars: Vec<char> = expr.chars().collect();
    let len = chars.len();
    let mut i = 0;

    while i < len {
        // Find a quoted string key
        if chars[i] == '"' {
            let key_start = i + 1;
            i += 1;
            while i < len && chars[i] != '"' { i += 1; }
            let key = chars[key_start..i].iter().collect::<String>();
            i += 1; // skip closing "

            // Skip to : {
            while i < len && chars[i] != '{' { i += 1; }
            if i >= len { break; }

            // Collect inner dict content
            let inner_start = i;
            let mut depth = 0;
            while i < len {
                match chars[i] {
                    '{' => depth += 1,
                    '}' => {
                        depth -= 1;
                        if depth == 0 { i += 1; break; }
                    }
                    _ => {}
                }
                i += 1;
            }
            let inner_str: String = chars[inner_start..i].iter().collect();

            // Parse inner dict entries
            let inner = parse_inner_jobs(&inner_str, variables);
            if !inner.is_empty() {
                result.insert(substitute_fstrings(key, variables), inner);
            }
        } else {
            i += 1;
        }
    }

    result
}

/// Parse inner job entries: "name": "path" or "name": {"data": ..., "output": ...}
fn parse_inner_jobs(expr: &str, variables: &HashMap<String, String>) -> HashMap<String, String> {
    let mut map = HashMap::new();

    // Match "key": "value" or "key": f"value"
    let re_simple = regex::Regex::new(r#""([^"]+)"\s*:\s*(?:f)?"([^"]*)"#).unwrap();
    // Match "key": {"data": "...", "output": "..."}
    let re_dict = regex::Regex::new(r#""([^"]+)"\s*:\s*\{([^}]*)\}"#).unwrap();

    // First try dict entries (they contain { })
    for caps in re_dict.captures_iter(expr) {
        let name = caps[1].to_string();
        let dict_content = &caps[2];
        // Reconstruct as a marker that resolve_jobs can parse
        let data = regex::Regex::new(r#""data"\s*:\s*(?:f")?"([^"]*)"#).unwrap()
            .captures(dict_content)
            .map(|c| c[1].to_string())
            .unwrap_or_default();
        let output = regex::Regex::new(r#""output"\s*:\s*(?:f")?"([^"]*)"#).unwrap()
            .captures(dict_content)
            .map(|c| c[1].to_string());

        let value = if let Some(out) = output {
            format!("{{\"data\":\"{}\",\"output\":\"{}\"}}", substitute_fstrings(data, variables), substitute_fstrings(out, variables))
        } else {
            substitute_fstrings(data, variables)
        };
        map.insert(name, value);
    }

    // Then simple string entries (skip those already captured as dicts)
    for caps in re_simple.captures_iter(expr) {
        let name = caps[1].to_string();
        if !map.contains_key(&name) && name != "data" && name != "output" {
            map.insert(name, substitute_fstrings(caps[2].to_string(), variables));
        }
    }

    map
}

/// Try to parse a job value as {"data": ..., "output": ...} marker.
fn parse_job_dict_value(value: &str) -> Option<JobValue> {
    if !value.starts_with("{\"data\"") { return None; }
    let data = regex::Regex::new(r#""data"\s*:\s*"([^"]*)"#).ok()?
        .captures(value)?.get(1)?.as_str().to_string();
    let output = regex::Regex::new(r#""output"\s*:\s*"([^"]*)"#).ok()?
        .captures(value)
        .map(|c| c[1].to_string());
    Some(JobValue::Detailed { data, output })
}

/// Substitute f-string variables: `{VARNAME}` → variable value.
fn substitute_fstrings(s: String, variables: &HashMap<String, String>) -> String {
    let mut result = s;
    for (name, value) in variables {
        result = result.replace(&format!("{{{name}}}"), value);
    }
    result
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
        let jobs = rf.jobs.get("template.pptx").unwrap();
        assert_eq!(jobs.len(), 2);
    }

    #[test]
    fn test_parse_with_detailed_job() {
        let toml_str = r#"
[jobs."template.pptx"]
us = "data/us.xlsx"
mx = { data = "data/mx.xlsx", output = "special/mx.pptx" }
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        let jobs = rf.jobs.get("template.pptx").unwrap();
        let mx = jobs.get("mx").unwrap();
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
    fn test_parse_multiple_templates() {
        let toml_str = r#"
default_output = "output/rpm_{name}.pptx"

[jobs."region1.pptx"]
australia = "data/au.xlsx"
japan = "data/jp.xlsx"

[jobs."region2.pptx"]
germany = "data/de.xlsx"
"#;
        let rf: RunFile = toml::from_str(toml_str).unwrap();
        assert_eq!(rf.jobs.len(), 2);
        assert_eq!(rf.jobs.get("region1.pptx").unwrap().len(), 2);
        assert_eq!(rf.jobs.get("region2.pptx").unwrap().len(), 1);
    }

    #[test]
    fn test_resolve_path_absolute() {
        let base = Path::new("C:/project");
        let result = resolve_path(base, "C:/data/file.xlsx");
        assert_eq!(result, PathBuf::from("C:/data/file.xlsx"));
    }

    #[test]
    fn test_resolve_path_relative() {
        let base = Path::new("C:/project");
        let result = resolve_path(base, "data/file.xlsx");
        assert_eq!(result, PathBuf::from("C:/project/data/file.xlsx"));
    }

    #[test]
    fn test_output_expansion() {
        let template = "output/rpm_2024_{name}.pptx";
        let expanded = template.replace("{name}", "australia");
        assert_eq!(expanded, "output/rpm_2024_australia.pptx");
    }

    #[test]
    fn test_parse_py_simple() {
        let py = r#"
jobs = {
    "template.pptx": {
        "us": "data/us.xlsx",
        "mx": "data/mx.xlsx",
    },
}
default_output = "output/{name}.pptx"
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.default_output.as_deref(), Some("output/{name}.pptx"));
        assert_eq!(rf.jobs.len(), 1);
        let jobs = rf.jobs.get("template.pptx").unwrap();
        assert_eq!(jobs.len(), 2);
        assert_eq!(jobs.get("us").unwrap().excel_path(), "data/us.xlsx");
    }

    #[test]
    fn test_parse_py_with_steps() {
        let py = r#"
jobs = {
    "t.pptx": {
        "a": "a.xlsx",
    },
}
steps = ["links", "tables"]
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.steps.as_ref().unwrap(), &["links", "tables"]);
    }

    #[test]
    fn test_parse_py_with_fstring() {
        let py = r#"
DATAPATH = "C:/data"

jobs = {
    "template.pptx": {
        "us": f"{DATAPATH}/us.xlsx",
    },
}
"#;
        let rf = parse_py_runfile(py).unwrap();
        let jobs = rf.jobs.get("template.pptx").unwrap();
        assert_eq!(jobs.get("us").unwrap().excel_path(), "C:/data/us.xlsx");
    }

    #[test]
    fn test_parse_py_with_docstring() {
        let py = r#"
"""
This is a docstring that should be ignored.
"""
jobs = {
    "t.pptx": {
        "a": "a.xlsx",
    },
}
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.jobs.len(), 1);
    }

    #[test]
    fn test_parse_py_multiple_templates() {
        let py = r#"
jobs = {
    "region1.pptx": {
        "au": "au.xlsx",
        "jp": "jp.xlsx",
    },
    "region2.pptx": {
        "de": "de.xlsx",
    },
}
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.jobs.len(), 2);
        assert_eq!(rf.jobs.get("region1.pptx").unwrap().len(), 2);
        assert_eq!(rf.jobs.get("region2.pptx").unwrap().len(), 1);
    }
}
