//! Python runfile parser — parse decx-compatible .py runfiles.
//!
//! Supports:
//! - `jobs = { "template.pptx": { "name": "data.xlsx" } }` — nested dicts
//! - `default_output = "output/{name}.pptx"` — strings
//! - `steps = ["links", "tables"]` — lists
//! - `config = { "key": "value" }` — flat dicts (with commented keys)
//! - `VARIABLE = "value"` + f-strings — variable interpolation
//! - `#` comments, `"""` docstrings

use std::collections::HashMap;

use crate::commands::run::{RunFile, JobValue};
use crate::error::OaResult;

/// Parse a Python runfile (.py) into a RunFile struct.
pub fn parse_py_runfile(content: &str) -> OaResult<RunFile> {
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

    // Step 3: Extract top-level assignments
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
        templates: HashMap::new(),
        data_path: None,
        default_output,
        steps,
        config,
        job: Vec::new(),
        jobs: Some(jobs),
    })
}

/// Extract top-level assignments from Python source.
fn extract_assignments(source: &str) -> HashMap<String, String> {
    let mut result: HashMap<String, String> = HashMap::new();
    let mut current_var: Option<String> = None;
    let mut current_expr = String::new();
    let mut brace_depth = 0i32;
    let mut bracket_depth = 0i32;

    for line in source.lines() {
        let trimmed = line.trim();
        if trimmed.starts_with('#') || trimmed.is_empty() {
            continue;
        }

        if brace_depth == 0 && bracket_depth == 0
            && let Some(eq_pos) = trimmed.find('=') {
                let var_name = trimmed[..eq_pos].trim();
                if var_name.chars().all(|c| c.is_alphanumeric() || c == '_') && !var_name.is_empty() {
                    if let Some(prev_var) = current_var.take() {
                        result.insert(prev_var, current_expr.trim().to_string());
                    }
                    current_var = Some(var_name.to_string());
                    current_expr = trimmed[eq_pos + 1..].to_string();
                    for ch in current_expr.chars() {
                        match ch { '{' => brace_depth += 1, '}' => brace_depth -= 1,
                                   '[' => bracket_depth += 1, ']' => bracket_depth -= 1, _ => {} }
                    }
                    continue;
                }
            }

        if current_var.is_some() {
            current_expr.push('\n');
            current_expr.push_str(trimmed);
            for ch in trimmed.chars() {
                match ch { '{' => brace_depth += 1, '}' => brace_depth -= 1,
                           '[' => bracket_depth += 1, ']' => bracket_depth -= 1, _ => {} }
            }
        }
    }

    if let Some(var) = current_var {
        result.insert(var, current_expr.trim().to_string());
    }
    result
}

fn extract_string_value(expr: &str) -> String {
    let trimmed = expr.trim().trim_matches('"').trim_matches('\'');
    let trimmed = trimmed.strip_prefix("f\"").unwrap_or(trimmed);
    let trimmed = trimmed.strip_suffix('"').unwrap_or(trimmed);
    trimmed.to_string()
}

fn extract_string_list(expr: &str) -> Vec<String> {
    let re = regex::Regex::new(r#""([^"]+)""#).unwrap();
    re.captures_iter(expr).map(|c| c[1].to_string()).collect()
}

fn extract_string_dict(expr: &str) -> HashMap<String, String> {
    let re = regex::Regex::new(r#""([^"]+)"\s*:\s*"([^"]*)"#).unwrap();
    let mut map = HashMap::new();
    for caps in re.captures_iter(expr) { map.insert(caps[1].to_string(), caps[2].to_string()); }
    let re_bool = regex::Regex::new(r#""([^"]+)"\s*:\s*(True|False|\d+)"#).unwrap();
    for caps in re_bool.captures_iter(expr) { map.entry(caps[1].to_string()).or_insert_with(|| caps[2].to_string()); }
    map
}

fn extract_nested_dict(expr: &str, variables: &HashMap<String, String>) -> HashMap<String, HashMap<String, String>> {
    let mut result = HashMap::new();
    let chars: Vec<char> = expr.chars().collect();
    let len = chars.len();
    let mut i = 0;

    while i < len {
        if chars[i] == '"' {
            let key_start = i + 1;
            i += 1;
            while i < len && chars[i] != '"' { i += 1; }
            let key = chars[key_start..i].iter().collect::<String>();
            i += 1;
            while i < len && chars[i] != '{' { i += 1; }
            if i >= len { break; }
            let inner_start = i;
            let mut depth = 0;
            while i < len {
                match chars[i] {
                    '{' => depth += 1,
                    '}' => { depth -= 1; if depth == 0 { i += 1; break; } }
                    _ => {}
                }
                i += 1;
            }
            let inner_str: String = chars[inner_start..i].iter().collect();
            let inner = parse_inner_jobs(&inner_str, variables);
            if !inner.is_empty() {
                result.insert(substitute_fstrings(key, variables), inner);
            }
        } else { i += 1; }
    }
    result
}

fn parse_inner_jobs(expr: &str, variables: &HashMap<String, String>) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let re_simple = regex::Regex::new(r#""([^"]+)"\s*:\s*(?:f)?"([^"]*)"#).unwrap();
    let re_dict = regex::Regex::new(r#""([^"]+)"\s*:\s*\{([^}]*)\}"#).unwrap();

    let re_data = regex::Regex::new(r#""data"\s*:\s*(?:f")?"([^"]*)"#).unwrap();
    let re_output = regex::Regex::new(r#""output"\s*:\s*(?:f")?"([^"]*)"#).unwrap();

    for caps in re_dict.captures_iter(expr) {
        let name = caps[1].to_string();
        let dict_content = &caps[2];
        let data = re_data
            .captures(dict_content).map(|c| c[1].to_string()).unwrap_or_default();
        let output = re_output
            .captures(dict_content).map(|c| c[1].to_string());
        let value = if let Some(out) = output {
            format!("{{\"data\":\"{}\",\"output\":\"{}\"}}", substitute_fstrings(data, variables), substitute_fstrings(out, variables))
        } else { substitute_fstrings(data, variables) };
        map.insert(name, value);
    }

    for caps in re_simple.captures_iter(expr) {
        let name = caps[1].to_string();
        if !map.contains_key(&name) && name != "data" && name != "output" {
            map.insert(name, substitute_fstrings(caps[2].to_string(), variables));
        }
    }
    map
}

fn parse_job_dict_value(value: &str) -> Option<JobValue> {
    if !value.starts_with("{\"data\"") { return None; }
    let data = regex::Regex::new(r#""data"\s*:\s*"([^"]*)"#).ok()?
        .captures(value)?.get(1)?.as_str().to_string();
    let output = regex::Regex::new(r#""output"\s*:\s*"([^"]*)"#).ok()?
        .captures(value).map(|c| c[1].to_string());
    Some(JobValue::Detailed { data, output })
}

fn substitute_fstrings(s: String, variables: &HashMap<String, String>) -> String {
    let mut result = s;
    for (name, value) in variables { result = result.replace(&format!("{{{name}}}"), value); }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

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
        assert_eq!(rf.jobs.as_ref().unwrap().len(), 1);
        let jobs = rf.jobs.as_ref().unwrap().get("template.pptx").unwrap();
        assert_eq!(jobs.len(), 2);
        assert_eq!(jobs.get("us").unwrap().excel_path(), "data/us.xlsx");
    }

    #[test]
    fn test_parse_py_with_steps() {
        let py = r#"
jobs = { "t.pptx": { "a": "a.xlsx" } }
steps = ["links", "tables"]
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.steps.as_ref().unwrap(), &["links", "tables"]);
    }

    #[test]
    fn test_parse_py_with_fstring() {
        let py = r#"
DATAPATH = "C:/data"
jobs = { "template.pptx": { "us": f"{DATAPATH}/us.xlsx" } }
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.jobs.as_ref().unwrap().get("template.pptx").unwrap().get("us").unwrap().excel_path(), "C:/data/us.xlsx");
    }

    #[test]
    fn test_parse_py_with_docstring() {
        let py = r#"
"""This is ignored."""
jobs = { "t.pptx": { "a": "a.xlsx" } }
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.jobs.as_ref().unwrap().len(), 1);
    }

    #[test]
    fn test_parse_py_multiple_templates() {
        let py = r#"
jobs = {
    "r1.pptx": { "au": "au.xlsx", "jp": "jp.xlsx" },
    "r2.pptx": { "de": "de.xlsx" },
}
"#;
        let rf = parse_py_runfile(py).unwrap();
        assert_eq!(rf.jobs.as_ref().unwrap().len(), 2);
    }
}
