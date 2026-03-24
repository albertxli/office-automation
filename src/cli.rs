use clap::{Parser, Subcommand};

#[derive(Parser, Debug)]
#[command(
    name = "oa",
    about = "Office Automation — update PowerPoint presentations with Excel data via COM",
    version,
    propagate_version = true
)]
pub struct Cli {
    #[command(subcommand)]
    pub command: Commands,
}

#[derive(Subcommand, Debug)]
pub enum Commands {
    /// Run the update pipeline on one or more PPTX files
    #[command(
        after_long_help = "PIPELINE STEPS (executed in this order):\n  \
            links     Re-point OLE links to a new Excel file\n  \
            tables    Populate PPT tables from Excel ranges\n  \
            deltas    Swap delta indicator arrows based on sign\n  \
            coloring  Apply sign-based color coding (_ccst shapes)\n  \
            charts    Update chart data links\n\n\
            All steps run by default. Use --steps or --skip to control which run."
    )]
    Update(UpdateArgs),

    /// Execute a TOML runfile for batch processing
    Run(RunArgs),

    /// Validate PPT values against Excel source data
    Check(CheckArgs),

    /// Compare two PPTX files side by side
    Diff(DiffArgs),

    /// Inspect a PPTX file (read-only)
    Info(InfoArgs),

    /// Kill zombie PowerPoint and Excel processes
    Clean(CleanArgs),

    /// Show all available --set config keys and their defaults
    Config,
}

#[derive(Parser, Debug)]
pub struct UpdateArgs {
    /// PPTX files to process (glob patterns supported)
    #[arg(value_name = "FILES")]
    pub files: Vec<String>,

    /// Path to the Excel data file (auto-detected from OLE links if omitted)
    #[arg(short, long, value_name = "PATH")]
    pub excel: Option<String>,

    /// Open a native file dialog to select the Excel file
    #[arg(short, long)]
    pub pick: bool,

    /// Explicit PPTX=XLSX pair (repeatable). Use '=' separator.
    #[arg(long, value_name = "PPT=XLSX")]
    pub pair: Vec<String>,

    /// Output file (.pptx) or directory
    #[arg(short, long, value_name = "PATH")]
    pub output: Option<String>,

    /// Run only these pipeline steps (comma-separated: links,tables,deltas,coloring,charts)
    #[arg(long, value_delimiter = ',', value_name = "STEP,...")]
    pub steps: Vec<String>,

    /// Skip these pipeline steps (comma-separated, mutually exclusive with --steps)
    #[arg(long, value_delimiter = ',', value_name = "STEP,...")]
    pub skip: Vec<String>,

    /// Override a config value (repeatable, e.g. --set ccst.positive_color=#FF0000)
    #[arg(long, value_name = "KEY=VALUE")]
    pub set: Vec<String>,

    /// Run validation against Excel after processing
    #[arg(long)]
    pub check: bool,

    /// Show what would happen without saving changes
    #[arg(long)]
    pub dry_run: bool,

    /// Enable debug logging
    #[arg(short, long)]
    pub verbose: bool,

    /// Suppress all output except errors
    #[arg(short, long)]
    pub quiet: bool,
}

#[derive(Parser, Debug)]
pub struct RunArgs {
    /// Path to the TOML runfile
    #[arg(value_name = "RUNFILE")]
    pub runfile: String,

    /// Run validation after each job
    #[arg(long)]
    pub check: bool,

    /// Show what would happen without saving changes
    #[arg(long)]
    pub dry_run: bool,

    /// Enable debug logging
    #[arg(short, long)]
    pub verbose: bool,

    /// Suppress all output except errors
    #[arg(short, long)]
    pub quiet: bool,
}

#[derive(Parser, Debug)]
pub struct CheckArgs {
    /// PPTX file or runfile (.toml/.py) to validate
    #[arg(value_name = "FILE")]
    pub file: String,

    /// Excel file to check against (auto-detected from OLE links if omitted)
    #[arg(short, long, value_name = "PATH")]
    pub excel: Option<String>,

    /// Override a config value (repeatable)
    #[arg(long, value_name = "KEY=VALUE")]
    pub set: Vec<String>,

    /// Enable debug logging
    #[arg(short, long)]
    pub verbose: bool,
}

#[derive(Parser, Debug)]
pub struct DiffArgs {
    /// First PPTX file
    #[arg(value_name = "A.pptx")]
    pub file_a: String,

    /// Second PPTX file
    #[arg(value_name = "B.pptx")]
    pub file_b: String,

    /// Enable debug logging
    #[arg(short, long)]
    pub verbose: bool,
}

#[derive(Parser, Debug)]
pub struct InfoArgs {
    /// PPTX file to inspect
    #[arg(value_name = "FILE")]
    pub file: String,

    /// Show per-slide breakdown
    #[arg(short, long)]
    pub verbose: bool,
}

#[derive(Parser, Debug)]
pub struct CleanArgs {
    /// Kill processes without prompting for confirmation
    #[arg(short, long)]
    pub force: bool,
}

/// The valid pipeline step names.
pub const VALID_STEPS: &[&str] = &["links", "tables", "deltas", "coloring", "charts"];

/// Resolve which steps to run from --steps and --skip flags.
/// Returns an error if both are specified, or if unknown step names are used.
pub fn resolve_steps(steps: &[String], skip: &[String]) -> Result<Vec<String>, String> {
    if !steps.is_empty() && !skip.is_empty() {
        return Err("Cannot use both --steps and --skip at the same time".into());
    }

    // Validate step names
    let validate = |names: &[String]| -> Result<(), String> {
        for name in names {
            if !VALID_STEPS.contains(&name.as_str()) {
                return Err(format!(
                    "Unknown step: {name:?}. Valid steps: {}",
                    VALID_STEPS.join(", ")
                ));
            }
        }
        Ok(())
    };

    if !steps.is_empty() {
        validate(steps)?;
        return Ok(steps.to_vec());
    }

    if !skip.is_empty() {
        validate(skip)?;
        return Ok(VALID_STEPS
            .iter()
            .filter(|s| !skip.iter().any(|sk| sk == *s))
            .map(|s| s.to_string())
            .collect());
    }

    // Default: all steps
    Ok(VALID_STEPS.iter().map(|s| s.to_string()).collect())
}

/// Parse a `--pair PPT=XLSX` value, handling Windows drive letters.
///
/// Examples:
/// - `report.pptx=data.xlsx` → ("report.pptx", "data.xlsx")
/// - `C:\report.pptx=C:\data.xlsx` → (`C:\report.pptx`, `C:\data.xlsx`)
pub fn parse_pair(pair: &str) -> Result<(String, String), String> {
    // Find '=' that isn't part of a Windows drive letter (X:\ pattern).
    // Strategy: split on '=' and rejoin pieces that look like drive prefixes.
    let parts: Vec<&str> = pair.split('=').collect();

    if parts.len() < 2 {
        return Err(format!("Invalid --pair format: {pair:?} (expected PPT=XLSX)"));
    }

    // Reassemble: a single-char part followed by something starting with '\' or '/'
    // is a drive letter that was split.
    let mut segments: Vec<String> = Vec::new();
    let mut i = 0;
    while i < parts.len() {
        if parts[i].len() == 1
            && parts[i].chars().next().unwrap().is_ascii_alphabetic()
            && i + 1 < parts.len()
            && (parts[i + 1].starts_with('\\') || parts[i + 1].starts_with('/'))
        {
            // Drive letter: rejoin "C" + "=" + "\path..."
            segments.push(format!("{}={}", parts[i], parts[i + 1]));
            i += 2;
        } else {
            segments.push(parts[i].to_string());
            i += 1;
        }
    }

    if segments.len() != 2 {
        return Err(format!(
            "Invalid --pair format: {pair:?} (expected PPT=XLSX, got {} segments)",
            segments.len()
        ));
    }

    Ok((segments[0].clone(), segments[1].clone()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_resolve_steps_default_all() {
        let steps = resolve_steps(&[], &[]).unwrap();
        assert_eq!(steps.len(), 5);
        assert_eq!(steps, vec!["links", "tables", "deltas", "coloring", "charts"]);
    }

    #[test]
    fn test_resolve_steps_include() {
        let steps = resolve_steps(&["links".into(), "tables".into()], &[]).unwrap();
        assert_eq!(steps, vec!["links", "tables"]);
    }

    #[test]
    fn test_resolve_steps_skip() {
        let steps = resolve_steps(&[], &["charts".into()]).unwrap();
        assert_eq!(steps, vec!["links", "tables", "deltas", "coloring"]);
    }

    #[test]
    fn test_resolve_steps_mutual_exclusion() {
        let result = resolve_steps(&["links".into()], &["charts".into()]);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Cannot use both"));
    }

    #[test]
    fn test_resolve_steps_unknown_step() {
        let result = resolve_steps(&["invalid".into()], &[]);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Unknown step"));
    }

    #[test]
    fn test_resolve_steps_unknown_skip() {
        let result = resolve_steps(&[], &["invalid".into()]);
        assert!(result.is_err());
    }

    #[test]
    fn test_parse_pair_simple() {
        let (pptx, xlsx) = parse_pair("report.pptx=data.xlsx").unwrap();
        assert_eq!(pptx, "report.pptx");
        assert_eq!(xlsx, "data.xlsx");
    }

    #[test]
    fn test_parse_pair_windows_paths() {
        let (pptx, xlsx) = parse_pair(r"C=\Users\report.pptx=C=\Data\file.xlsx").unwrap();
        assert_eq!(pptx, r"C=\Users\report.pptx");
        assert_eq!(xlsx, r"C=\Data\file.xlsx");
    }

    #[test]
    fn test_parse_pair_no_separator() {
        let result = parse_pair("no_equals_here");
        assert!(result.is_err());
    }

    #[test]
    fn test_parse_pair_relative_paths() {
        let (pptx, xlsx) = parse_pair("templates/report.pptx=data/us.xlsx").unwrap();
        assert_eq!(pptx, "templates/report.pptx");
        assert_eq!(xlsx, "data/us.xlsx");
    }
}
