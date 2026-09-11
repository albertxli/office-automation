//! End-to-end CLI tests using assert_cmd.
//! These test argument parsing and help output — no Office needed.

use std::path::PathBuf;

use assert_cmd::Command;
use predicates::prelude::*;

fn oa() -> Command {
    Command::cargo_bin("oa").unwrap()
}

/// A PPTX fixture from quick_test_files/, or None when the fixture is absent (CI runners).
fn fixture(name: &str) -> Option<PathBuf> {
    let p = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("quick_test_files").join(name);
    if p.exists() { Some(p) } else { eprintln!("Skipping: {name} not found"); None }
}

// ── oa find (pure ZIP, no Office needed — these run against real fixtures) ──

#[test]
fn test_find_help() {
    oa().args(["find", "--help"])
        .assert()
        .success()
        .stdout(predicate::str::contains("--text"))
        .stdout(predicate::str::contains("--ignore-case"))
        .stdout(predicate::str::contains("EXIT CODES"));
}

#[test]
fn test_find_single_run_hit() {
    let Some(pptx) = fixture("test_template.pptx") else { return };
    oa().arg("find").arg(&pptx).args(["-t", "Market Report"])
        .assert()
        .success()
        .stdout(predicate::str::contains("Slide  1"))
        .stdout(predicate::str::contains("Market Report"))
        .stdout(predicate::str::contains("✓"));
}

#[test]
fn test_find_reports_shape_name_and_counts_per_shape() {
    // "change" appears once in each of three text boxes on slide 1
    let Some(pptx) = fixture("linked_chart_validation_no_ole.pptx") else { return };
    oa().arg("find").arg(&pptx).args(["-t", "change"])
        .assert()
        .success()
        .stdout(predicate::str::contains("TextBox 4"))
        .stdout(predicate::str::contains("TextBox 10"))
        .stdout(predicate::str::contains("TextBox 13"))
        .stdout(predicate::str::contains("3 hits for \"change\""));
}

#[test]
fn test_find_ignore_case() {
    let Some(pptx) = fixture("test_template.pptx") else { return };
    oa().arg("find").arg(&pptx).args(["-t", "market report"]).assert().code(1);
    oa().arg("find").arg(&pptx).args(["-i", "-t", "market report"]).assert().success();
}

#[test]
fn test_find_multiple_needles_and_no_hit_exit_code() {
    let Some(pptx) = fixture("test_template.pptx") else { return };
    // one hit + one miss → still exit 0, both sections printed
    oa().arg("find").arg(&pptx).args(["-t", "Market Report", "-t", "zzz_not_here_zzz"])
        .assert()
        .success()
        .stdout(predicate::str::contains("no hits for \"zzz_not_here_zzz\""))
        .stdout(predicate::str::contains("hits total"));
    // nothing at all → exit 1
    oa().arg("find").arg(&pptx).args(["-t", "zzz_not_here_zzz"]).assert().code(1);
}

#[test]
fn test_find_missing_file_exits_2() {
    oa().args(["find", "nonexistent.pptx", "-t", "x"]).assert().code(2);
}

#[test]
fn test_find_requires_text() {
    oa().args(["find", "whatever.pptx"]).assert().failure();
}

#[test]
fn test_help_exits_zero() {
    oa().arg("--help")
        .assert()
        .success()
        .stdout(predicate::str::contains("update"))
        .stdout(predicate::str::contains("find"))
        .stdout(predicate::str::contains("info"))
        .stdout(predicate::str::contains("check"))
        .stdout(predicate::str::contains("diff"))
        .stdout(predicate::str::contains("clean"))
        .stdout(predicate::str::contains("config"))
        .stdout(predicate::str::contains("run"));
}

#[test]
fn test_version() {
    oa().arg("--version")
        .assert()
        .success()
        .stdout(predicate::str::contains("oa"));
}

#[test]
fn test_update_help() {
    oa().args(["update", "--help"])
        .assert()
        .success()
        .stdout(predicate::str::contains("--steps"))
        .stdout(predicate::str::contains("--skip"))
        .stdout(predicate::str::contains("--dry-run"))
        .stdout(predicate::str::contains("--check"))
        .stdout(predicate::str::contains("--pair"))
        .stdout(predicate::str::contains("--excel"))
        .stdout(predicate::str::contains("--pick"))
        .stdout(predicate::str::contains("--quiet"))
        .stdout(predicate::str::contains("PIPELINE STEPS"));
}

#[test]
fn test_config_shows_all_keys() {
    oa().arg("config")
        .assert()
        .success()
        .stdout(predicate::str::contains("heatmap.color_minimum"))
        .stdout(predicate::str::contains("ccst.positive_color"))
        .stdout(predicate::str::contains("delta.template_positive"))
        .stdout(predicate::str::contains("links.set_manual"))
        .stdout(predicate::str::contains("#F8696B"))
        .stdout(predicate::str::contains("#33CC33"));
}

#[test]
fn test_unknown_command() {
    oa().arg("foobar")
        .assert()
        .failure()
        .stderr(predicate::str::contains("unrecognized subcommand"));
}

#[test]
fn test_update_no_files() {
    oa().arg("update")
        .assert()
        .failure();
    // Error message varies: "No files to process" locally, COM error on CI without Office
}

#[test]
fn test_update_nonexistent_file() {
    oa().args(["update", "nonexistent.pptx", "-e", "data.xlsx"])
        .assert()
        .failure();
}

#[test]
fn test_info_nonexistent_file() {
    oa().args(["info", "nonexistent.pptx"])
        .assert()
        .failure();
}

/// `oa clean -f` force-kills every PowerPoint/Excel process on the machine, so this
/// must never run in the default suite. Run with:
/// `OA_INTEGRATION=1 cargo test --test cli_e2e -- --ignored`
#[test]
#[ignore]
fn test_clean_no_processes() {
    if std::env::var("OA_INTEGRATION").is_err() {
        eprintln!("Skipping: OA_INTEGRATION not set");
        return;
    }
    oa().args(["clean", "-f"])
        .assert()
        .success()
        .stdout(predicate::str::contains("No Office processes found"));
}

#[test]
fn test_check_help() {
    oa().args(["check", "--help"])
        .assert()
        .success()
        .stdout(predicate::str::contains("--excel"))
        .stdout(predicate::str::contains("--set"));
}

#[test]
fn test_diff_help() {
    oa().args(["diff", "--help"])
        .assert()
        .success()
        .stdout(predicate::str::contains("A.pptx"))
        .stdout(predicate::str::contains("B.pptx"));
}

#[test]
fn test_run_help() {
    oa().args(["run", "--help"])
        .assert()
        .success()
        .stdout(predicate::str::contains("RUNFILE"))
        .stdout(predicate::str::contains("--check"))
        .stdout(predicate::str::contains("--dry-run"));
}
