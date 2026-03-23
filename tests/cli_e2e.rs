//! End-to-end CLI tests using assert_cmd.
//! These test argument parsing and help output — no Office needed.

use assert_cmd::Command;
use predicates::prelude::*;

fn oa() -> Command {
    Command::cargo_bin("oa").unwrap()
}

#[test]
fn test_help_exits_zero() {
    oa().arg("--help")
        .assert()
        .success()
        .stdout(predicate::str::contains("update"))
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
        .stdout(predicate::str::contains("0.1.0"));
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
        .failure()
        .stderr(predicate::str::contains("No files to process"));
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

#[test]
fn test_clean_no_processes() {
    oa().args(["clean", "-f"])
        .assert()
        .success()
        .stdout(predicate::str::contains("No zombie Office processes found"));
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
