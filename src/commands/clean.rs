//! `oa clean` — kill zombie PowerPoint and Excel processes.

use std::io::{self, Write};
use sysinfo::System;

use console::Style;

use crate::error::OaResult;

const OFFICE_PROCESSES: &[&str] = &["POWERPNT.EXE", "EXCEL.EXE"];

/// Fixed column for PID alignment (shared by discovery + kill rows).
const PID_COL: usize = 40;

/// Run the `oa clean` command.
pub fn run_clean(force: bool) -> OaResult<()> {
    let s_dim = Style::new().dim();
    let s_ok = Style::new().green();
    let s_count = Style::new().white().bold();
    let s_name = Style::new().yellow();

    let mut sys = System::new();
    sys.refresh_processes(sysinfo::ProcessesToUpdate::All, true);

    let mut found: Vec<(sysinfo::Pid, String)> = Vec::new();

    for process in sys.processes().values() {
        let name = process.name().to_string_lossy().to_string();
        if OFFICE_PROCESSES.iter().any(|&p| name.eq_ignore_ascii_case(p)) {
            found.push((process.pid(), name));
        }
    }

    // Discovery phase
    let count_label = if found.len() == 1 { "Office process" } else { "Office processes" };
    println!();
    println!("  {} {} {}",
        s_dim.apply_to("Found"),
        s_count.apply_to(found.len()),
        s_dim.apply_to(count_label));

    if found.is_empty() {
        println!();
        println!("  {} {}", s_ok.apply_to("✓"), s_ok.apply_to("nothing to clean"));
        return Ok(());
    }

    // List processes
    println!();
    for (pid, name) in &found {
        print_process_row("", name, pid, &s_name, &s_dim, &s_count);
    }

    // Confirmation phase
    if !force {
        println!();
        print!("  {} ", s_name.apply_to("Kill all? [y/N]"));
        io::stdout().flush()?;
        let mut input = String::new();
        io::stdin().read_line(&mut input)?;
        if !input.trim().eq_ignore_ascii_case("y") {
            println!();
            println!("  {} {}",
                s_ok.apply_to("✓"),
                s_dim.apply_to("cancelled — no processes killed"));
            return Ok(());
        }
    }

    // Kill phase
    let mut killed = 0;
    let mut sys2 = System::new();
    sys2.refresh_processes(sysinfo::ProcessesToUpdate::All, true);

    println!();
    for (pid, name) in &found {
        if let Some(process) = sys2.process(*pid) {
            process.kill();
            killed += 1;
            print_process_row(
                &format!("{} ", s_ok.apply_to("✓ Killed")),
                name, pid, &s_name, &s_dim, &s_count,
            );
        }
    }

    // Summary
    let kill_label = if killed == 1 { "process killed" } else { "processes killed" };
    println!();
    println!("  {} {} {} {} {}",
        s_ok.apply_to("✓"),
        s_ok.apply_to("cleaned"),
        s_dim.apply_to("·"),
        s_count.apply_to(killed),
        s_dim.apply_to(kill_label));

    Ok(())
}

/// Print a process row with dot-leader alignment to PID column.
fn print_process_row(
    prefix: &str,
    name: &str,
    pid: &sysinfo::Pid,
    s_name: &Style,
    s_dim: &Style,
    s_count: &Style,
) {
    // Calculate display width: "  " + prefix + name
    let indent = "  ";
    // prefix might contain ANSI codes, so measure visible chars separately
    let prefix_visible_len = if prefix.is_empty() { 0 } else {
        // "✓ Killed " = 9 visible chars
        console::measure_text_width(prefix)
    };
    let display_len = indent.len() + prefix_visible_len + name.len();
    let leader_len = PID_COL.saturating_sub(display_len + 1);
    let leaders = "·".repeat(leader_len);

    println!("{}{}{} {} {} {}",
        indent,
        prefix,
        s_name.apply_to(name),
        s_dim.apply_to(&leaders),
        s_dim.apply_to("PID"),
        s_count.apply_to(pid));
}
