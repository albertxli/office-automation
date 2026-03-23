//! `oa clean` — kill zombie PowerPoint and Excel processes.

use std::io::{self, Write};
use sysinfo::System;

use crate::error::OaResult;

const OFFICE_PROCESSES: &[&str] = &["POWERPNT.EXE", "EXCEL.EXE"];

/// Run the `oa clean` command.
pub fn run_clean(force: bool) -> OaResult<()> {
    let mut sys = System::new();
    sys.refresh_processes(sysinfo::ProcessesToUpdate::All, true);

    let mut found: Vec<(sysinfo::Pid, String)> = Vec::new();

    for process in sys.processes().values() {
        let name = process.name().to_string_lossy().to_string();
        if OFFICE_PROCESSES.iter().any(|&p| name.eq_ignore_ascii_case(p)) {
            found.push((process.pid(), name));
        }
    }

    if found.is_empty() {
        println!("No zombie Office processes found.");
        return Ok(());
    }

    println!("Found {} Office process(es):", found.len());
    for (pid, name) in &found {
        println!("  {} (PID {})", name, pid);
    }

    if !force {
        print!("\nKill all? [y/N] ");
        io::stdout().flush()?;
        let mut input = String::new();
        io::stdin().read_line(&mut input)?;
        if !input.trim().eq_ignore_ascii_case("y") {
            println!("Cancelled.");
            return Ok(());
        }
    }

    let mut killed = 0;
    let mut sys2 = System::new();
    sys2.refresh_processes(sysinfo::ProcessesToUpdate::All, true);
    for (pid, name) in &found {
        if let Some(process) = sys2.process(*pid) {
            process.kill();
            killed += 1;
            println!("  Killed {} (PID {})", name, pid);
        }
    }

    println!("Killed {} process(es).", killed);
    Ok(())
}
