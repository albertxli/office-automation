use console::Style;

use crate::config::Config;

/// Print all config keys and their default values.
pub fn run_config() {
    let config = Config::default();
    println!("{:<30} DEFAULT", "KEY");
    println!("{:<30} -------", "---");
    for (key, value) in config.all_keys() {
        let display = if value.is_empty() {
            "(empty)".to_string()
        } else {
            value
        };
        println!("{key:<30} {display}");
    }

    // Open-ended key family (GOTCHA #45) — cannot be listed exhaustively
    let s_dim = Style::new().dim();
    println!("{:<30} {}", "delta.threshold.<token>",
        s_dim.apply_to("per-OLE-name dead band, e.g. --set delta.threshold.globalnet=0.02"));
}
