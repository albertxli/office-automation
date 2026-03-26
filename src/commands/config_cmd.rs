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
        println!("{:<30} {}", key, display);
    }
}
