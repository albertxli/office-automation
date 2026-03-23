//! Verbose logging for pipeline steps.
//!
//! Prints dim, indented detail lines when verbose mode is active.
//! Zero cost when verbose=false (just a bool check).

use std::sync::atomic::{AtomicBool, Ordering};

use console::Style;

/// Global verbose flag — set once at pipeline start.
static VERBOSE: AtomicBool = AtomicBool::new(false);

/// Enable verbose output.
pub fn set_verbose(on: bool) {
    VERBOSE.store(on, Ordering::Relaxed);
}

/// Check if verbose is enabled.
pub fn is_verbose() -> bool {
    VERBOSE.load(Ordering::Relaxed)
}

/// Print a verbose detail line if verbose mode is active.
///
/// Format: `      Slide {n} │ {shape_name:<24} {detail}`
/// All in dim style, 6-space indent.
pub fn detail(slide: i32, shape: &str, info: &str) {
    if !is_verbose() {
        return;
    }
    let s = Style::new().dim();
    println!("      {} {:>2} {} {:<24} {}",
        s.apply_to("Slide"),
        s.apply_to(slide),
        s.apply_to("│"),
        s.apply_to(shape),
        s.apply_to(info));
}

/// Print a verbose check detail line with type column and colored checkmark.
///
/// Format: `      Slide  3 │ table │ ntbl_Object_edu          ✓ (1,1) '74%'`
/// Checkmark is green for pass, red for fail. Values in default white.
pub fn check_detail(slide: i32, check_type: &str, shape: &str, passed: bool, info: &str) {
    if !is_verbose() {
        return;
    }
    let s = Style::new().dim();
    let mark = if passed {
        Style::new().green().apply_to("✓")
    } else {
        Style::new().red().apply_to("✗")
    };
    println!("      {} {:>2} {} {:<5} {} {:<24} {} {}",
        s.apply_to("Slide"),
        s.apply_to(slide),
        s.apply_to("│"),
        s.apply_to(check_type),
        s.apply_to("│"),
        s.apply_to(shape),
        mark,
        info);
}

/// Print a verbose line without slide context.
pub fn note(msg: &str) {
    if !is_verbose() {
        return;
    }
    let s = Style::new().dim();
    println!("      {}", s.apply_to(msg));
}
