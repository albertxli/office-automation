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

/// Print a chart series mismatch continuation line (no slide/chart prefix).
///
/// Format: `           ╰ 'Series1'  3/4 differ   .28→.26  .47→.56  ...`
/// Coloring: ╰=dim, name=yellow, count=dim, PPT=red, →=dim, Excel=white.bold, ...=dim
/// `name_pad` aligns diff counts across series within the same chart.
pub fn check_chart_series_diff(name: &str, name_pad: usize, diff_count: usize, total: usize, pairs: &[(f64, f64)], has_more: bool) {
    if !is_verbose() {
        return;
    }
    let s_dim = Style::new().dim();
    let s_name = Style::new().yellow();
    let s_ppt = Style::new().red();
    let s_arrow = Style::new().dim();
    let s_excel = Style::new().white().bold();

    let mut pair_strs = Vec::new();
    for (ppt, excel) in pairs {
        pair_strs.push(format!("{}{}{}",
            s_ppt.apply_to(fmt_short(*ppt)),
            s_arrow.apply_to("→"),
            s_excel.apply_to(fmt_short(*excel))));
    }
    let values = pair_strs.join("  ");
    let overflow = if has_more { format!("  {}", s_dim.apply_to("...")) } else { String::new() };

    // Build prefix: "╰ 'name'{pad} N/M differ"
    // Pad to 41 chars so values start at col 52 (continuation starts at col 11).
    let padded_name = format!("'{name}'{}", " ".repeat(name_pad.saturating_sub(name.len())));
    let diff_label = format!("{diff_count}/{total} differ");
    let prefix_plain_len = 2 + padded_name.len() + 1 + diff_label.len(); // "╰ " + name + " " + label
    let gap = 41usize.saturating_sub(prefix_plain_len);

    print!("           {} {} {}{}",
        s_dim.apply_to("╰"),
        s_name.apply_to(&padded_name),
        s_dim.apply_to(&diff_label),
        " ".repeat(gap));

    if !pair_strs.is_empty() {
        print!("{values}");
    }
    println!("{overflow}");
}

/// Middle-truncate a name to max 14 display chars: first 7 + … + last 6.
pub fn truncate_middle(name: &str) -> String {
    if name.len() <= 14 {
        return name.to_string();
    }
    format!("{}…{}", &name[..7], &name[name.len() - 6..])
}

/// Format a float compactly: drop leading zero for decimals (0.28 → .28).
fn fmt_short(v: f64) -> String {
    let s = format!("{v:.2}");
    if s.starts_with("0.") {
        s[1..].to_string()  // "0.28" → ".28"
    } else if s.starts_with("-0.") {
        format!("-{}", &s[2..])  // "-0.28" → "-.28"
    } else {
        s
    }
}

/// Print a verbose line without slide context.
pub fn note(msg: &str) {
    if !is_verbose() {
        return;
    }
    let s = Style::new().dim();
    println!("      {}", s.apply_to(msg));
}
