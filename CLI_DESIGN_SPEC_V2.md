# CLI Output Design Spec V2 (Adaptive Colors)

## Overview

This document defines the terminal output style for the `oa` CLI tool. All command output should follow this design language.

This version uses **ANSI 16 standard colors** instead of hardcoded RGB values, so the output adapts to any terminal theme (Dracula, Tokyo Night, Solarized, One Dark, default CMD, etc.). The visual hierarchy is identical to V1 — only the color implementation changes.

## Rust crates to use

```toml
[dependencies]
console = "0.16"
indicatif = "0.18"
```

- **`console`** — terminal colors, styling, and Unicode support
- **`indicatif`** — spinner animations during real-time processing steps

## Color mapping

Use `console::Style` with standard ANSI color methods. These map to the user's terminal theme automatically.

| Role              | ANSI style            | Tokyo Night equivalent | Usage                              |
|-------------------|----------------------|------------------------|------------------------------------|
| Target file       | `.cyan()`            | #7dcfff                | Primary file name (pptx)           |
| Source file       | `.yellow()`          | #ff9e64                | Source file name (xlsx data)        |
| Success indicator | `.green()`           | #9ece6a                | Dot ●, checkmark ✓, total time     |
| Counts / emphasis | `.white().bold()`    | #c0caf5                | Numeric counts                     |
| Secondary info    | `.dim()`             | #565f89                | Dot leaders, per-step timing, labels |
| Divider line      | `.dim()`             | #3b4261                | Thin dash divider only             |

### Why ANSI 16 over RGB

- Dracula "cyan" is readable on Dracula's purple background
- Tokyo Night "cyan" is readable on Tokyo Night's dark blue background
- Solarized Light "cyan" is readable on Solarized's light background
- Hardcoded `#7dcfff` would be invisible on a white terminal

The user's theme guarantees contrast. We just declare *semantic intent* (cyan = primary, green = success, dim = secondary) and the theme handles the rest.

## Output structure

The output has 3 sections: file header, step table, and completion summary.

```
  ▸ {target_file.pptx}                    ← cyan
    ← {source_file.xlsx}                  ← yellow
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌  ← dim (tight to header)
                                             ← blank line below divider
  ● {Step} ····················· {count}   {time}
  ● {Step} ····················· {count}   {time}
  ...
                                             ← blank line
  ✓ completed · {total} objects · {total_time}
```

## Detailed rules

### Section 1: File header
- Line 1: 2-space indent + `▸` + space + target filename in **cyan**
- Line 2: 4-space indent + `←` + space + source filename in **yellow**
- Divider line immediately below (no blank line between header and divider)

### Divider
- 2-space indent + thin dashes `╌` repeated to fill width (~39 chars)
- Style: `.dim()` — fades into background, does not compete with content
- No blank line above, **one blank line below**

### Section 2: Step rows
- 2-space indent + green `●` + space + step name (default terminal color)
- Dot leaders `···` in dim fill the gap between step name and count
- Right-aligned count in **white bold** (3-char wide column)
- 3-space gap
- Right-aligned time in **dim** (format: `{n.n}s`, 4-char wide column)
- **Counts and times must be vertically aligned across all rows**

### Section 3: Completion
- One blank line above
- 2-space indent + green `✓` + space + `completed` in **green**
- ` · ` separator in dim
- Total count in **white bold** + `objects` in dim
- ` · ` separator in dim
- Total time in **green**

## Real-time behavior

During processing, each step row should:
1. Start with a **spinner animation** (using `indicatif`) instead of the green `●`
2. Show the step name and dot leaders while processing
3. On completion, replace the spinner with green `●` and append count + time
4. Move to the next step

After all steps complete, print the blank line and completion summary.

## Unicode characters used

- `▸` (U+25B8) — file indicator
- `←` (U+2190) — source arrow
- `╌` (U+254C) — box drawings light double dash horizontal (divider)
- `●` (U+25CF) — filled circle (completed step)
- `·` (U+00B7) — middle dot (leaders and separators)
- `✓` (U+2713) — check mark (completion)

## Implementation reference

```rust
use console::Style;
use indicatif::{ProgressBar, ProgressStyle};
use std::time::Duration;

// ── Styles (ANSI 16, theme-adaptive) ──────────────────────
let s_target  = Style::new().cyan();           // target pptx file
let s_source  = Style::new().yellow();         // source xlsx file
let s_ok      = Style::new().green();          // ●, ✓, total time
let s_count   = Style::new().white().bold();   // numeric counts
let s_dim     = Style::new().dim();            // leaders, timing, divider

// ── Section 1: File header ────────────────────────────────
println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&target_file));
println!("    {} {}", s_source.apply_to("←"), s_source.apply_to(&source_file));
println!("  {}", s_dim.apply_to("╌".repeat(39)));
println!();

// ── Section 2: Steps with spinner ─────────────────────────
// While processing: show spinner
let spinner = ProgressBar::new_spinner();
spinner.enable_steady_tick(Duration::from_millis(80));
spinner.set_style(
    ProgressStyle::default_spinner()
        .tick_strings(&["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"])
        .template("  {spinner} {msg}")
        .unwrap()
);
spinner.set_message("Links ···················");

// After step completes: replace spinner with final line
spinner.finish_and_clear();
println!(
    "  {} Links {} {:>3}   {}",
    s_ok.apply_to("●"),
    s_dim.apply_to("···················"),
    s_count.apply_to(86),
    s_dim.apply_to("0.0s")
);

// ── Section 3: Completion ─────────────────────────────────
println!();
println!(
    "  {} {} {} {} {} {}",
    s_ok.apply_to("✓ completed"),
    s_dim.apply_to("·"),
    s_count.apply_to(296),
    s_dim.apply_to("objects"),
    s_dim.apply_to("·"),
    s_ok.apply_to("9.9s")
);
```

## Column alignment helper

To ensure counts and times align vertically, pad step names + dot leaders to a fixed width. Example approach:

```rust
fn format_step_row(name: &str, count: usize, time_secs: f64) -> String {
    let leader_len = 30 - name.len(); // adjust 30 to your max step name width
    let leaders = "·".repeat(leader_len);
    format!(
        "  {} {} {} {:>3}   {}",
        green("●"),
        name,
        dim(&leaders),
        white_bold(count),
        dim(&format!("{:.1}s", time_secs))
    )
}
```

## Notes

- Never use hardcoded RGB/hex colors — ANSI 16 only
- `console` crate handles Windows terminal compatibility automatically
- `.dim()` is the workhorse — use it for anything that should recede
- `.bold()` on counts gives just enough emphasis without shouting
- Test output on both dark and light terminals if possible
