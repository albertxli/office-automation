# CLI Output Design Spec V3 (Adaptive Colors)

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

| Role              | ANSI style            | Tokyo Night equivalent | Usage                                      |
|-------------------|----------------------|------------------------|--------------------------------------------|
| Target file       | `.cyan()`            | #7dcfff                | Primary file name (pptx)                   |
| Source file       | `.yellow()`          | #ff9e64                | Source file name (xlsx data)                |
| Success indicator | `.green()`           | #9ece6a                | Dot •, checkmark ✓, PASS, total time       |
| Failure indicator | `.red()`             | #ff79c6                | Cross ✗, FAIL, mismatch counts when > 0    |
| Counts / emphasis | `.white().bold()`    | #c0caf5                | Numeric counts                             |
| Secondary info    | `.dim()`             | #565f89                | Dot leaders, per-step timing, labels        |
| Divider line      | `.dim()`             | #3b4261                | Thin dash divider only                     |

### Why ANSI 16 over RGB

- Dracula "cyan" is readable on Dracula's purple background
- Tokyo Night "cyan" is readable on Tokyo Night's dark blue background
- Solarized Light "cyan" is readable on Solarized's light background
- Hardcoded `#7dcfff` would be invisible on a white terminal

The user's theme guarantees contrast. We just declare *semantic intent* (cyan = primary, green = success, red = failure, dim = secondary) and the theme handles the rest.

---

## `oa update` — Output structure

The update output has 3 sections: file header, step table, and completion summary.

```
  ▸ {target_file.pptx}                    ← cyan
    ← {source_file.xlsx}                  ← yellow
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌  ← dim (tight to header)
                                             ← blank line below divider
  • {Step} ····················· {count}   {time}
  • {Step} ····················· {count}   {time}
  ...
                                             ← blank line
  ✓ completed · {total} objects · {total_time}
```

### Section 1: File header
- Line 1: 2-space indent + `▸` + space + target filename in **cyan**
- Line 2: 4-space indent + `←` + space + source filename in **yellow**
- Divider line immediately below (no blank line between header and divider)

### Divider
- 2-space indent + thin dashes `╌` repeated to fill width (~39 chars)
- Style: `.dim()` — fades into background, does not compete with content
- No blank line above, **one blank line below**

### Section 2: Step rows
- 2-space indent + green `•` + space + step name (default terminal color)
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

### Real-time behavior (oa update)

During processing, each step row should:
1. Start with a **spinner animation** (using `indicatif`) instead of the green `•`
2. Show the step name and dot leaders while processing
3. On completion, replace the spinner with green `•` and append count + time
4. Move to the next step

After all steps complete, print the blank line and completion summary.

### Example: `oa update` output

```
  ▸ rpm_2024_market_report_template.pptx
    ← rpm_tracking_Japan_(05_07).xlsx
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌

  • Links ···················  86   0.0s
  • Tables ··················  80   1.0s
  • Deltas ··················   6   0.3s
  • Coloring ················  24   0.2s
  • Charts ··················  100   0.9s

  ✓ completed · 296 objects · 9.9s
```

---

## `oa check` — Output structure

The check output has 3 sections: file header, check results table, and summary.

Each row has a **status badge** (`PASS` or `FAIL`) at the end. The status and all associated indicators change color based on whether mismatches were found.

```
  ▸ {target_file.pptx}                                              ← cyan
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌  ← dim
                                                                       ← blank line
  ✓ {Step} ·········· {count} checked                · {n} mismatches  PASS
  ✓ {Step} ·········· {count} checked                · {n} mismatches  PASS
  ✓ {Step} ·········· {count} checked ({n} series)   · {n} mismatches  PASS
  ...
                                                                       ← blank line
  ✓ check passed · {total} checked · {n} mismatches · {time}
```

### Section 1: File header (check)
- Line 1: 2-space indent + `▸` + space + target filename in **cyan**
- No source file line (check only reads the pptx, no xlsx needed)
- Divider line immediately below

### Section 2: Check result rows

Each row has 5 aligned columns:

| Column            | Alignment  | Style                          | Example                  |
|-------------------|------------|--------------------------------|--------------------------|
| Status icon       | fixed      | green `✓` or red `✗`           | `✓` / `✗`                |
| Step name         | left       | default terminal color          | `Tables`                 |
| Dot leaders       | fill       | `.dim()`                       | `··········`             |
| Count + detail    | fixed-width| **white bold** count + dim text | `147 checked (234 series)` |
| Mismatches        | right      | see rules below                | `0 mismatches`           |
| Status badge      | right      | green `PASS` or red `FAIL`     | `PASS` / `FAIL`          |

**Mismatch coloring rules:**
- When mismatches = 0: mismatch count in **white bold**, `mismatches` in dim
- When mismatches > 0: entire mismatch value in **red** (e.g. `2`)

**Status badge rules:**
- When mismatches = 0: `PASS` in **green**
- When mismatches > 0: `FAIL` in **red**

**Row icon rules:**
- When mismatches = 0: green `✓`
- When mismatches > 0: red `✗`

**Count + detail column (fixed-width for alignment):**
- All rows use `checked` as the primary label
- Charts appends `({n} series)` in dim after `checked`, with the series count in **white bold**
- Tables and Deltas pad with whitespace to the same width as the Charts row
- This ensures `· {n} mismatches  PASS` starts at the exact same horizontal position on every row
- The fixed width of this column should accommodate the longest possible content: `{3-digit count} checked ({3-digit count} series)`

### Section 3: Check summary

- One blank line above
- When all pass: green `✓` + `check passed` in **green**
- When any fail: red `✗` + `check failed` in **red**
- ` · ` separator in dim
- Total count in **white bold** + `checked` in dim
- ` · ` separator in dim
- Mismatch total: green `0 mismatches` when passing, red `{n} mismatches` when failing
- ` · ` separator in dim
- Total time in **green** (always green regardless of pass/fail)

### Example: `oa check` — all passing

```
  ▸ rpm_2024_market_report_template.pptx
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌

  ✓ Tables ·········· 147 checked                ·  0 mismatches  PASS
  ✓ Deltas ··········   6 checked                ·  0 mismatches  PASS
  ✓ Charts ·········· 100 checked (234 series)   ·  0 mismatches  PASS

  ✓ check passed · 387 checked · 0 mismatches · 3.5s
```

Color breakdown for all-passing:
- `✓` = green
- Step names = default
- Dot leaders = dim
- `147`, `6`, `100`, `234` = white bold
- `checked`, `(`, `series)` = dim
- `·` separators = dim
- `0` = white bold
- `mismatches` = dim
- `PASS` = green
- `check passed` = green
- `387` = white bold
- `0 mismatches` = green (summary line)
- `3.5s` = green

### Example: `oa check` — with failures

```
  ▸ rpm_2024_market_report_template.pptx
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌

  ✓ Tables ·········· 147 checked                ·  0 mismatches  PASS
  ✗ Deltas ··········   6 checked                ·  2 mismatches  FAIL
  ✓ Charts ·········· 100 checked (234 series)   ·  0 mismatches  PASS

  ✗ check failed · 387 checked · 2 mismatches · 3.5s
```

Color breakdown for failure:
- Passing rows are identical to the all-passing example
- Failing row: `✗` = red, `2` = red, `FAIL` = red
- Summary: `✗ check failed` = red, `2 mismatches` = red
- `3.5s` = green (always green — timing is neutral)

---

## Unicode characters used

- `▸` (U+25B8) — file indicator
- `←` (U+2190) — source arrow (oa update only)
- `╌` (U+254C) — box drawings light double dash horizontal (divider)
- `•` (U+2022) — bullet (completed update step)
- `·` (U+00B7) — middle dot (leaders and separators)
- `✓` (U+2713) — check mark (success)
- `✗` (U+2717) — ballot x (failure)

---

## Implementation reference

```rust
use console::Style;
use indicatif::{ProgressBar, ProgressStyle};
use std::time::Duration;

// ── Styles (ANSI 16, theme-adaptive) ──────────────────────
let s_target  = Style::new().cyan();           // target pptx file
let s_source  = Style::new().yellow();         // source xlsx file
let s_ok      = Style::new().green();          // •, ✓, PASS, total time
let s_fail    = Style::new().red();            // ✗, FAIL, mismatch counts
let s_count   = Style::new().white().bold();   // numeric counts
let s_dim     = Style::new().dim();            // leaders, timing, divider

// ── oa update: File header ────────────────────────────────
println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&target_file));
println!("    {} {}", s_source.apply_to("←"), s_source.apply_to(&source_file));
println!("  {}", s_dim.apply_to("╌".repeat(39)));
println!();

// ── oa update: Steps with spinner ─────────────────────────
let spinner = ProgressBar::new_spinner();
spinner.enable_steady_tick(Duration::from_millis(80));
spinner.set_style(
    ProgressStyle::default_spinner()
        .tick_strings(&["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"])
        .template("  {spinner} {msg}")
        .unwrap()
);
spinner.set_message("Links ···················");

// After step completes
spinner.finish_and_clear();
println!(
    "  {} Links {} {:>3}   {}",
    s_ok.apply_to("•"),
    s_dim.apply_to("···················"),
    s_count.apply_to(86),
    s_dim.apply_to("0.0s")
);

// ── oa update: Completion ─────────────────────────────────
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

// ── oa check: File header ─────────────────────────────────
println!("  {} {}", s_target.apply_to("▸"), s_target.apply_to(&target_file));
println!("  {}", s_dim.apply_to("╌".repeat(64)));
println!();

// ── oa check: Result row (passing, no series detail) ─────
// The detail column is fixed-width: pad to match the longest variant "(234 series)"
println!(
    "  {} Tables {} {:>3} {}  {}  {} {}  {}",
    s_ok.apply_to("✓"),
    s_dim.apply_to("··········"),
    s_count.apply_to(147),
    s_dim.apply_to("checked              "),  // padded to match "(234 series)"
    s_dim.apply_to("·"),
    s_count.apply_to(0),
    s_dim.apply_to("mismatches"),
    s_ok.apply_to("PASS")
);

// ── oa check: Result row (passing, with series detail) ───
println!(
    "  {} Charts {} {:>3} {}  {}  {} {}  {}",
    s_ok.apply_to("✓"),
    s_dim.apply_to("··········"),
    s_count.apply_to(100),
    format!(
        "{} ({} {})",
        s_dim.apply_to("checked"),
        s_count.apply_to(234),
        s_dim.apply_to("series)  ")
    ),
    s_dim.apply_to("·"),
    s_count.apply_to(0),
    s_dim.apply_to("mismatches"),
    s_ok.apply_to("PASS")
);

// ── oa check: Result row (failing) ────────────────────────
println!(
    "  {} Deltas {} {:>3} {}  {}  {} {}  {}",
    s_fail.apply_to("✗"),
    s_dim.apply_to("··········"),
    s_count.apply_to(6),
    s_dim.apply_to("checked              "),  // padded to match "(234 series)"
    s_dim.apply_to("·"),
    s_fail.apply_to(2),
    s_dim.apply_to("mismatches"),
    s_fail.apply_to("FAIL")
);

// ── oa check: Summary (all passing) ──────────────────────
println!();
println!(
    "  {} {} {} {} {} {} {} {}",
    s_ok.apply_to("✓ check passed"),
    s_dim.apply_to("·"),
    s_count.apply_to(387),
    s_dim.apply_to("checked"),
    s_dim.apply_to("·"),
    s_ok.apply_to("0 mismatches"),
    s_dim.apply_to("·"),
    s_ok.apply_to("3.5s")
);

// ── oa check: Summary (with failures) ────────────────────
println!();
println!(
    "  {} {} {} {} {} {} {} {}",
    s_fail.apply_to("✗ check failed"),
    s_dim.apply_to("·"),
    s_count.apply_to(387),
    s_dim.apply_to("checked"),
    s_dim.apply_to("·"),
    s_fail.apply_to("2 mismatches"),
    s_dim.apply_to("·"),
    s_ok.apply_to("3.5s")  // time is always green
);
```

## Column alignment helper

To ensure counts and times align vertically, pad step names + dot leaders to a fixed width.

```rust
// oa update: step row
fn format_update_row(name: &str, count: usize, time_secs: f64) -> String {
    let leader_len = 30 - name.len();
    let leaders = "·".repeat(leader_len);
    format!(
        "  {} {} {} {:>3}   {}",
        green("•"),
        name,
        dim(&leaders),
        white_bold(count),
        dim(&format!("{:.1}s", time_secs))
    )
}

// oa check: result row with fixed-width detail column
fn format_check_row(
    name: &str,
    count: usize,
    series: Option<usize>,  // Some(234) for Charts, None for others
    mismatches: usize,
) -> String {
    let leader_len = 20 - name.len();
    let leaders = "·".repeat(leader_len);

    // Fixed-width detail column: "checked (234 series)" or "checked" + padding
    let detail = match series {
        Some(n) => format!("checked ({} series)", n),   // e.g. "checked (234 series)"
        None    => format!("checked              "),     // padded to same width
    };

    let (icon, badge, mm_style) = if mismatches == 0 {
        (green("✓"), green("PASS"), white_bold(0))
    } else {
        (red("✗"), red("FAIL"), red(mismatches))
    };

    format!(
        "  {} {} {} {:>3} {}  {}  {} {}  {}",
        icon, name, dim(&leaders),
        white_bold(count), dim(&detail),
        dim("·"),
        mm_style, dim("mismatches"),
        badge
    )
}
```

## Notes

- Never use hardcoded RGB/hex colors — ANSI 16 only
- `console` crate handles Windows terminal compatibility automatically
- `.dim()` is the workhorse — use it for anything that should recede
- `.bold()` on counts gives just enough emphasis without shouting
- `.red()` is only used for failure states — never for decoration
- Time is always green regardless of pass/fail (timing is neutral info)
- Test output on both dark and light terminals if possible
