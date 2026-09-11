//! `oa find` — search text inside a PPTX at ZIP level (read-only, no PowerPoint).
//!
//! Lists every occurrence with its location (slide / notes / layout / master), shape name
//! and a snippet. Exit codes follow grep: 0 found, 1 nothing found, 2 error. `-t` is the
//! first selector; charts or shapes by name may join it later.

use std::path::Path;
use std::time::Instant;

use console::Style;

use crate::error::{OaError, OaResult};
use crate::zip_ops::slide_text::extract_text_blocks;

/// Characters of context kept on each side of a match in the snippet.
const CONTEXT: usize = 30;

/// Run `oa find`. Returns the total number of hits across all needles.
pub fn run_find(file: &str, needles: &[String], ignore_case: bool) -> OaResult<usize> {
    let path = Path::new(file);
    if !path.exists() {
        return Err(OaError::Other(format!("File not found: {file}")));
    }
    if let Some(empty) = needles.iter().find(|n| n.is_empty()) {
        return Err(OaError::Config(format!("Empty search text {empty:?} (-t needs a value)")));
    }

    let t = Instant::now();
    let (blocks, stats) = extract_text_blocks(path).map_err(OaError::Other)?;

    let s_cyan = Style::new().cyan();
    let s_dim = Style::new().dim();
    let s_ok = Style::new().green();
    let s_none = Style::new().yellow();
    let s_bold = Style::new().bold();

    let file_name = path.file_name().unwrap_or_default().to_string_lossy();
    println!();
    println!("  {} {}", s_cyan.apply_to("▸"), s_cyan.apply_to(&*file_name));
    println!("  {}", s_dim.apply_to("╌".repeat(61)));
    println!();

    let stats_line = format!(
        "{} slides · {} layouts · {} master{} · {} notes · {:.2}s",
        stats.slides, stats.layouts, stats.masters,
        if stats.masters == 1 { "" } else { "s" },
        stats.notes, t.elapsed().as_secs_f64(),
    );

    let mut total = 0usize;
    for (ni, needle) in needles.iter().enumerate() {
        if needles.len() > 1 {
            if ni > 0 {
                println!();
            }
            println!("  {}", s_bold.apply_to(format!("\"{needle}\"")));
        }

        // Collect first so the location column is only as wide as the labels actually shown
        let rows: Vec<(String, &str, (String, String, String))> = blocks.iter()
            .flat_map(|block| {
                find_matches(&block.text, needle, ignore_case).into_iter().map(move |(start, len)| {
                    (block.location.to_string(), block.shape.as_str(), snippet_parts(&block.text, start, len))
                })
            })
            .collect();
        let loc_width = rows.iter().map(|(loc, _, _)| loc.chars().count()).max().unwrap_or(8).max(8);
        for (loc, shape, (before, matched, after)) in &rows {
            println!("  {} {} {} {}{}{}",
                s_dim.apply_to(format!("{loc:<loc_width$}")),
                s_dim.apply_to("│"),
                s_dim.apply_to(format!("{shape:<24}")),
                before,
                Style::new().yellow().bold().apply_to(matched),
                after);
        }
        let hits = rows.len();
        total += hits;

        let summary = if hits == 0 {
            format!("{} {}", s_none.apply_to("○"), s_none.apply_to(format!("no hits for \"{needle}\"")))
        } else {
            format!("{} {} hit{} for \"{needle}\"", s_ok.apply_to("✓"), s_ok.apply_to(hits), if hits == 1 { "" } else { "s" })
        };
        if needles.len() == 1 {
            println!("  {summary} {} {}", s_dim.apply_to("·"), s_dim.apply_to(&stats_line));
        } else {
            println!("  {summary}");
        }
    }

    if needles.len() > 1 {
        println!("  {}", s_dim.apply_to("╌".repeat(61)));
        let mark = if total > 0 { s_ok.apply_to("✓") } else { s_none.apply_to("○") };
        println!("  {mark} {} hits total {} {}", total, s_dim.apply_to("·"), s_dim.apply_to(&stats_line));
    }

    Ok(total)
}

/// Non-overlapping matches of `needle` in `text` as `(char index, char length)`.
/// Case-insensitive mode folds each char independently so indices still map to `text`.
fn find_matches(text: &str, needle: &str, ignore_case: bool) -> Vec<(usize, usize)> {
    let fold = |c: char| if ignore_case { c.to_lowercase().next().unwrap_or(c) } else { c };
    let hay: Vec<char> = text.chars().map(fold).collect();
    let nd: Vec<char> = needle.chars().map(fold).collect();
    let mut out = Vec::new();
    if nd.is_empty() || nd.len() > hay.len() {
        return out;
    }
    let mut i = 0;
    while i + nd.len() <= hay.len() {
        if hay[i..i + nd.len()] == nd[..] {
            out.push((i, nd.len()));
            i += nd.len();
        } else {
            i += 1;
        }
    }
    out
}

/// `(before, matched, after)` around a match, each side trimmed to `CONTEXT` chars with `…`
/// marking a cut; line breaks are shown as `⏎`.
fn snippet_parts(text: &str, start: usize, len: usize) -> (String, String, String) {
    let chars: Vec<char> = text.chars().collect();
    let end = (start + len).min(chars.len());
    let show = |cs: &[char]| -> String {
        cs.iter().map(|&c| if c == '\n' || c == '\r' || c == '\u{b}' { '⏎' } else { c }).collect()
    };

    let before_start = start.saturating_sub(CONTEXT);
    let mut before = show(&chars[before_start..start]);
    if before_start > 0 {
        before.insert(0, '…');
    }

    let matched = show(&chars[start..end]);

    let after_end = (end + CONTEXT).min(chars.len());
    let mut after = show(&chars[end..after_end]);
    if after_end < chars.len() {
        after.push('…');
    }

    (before, matched, after)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_find_matches_literal() {
        assert_eq!(find_matches("Market Report: [country]", "[country]", false), vec![(15, 9)]);
        assert_eq!(find_matches("[country] B and [country] C", "[country]", false), vec![(0, 9), (16, 9)]);
        assert!(find_matches("Market Report", "market", false).is_empty(), "case-sensitive by default");
        assert!(find_matches("short", "much longer needle", false).is_empty());
        assert!(find_matches("anything", "", false).is_empty());
    }

    #[test]
    fn test_find_matches_ignore_case_and_unicode() {
        assert_eq!(find_matches("Market Report: Japan", "japan", true), vec![(15, 5)]);
        assert_eq!(find_matches("Étude Étude", "étude", true), vec![(0, 5), (6, 5)]);
        // indices are char-based, so multi-byte text before the match does not shift them
        assert_eq!(find_matches("日本語 [country]", "[country]", false), vec![(4, 9)]);
    }

    #[test]
    fn test_find_matches_no_overlap() {
        assert_eq!(find_matches("aaaa", "aa", false), vec![(0, 2), (2, 2)]);
    }

    #[test]
    fn test_snippet_parts_trims_both_sides() {
        let text = format!("{}[country]{}", "x".repeat(50), "y".repeat(50));
        let (before, matched, after) = snippet_parts(&text, 50, 9);
        assert_eq!(before, format!("…{}", "x".repeat(CONTEXT)));
        assert_eq!(matched, "[country]");
        assert_eq!(after, format!("{}…", "y".repeat(CONTEXT)));
    }

    #[test]
    fn test_snippet_parts_short_text_and_line_breaks() {
        let (before, matched, after) = snippet_parts("Key\nfindings: [country]", 14, 9);
        assert_eq!(before, "Key⏎findings: ");
        assert_eq!(matched, "[country]");
        assert_eq!(after, "");
    }
}
