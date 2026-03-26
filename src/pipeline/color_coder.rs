//! Step 4: Sign-based color coding for _ccst shapes.
//!
//! Finds all table shapes with `_ccst` in their name and colors cells based
//! on their numeric value: positive → green, negative → red, zero/text → grey.
//!
//! Improvements over Python:
//! - Single cell write per iteration (Python wrote twice: prefix + symbol removal)
//! - Cache cell references to avoid repeated COM lookups

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::shapes::inventory::SlideInventory;
use crate::utils::color::hex_to_bgr;

/// Parse a cell text value to extract its numeric value.
///
/// Handles percentages, plus/minus prefixes. Returns `None` for non-numeric text.
pub fn parse_numeric(text: &str) -> Option<f64> {
    let mut s = text.trim().to_string();

    // Strip trailing %
    if s.ends_with('%') {
        s.pop();
        s = s.trim().to_string();
    }

    // Strip leading + (but keep -)
    if s.starts_with('+') {
        s.remove(0);
        s = s.trim().to_string();
    }

    s.parse::<f64>().ok()
}

/// Determine the sign category of a numeric value.
pub fn sign_category(value: f64) -> SignCategory {
    if value > 0.0 {
        SignCategory::Positive
    } else if value < 0.0 {
        SignCategory::Negative
    } else {
        SignCategory::Neutral
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SignCategory {
    Positive,
    Negative,
    Neutral,
}

/// Apply color coding to all _ccst tables in the presentation.
///
/// Returns the count of tables processed.
pub fn apply_color_coding(
    inventory: &SlideInventory,
    config: &Config,
) -> OaResult<usize> {
    if inventory.ccst_tables.is_empty() {
        return Ok(0);
    }

    let positive_color = hex_to_bgr(&config.ccst.positive_color);
    let negative_color = hex_to_bgr(&config.ccst.negative_color);
    let neutral_color = hex_to_bgr(&config.ccst.neutral_color);
    let positive_prefix = &config.ccst.positive_prefix;
    let symbol_removal = &config.ccst.symbol_removal;

    let mut total_tables = 0;

    for shape_ref in &inventory.ccst_tables {
        let mut shape = shape_ref.dispatch.clone();

        let mut table = match shape.get("Table") {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Dispatch::new(d),
                Err(_) => continue,
            },
            Err(_) => continue,
        };

        let rows = table.get("Rows")
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).get("Count"))
            .and_then(|v| v.as_i32())
            .unwrap_or(0);

        let cols = table.get("Columns")
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).get("Count"))
            .and_then(|v| v.as_i32())
            .unwrap_or(0);

        for row in 1..=rows {
            for col in 1..=cols {
                // Get cell dispatch
                let cell_variant = match table.call("Cell", &[Variant::from(row), Variant::from(col)]) {
                    Ok(v) => v,
                    Err(_) => continue,
                };
                let mut cell = match cell_variant.as_dispatch() {
                    Ok(d) => Dispatch::new(d),
                    Err(_) => continue,
                };

                // Navigate to text and font
                let mut cell_shape = match cell.get("Shape") {
                    Ok(v) => match v.as_dispatch() {
                        Ok(d) => Dispatch::new(d),
                        Err(_) => continue,
                    },
                    Err(_) => continue,
                };

                let cell_text = cell_shape.nav("TextFrame.TextRange")
                    .and_then(|mut tr| tr.get("Text"))
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                let parsed = parse_numeric(&cell_text);

                // Determine color and final text in ONE pass (fix Python's double-write bug)
                let (color, final_text) = match parsed {
                    Some(value) => {
                        let category = sign_category(value);
                        let color = match category {
                            SignCategory::Positive => positive_color,
                            SignCategory::Negative => negative_color,
                            SignCategory::Neutral => neutral_color,
                        };

                        // Build final text: apply prefix + symbol removal in one pass
                        let mut text = cell_text.clone();

                        // Add positive prefix if needed
                        if category == SignCategory::Positive
                            && !positive_prefix.is_empty()
                            && !text.starts_with(positive_prefix.as_str())
                        {
                            text = format!("{positive_prefix}{text}");
                        }

                        // Symbol removal (applied AFTER prefix)
                        if !symbol_removal.is_empty() {
                            if symbol_removal.contains('%') && text.ends_with('%') {
                                text.pop(); // Remove trailing %
                            }
                            if symbol_removal.contains('+') && text.starts_with('+') {
                                text.remove(0);
                            }
                            if symbol_removal.contains('-') && text.starts_with('-') {
                                text.remove(0);
                            }
                        }

                        (color, text)
                    }
                    None => {
                        // Non-numeric: neutral color, no text change
                        (neutral_color, cell_text.clone())
                    }
                };

                // Single write for text (fix Python's double-write)
                if final_text != cell_text {
                    let _ = cell_shape.nav("TextFrame.TextRange")
                        .and_then(|mut tr| tr.put("Text", Variant::from(final_text.as_str())));
                }

                // Set font color
                let _ = cell_shape.nav("TextFrame.TextRange.Font.Color")
                    .and_then(|mut fc| fc.put("RGB", Variant::from(color)));
            }
        }

        total_tables += 1;
        super::verbose::detail(
            shape_ref.slide_index,
            &shape_ref.name,
            &format!("{rows}×{cols} cells"),
        );
    }

    Ok(total_tables)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_numeric_positive() {
        assert_eq!(parse_numeric("1.5"), Some(1.5));
    }

    #[test]
    fn test_parse_numeric_negative() {
        assert_eq!(parse_numeric("-0.3"), Some(-0.3));
    }

    #[test]
    fn test_parse_numeric_zero() {
        assert_eq!(parse_numeric("0"), Some(0.0));
    }

    #[test]
    fn test_parse_numeric_with_percent() {
        assert_eq!(parse_numeric("1.5%"), Some(1.5));
    }

    #[test]
    fn test_parse_numeric_with_plus() {
        assert_eq!(parse_numeric("+1.5"), Some(1.5));
    }

    #[test]
    fn test_parse_numeric_with_plus_percent() {
        assert_eq!(parse_numeric("+1.5%"), Some(1.5));
    }

    #[test]
    fn test_parse_numeric_text() {
        assert_eq!(parse_numeric("N/A"), None);
    }

    #[test]
    fn test_parse_numeric_empty() {
        assert_eq!(parse_numeric(""), None);
    }

    #[test]
    fn test_parse_numeric_whitespace() {
        assert_eq!(parse_numeric("  1.5  "), Some(1.5));
    }

    #[test]
    fn test_sign_category() {
        assert_eq!(sign_category(1.5), SignCategory::Positive);
        assert_eq!(sign_category(-0.3), SignCategory::Negative);
        assert_eq!(sign_category(0.0), SignCategory::Neutral);
        assert_eq!(sign_category(-0.0), SignCategory::Neutral);
    }
}
