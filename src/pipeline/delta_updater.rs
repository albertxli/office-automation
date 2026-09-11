//! Step 3: Swap delta indicator arrows based on value sign.
//!
//! Uses a two-pass algorithm:
//! 1. Collect all OLE+delt pairs and their metadata (safe for iteration)
//! 2. Process each pair: delete old shape, copy template, reposition
//!
//! Value sign is determined from the PPT table cell (primary) or Excel (fallback).
//! Template shapes on slide 1: tmpl_delta_pos, tmpl_delta_neg, tmpl_delta_none.
//!
//! Thresholds (GOTCHA #45): when `config.delta.threshold_for(ole_name)` is `> 0`, the value
//! is read numerically from Excel (`Range.Value2`, decimal — a `2%` cell is `0.02`) and
//! `|v| < threshold` → "none". Deltas with threshold `0` keep the legacy text path above.
//!
//! Template sets: a `delt<N>_` shape (N ≥ 2) copies from `tmpl<N>_delta_{pos,neg,none}`
//! instead of the set-1 templates. Set number comes from `matcher::delta_set`; template
//! names from `matcher::template_name_for_set`. A set whose templates are missing is
//! skipped with a warning — never silently mapped to set 1.

use std::collections::HashMap;

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::shapes::inventory::SlideInventory;
use crate::shapes::matcher::{delta_set, strip_sign_suffix, template_name_for_set};
use crate::utils::link_parser::parse_source_full_name;

/// The three template shapes for one delta set.
struct TemplateTriple {
    pos: Dispatch,
    neg: Dispatch,
    none: Dispatch,
}

/// Metadata collected in Pass 1 for processing in Pass 2.
struct DeltaItem {
    slide_index: i32,
    ole_name: String,
    ole_source_full: String,
    delt_base_name: String, // Name with _pos/_neg/_none suffix stripped
    set: u32,               // Template set number (1 = default tmpl_delta_*)
    delt_left: f64,
    delt_top: f64,
    delt_width: f64,
    delt_height: f64,
}

/// Determine the sign of a cell value string.
///
/// Returns "pos", "neg", or "none".
pub fn determine_sign(value: &str) -> &'static str {
    let mut s = value.trim().to_string();

    // Strip trailing %
    if s.ends_with('%') {
        s.pop();
        s = s.trim().to_string();
    }

    match s.parse::<f64>() {
        Ok(num) if num > 0.0 => "pos",
        Ok(num) if num < 0.0 => "neg",
        _ => "none",
    }
}

/// Sign of a numeric value with a dead band (GOTCHA #45): `v >= t` → "pos",
/// `v <= -t` → "neg", anything strictly in between → "none". Boundaries are inclusive
/// with a 1e-9 tolerance so `0.02` read from Excel matches a `0.02` threshold.
pub fn sign_with_threshold(v: f64, t: f64) -> &'static str {
    const EPS: f64 = 1e-9;
    if v >= t - EPS && v > 0.0 {
        "pos"
    } else if v <= -t + EPS && v < 0.0 {
        "neg"
    } else {
        "none"
    }
}

/// Read the delta's Excel cell as a number (`Range.Value2`).
///
/// `Ok(Some(v))` numeric (a percentage-typed cell is already the decimal),
/// `Ok(None)` text / blank / error cell, `Err` no usable range on the OLE or COM failure.
fn get_delta_number(
    item: &DeltaItem,
    excel_app: &mut Dispatch,
    excel_path: &str,
) -> Result<Option<f64>, String> {
    if item.ole_source_full.is_empty() {
        return Err("OLE has no SourceFullName".into());
    }
    let parts = parse_source_full_name(&item.ole_source_full);
    if parts.range_address == "Not Specified" || parts.sheet_name == "Not Specified" {
        return Err(format!("OLE link has no sheet/range: {}", item.ole_source_full));
    }
    // Use CLI excel_path, not the old SourceFullName path (GOTCHA #29)
    let mut workbooks = excel_app.get("Workbooks")
        .and_then(|v| v.as_dispatch())
        .map(Dispatch::new)
        .map_err(|e| e.to_string())?;
    let mut wb = crate::pipeline::table_updater::open_or_get_workbook(&mut workbooks, excel_path)
        .map_err(|e| e.to_string())?;
    let val = wb.get("Worksheets")
        .and_then(|v| v.as_dispatch())
        .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(parts.sheet_name.as_str())]))
        .and_then(|v| v.as_dispatch())
        .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(parts.range_address.as_str())]))
        .and_then(|v| v.as_dispatch())
        .and_then(|d| Dispatch::new(d).call("Cells", &[Variant::from(1i32), Variant::from(1i32)]))
        .and_then(|v| v.as_dispatch())
        .and_then(|d| Dispatch::new(d).get("Value2"))
        .map_err(|e| format!("{}!{}: {e}", parts.sheet_name, parts.range_address))?;
    let values = val.as_flat_opt_f64_vec().map_err(|e| e.to_string())?;
    Ok(values.first().copied().flatten())
}

/// Update all delta indicator shapes in the presentation.
///
/// Two-pass: collect metadata (Pass 1), then process (Pass 2).
/// Returns the count of deltas updated.
pub fn update_deltas(
    inventory: &SlideInventory,
    config: &Config,
    presentation: &mut Dispatch,
    excel_path: &str,
    excel_app: &mut Dispatch,
) -> OaResult<usize> {
    let template_slide = config.delta.template_slide;

    // --- Pass 1: Collect metadata ---
    let mut items: Vec<DeltaItem> = Vec::new();

    for ole_ref in &inventory.ole_shapes {
        let key = (ole_ref.slide_index, ole_ref.name.clone());

        if let Some(delt_ref) = inventory.delts.get(&key) {
            // Skip template slide
            if ole_ref.slide_index <= template_slide {
                continue;
            }

            let mut delt_shape = delt_ref.dispatch.clone();
            let delt_base = strip_sign_suffix(&delt_ref.name).to_string();
            let set = delta_set(&delt_ref.name).unwrap_or(1);

            let left = delt_shape.get("Left").and_then(|v| v.as_f64()).unwrap_or(0.0);
            let top = delt_shape.get("Top").and_then(|v| v.as_f64()).unwrap_or(0.0);
            let width = delt_shape.get("Width").and_then(|v| v.as_f64()).unwrap_or(0.0);
            let height = delt_shape.get("Height").and_then(|v| v.as_f64()).unwrap_or(0.0);

            // Get OLE source link for Excel fallback
            let ole_source = {
                let mut ole_shape = ole_ref.dispatch.clone();
                ole_shape.nav("LinkFormat")
                    .and_then(|mut lf| lf.get("SourceFullName"))
                    .and_then(|v| v.as_string())
                    .unwrap_or_default()
            };

            items.push(DeltaItem {
                slide_index: ole_ref.slide_index,
                ole_name: ole_ref.name.clone(),
                ole_source_full: ole_source,
                delt_base_name: delt_base,
                set,
                delt_left: left,
                delt_top: top,
                delt_width: width,
                delt_height: height,
            });
        }
    }

    if items.is_empty() {
        return Ok(0);
    }

    // --- Resolve template triples for every set present ---
    // One scan of the template slide, then per-set lookup by derived name.
    let slide_shapes_by_name = collect_slide_shapes_by_name(presentation, template_slide);
    let mut templates: HashMap<u32, TemplateTriple> = HashMap::new();

    let mut sets_present: Vec<u32> = items.iter().map(|i| i.set).collect();
    sets_present.sort_unstable();
    sets_present.dedup();

    for set in sets_present {
        match resolve_template_set(&slide_shapes_by_name, config, set) {
            Ok(triple) => {
                templates.insert(set, triple);
            }
            Err(missing) => {
                if set == 1 {
                    eprintln!("Warning: missing delta template shapes on slide {} — skipping deltas", template_slide);
                } else {
                    eprintln!(
                        "Warning: missing delta template shapes for set {set} on slide {} ({}) — skipping delt{set}_ deltas",
                        template_slide,
                        missing.join(", ")
                    );
                }
            }
        }
    }

    if templates.is_empty() {
        return Ok(0);
    }

    // --- Pass 2: Process each delta ---
    let mut slides = Dispatch::new(presentation.get("Slides")?.as_dispatch()?);
    let mut count = 0;

    for item in &items {
        // Skip items whose template set is unavailable (already warned above)
        let Some(triple) = templates.get_mut(&item.set) else {
            continue;
        };

        // Dead band for this delta's category (GOTCHA #45); 0 = legacy sign test.
        let (threshold, token) = config.delta.threshold_for(&item.ole_name);

        let (sign, display_value) = if threshold > 0.0 {
            // Numeric Excel read: Value2 is already decimal (2% → 0.02)
            match get_delta_number(item, &mut *excel_app, excel_path) {
                Ok(Some(v)) => (sign_with_threshold(v, threshold), format!("{v}")),
                Ok(None) => {
                    super::verbose::warn(&format!(
                        "Slide {:>2} │ {} · Excel cell is text or blank, delta set to none (thr {threshold})",
                        item.slide_index, item.delt_base_name));
                    ("none", "(text)".to_string())
                }
                Err(e) => {
                    super::verbose::warn(&format!(
                        "Slide {:>2} │ {} · cannot read Excel value ({e}), delta set to none",
                        item.slide_index, item.delt_base_name));
                    ("none", "(unreadable)".to_string())
                }
            }
        } else {
            // Legacy path — unchanged. Get the cell value (primary: from PPT table,
            // fallback: from Excel). Empty/missing data → "none" (no change indicator);
            // previously this skipped the delta, leaving stale _pos/_neg shapes.
            let cell_value = get_delta_value(inventory, item, Some(&mut *excel_app), excel_path);
            let sign = match cell_value {
                Some(ref v) if !v.is_empty() => determine_sign(v),
                _ => "none",
            };
            (sign, cell_value.unwrap_or_else(|| "(empty)".to_string()))
        };

        // Pick template by set, then by sign
        let template = match sign {
            "pos" => &mut triple.pos,
            "neg" => &mut triple.neg,
            _ => &mut triple.none,
        };

        // Get slide
        let slide_variant = match slides.call("Item", &[Variant::from(item.slide_index)]) {
            Ok(v) => v,
            Err(_) => continue,
        };
        let mut slide = match slide_variant.as_dispatch() {
            Ok(d) => Dispatch::new(d),
            Err(_) => continue,
        };

        // Delete old delt_ shape (find by base name, ignoring sign suffix)
        delete_old_delta(&mut slide, &item.delt_base_name);

        // Copy template to slide
        if template.call0("Copy").is_err() {
            continue;
        }

        let mut slide_shapes = match slide.get("Shapes") {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Dispatch::new(d),
                Err(_) => continue,
            },
            Err(_) => continue,
        };

        if slide_shapes.call0("Paste").is_err() {
            continue;
        }

        // The pasted shape is the last one
        let shape_count = slide_shapes.get("Count")
            .and_then(|v| v.as_i32())
            .unwrap_or(0);

        let new_variant = match slide_shapes.call("Item", &[Variant::from(shape_count)]) {
            Ok(v) => v,
            Err(_) => continue,
        };
        let mut new_shape = match new_variant.as_dispatch() {
            Ok(d) => Dispatch::new(d),
            Err(_) => continue,
        };

        // Reposition and rename
        let _ = new_shape.put("Left", Variant::from(item.delt_left));
        let _ = new_shape.put("Top", Variant::from(item.delt_top));
        let _ = new_shape.put("Width", Variant::from(item.delt_width));
        let _ = new_shape.put("Height", Variant::from(item.delt_height));
        let new_name = format!("{}_{sign}", item.delt_base_name);
        let _ = new_shape.put("Name", Variant::from(new_name.as_str()));

        count += 1;
        let mut detail = format!("{display_value} → {sign}");
        if item.set != 1 {
            detail.push_str(&format!(" (set {})", item.set));
        }
        if threshold > 0.0 {
            match token {
                Some(tok) => detail.push_str(&format!(" · thr {threshold} via {tok}")),
                None => detail.push_str(&format!(" · thr {threshold}")),
            }
        }
        super::verbose::detail(item.slide_index, &item.delt_base_name, &detail);
    }

    Ok(count)
}

/// Resolve the three template shapes for `set` from the pre-scanned template slide.
///
/// Set 1 uses the configured names verbatim; set N ≥ 2 derives `tmpl<N>_delta_*`
/// via `template_name_for_set`. On failure returns the list of missing names.
fn resolve_template_set(
    shapes_by_name: &HashMap<String, Dispatch>,
    config: &Config,
    set: u32,
) -> Result<TemplateTriple, Vec<String>> {
    let name_pos = template_name_for_set(&config.delta.template_positive, set, "pos");
    let name_neg = template_name_for_set(&config.delta.template_negative, set, "neg");
    let name_none = template_name_for_set(&config.delta.template_none, set, "none");

    let mut missing = Vec::new();
    for n in [&name_pos, &name_neg, &name_none] {
        if !shapes_by_name.contains_key(n.as_str()) {
            missing.push(n.clone());
        }
    }
    if !missing.is_empty() {
        return Err(missing);
    }

    Ok(TemplateTriple {
        pos: shapes_by_name[&name_pos].clone(),
        neg: shapes_by_name[&name_neg].clone(),
        none: shapes_by_name[&name_none].clone(),
    })
}

/// Scan one slide once and index its top-level shapes by name.
///
/// Returns an empty map if the slide cannot be read (caller reports missing templates).
fn collect_slide_shapes_by_name(presentation: &mut Dispatch, slide_index: i32) -> HashMap<String, Dispatch> {
    let mut map = HashMap::new();

    let Some(mut shapes) = presentation.get("Slides").ok()
        .and_then(|v| v.as_dispatch().ok())
        .map(Dispatch::new)
        .and_then(|mut slides| slides.call("Item", &[Variant::from(slide_index)]).ok())
        .and_then(|v| v.as_dispatch().ok())
        .map(Dispatch::new)
        .and_then(|mut slide| slide.get("Shapes").ok())
        .and_then(|v| v.as_dispatch().ok())
        .map(Dispatch::new)
    else {
        return map;
    };

    let count = shapes.get("Count").and_then(|v| v.as_i32()).unwrap_or(0);
    for i in 1..=count {
        if let Ok(v) = shapes.call("Item", &[Variant::from(i)])
            && let Ok(d) = v.as_dispatch()
        {
            let mut shape = Dispatch::new(d);
            if let Ok(name) = shape.get("Name").and_then(|v| v.as_string()) {
                // First occurrence wins, matching the old find_template linear scan
                map.entry(name).or_insert(shape);
            }
        }
    }

    map
}

/// Try to read the delta value from the associated PPT table, then fall back to Excel.
fn get_delta_value(
    inventory: &SlideInventory,
    item: &DeltaItem,
    excel_app: Option<&mut Dispatch>,
    excel_path: &str,
) -> Option<String> {
    let key = (item.slide_index, item.ole_name.clone());

    // Primary: read from PPT table cell (1,1)
    if let Some(table_info) = inventory.tables.get(&key) {
        let mut tbl_shape = table_info.dispatch.clone();
        let value = tbl_shape.get("Table")
            .and_then(|v| v.as_dispatch())
            .and_then(|d| {
                let mut tbl = Dispatch::new(d);
                tbl.call("Cell", &[Variant::from(1i32), Variant::from(1i32)])
            })
            .and_then(|v| v.as_dispatch())
            .and_then(|d| Dispatch::new(d).nav("Shape.TextFrame.TextRange"))
            .and_then(|mut tr| tr.get("Text"))
            .and_then(|v| v.as_string())
            .ok();

        if let Some(v) = value {
            let trimmed = v.trim().to_string();
            if !trimmed.is_empty() {
                return Some(trimmed);
            }
        }
    }

    // Fallback: read from Excel (for delt-only OLE shapes with no table)
    if let Some(excel) = excel_app
        && !item.ole_source_full.is_empty() && !excel_path.is_empty() {
            let parts = parse_source_full_name(&item.ole_source_full);
            if parts.range_address != "Not Specified" && parts.sheet_name != "Not Specified" {
                // Use CLI excel_path, not the old SourceFullName path (GOTCHA #29)
                if let Ok(mut workbooks) = excel.get("Workbooks")
                    .and_then(|v| v.as_dispatch())
                    .map(Dispatch::new)
                    && let Ok(mut wb) = crate::pipeline::table_updater::open_or_get_workbook(&mut workbooks, excel_path) {
                        let cell_text = wb.get("Worksheets")
                            .and_then(|v| v.as_dispatch())
                            .and_then(|d| Dispatch::new(d).call("Item", &[Variant::from(parts.sheet_name.as_str())]))
                            .and_then(|v| v.as_dispatch())
                            .and_then(|d| Dispatch::new(d).call("Range", &[Variant::from(parts.range_address.as_str())]))
                            .and_then(|v| v.as_dispatch())
                            .and_then(|d| Dispatch::new(d).get("Text"))
                            .and_then(|v| v.as_string())
                            .ok();

                        if let Some(text) = cell_text {
                            let trimmed = text.trim().to_string();
                            if !trimmed.is_empty() {
                                return Some(trimmed);
                            }
                        }
                    }
            }
        }

    None
}

/// Delete the old delta shape from a slide (find by base name, ignoring sign suffix).
fn delete_old_delta(slide: &mut Dispatch, base_name: &str) {
    let mut shapes = match slide.get("Shapes") {
        Ok(v) => match v.as_dispatch() {
            Ok(d) => Dispatch::new(d),
            Err(_) => return,
        },
        Err(_) => return,
    };

    let count = shapes.get("Count")
        .and_then(|v| v.as_i32())
        .unwrap_or(0);

    for i in 1..=count {
        if let Ok(v) = shapes.call("Item", &[Variant::from(i)])
            && let Ok(d) = v.as_dispatch() {
                let mut shp = Dispatch::new(d);
                let name = shp.get("Name")
                    .and_then(|v| v.as_string())
                    .unwrap_or_default();

                if strip_sign_suffix(&name) == base_name {
                    let _ = shp.call0("Delete");
                    return; // Only delete the first match
                }
            }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── GOTCHA #45: sign_with_threshold ────────────────────

    #[test]
    fn test_sign_with_threshold_boundaries() {
        assert_eq!(sign_with_threshold(0.02, 0.02), "pos", "boundary is inclusive");
        assert_eq!(sign_with_threshold(0.019, 0.02), "none");
        assert_eq!(sign_with_threshold(-0.02, 0.02), "neg", "boundary is inclusive");
        assert_eq!(sign_with_threshold(-0.019, 0.02), "none");
        assert_eq!(sign_with_threshold(0.0, 0.02), "none");
        assert_eq!(sign_with_threshold(0.05, 0.05), "pos");
        assert_eq!(sign_with_threshold(-0.03, 0.05), "none");
    }

    #[test]
    fn test_sign_with_threshold_float_noise() {
        // 2% read from Excel may carry binary noise; must still count as reaching 0.02
        assert_eq!(sign_with_threshold(0.020_000_000_000_000_004, 0.02), "pos");
        assert_eq!(sign_with_threshold(0.019_999_999_999_999_997, 0.02), "pos");
        assert_eq!(sign_with_threshold(-0.019_999_999_999_999_997, 0.02), "neg");
    }

    #[test]
    fn test_sign_with_threshold_zero_matches_legacy_rule() {
        assert_eq!(sign_with_threshold(0.0001, 0.0), "pos");
        assert_eq!(sign_with_threshold(-0.0001, 0.0), "neg");
        assert_eq!(sign_with_threshold(0.0, 0.0), "none");
    }

    #[test]
    fn test_determine_sign_positive() {
        assert_eq!(determine_sign("1.5"), "pos");
        assert_eq!(determine_sign("+0.3"), "pos");
        assert_eq!(determine_sign("1.5%"), "pos");
    }

    #[test]
    fn test_determine_sign_negative() {
        assert_eq!(determine_sign("-0.3"), "neg");
        assert_eq!(determine_sign("-100"), "neg");
        assert_eq!(determine_sign("-0.5%"), "neg");
    }

    #[test]
    fn test_determine_sign_zero() {
        assert_eq!(determine_sign("0"), "none");
        assert_eq!(determine_sign("0.0"), "none");
        assert_eq!(determine_sign("0%"), "none");
    }

    #[test]
    fn test_determine_sign_non_numeric() {
        assert_eq!(determine_sign("N/A"), "none");
        assert_eq!(determine_sign(""), "none");
        assert_eq!(determine_sign("text"), "none");
    }
}
