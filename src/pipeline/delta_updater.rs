//! Step 3: Swap delta indicator arrows based on value sign.
//!
//! Uses a two-pass algorithm:
//! 1. Collect all OLE+delt pairs and their metadata (safe for iteration)
//! 2. Process each pair: delete old shape, copy template, reposition
//!
//! Value sign is determined from the PPT table cell (primary) or Excel (fallback).
//! Template shapes on slide 1: tmpl_delta_pos, tmpl_delta_neg, tmpl_delta_none.

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::shapes::inventory::SlideInventory;
use crate::shapes::matcher::strip_sign_suffix;
use crate::utils::link_parser::parse_source_full_name;

/// Metadata collected in Pass 1 for processing in Pass 2.
struct DeltaItem {
    slide_index: i32,
    ole_name: String,
    ole_source_full: String,
    delt_base_name: String, // Name with _pos/_neg/_none suffix stripped
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
    // Find template shapes on slide 1
    let template_slide = config.delta.template_slide;
    let tmpl_pos = find_template(presentation, &config.delta.template_positive, template_slide);
    let tmpl_neg = find_template(presentation, &config.delta.template_negative, template_slide);
    let tmpl_none = find_template(presentation, &config.delta.template_none, template_slide);

    let (Some(mut tmpl_pos), Some(mut tmpl_neg), Some(mut tmpl_none)) = (tmpl_pos, tmpl_neg, tmpl_none) else {
        eprintln!("Warning: missing delta template shapes on slide {} — skipping deltas", template_slide);
        return Ok(0);
    };

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

    // --- Pass 2: Process each delta ---
    let mut slides = Dispatch::new(presentation.get("Slides")?.as_dispatch()?);
    let mut count = 0;

    for item in &items {
        // Get the cell value (primary: from PPT table, fallback: from Excel)
        let cell_value = get_delta_value(
            inventory,
            item,
            Some(&mut *excel_app),
            excel_path,
        );

        // Empty/missing data → treat as "none" (no change indicator).
        // Previously this skipped the delta entirely, leaving stale _pos/_neg
        // shapes from the template or a previous run.
        let sign = match cell_value {
            Some(ref v) if !v.is_empty() => determine_sign(v),
            _ => "none",
        };

        // Pick template by sign
        let template = match sign {
            "pos" => &mut tmpl_pos,
            "neg" => &mut tmpl_neg,
            _ => &mut tmpl_none,
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
        let display_value = cell_value.as_deref().unwrap_or("(empty)");
        super::verbose::detail(
            item.slide_index,
            &item.delt_base_name,
            &format!("{display_value} → {sign}"),
        );
    }

    Ok(count)
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

/// Find a template shape by name on the specified slide.
fn find_template(presentation: &mut Dispatch, name: &str, slide_index: i32) -> Option<Dispatch> {
    let mut slides = Dispatch::new(presentation.get("Slides").ok()?.as_dispatch().ok()?);
    let slide_variant = slides.call("Item", &[Variant::from(slide_index)]).ok()?;
    let mut slide = Dispatch::new(slide_variant.as_dispatch().ok()?);

    let mut shapes = Dispatch::new(slide.get("Shapes").ok()?.as_dispatch().ok()?);
    let count = shapes.get("Count").ok()?.as_i32().ok()?;

    for i in 1..=count {
        let shape_variant = shapes.call("Item", &[Variant::from(i)]).ok()?;
        let mut shape = Dispatch::new(shape_variant.as_dispatch().ok()?);
        let shape_name = shape.get("Name").ok()?.as_string().ok()?;
        if shape_name == name {
            return Some(shape);
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
