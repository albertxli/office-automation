//! Step 2: Populate PPT tables from Excel ranges.
//!
//! For each linked OLE object, finds the associated table shape and fills it
//! with values from the linked Excel range.
//!
//! Table types:
//! - ntbl_ (normal): preserve formatting, only update cell text — skip formatting extract/apply
//! - htmp_ (heatmap): recalculate 3-color scale from Excel, apply fill + contrast fonts
//! - trns_ (transposed): swap rows/cols, preserve formatting
//!
//! GOTCHA #16: Use .Text (not .Value2) for display strings from Excel.
//! Python optimization: skipping formatting on ntbl_/trns_ saves ~232k COM calls.

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::shapes::inventory::{SlideInventory, TableInfo};
use crate::shapes::matcher::TableType;
use crate::utils::color::{contrast_font_color, hex_to_bgr};
use crate::utils::link_parser::parse_source_full_name;

/// Update all tables in the presentation from Excel data.
///
/// `excel_path` is the NEW Excel file path (from the command line), used to open workbooks.
/// The sheet name and range are extracted from each OLE's SourceFullName.
///
/// Returns the count of tables updated.
pub fn update_tables(
    inventory: &SlideInventory,
    config: &Config,
    excel_app: &mut Dispatch,
    excel_path: &str,
) -> OaResult<usize> {
    if inventory.ole_shapes.is_empty() {
        return Ok(0);
    }

    let mut count = 0;

    for ole_ref in &inventory.ole_shapes {
        let mut ole_shape = ole_ref.dispatch.clone();

        // Get link parts
        let source_full = match ole_shape.nav("LinkFormat")
            .and_then(|mut lf| lf.get("SourceFullName"))
            .and_then(|v| v.as_string().map_err(|e| e))
        {
            Ok(s) => s,
            Err(_) => continue,
        };

        let parts = parse_source_full_name(&source_full);
        if parts.range_address == "Not Specified" {
            continue;
        }

        let key = (ole_ref.slide_index, ole_ref.name.clone());

        // Find associated table shape
        let table_info = inventory.tables.get(&key);

        // Skip if this OLE has a delta but no table (delt-only shape)
        if table_info.is_none() && inventory.delts.contains_key(&key) {
            continue;
        }

        // Skip if no table shape exists and we can't create one (need a table to fill)
        let Some(table_info) = table_info else {
            continue;
        };

        // Use the NEW excel_path (from command line), not the old path from SourceFullName.
        // The sheet name and range still come from SourceFullName.
        match process_table(
            table_info,
            excel_path,
            &parts.sheet_name,
            &parts.range_address,
            config,
            excel_app,
        ) {
            Ok(true) => count += 1,
            Ok(false) => {} // Skipped
            Err(e) => {
                eprintln!("Warning: failed to update table for '{}': {e}", ole_ref.name);
            }
        }
    }

    Ok(count)
}

/// Process a single table: read Excel data and fill PPT table cells.
fn process_table(
    table_info: &TableInfo,
    file_path: &str,
    sheet_name: &str,
    range_address: &str,
    config: &Config,
    excel_app: &mut Dispatch,
) -> OaResult<bool> {
    let do_transpose = table_info.table_type == TableType::Transposed;

    // For ntbl_ and trns_: skip all formatting (preserve template formatting)
    // Only htmp_ tables need formatting work
    let skip_formatting = table_info.table_type != TableType::Heatmap;

    // Open workbook and get range
    let mut workbooks = Dispatch::new(excel_app.get("Workbooks")?.as_dispatch()?);
    let mut wb = open_or_get_workbook(&mut workbooks, file_path)?;
    let mut sheets = Dispatch::new(wb.get("Worksheets")?.as_dispatch()?);

    let sheet_variant = sheets.call("Item", &[Variant::from(sheet_name)])?;
    let mut sheet = Dispatch::new(sheet_variant.as_dispatch()?);

    let range_variant = sheet.call("Range", &[Variant::from(range_address)])?;
    let mut cell_range = Dispatch::new(range_variant.as_dispatch()?);

    // Get range dimensions
    let max_rows = get_count_prop(&mut cell_range, "Rows")?;
    let max_cols = get_count_prop(&mut cell_range, "Columns")?;

    // For heatmap: apply color scale and contrast fonts in Excel
    if !skip_formatting {
        apply_heatmap_to_excel(&mut cell_range, &mut sheet, excel_app, config)?;
    }

    // Fill PPT table cells
    let mut table_shape = table_info.dispatch.clone();
    let mut tbl = Dispatch::new(table_shape.get("Table")?.as_dispatch()?);

    let total_rows = get_count_prop(&mut tbl, "Rows")?;
    let total_cols = get_count_prop(&mut tbl, "Columns")?;

    let mut cells_dispatch = Dispatch::new(cell_range.get("Cells")?.as_dispatch()?);

    for excel_row in 1..=max_rows {
        for excel_col in 1..=max_cols {
            // Apply transpose if needed
            let (ppt_row, ppt_col) = if do_transpose {
                (excel_col, excel_row)
            } else {
                (excel_row, excel_col)
            };

            // Bounds check
            if ppt_row > total_rows || ppt_col > total_cols {
                continue;
            }

            // Read Excel cell text (GOTCHA #16: .Text preserves formatting)
            let excel_cell_variant = cells_dispatch.call("Item", &[
                Variant::from(excel_row),
                Variant::from(excel_col),
            ])?;
            let mut excel_cell = Dispatch::new(excel_cell_variant.as_dispatch()?);
            let cell_text = excel_cell.get("Text")
                .and_then(|v| v.as_string().map_err(|e| e))
                .unwrap_or_default();

            // Get PPT table cell
            let ppt_cell_variant = tbl.call("Cell", &[
                Variant::from(ppt_row),
                Variant::from(ppt_col),
            ])?;
            let mut ppt_cell = Dispatch::new(ppt_cell_variant.as_dispatch()?);

            // Set cell text
            let mut cell_shape = Dispatch::new(ppt_cell.get("Shape")?.as_dispatch()?);
            let _ = cell_shape.nav("TextFrame.TextRange")
                .and_then(|mut tr| tr.put("Text", Variant::from(cell_text.as_str())).map_err(|e| e));

            // For heatmap: copy fill color and font from Excel
            if !skip_formatting {
                // Fill color from Excel's DisplayFormat
                if let Ok(bg_color) = excel_cell.nav("DisplayFormat.Interior")
                    .and_then(|mut int| int.get("Color"))
                    .and_then(|v| v.as_i32().map_err(|e| e))
                {
                    let _ = cell_shape.nav("Fill")
                        .and_then(|mut fill| {
                            fill.call0("Solid")?;
                            fill.nav("ForeColor")
                                .and_then(|mut fc| fc.put("RGB", Variant::from(bg_color)).map_err(|e| e))
                        });
                }

                // Font properties from Excel
                if let Ok(mut excel_font) = excel_cell.get("Font")
                    .and_then(|v| v.as_dispatch().map_err(|e| e))
                    .map(Dispatch::new)
                {
                    if let Ok(mut ppt_font) = cell_shape.nav("TextFrame.TextRange.Font") {
                        let _ = ppt_font.put("Name", Variant::from(
                            excel_font.get("Name").and_then(|v| v.as_string().map_err(|e| e)).unwrap_or_default().as_str()
                        ));
                        let _ = ppt_font.put("Size", excel_font.get("Size").unwrap_or_default());
                        let _ = ppt_font.put("Bold", excel_font.get("Bold").unwrap_or_default());
                        let _ = ppt_font.put("Italic", excel_font.get("Italic").unwrap_or_default());
                        let _ = ppt_font.nav("Color").and_then(|mut c| {
                            let color = excel_font.get("Color").unwrap_or_default();
                            c.put("RGB", color).map_err(|e| e)
                        });
                    }
                }
            }
        }
    }

    Ok(true)
}

/// Apply a 3-color heatmap scale to an Excel range and calculate contrast fonts.
fn apply_heatmap_to_excel(
    cell_range: &mut Dispatch,
    sheet: &mut Dispatch,
    excel_app: &mut Dispatch,
    config: &Config,
) -> OaResult<()> {
    let color_min = hex_to_bgr(&config.heatmap.color_minimum);
    let color_mid = hex_to_bgr(&config.heatmap.color_midpoint);
    let color_max = hex_to_bgr(&config.heatmap.color_maximum);
    let dark_font = hex_to_bgr(&config.heatmap.dark_font);
    let light_font = hex_to_bgr(&config.heatmap.light_font);

    // Clear existing conditional formatting and add 3-color scale
    let mut fmt_conditions = Dispatch::new(cell_range.get("FormatConditions")?.as_dispatch()?);
    let _ = fmt_conditions.call0("Delete");

    let cs_variant = fmt_conditions.call("AddColorScale", &[Variant::from(3i32)])?;
    let mut cs = Dispatch::new(cs_variant.as_dispatch()?);

    // Criterion 1: Lowest → minimum color
    let mut crit1 = Dispatch::new(cs.call("ColorScaleCriteria", &[Variant::from(1i32)])?.as_dispatch()?);
    crit1.put("Type", Variant::from(1i32))?; // xlConditionValueLowestValue
    crit1.nav("FormatColor")
        .and_then(|mut fc| fc.put("Color", Variant::from(color_min)).map_err(|e| e))?;

    // Criterion 2: 50th percentile → midpoint color
    let mut crit2 = Dispatch::new(cs.call("ColorScaleCriteria", &[Variant::from(2i32)])?.as_dispatch()?);
    crit2.put("Type", Variant::from(4i32))?; // xlConditionValuePercentile
    crit2.put("Value", Variant::from(50i32))?;
    crit2.nav("FormatColor")
        .and_then(|mut fc| fc.put("Color", Variant::from(color_mid)).map_err(|e| e))?;

    // Criterion 3: Highest → maximum color
    let mut crit3 = Dispatch::new(cs.call("ColorScaleCriteria", &[Variant::from(3i32)])?.as_dispatch()?);
    crit3.put("Type", Variant::from(2i32))?; // xlConditionValueHighestValue
    crit3.nav("FormatColor")
        .and_then(|mut fc| fc.put("Color", Variant::from(color_max)).map_err(|e| e))?;

    // Recalculate to apply conditional formatting
    let _ = excel_app.call0("Calculate");
    let _ = sheet.call0("Calculate");

    // Apply contrast font colors based on background
    let mut cells = Dispatch::new(cell_range.get("Cells")?.as_dispatch()?);
    let rows = get_count_prop(cell_range, "Rows")?;
    let cols = get_count_prop(cell_range, "Columns")?;

    for r in 1..=rows {
        for c in 1..=cols {
            if let Ok(cell_v) = cells.call("Item", &[Variant::from(r), Variant::from(c)]) {
                if let Ok(d) = cell_v.as_dispatch() {
                    let mut cell = Dispatch::new(d);
                    let bg_color = cell.nav("DisplayFormat.Interior")
                        .and_then(|mut int| int.get("Color"))
                        .and_then(|v| v.as_i32().map_err(|e| e))
                        .unwrap_or(0);

                    let font_color = contrast_font_color(bg_color, dark_font, light_font);
                    let _ = cell.nav("Font.Color")
                        .and_then(|mut fc| fc.put("", Variant::from(font_color)).map_err(|e| e))
                        // Fallback: set via Font directly
                        .or_else(|_| cell.nav("Font").and_then(|mut f| f.put("Color", Variant::from(font_color)).map_err(|e| e)));
                }
            }
        }
    }

    Ok(())
}

/// Open a workbook or return an already-open one.
pub fn open_or_get_workbook(workbooks: &mut Dispatch, file_path: &str) -> OaResult<Dispatch> {
    // Try to find already-open workbook by filename
    let filename = std::path::Path::new(file_path)
        .file_name()
        .map(|f| f.to_string_lossy().to_string())
        .unwrap_or_default();

    let count = workbooks.get("Count")
        .and_then(|v| v.as_i32().map_err(|e| e))
        .unwrap_or(0);

    for i in 1..=count {
        if let Ok(v) = workbooks.call("Item", &[Variant::from(i)]) {
            if let Ok(d) = v.as_dispatch() {
                let mut wb = Dispatch::new(d);
                let wb_name = wb.get("Name")
                    .and_then(|v| v.as_string().map_err(|e| e))
                    .unwrap_or_default();
                if wb_name.eq_ignore_ascii_case(&filename) {
                    return Ok(wb);
                }
            }
        }
    }

    // Not open — open it
    let wb_variant = workbooks.call("Open", &[
        Variant::from(file_path),
        Variant::from(0i32), // UpdateLinks = 0
    ])?;
    Ok(Dispatch::new(wb_variant.as_dispatch()?))
}

fn get_count_prop(parent: &mut Dispatch, collection_name: &str) -> OaResult<i32> {
    let coll = parent.get(collection_name)?.as_dispatch()?;
    let count = Dispatch::new(coll).get("Count")?.as_i32()?;
    Ok(count)
}
