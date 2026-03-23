//! CellFormatting and TableFormatting — extract/apply table formatting via COM.
//!
//! Two extraction modes:
//! - `extract_formatting()` — full extraction (27+ properties per cell)
//! - `extract_formatting_minimal()` — fill-only (4 properties per cell, 75% fewer COM calls)
//!
//! The minimal version is used for htmp_ tables where only fill colors change.
//! For ntbl_/trns_ tables, formatting is SKIPPED entirely (they preserve template formatting).

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::office::constants::MsoTriState;

/// Border type indices for COM Cell.Borders() calls.
const BORDER_TYPES: [i32; 4] = [1, 2, 3, 4]; // Left, Top, Right, Bottom

/// Per-cell formatting data.
#[derive(Debug, Clone, Default)]
pub struct CellFormatting {
    // Font
    pub font_name: String,
    pub font_size: f64,
    pub font_bold: i32,
    pub font_italic: i32,
    pub font_color: i32,
    pub font_underline: i32,
    pub font_shadow: i32,
    // Paragraph / alignment
    pub h_alignment: i32,
    pub v_alignment: i32,
    pub margin_left: f64,
    pub margin_right: f64,
    pub margin_top: f64,
    pub margin_bottom: f64,
    // Fill
    pub fill_visible: bool,
    pub fill_type: i32,
    pub fill_color: i32,
    pub fill_transparency: f64,
    // Borders: (visible, weight, dash_style, color)
    pub borders: [(i32, f64, i32, i32); 4],
}

/// Complete formatting snapshot for a PowerPoint table shape.
#[derive(Debug, Clone, Default)]
pub struct TableFormatting {
    pub shape_left: f64,
    pub shape_top: f64,
    pub shape_width: f64,
    pub shape_height: f64,
    pub row_heights: Vec<f64>,
    pub column_widths: Vec<f64>,
    pub cells: Vec<Vec<CellFormatting>>,
}

/// Extract full formatting from a table shape (all 27+ properties per cell).
///
/// Only used when creating brand new tables or reformatting htmp_ tables.
pub fn extract_formatting(table_shape: &mut Dispatch) -> Option<TableFormatting> {
    let mut fmt = extract_geometry(table_shape)?;
    let mut tbl = Dispatch::new(table_shape.get("Table").ok()?.as_dispatch().ok()?);
    let rows = get_count(&mut tbl, "Rows")?;
    let cols = get_count(&mut tbl, "Columns")?;

    for i in 1..=rows {
        let mut row_cells = Vec::with_capacity(cols as usize);
        for j in 1..=cols {
            let mut cell = get_cell(&mut tbl, i, j)?;
            let mut cell_shape = Dispatch::new(cell.get("Shape").ok()?.as_dispatch().ok()?);
            let mut tf = cell_shape.nav("TextFrame").ok()?;
            let mut font = tf.nav("TextRange.Font").ok()?;
            let mut fill = Dispatch::new(cell_shape.get("Fill").ok()?.as_dispatch().ok()?);

            let cf = CellFormatting {
                font_name: font.get("Name").and_then(|v| v.as_string().map_err(|e| e)).unwrap_or_default(),
                font_size: font.get("Size").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                font_bold: font.get("Bold").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                font_italic: font.get("Italic").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                font_color: font.nav("Color").and_then(|mut c| c.get("RGB")).and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                font_underline: font.get("Underline").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                font_shadow: font.get("Shadow").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                h_alignment: tf.nav("TextRange.ParagraphFormat").and_then(|mut pf| pf.get("Alignment")).and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                v_alignment: tf.get("VerticalAnchor").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                margin_left: tf.get("MarginLeft").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                margin_right: tf.get("MarginRight").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                margin_top: tf.get("MarginTop").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                margin_bottom: tf.get("MarginBottom").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                fill_visible: fill.get("Visible").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0) != MsoTriState::False as i32,
                fill_type: fill.get("Type").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                fill_color: fill.nav("ForeColor").and_then(|mut fc| fc.get("RGB")).and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                fill_transparency: fill.get("Transparency").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                borders: extract_borders(&mut cell),
            };

            row_cells.push(cf);
        }
        fmt.cells.push(row_cells);
    }

    Some(fmt)
}

/// Extract minimal formatting: geometry + fill colors only (skip borders/fonts/margins).
///
/// ~75% fewer COM calls than full extract. Used for htmp_ tables.
pub fn extract_formatting_minimal(table_shape: &mut Dispatch) -> Option<TableFormatting> {
    let mut fmt = extract_geometry(table_shape)?;
    let mut tbl = Dispatch::new(table_shape.get("Table").ok()?.as_dispatch().ok()?);
    let rows = get_count(&mut tbl, "Rows")?;
    let cols = get_count(&mut tbl, "Columns")?;

    for i in 1..=rows {
        let mut row_cells = Vec::with_capacity(cols as usize);
        for j in 1..=cols {
            let mut cell = get_cell(&mut tbl, i, j)?;
            let mut cell_shape = Dispatch::new(cell.get("Shape").ok()?.as_dispatch().ok()?);
            let mut fill = Dispatch::new(cell_shape.get("Fill").ok()?.as_dispatch().ok()?);

            let cf = CellFormatting {
                fill_visible: fill.get("Visible").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0) != MsoTriState::False as i32,
                fill_type: fill.get("Type").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                fill_color: fill.nav("ForeColor").and_then(|mut fc| fc.get("RGB")).and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                fill_transparency: fill.get("Transparency").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                ..Default::default() // Leave font/margins/borders at defaults
            };

            row_cells.push(cf);
        }
        fmt.cells.push(row_cells);
    }

    Some(fmt)
}

/// Apply saved formatting back to a table shape.
///
/// If `preserve_fill` is false, skips fill restoration (used for htmp_ where colors come from Excel).
pub fn apply_formatting(table_shape: &mut Dispatch, fmt: &TableFormatting, preserve_fill: bool) {
    let mut tbl = match table_shape.get("Table") {
        Ok(v) => match v.as_dispatch() {
            Ok(d) => Dispatch::new(d),
            Err(_) => return,
        },
        Err(_) => return,
    };

    let rows = get_count(&mut tbl, "Rows").unwrap_or(0);
    let cols = get_count(&mut tbl, "Columns").unwrap_or(0);

    // Restore row heights
    for (i, &height) in fmt.row_heights.iter().enumerate() {
        if (i as i32) < rows {
            let mut rows_coll = Dispatch::new(tbl.get("Rows").unwrap().as_dispatch().unwrap());
            if let Ok(v) = rows_coll.call("Item", &[Variant::from((i + 1) as i32)]) {
                if let Ok(d) = v.as_dispatch() {
                    let _ = Dispatch::new(d).put("Height", Variant::from(height));
                }
            }
        }
    }

    // Restore column widths
    for (j, &width) in fmt.column_widths.iter().enumerate() {
        if (j as i32) < cols {
            let mut cols_coll = Dispatch::new(tbl.get("Columns").unwrap().as_dispatch().unwrap());
            if let Ok(v) = cols_coll.call("Item", &[Variant::from((j + 1) as i32)]) {
                if let Ok(d) = v.as_dispatch() {
                    let _ = Dispatch::new(d).put("Width", Variant::from(width));
                }
            }
        }
    }

    // Restore per-cell formatting
    for (i, row) in fmt.cells.iter().enumerate() {
        if (i as i32) >= rows {
            break;
        }
        for (j, cf) in row.iter().enumerate() {
            if (j as i32) >= cols {
                break;
            }

            let Some(mut cell) = get_cell(&mut tbl, (i + 1) as i32, (j + 1) as i32) else {
                continue;
            };

            let mut cell_shape = match cell.get("Shape") {
                Ok(v) => match v.as_dispatch() {
                    Ok(d) => Dispatch::new(d),
                    Err(_) => continue,
                },
                Err(_) => continue,
            };

            // Fill
            if preserve_fill {
                apply_fill(&mut cell_shape, cf);
            }

            // Font (skip if default — indicates minimal extraction was used)
            if !cf.font_name.is_empty() {
                apply_font(&mut cell_shape, cf);
            }

            // Margins (skip if all zero — indicates minimal extraction)
            if cf.margin_left != 0.0 || cf.margin_right != 0.0 || cf.margin_top != 0.0 || cf.margin_bottom != 0.0 {
                apply_margins(&mut cell_shape, cf);
            }

            // Borders (skip if all default)
            if cf.borders.iter().any(|(v, _, _, _)| *v != 0) {
                apply_borders(&mut cell, cf);
            }
        }
    }
}

// --- Helpers ---

fn extract_geometry(table_shape: &mut Dispatch) -> Option<TableFormatting> {
    let mut fmt = TableFormatting {
        shape_left: table_shape.get("Left").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
        shape_top: table_shape.get("Top").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
        shape_width: table_shape.get("Width").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
        shape_height: table_shape.get("Height").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
        ..Default::default()
    };

    let mut tbl = Dispatch::new(table_shape.get("Table").ok()?.as_dispatch().ok()?);
    let rows = get_count(&mut tbl, "Rows")?;
    let cols = get_count(&mut tbl, "Columns")?;

    // Row heights
    let mut rows_coll = Dispatch::new(tbl.get("Rows").ok()?.as_dispatch().ok()?);
    for i in 1..=rows {
        let h = rows_coll.call("Item", &[Variant::from(i)])
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).get("Height"))
            .and_then(|v| v.as_f64().map_err(|e| e))
            .unwrap_or(0.0);
        fmt.row_heights.push(h);
    }

    // Column widths
    let mut cols_coll = Dispatch::new(tbl.get("Columns").ok()?.as_dispatch().ok()?);
    for j in 1..=cols {
        let w = cols_coll.call("Item", &[Variant::from(j)])
            .and_then(|v| v.as_dispatch().map_err(|e| e))
            .and_then(|d| Dispatch::new(d).get("Width"))
            .and_then(|v| v.as_f64().map_err(|e| e))
            .unwrap_or(0.0);
        fmt.column_widths.push(w);
    }

    Some(fmt)
}

fn get_count(collection_parent: &mut Dispatch, name: &str) -> Option<i32> {
    collection_parent.get(name)
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .and_then(|d| Dispatch::new(d).get("Count"))
        .and_then(|v| v.as_i32().map_err(|e| e))
        .ok()
}

fn get_cell(tbl: &mut Dispatch, row: i32, col: i32) -> Option<Dispatch> {
    tbl.call("Cell", &[Variant::from(row), Variant::from(col)])
        .ok()
        .and_then(|v| v.as_dispatch().ok())
        .map(Dispatch::new)
}

fn extract_borders(cell: &mut Dispatch) -> [(i32, f64, i32, i32); 4] {
    let mut borders = [(0i32, 0.0f64, 1i32, 0x00FFFFi32); 4]; // defaults
    for (k, &bt) in BORDER_TYPES.iter().enumerate() {
        if let Ok(v) = cell.call("Borders", &[Variant::from(bt)]) {
            if let Ok(d) = v.as_dispatch() {
                let mut border = Dispatch::new(d);
                borders[k] = (
                    border.get("Visible").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0),
                    border.get("Weight").and_then(|v| v.as_f64().map_err(|e| e)).unwrap_or(0.0),
                    border.get("DashStyle").and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(1),
                    border.nav("ForeColor").and_then(|mut fc| fc.get("RGB")).and_then(|v| v.as_i32().map_err(|e| e)).unwrap_or(0x00FFFF),
                );
            }
        }
    }
    borders
}

fn apply_fill(cell_shape: &mut Dispatch, cf: &CellFormatting) {
    if let Ok(v) = cell_shape.get("Fill") {
        if let Ok(d) = v.as_dispatch() {
            let mut fill = Dispatch::new(d);
            if !cf.fill_visible {
                let _ = fill.put("Visible", Variant::from(MsoTriState::False as i32));
            } else {
                let _ = fill.put("Visible", Variant::from(MsoTriState::True as i32));
                let _ = fill.call0("Solid"); // Reset to solid fill
                let _ = fill.nav("ForeColor").and_then(|mut fc| fc.put("RGB", Variant::from(cf.fill_color)).map_err(|e| e));
                let _ = fill.put("Transparency", Variant::from(cf.fill_transparency));
            }
        }
    }
}

fn apply_font(cell_shape: &mut Dispatch, cf: &CellFormatting) {
    if let Ok(mut font) = cell_shape.nav("TextFrame.TextRange.Font") {
        let _ = font.put("Name", Variant::from(cf.font_name.as_str()));
        let _ = font.put("Size", Variant::from(cf.font_size));
        let _ = font.put("Bold", Variant::from(cf.font_bold));
        let _ = font.put("Italic", Variant::from(cf.font_italic));
        let _ = font.nav("Color").and_then(|mut c| c.put("RGB", Variant::from(cf.font_color)).map_err(|e| e));
        let _ = font.put("Underline", Variant::from(cf.font_underline));
        let _ = font.put("Shadow", Variant::from(cf.font_shadow));
    }
}

fn apply_margins(cell_shape: &mut Dispatch, cf: &CellFormatting) {
    if let Ok(mut tf) = cell_shape.nav("TextFrame") {
        let _ = tf.put("MarginLeft", Variant::from(cf.margin_left));
        let _ = tf.put("MarginRight", Variant::from(cf.margin_right));
        let _ = tf.put("MarginTop", Variant::from(cf.margin_top));
        let _ = tf.put("MarginBottom", Variant::from(cf.margin_bottom));
        if let Ok(mut pf) = tf.nav("TextRange.ParagraphFormat") {
            let _ = pf.put("Alignment", Variant::from(cf.h_alignment));
        }
        let _ = tf.put("VerticalAnchor", Variant::from(cf.v_alignment));
    }
}

fn apply_borders(cell: &mut Dispatch, cf: &CellFormatting) {
    for (k, &bt) in BORDER_TYPES.iter().enumerate() {
        if let Ok(v) = cell.call("Borders", &[Variant::from(bt)]) {
            if let Ok(d) = v.as_dispatch() {
                let mut border = Dispatch::new(d);
                let (visible, weight, dash, color) = cf.borders[k];
                if visible == MsoTriState::False as i32 {
                    let _ = border.put("Visible", Variant::from(MsoTriState::False as i32));
                } else {
                    let _ = border.put("Visible", Variant::from(visible));
                    let _ = border.put("Weight", Variant::from(weight));
                    let _ = border.put("DashStyle", Variant::from(dash));
                    let _ = border.nav("ForeColor").and_then(|mut fc| fc.put("RGB", Variant::from(color)).map_err(|e| e));
                }
            }
        }
    }
}
