//! Typed Excel range reads shared by the chart pre-update and `oa check` (GOTCHA #48).
//!
//! Two readers, chosen by the chart's own reference type:
//! - `numRef`  → [`read_range_numbers`]: `Value2`, blanks/text → `None` (GOTCHA #43)
//! - `strRef`  → [`read_range_texts`]: the cells' display text, which is what PowerPoint
//!   stores for category labels and series names

use crate::com::dispatch::Dispatch;
use crate::com::variant::{CellValue, Variant};
use crate::error::OaResult;

/// Resolve a normalised `Sheet!A1:B2` reference (no `$`, no `[book]`) to a Range dispatch.
/// A reference without `!` is looked up on the `Tables` sheet.
pub fn range_dispatch(wb: &mut Dispatch, range_ref: &str) -> OaResult<Dispatch> {
    let (sheet, addr) = match range_ref.find('!') {
        Some(p) => (&range_ref[..p], &range_ref[p + 1..]),
        None => ("Tables", range_ref),
    };
    let mut sheets = Dispatch::new(wb.get("Worksheets")?.as_dispatch()?);
    let mut ws = Dispatch::new(sheets.call("Item", &[Variant::from(sheet)])?.as_dispatch()?);
    Ok(Dispatch::new(ws.call("Range", &[Variant::from(addr)])?.as_dispatch()?))
}

/// Numbers via `Value2` — one COM call for the whole range; blank → `None`.
pub fn read_range_numbers(range: &mut Dispatch) -> OaResult<Vec<Option<f64>>> {
    range.get("Value2")?.as_flat_opt_f64_vec()
}

/// Display text per cell, row-major. One `Value2` read for the range; only cells that
/// came back numeric (e.g. a `2024` header or a `54%` label) are re-read with `.Text`
/// so their number format is kept. Blank cells → `None`.
pub fn read_range_texts(range: &mut Dispatch) -> OaResult<Vec<Option<String>>> {
    let cells = range.get("Value2")?.as_flat_cell_vec()?;
    let mut out = Vec::with_capacity(cells.len());
    let mut cells_coll: Option<Dispatch> = None;
    for (i, cell) in cells.into_iter().enumerate() {
        out.push(match cell {
            CellValue::Str(s) => if s.is_empty() { None } else { Some(s) },
            CellValue::Empty => None,
            CellValue::F64(_) | CellValue::I32(_) => {
                if cells_coll.is_none() {
                    cells_coll = Some(Dispatch::new(range.get("Cells")?.as_dispatch()?));
                }
                let coll = cells_coll.as_mut().expect("initialised above");
                let text = Dispatch::new(coll.call("Item", &[Variant::from((i + 1) as i32)])?.as_dispatch()?)
                    .get("Text")?
                    .as_string()?;
                if text.is_empty() { None } else { Some(text) }
            }
        });
    }
    Ok(out)
}
