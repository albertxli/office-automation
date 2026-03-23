//! Step 5: Update chart data links via COM.
//!
//! Unlike OLE links (where we skip Update), charts NEED LinkFormat.Update()
//! to refresh their data series. This is slower but necessary for correct output.
//!
//! GOTCHA #25: Some charts have IsLinked=True but SourceFullName="NULL" —
//! these are broken links that we skip gracefully.

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::error::OaResult;
use crate::office::constants::PpUpdateOption;
use crate::shapes::inventory::SlideInventory;

/// Update all linked chart shapes to point to the new Excel file.
///
/// Sets SourceFullName, calls Update() to refresh data, then sets AutoUpdate=Manual.
///
/// Returns the count of charts successfully updated.
pub fn update_charts(
    inventory: &SlideInventory,
    excel_path: &str,
) -> OaResult<usize> {
    if inventory.charts.is_empty() {
        return Ok(0);
    }

    let mut updated = 0;

    for chart_ref in &inventory.charts {
        let mut shape = chart_ref.dispatch.clone();

        let mut link_format = match shape.nav("LinkFormat") {
            Ok(lf) => lf,
            Err(_) => continue,
        };

        // GOTCHA #25: Check for broken links (SourceFullName = "NULL")
        let current_source = link_format.get("SourceFullName")
            .and_then(|v| v.as_string().map_err(|e| e))
            .unwrap_or_default();

        if current_source == "NULL" || current_source.is_empty() {
            // Broken link — skip gracefully
            continue;
        }

        // Set new source
        if let Err(e) = link_format.put("SourceFullName", Variant::from(excel_path)) {
            eprintln!("Warning: failed to set chart SourceFullName for '{}': {e}", chart_ref.name);
            continue;
        }

        // Update chart data (charts need this, unlike OLE links)
        if let Err(e) = link_format.call0("Update") {
            eprintln!("Warning: failed to update chart data for '{}': {e}", chart_ref.name);
            // Continue anyway — the source was set even if update failed
        }

        // Set to manual update
        let _ = link_format.put("AutoUpdate", Variant::from(PpUpdateOption::Manual as i32));

        updated += 1;
    }

    Ok(updated)
}
