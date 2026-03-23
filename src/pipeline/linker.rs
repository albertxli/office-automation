//! Step 1: OLE link re-pointing via COM.
//!
//! ZIP pre-relink (Step 0) already rewrites all link paths in the PPTX XML.
//! When PowerPoint opens the file, it reads the ALREADY-CORRECT paths.
//!
//! This COM step only sets AutoUpdate=Manual if configured — it does NOT
//! touch SourceFullName (which costs 0.5s/shape due to link resolution).
//!
//! GOTCHA #15: LinkFormat.Update() is never called — too slow.
//! The ZIP pre-relink + AutoUpdate=Manual combo is 200x faster than COM path setting.

use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::office::constants::PpUpdateOption;
use crate::shapes::inventory::SlideInventory;

/// Set AutoUpdate=Manual on all linked OLE objects.
///
/// ZIP pre-relink already handled the path changes — this only sets the update mode.
/// Returns the count of shapes processed.
pub fn update_links(
    inventory: &SlideInventory,
    _excel_path: &str,
    config: &Config,
) -> OaResult<usize> {
    if inventory.ole_shapes.is_empty() || !config.links.set_manual {
        return Ok(0);
    }

    let mut updated = 0;

    for ole_ref in &inventory.ole_shapes {
        let mut shape = ole_ref.dispatch.clone();

        let result = shape.nav("LinkFormat")
            .and_then(|mut lf| lf.put("AutoUpdate", Variant::from(PpUpdateOption::Manual as i32)).map_err(|e| e));

        if result.is_ok() {
            updated += 1;
            super::verbose::detail(ole_ref.slide_index, &ole_ref.name, "AutoUpdate=Manual");
        }
    }

    Ok(updated)
}
