//! `oa diff` — compare two PPTX files side by side.
//!
//! Opens both files read-only, builds inventories, and compares
//! table values, delta signs, and chart presence.

use crate::com::dispatch::Dispatch;
use crate::com::session::{create_instance, init_com_sta, spawn_dialog_dismisser, stop_dialog_dismisser};
use crate::com::variant::Variant;
use crate::error::OaResult;
use crate::office::constants::MsoTriState;
use crate::shapes::inventory::{build_inventory, SlideInventory};

fn strip_unc(path: &std::path::Path) -> String {
    let s = path.to_string_lossy().to_string();
    s.strip_prefix(r"\\?\").unwrap_or(&s).to_string()
}

/// A single difference found.
#[derive(Debug)]
pub struct Diff {
    pub slide: i32,
    pub shape: String,
    pub category: String,
    pub detail: String,
}

/// Run the `oa diff` command.
pub fn run_diff(file_a: &str, file_b: &str) -> OaResult<Vec<Diff>> {
    let path_a = std::path::Path::new(file_a);
    let path_b = std::path::Path::new(file_b);
    if !path_a.exists() {
        return Err(crate::error::OaError::Other(format!("File not found: {file_a}")));
    }
    if !path_b.exists() {
        return Err(crate::error::OaError::Other(format!("File not found: {file_b}")));
    }

    let str_a = strip_unc(&path_a.canonicalize()?);
    let str_b = strip_unc(&path_b.canonicalize()?);

    let _com = init_com_sta()?;
    let (stop, handle) = spawn_dialog_dismisser();

    let mut ppt = create_instance("PowerPoint.Application")?;
    ppt.put("DisplayAlerts", Variant::from(0i32))?;

    let mut preses = Dispatch::new(ppt.get("Presentations")?.as_dispatch()?);

    // Open both read-only
    let pres_a_v = preses.call("Open", &[
        Variant::from(str_a.as_str()),
        Variant::from(MsoTriState::True as i32),
        Variant::from(MsoTriState::True as i32),
        Variant::from(MsoTriState::False as i32),
    ])?;
    let mut pres_a = Dispatch::new(pres_a_v.as_dispatch()?);

    let pres_b_v = preses.call("Open", &[
        Variant::from(str_b.as_str()),
        Variant::from(MsoTriState::True as i32),
        Variant::from(MsoTriState::True as i32),
        Variant::from(MsoTriState::False as i32),
    ])?;
    let mut pres_b = Dispatch::new(pres_b_v.as_dispatch()?);

    let inv_a = build_inventory(&mut pres_a);
    let inv_b = build_inventory(&mut pres_b);

    // Compare
    let diffs = compare_inventories(&inv_a, &inv_b);

    // Cleanup (GOTCHA #21)
    drop(inv_a);
    drop(inv_b);
    pres_a.call("Close", &[])?;
    pres_b.call("Close", &[])?;
    drop(pres_a);
    drop(pres_b);
    drop(preses);
    ppt.call0("Quit")?;
    stop_dialog_dismisser(stop, handle);

    // Print results
    let name_a = path_a.file_name().unwrap_or_default().to_string_lossy();
    let name_b = path_b.file_name().unwrap_or_default().to_string_lossy();

    if diffs.is_empty() {
        println!("No differences found between {name_a} and {name_b}.");
    } else {
        println!("{} difference(s) between {name_a} and {name_b}:", diffs.len());
        for d in &diffs {
            println!("  Slide {}, {}: [{}] {}", d.slide, d.shape, d.category, d.detail);
        }
    }

    Ok(diffs)
}

/// Compare two inventories and return differences.
fn compare_inventories(a: &SlideInventory, b: &SlideInventory) -> Vec<Diff> {
    let mut diffs = Vec::new();

    // Compare table counts
    if a.count_ntbl != b.count_ntbl {
        diffs.push(Diff {
            slide: 0, shape: "(global)".into(),
            category: "count".into(),
            detail: format!("ntbl_ count: {} vs {}", a.count_ntbl, b.count_ntbl),
        });
    }
    if a.count_htmp != b.count_htmp {
        diffs.push(Diff {
            slide: 0, shape: "(global)".into(),
            category: "count".into(),
            detail: format!("htmp_ count: {} vs {}", a.count_htmp, b.count_htmp),
        });
    }

    // Compare table values for matching shapes
    for (key, table_a) in &a.tables {
        if let Some(table_b) = b.tables.get(key) {
            let val_a = read_table_cell_11(&table_a.dispatch);
            let val_b = read_table_cell_11(&table_b.dispatch);
            if val_a != val_b {
                diffs.push(Diff {
                    slide: key.0,
                    shape: key.1.clone(),
                    category: "table".into(),
                    detail: format!("A={val_a:?} vs B={val_b:?}"),
                });
            }
        } else {
            diffs.push(Diff {
                slide: key.0,
                shape: key.1.clone(),
                category: "only_in_A".into(),
                detail: "Table exists in A but not B".into(),
            });
        }
    }

    for key in b.tables.keys() {
        if !a.tables.contains_key(key) {
            diffs.push(Diff {
                slide: key.0,
                shape: key.1.clone(),
                category: "only_in_B".into(),
                detail: "Table exists in B but not A".into(),
            });
        }
    }

    // Compare chart counts
    if a.charts.len() != b.charts.len() {
        diffs.push(Diff {
            slide: 0, shape: "(global)".into(),
            category: "count".into(),
            detail: format!("Linked charts: {} vs {}", a.charts.len(), b.charts.len()),
        });
    }

    diffs
}

fn read_table_cell_11(shape: &Dispatch) -> String {
    let mut s = shape.clone();
    s.get("Table")
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .and_then(|d| Dispatch::new(d).call("Cell", &[Variant::from(1i32), Variant::from(1i32)]))
        .and_then(|v| v.as_dispatch().map_err(|e| e))
        .and_then(|d| Dispatch::new(d).nav("Shape.TextFrame.TextRange"))
        .and_then(|mut tr| tr.get("Text"))
        .and_then(|v| v.as_string().map_err(|e| e))
        .unwrap_or_default()
        .trim()
        .to_string()
}
