//! SlideInventory — pre-scanned index of all interesting shapes in a presentation.
//!
//! Built once by `build_inventory()`, used by all pipeline steps for O(1) lookups
//! instead of repeated slide enumeration.

use std::collections::HashMap;

use crate::com::dispatch::Dispatch;
use crate::com::variant::Variant;
use crate::office::constants::MsoShapeType;
use crate::shapes::matcher::{self, ShapePrefix, TableType};

/// A reference to a shape discovered during inventory scan.
#[derive(Debug)]
pub struct ShapeRef {
    pub dispatch: Dispatch,
    pub name: String,
    pub slide_index: i32,
}

/// Information about a table shape associated with an OLE object.
#[derive(Debug)]
pub struct TableInfo {
    pub dispatch: Dispatch,
    pub table_type: TableType,
    pub name: String,
}

/// Pre-scanned index of all interesting shapes in a presentation.
#[derive(Debug)]
pub struct SlideInventory {
    /// All linked OLE Excel.Sheet shapes: (slide_index, shape dispatch, shape name)
    pub ole_shapes: Vec<ShapeRef>,
    /// Table shapes indexed by (slide_idx, ole_name) → TableInfo
    pub tables: HashMap<(i32, String), TableInfo>,
    /// Delta shapes indexed by (slide_idx, ole_name) → shape dispatch
    pub delts: HashMap<(i32, String), ShapeRef>,
    /// Shapes with _ccst in name and HasTable
    pub ccst_tables: Vec<ShapeRef>,
    /// Linked chart shapes
    pub charts: Vec<ShapeRef>,
    /// Raw counts for `oa info`
    pub count_ntbl: usize,
    pub count_htmp: usize,
    pub count_trns: usize,
    pub count_delt: usize,
    pub count_ccst: usize,
    /// `delt*_` shapes no OLE object on their slide matched (GOTCHA #49) — human naming errors.
    pub unpaired_delts: Vec<ShapeRef>,
    /// `ntbl_/htmp_/trns_` shapes no OLE object on their slide matched (GOTCHA #49).
    pub unpaired_tables: Vec<(ShapeRef, TableType)>,
    /// OLE objects that drive neither a table nor a delta (legitimate for standalone pictures).
    pub unpaired_oles: Vec<ShapeRef>,
    /// OLE names per slide, for "closest name" hints.
    pub ole_names_by_slide: HashMap<i32, Vec<String>>,
}

impl SlideInventory {
    fn new() -> Self {
        Self {
            ole_shapes: Vec::new(),
            tables: HashMap::new(),
            delts: HashMap::new(),
            ccst_tables: Vec::new(),
            charts: Vec::new(),
            count_ntbl: 0,
            count_htmp: 0,
            count_trns: 0,
            count_delt: 0,
            count_ccst: 0,
            unpaired_delts: Vec::new(),
            unpaired_tables: Vec::new(),
            unpaired_oles: Vec::new(),
            ole_names_by_slide: HashMap::new(),
        }
    }

    /// The OLE name on `slide` closest to the token of an unpaired special shape
    /// (`delt2_globalnet_f` → token `globalnet_f` → maybe `globalnet_g`). `None` when nothing is near.
    pub fn closest_ole_name(&self, slide: i32, shape_name: &str) -> Option<&str> {
        let names = self.ole_names_by_slide.get(&slide)?;
        let candidates: Vec<&str> = names.iter().map(|s| s.as_str()).collect();
        matcher::closest_name(matcher::special_token(shape_name), &candidates)
    }

    /// Every unpaired special shape as `(slide, name, closest OLE name)`, deltas first.
    pub fn unpaired_special_shapes(&self) -> Vec<(i32, String, Option<String>)> {
        self.unpaired_delts.iter()
            .map(|r| (r.slide_index, r.name.clone()))
            .chain(self.unpaired_tables.iter().map(|(r, _)| (r.slide_index, r.name.clone())))
            .map(|(slide, name)| {
                let hint = self.closest_ole_name(slide, &name).map(|s| s.to_string());
                (slide, name, hint)
            })
            .collect()
    }
}

/// Intermediate collection of table candidate shapes (before OLE matching).
struct TableCandidate {
    slide_index: i32,
    dispatch: Dispatch,
    table_type: TableType,
    name: String,
}

/// Intermediate collection of delta candidate shapes (before OLE matching).
struct DeltCandidate {
    slide_index: i32,
    dispatch: Dispatch,
    name: String,
}

/// Scan all slides and shapes ONCE to build a complete inventory.
///
/// This replaces multiple per-step slide enumerations with a single pass.
pub fn build_inventory(presentation: &mut Dispatch) -> SlideInventory {
    let mut inventory = SlideInventory::new();
    let mut table_candidates: Vec<TableCandidate> = Vec::new();
    let mut delt_candidates: Vec<DeltCandidate> = Vec::new();

    // Get slide count
    let mut slides = match presentation.get("Slides") {
        Ok(v) => match v.as_dispatch() {
            Ok(d) => Dispatch::new(d),
            Err(_) => return inventory,
        },
        Err(_) => return inventory,
    };

    let slide_count = match slides.get("Count") {
        Ok(v) => v.as_i32().unwrap_or(0),
        Err(_) => return inventory,
    };

    // Iterate slides
    for s in 1..=slide_count {
        let mut slide = match slides.call("Item", &[Variant::from(s)]) {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Dispatch::new(d),
                Err(_) => continue,
            },
            Err(_) => continue,
        };

        let slide_index = slide.get("SlideIndex")
            .and_then(|v| v.as_i32())
            .unwrap_or(s);

        let mut shapes = match slide.get("Shapes") {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Dispatch::new(d),
                Err(_) => continue,
            },
            Err(_) => continue,
        };

        let shape_count = match shapes.get("Count") {
            Ok(v) => v.as_i32().unwrap_or(0),
            Err(_) => continue,
        };

        // Shared DISPID cache for all shapes on this slide (same COM class)
        let shape_dispid_cache = shapes.cache();

        for i in 1..=shape_count {
            if let Ok(v) = shapes.call("Item", &[Variant::from(i)])
                && let Ok(d) = v.as_dispatch() {
                    let shape = Dispatch::new_with_cache(d, shape_dispid_cache.clone());
                    scan_shape_recursive(
                        shape,
                        slide_index,
                        &mut inventory,
                        &mut table_candidates,
                        &mut delt_candidates,
                    );
                }
        }
    }

    // Index tables and delts by (slide_index, ole_name).
    // Candidates a later OLE also matches are reused (first match wins, GOTCHA #42);
    // candidates nobody matches are kept as `unpaired_*` (GOTCHA #49).
    let mut used_tables: std::collections::HashSet<usize> = std::collections::HashSet::new();
    let mut used_delts: std::collections::HashSet<usize> = std::collections::HashSet::new();

    for ole_ref in &inventory.ole_shapes {
        let key = (ole_ref.slide_index, ole_ref.name.clone());
        inventory.ole_names_by_slide.entry(ole_ref.slide_index).or_default().push(ole_ref.name.clone());

        // Tables: search priority ntbl -> htmp -> trns (same slide only)
        if !inventory.tables.contains_key(&key) {
            for priority in &[TableType::Normal, TableType::Heatmap, TableType::Transposed] {
                let found = table_candidates.iter().position(|tc| {
                    tc.slide_index == ole_ref.slide_index
                        && tc.table_type == *priority
                        && matcher::is_exact_token_match(&tc.name, &ole_ref.name)
                });
                if let Some(idx) = found {
                    let tc = &table_candidates[idx];
                    inventory.tables.insert(key.clone(), TableInfo {
                        dispatch: tc.dispatch.clone(),
                        table_type: tc.table_type,
                        name: tc.name.clone(),
                    });
                    used_tables.insert(idx);
                    break;
                }
            }
        }

        // Delts (same slide only)
        if let std::collections::hash_map::Entry::Vacant(e) = inventory.delts.entry(key) {
            let found = delt_candidates.iter().position(|dc| {
                dc.slide_index == ole_ref.slide_index
                    && matcher::is_exact_token_match(&dc.name, &ole_ref.name)
            });
            if let Some(idx) = found {
                let dc = &delt_candidates[idx];
                e.insert(ShapeRef {
                    dispatch: dc.dispatch.clone(),
                    name: dc.name.clone(),
                    slide_index: dc.slide_index,
                });
                used_delts.insert(idx);
            }
        }
    }

    // Leftovers: special shapes without an OLE partner (naming errors), and OLE objects
    // that drive nothing (usually fine — standalone linked pictures).
    for (i, tc) in table_candidates.into_iter().enumerate() {
        if !used_tables.contains(&i) {
            inventory.unpaired_tables.push((
                ShapeRef { dispatch: tc.dispatch, name: tc.name, slide_index: tc.slide_index },
                tc.table_type,
            ));
        }
    }
    for (i, dc) in delt_candidates.into_iter().enumerate() {
        if !used_delts.contains(&i) {
            inventory.unpaired_delts.push(ShapeRef { dispatch: dc.dispatch, name: dc.name, slide_index: dc.slide_index });
        }
    }
    let lonely: Vec<ShapeRef> = inventory.ole_shapes.iter()
        .filter(|o| {
            let key = (o.slide_index, o.name.clone());
            !inventory.tables.contains_key(&key) && !inventory.delts.contains_key(&key)
        })
        .map(|o| ShapeRef { dispatch: o.dispatch.clone(), name: o.name.clone(), slide_index: o.slide_index })
        .collect();
    inventory.unpaired_oles = lonely;

    inventory
}

/// Recursively scan a shape (and groups) to populate inventory lists.
fn scan_shape_recursive(
    mut shape: Dispatch,
    slide_index: i32,
    inventory: &mut SlideInventory,
    table_candidates: &mut Vec<TableCandidate>,
    delt_candidates: &mut Vec<DeltCandidate>,
) {
    let name = shape.get("Name")
        .and_then(|v| v.as_string())
        .unwrap_or_default();

    let shape_type = shape.get("Type")
        .and_then(|v| v.as_i32())
        .unwrap_or(0);

    // Group shapes
    if shape_type == MsoShapeType::Group as i32 {
        // GOTCHA #13: Check the group's name for delt_ BEFORE recursing
        if matcher::delta_set(&name).is_some() {
            delt_candidates.push(DeltCandidate {
                slide_index,
                dispatch: shape.clone(),
                name: name.clone(),
            });
            inventory.count_delt += 1;
        }

        // Recurse into group items
        if let Ok(gi_variant) = shape.get("GroupItems")
            && let Ok(gi_dispatch) = gi_variant.as_dispatch() {
                let mut group_items = Dispatch::new(gi_dispatch);
                let count = group_items.get("Count")
                    .and_then(|v| v.as_i32())
                    .unwrap_or(0);

                for i in 1..=count {
                    if let Ok(v) = group_items.call("Item", &[Variant::from(i)])
                        && let Ok(d) = v.as_dispatch() {
                            scan_shape_recursive(
                                Dispatch::new(d),
                                slide_index,
                                inventory,
                                table_candidates,
                                delt_candidates,
                            );
                        }
                }
            }
        return;
    }

    // Linked OLE Excel.Sheet — skip HasChart/HasTable checks (OLE is never a chart or table)
    if shape_type == MsoShapeType::LinkedOleObject as i32 {
        if let Ok(lf_variant) = shape.get("LinkFormat")
            && lf_variant.as_dispatch().is_ok() {
                let prog_id = shape.nav("OLEFormat")
                    .and_then(|mut ole| ole.get("ProgID"))
                    .and_then(|v| v.as_string())
                    .unwrap_or_default();

                if prog_id.contains("Excel.Sheet") {
                    inventory.ole_shapes.push(ShapeRef {
                        dispatch: shape.clone(),
                        name: name.clone(),
                        slide_index,
                    });
                }
            }
        // delt_ check for non-group OLE shapes still needed
        if matcher::delta_set(&name).is_some() {
            delt_candidates.push(DeltCandidate {
                slide_index,
                dispatch: shape.clone(),
                name,
            });
            inventory.count_delt += 1;
        }
        return; // Short-circuit: OLE is never a chart or table shape
    }

    // Linked charts
    let has_chart = shape.get("HasChart")
        .and_then(|v| v.as_i32())
        .unwrap_or(0);

    if has_chart != 0 {
        let is_linked = shape.nav("Chart.ChartData")
            .and_then(|mut cd| cd.get("IsLinked"))
            .and_then(|v| v.as_bool())
            .unwrap_or(false);

        if is_linked {
            inventory.charts.push(ShapeRef {
                dispatch: shape.clone(),
                name: name.clone(),
                slide_index,
            });
        }
        return; // Charts are never tables
    }

    // Table shapes
    let has_table = shape.get("HasTable")
        .and_then(|v| v.as_i32())
        .unwrap_or(0);

    if has_table != 0 {
        // _ccst tables
        if name.contains("_ccst") {
            inventory.ccst_tables.push(ShapeRef {
                dispatch: shape.clone(),
                name: name.clone(),
                slide_index,
            });
            inventory.count_ccst += 1;
        }

        // ntbl_/htmp_/trns_ candidates
        if let Some(prefix) = matcher::classify_shape_name(&name)
            && let Some(table_type) = matcher::prefix_to_table_type(prefix) {
                table_candidates.push(TableCandidate {
                    slide_index,
                    dispatch: shape.clone(),
                    table_type,
                    name: name.clone(),
                });
                match prefix {
                    ShapePrefix::NormalTable => inventory.count_ntbl += 1,
                    ShapePrefix::Heatmap => inventory.count_htmp += 1,
                    ShapePrefix::Transposed => inventory.count_trns += 1,
                    _ => {}
                }
            }
    }

    // delt_ candidates (non-group shapes)
    if matcher::delta_set(&name).is_some() && shape_type != MsoShapeType::Group as i32 {
        delt_candidates.push(DeltCandidate {
            slide_index,
            dispatch: shape.clone(),
            name,
        });
        inventory.count_delt += 1;
    }
}
