//! `oa info` — inspect a PPTX file and show shape/link/chart counts.

use std::collections::HashMap;
use std::path::Path;

use crate::com::dispatch::Dispatch;
use crate::com::session::{create_instance, init_com_sta, spawn_dialog_dismisser, stop_dialog_dismisser};
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::office::constants::MsoTriState;
use crate::shapes::inventory::build_inventory;
use crate::utils::link_parser::extract_file_path;

/// Run the `oa info` command — inspect a PPTX file.
pub fn run_info(pptx_path: &str) -> OaResult<()> {
    let path = Path::new(pptx_path);
    if !path.exists() {
        eprintln!("File not found: {pptx_path}");
        std::process::exit(1);
    }

    let abs_path = path.canonicalize()?;
    // GOTCHA #26: Strip \\?\ UNC prefix
    let path_str = abs_path.to_string_lossy().to_string();
    let path_str = path_str.strip_prefix(r"\\?\").unwrap_or(&path_str);

    // Initialize COM
    let _com = init_com_sta()?;
    let (stop, handle) = spawn_dialog_dismisser();

    // Create PowerPoint
    let mut ppt = create_instance("PowerPoint.Application")?;
    ppt.put("DisplayAlerts", Variant::from(0i32))?;

    // Open presentation read-only
    let mut presentations = Dispatch::new(ppt.get("Presentations")?.as_dispatch()?);
    let pres_variant = presentations.call("Open", &[
        Variant::from(path_str),
        Variant::from(MsoTriState::True as i32),  // ReadOnly
        Variant::from(MsoTriState::True as i32),  // Untitled
        Variant::from(MsoTriState::False as i32), // WithWindow
    ])?;
    let mut presentation = Dispatch::new(pres_variant.as_dispatch()?);

    // Get slide count
    let mut slides_obj = Dispatch::new(presentation.get("Slides")?.as_dispatch()?);
    let slide_count = slides_obj.get("Count")?.as_i32()?;

    // Build inventory
    let inventory = build_inventory(&mut presentation);

    // Collect OLE source file paths
    let mut ole_sources: HashMap<String, usize> = HashMap::new();
    for ole_ref in &inventory.ole_shapes {
        let mut shape = ole_ref.dispatch.clone();
        let source = shape.nav("LinkFormat")
            .and_then(|mut lf| lf.get("SourceFullName"))
            .and_then(|v| v.as_string().map_err(|e| e))
            .map(|s| extract_file_path(&s))
            .unwrap_or_else(|_| "(unknown)".to_string());
        *ole_sources.entry(source).or_insert(0) += 1;
    }

    // Find template shapes on slide 1
    let config = Config::default();
    let template_names = [
        &config.delta.template_positive,
        &config.delta.template_negative,
        &config.delta.template_none,
    ];
    let mut template_found = Vec::new();
    for name in &template_names {
        let found = find_template_shape(&mut presentation, name, 1);
        template_found.push((name.to_string(), found));
    }

    // Count unlinked charts
    let unlinked_charts = count_unlinked_charts(&mut presentation);

    // --- Cleanup COM before printing ---
    // GOTCHA #21: Drop all refs before Quit
    drop(slides_obj);
    presentation.call("Close", &[])?;
    drop(presentation);
    drop(presentations);
    ppt.call0("Quit")?;
    drop(ppt);
    stop_dialog_dismisser(stop, handle);

    // --- Print results ---
    println!();
    println!("Presentation");
    println!("  File:   {}", Path::new(pptx_path).file_name().unwrap_or_default().to_string_lossy());
    println!("  Slides: {slide_count}");

    println!();
    println!("OLE Links");
    let total_ole: usize = ole_sources.values().sum();
    if ole_sources.is_empty() {
        println!("  (none)");
    } else {
        let mut sorted: Vec<_> = ole_sources.iter().collect();
        sorted.sort_by(|a, b| b.1.cmp(a.1)); // most common first
        for (src, count) in &sorted {
            println!("  {:<60} {:>4}", src, count);
        }
        println!("  {:<60} {:>4}", "Total", total_ole);
    }

    println!();
    println!("Charts");
    println!("  Linked:   {}", inventory.charts.len());
    println!("  Unlinked: {unlinked_charts}");

    println!();
    println!("Special Shapes");
    println!("  ntbl_ (normal tables):    {}", inventory.count_ntbl);
    println!("  htmp_ (heatmap tables):   {}", inventory.count_htmp);
    println!("  trns_ (transposed):       {}", inventory.count_trns);
    println!("  delt_ (delta indicators): {}", inventory.count_delt);
    println!("  _ccst (color-coded):      {}", inventory.count_ccst);

    println!();
    println!("Delta Templates (Slide 1)");
    for (name, found) in &template_found {
        let marker = if *found { "\u{2713}" } else { "\u{2717}" };
        println!("  {name:<30} {marker}");
    }

    Ok(())
}

/// Check if a template shape exists on the given slide.
fn find_template_shape(presentation: &mut Dispatch, name: &str, slide_index: i32) -> bool {
    let slide = {
        let mut slides_disp = match presentation.get("Slides") {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Dispatch::new(d),
                Err(_) => return false,
            },
            Err(_) => return false,
        };
        match slides_disp.call("Item", &[Variant::from(slide_index)]) {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Ok(Dispatch::new(d)),
                Err(e) => Err(e),
            },
            Err(e) => Err(e),
        }
    };

    let mut slide = match slide {
        Ok(s) => s,
        Err(_) => return false,
    };

    let mut shapes = match slide.get("Shapes") {
        Ok(v) => match v.as_dispatch() {
            Ok(d) => Dispatch::new(d),
            Err(_) => return false,
        },
        Err(_) => return false,
    };

    let count = shapes.get("Count")
        .and_then(|v| v.as_i32().map_err(|e| e))
        .unwrap_or(0);

    for i in 1..=count {
        if let Ok(v) = shapes.call("Item", &[Variant::from(i)]) {
            if let Ok(d) = v.as_dispatch() {
                let mut shp = Dispatch::new(d);
                if let Ok(n) = shp.get("Name") {
                    if let Ok(shape_name) = n.as_string() {
                        if shape_name == name {
                            return true;
                        }
                    }
                }
            }
        }
    }

    false
}

/// Count unlinked charts in the presentation.
fn count_unlinked_charts(presentation: &mut Dispatch) -> usize {
    let mut count = 0;
    let mut slides = match presentation.get("Slides") {
        Ok(v) => match v.as_dispatch() {
            Ok(d) => Dispatch::new(d),
            Err(_) => return 0,
        },
        Err(_) => return 0,
    };

    let slide_count = slides.get("Count")
        .and_then(|v| v.as_i32().map_err(|e| e))
        .unwrap_or(0);

    for s in 1..=slide_count {
        let mut slide = match slides.call("Item", &[Variant::from(s)]) {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Dispatch::new(d),
                Err(_) => continue,
            },
            Err(_) => continue,
        };

        let mut shapes = match slide.get("Shapes") {
            Ok(v) => match v.as_dispatch() {
                Ok(d) => Dispatch::new(d),
                Err(_) => continue,
            },
            Err(_) => continue,
        };

        let shape_count = shapes.get("Count")
            .and_then(|v| v.as_i32().map_err(|e| e))
            .unwrap_or(0);

        for i in 1..=shape_count {
            if let Ok(v) = shapes.call("Item", &[Variant::from(i)]) {
                if let Ok(d) = v.as_dispatch() {
                    let mut shp = Dispatch::new(d);
                    let has_chart = shp.get("HasChart")
                        .and_then(|v| v.as_i32().map_err(|e| e))
                        .unwrap_or(0);

                    if has_chart != 0 {
                        let is_linked = shp.nav("Chart.ChartData")
                            .and_then(|mut cd| cd.get("IsLinked"))
                            .and_then(|v| v.as_bool().map_err(|e| e))
                            .unwrap_or(false);

                        if !is_linked {
                            count += 1;
                        }
                    }
                }
            }
        }
    }

    count
}
