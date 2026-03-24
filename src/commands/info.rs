//! `oa info` — inspect a PPTX file and show shape/link/chart counts.

use std::collections::HashMap;
use std::path::Path;

use console::Style;

use crate::com::dispatch::Dispatch;
use crate::com::session::{create_instance, init_com_sta, spawn_dialog_dismisser, stop_dialog_dismisser};
use crate::com::variant::Variant;
use crate::config::Config;
use crate::error::OaResult;
use crate::office::constants::MsoTriState;
use crate::shapes::inventory::build_inventory;
use crate::utils::link_parser::extract_file_path;

/// Per-slide shape counts for verbose output.
struct SlideBreakdown {
    slide: i32,
    ole: usize,
    chart: usize,
    ntbl: usize,
    htmp: usize,
    trns: usize,
    delt: usize,
    ccst: usize,
}

impl SlideBreakdown {
    fn total(&self) -> usize {
        self.ole + self.chart + self.ntbl + self.htmp + self.trns + self.delt + self.ccst
    }
}

/// Run the `oa info` command — inspect a PPTX file.
pub fn run_info(pptx_path: &str, verbose: bool) -> OaResult<()> {
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

    // Collect per-slide breakdown from inventory (before dropping COM refs)
    let per_slide = if verbose {
        collect_per_slide_breakdown(&inventory, slide_count)
    } else {
        Vec::new()
    };

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
    let s_cyan = Style::new().cyan();
    let s_dim = Style::new().dim();

    let file_name = Path::new(pptx_path).file_name().unwrap_or_default().to_string_lossy();
    println!();
    println!("  {} {}", s_cyan.apply_to("▸"), s_cyan.apply_to(&*file_name));
    println!("  {}", s_dim.apply_to("╌".repeat(61)));

    // File size
    let file_size = std::fs::metadata(pptx_path).map(|m| m.len()).unwrap_or(0);
    println!();
    info_row_str("File size", &format_file_size(file_size), false);

    // Slides
    info_row("Slides", slide_count as usize, false);

    // OLE links
    let total_ole: usize = ole_sources.values().sum();
    println!();
    info_row("OLE links", total_ole, false);
    let mut sorted: Vec<_> = ole_sources.iter().collect();
    sorted.sort_by(|a, b| b.1.cmp(a.1));
    for (src, count) in &sorted {
        let display_name = Path::new(src.as_str())
            .file_name()
            .map(|f| f.to_string_lossy().to_string())
            .unwrap_or_else(|| src.to_string());
        info_row(&display_name, **count, true);
    }

    // Charts
    let total_charts = inventory.charts.len() + unlinked_charts;
    println!();
    info_row("Charts", total_charts, false);
    info_row("Linked", inventory.charts.len(), true);
    info_row("Unlinked", unlinked_charts, true);
    let empty_cache = count_empty_cache_charts(path);
    if empty_cache > 0 {
        info_row("Empty cache", empty_cache, true);
    }

    // Special shapes
    let total_special = inventory.count_ntbl + inventory.count_htmp
        + inventory.count_trns + inventory.count_delt + inventory.count_ccst;
    println!();
    info_row("Special shapes", total_special, false);
    info_row("ntbl_ normal tables", inventory.count_ntbl, true);
    info_row("htmp_ heatmap tables", inventory.count_htmp, true);
    info_row("trns_ transposed tables", inventory.count_trns, true);
    info_row("delt_ delta indicators", inventory.count_delt, true);
    info_row("_ccst color-coded", inventory.count_ccst, true);

    // Delta templates
    println!();
    println!("  {}", s_dim.apply_to("Delta templates"));
    for (name, found) in &template_found {
        info_row_status(name, *found);
    }

    // Per-slide breakdown (verbose only)
    if verbose && !per_slide.is_empty() {
        print_per_slide_breakdown(&per_slide, slide_count);
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

/// Print a dot-leader row with a right-aligned string value (e.g., "12.4 MB").
fn info_row_str(label: &str, value: &str, indent: bool) {
    let s_dim = Style::new().dim();
    let s_count = Style::new().white().bold();
    let prefix = if indent { "    ╰ " } else { "  " };
    let target_col: usize = 48;

    let display_len = prefix.chars().count() + label.chars().count() + 1;
    let leader_len = target_col.saturating_sub(display_len);
    let padded = format!("{prefix}{label} {}", "·".repeat(leader_len));

    // Right-align the string in 4+ chars
    println!("{} {:>4}",
        s_dim.apply_to(&padded),
        s_count.apply_to(value));
}

/// Format a byte count as human-readable file size.
fn format_file_size(bytes: u64) -> String {
    if bytes < 1024 {
        format!("{bytes} B")
    } else if bytes < 1024 * 1024 {
        format!("{:.1} KB", bytes as f64 / 1024.0)
    } else if bytes < 1024 * 1024 * 1024 {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    } else {
        format!("{:.1} GB", bytes as f64 / (1024.0 * 1024.0 * 1024.0))
    }
}

/// Count charts with empty numCache (ptCount but no pt elements).
fn count_empty_cache_charts(pptx_path: &Path) -> usize {
    let cache = match crate::zip_ops::chart_data::read_all_chart_cache(pptx_path) {
        Ok(m) => m,
        Err(_) => return 0,
    };
    let mut count = 0;
    for series_list in cache.values() {
        for (_ref_str, values) in series_list {
            if values.is_empty() {
                count += 1;
                break; // Count the chart once, not per-series
            }
        }
    }
    count
}

/// Print a dot-leader row with right-aligned number.
///
/// `indent=true` adds `╰` prefix for sub-items.
fn info_row(label: &str, count: usize, indent: bool) {
    let s_dim = Style::new().dim();
    let s_count = Style::new().white().bold();
    let prefix = if indent { "    ╰ " } else { "  " };
    let target_col: usize = 48;

    // Use char count (not byte count) — ╰ is 3 bytes but 1 display char
    let display_len = prefix.chars().count() + label.chars().count() + 1; // +1 for trailing space
    let leader_len = target_col.saturating_sub(display_len);
    let padded = format!("{prefix}{label} {}", "·".repeat(leader_len));

    println!("{} {:>4}",
        s_dim.apply_to(&padded),
        s_count.apply_to(count));
}

/// Print a dot-leader row with ✓/✗ status instead of a number.
fn info_row_status(label: &str, found: bool) {
    let s_dim = Style::new().dim();
    let prefix = "    ╰ ";
    let target_col: usize = 48;

    let display_len = prefix.chars().count() + label.chars().count() + 1;
    let leader_len = target_col.saturating_sub(display_len);
    let padded = format!("{prefix}{label} {}", "·".repeat(leader_len));

    let icon = if found {
        Style::new().green().apply_to("✓")
    } else {
        Style::new().red().apply_to("✗")
    };

    println!("{}    {}",
        s_dim.apply_to(&padded),
        icon);
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

/// Collect per-slide shape counts from the inventory.
fn collect_per_slide_breakdown(
    inventory: &crate::shapes::inventory::SlideInventory,
    slide_count: i32,
) -> Vec<SlideBreakdown> {
    use crate::shapes::matcher::TableType;

    let mut slides: HashMap<i32, SlideBreakdown> = HashMap::new();

    for s in 1..=slide_count {
        slides.insert(s, SlideBreakdown {
            slide: s, ole: 0, chart: 0, ntbl: 0, htmp: 0, trns: 0, delt: 0, ccst: 0,
        });
    }

    for ole in &inventory.ole_shapes {
        if let Some(entry) = slides.get_mut(&ole.slide_index) {
            entry.ole += 1;
        }
    }

    for chart in &inventory.charts {
        if let Some(entry) = slides.get_mut(&chart.slide_index) {
            entry.chart += 1;
        }
    }

    for ((slide_idx, _), table_info) in &inventory.tables {
        if let Some(entry) = slides.get_mut(slide_idx) {
            match table_info.table_type {
                TableType::Normal => entry.ntbl += 1,
                TableType::Heatmap => entry.htmp += 1,
                TableType::Transposed => entry.trns += 1,
            }
        }
    }

    for ((slide_idx, _), _) in &inventory.delts {
        if let Some(entry) = slides.get_mut(slide_idx) {
            entry.delt += 1;
        }
    }

    for ccst in &inventory.ccst_tables {
        if let Some(entry) = slides.get_mut(&ccst.slide_index) {
            entry.ccst += 1;
        }
    }

    let mut result: Vec<SlideBreakdown> = slides.into_values().collect();
    result.sort_by_key(|s| s.slide);
    result
}

/// Print the per-slide breakdown table.
fn print_per_slide_breakdown(per_slide: &[SlideBreakdown], total_slides: i32) {
    let s_dim = Style::new().dim();
    let s_count = Style::new().white().bold();

    let divider = "╌".repeat(70);

    println!();
    println!("  {}", s_dim.apply_to("Per-slide breakdown"));
    println!("  {}", s_dim.apply_to(&divider));
    println!("  {}",
        s_dim.apply_to("  slide    ole  chart   ntbl   htmp   trns   delt   ccst  total"));
    println!();

    let mut active = 0usize;
    for row in per_slide {
        let total = row.total();
        if total == 0 { continue; }
        active += 1;

        print!("  ");
        print!("{}", s_dim.apply_to(format!("{:>6}", row.slide)));
        print_cell(row.ole, &s_dim, &s_count);
        print_cell(row.chart, &s_dim, &s_count);
        print_cell(row.ntbl, &s_dim, &s_count);
        print_cell(row.htmp, &s_dim, &s_count);
        print_cell(row.trns, &s_dim, &s_count);
        print_cell(row.delt, &s_dim, &s_count);
        print_cell(row.ccst, &s_dim, &s_count);
        print!("  {}", s_count.apply_to(format!("{:>4}", total)));
        println!();
    }

    let empty = total_slides as usize - active;
    println!("  {}", s_dim.apply_to(&divider));
    println!("  {} {} {} {} {} {}",
        s_count.apply_to(total_slides),
        s_dim.apply_to("slides ·"),
        s_count.apply_to(active),
        s_dim.apply_to("active ·"),
        s_count.apply_to(empty),
        s_dim.apply_to("empty"),
    );
}

/// Print a single cell in the per-slide table.
fn print_cell(value: usize, s_dim: &Style, s_count: &Style) {
    if value == 0 {
        print!("{}", s_dim.apply_to(format!("{:>6}", "·")));
    } else {
        print!("{}", s_count.apply_to(format!("{:>6}", value)));
    }
}
