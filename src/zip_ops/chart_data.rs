//! ZIP-level chart data pre-update.
//!
//! Rewrites `<c:numCache>` values in chart XML files directly in the PPTX ZIP,
//! using fresh values read from Excel. This bypasses the extremely slow
//! `LinkFormat.Update()` COM call (~25ms/chart local, ~4s/chart network).
//!
//! GOTCHA #23: Only update `<c:val>` (value axis), NOT `<c:cat>` (category axis).
//! GOTCHA #20: Handle non-contiguous ranges (comma-separated in `<c:f>`).
//! GOTCHA #43: Values are `Option<f64>` — `None` is a blank Excel cell and produces
//! NO `<c:pt>` (PowerPoint draws nothing, no label); `Some(0.0)` is a real zero point.
//! Each series cache is rebuilt from scratch so holes at any position are filled and
//! stale points are removed (supersedes the trailing-only injection of #36/#37).

use std::collections::HashMap;
use std::io::{Read, Write};
use std::path::Path;

/// One chart value: `Some(v)` = a data point, `None` = no point (blank cell).
pub type ChartValue = Option<f64>;

/// Series data: Vec of (range_ref, cached_values) per chart.
pub type ChartSeriesData = Vec<(String, Vec<ChartValue>)>;

/// Result of chart data pre-update.
pub struct ChartDataResult {
    pub charts_updated: usize,
    pub series_updated: usize,
}

/// Scan all chart XML files in a PPTX and collect unique range references.
///
/// Returns a map of chart XML path → list of (series_index, range_ref) pairs.
/// Only includes charts with external links (checks chart .rels for TargetMode="External").
pub fn scan_chart_ranges(pptx_path: &Path) -> Result<HashMap<String, Vec<String>>, String> {
    let data = std::fs::read(pptx_path).map_err(|e| format!("Failed to read PPTX: {e}"))?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&data))
        .map_err(|e| format!("Failed to open ZIP: {e}"))?;

    let mut result: HashMap<String, Vec<String>> = HashMap::new();

    // Collect all chart XML filenames
    let chart_names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let entry = archive.by_index(i).ok()?;
            let name = entry.name().to_string();
            if name.starts_with("ppt/charts/chart") && name.ends_with(".xml") && !name.contains(".rels") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for chart_name in &chart_names {
        // Check if this chart has an external link
        let chart_filename = chart_name.rsplit('/').next().unwrap_or(chart_name);
        let rels_path = format!("ppt/charts/_rels/{chart_filename}.rels");
        if !has_external_link(&mut archive, &rels_path) {
            continue;
        }

        // Parse chart XML for series value range references
        let xml = match read_entry(&mut archive, chart_name) {
            Some(data) => data,
            None => continue,
        };

        let refs = extract_val_refs(&xml);
        if !refs.is_empty() {
            result.insert(chart_name.clone(), refs);
        }
    }

    Ok(result)
}

/// Update chart numCache values in the PPTX ZIP.
///
/// `range_values` maps normalized range ref (e.g., "Tables!C388:C390") → `Vec<ChartValue>`
/// (`None` = blank cell = no point). The PPTX is modified in-place via temp file + rename.
///
/// Returns the count of charts and series updated.
pub fn update_chart_data(
    pptx_path: &Path,
    range_values: &HashMap<String, Vec<ChartValue>>,
) -> Result<ChartDataResult, String> {
    let data = std::fs::read(pptx_path).map_err(|e| format!("Failed to read PPTX: {e}"))?;
    let mut reader = zip::ZipArchive::new(std::io::Cursor::new(&data))
        .map_err(|e| format!("Failed to open ZIP: {e}"))?;

    let tmp_path = pptx_path.with_extension("pptx.chartdata.tmp");
    let tmp_file = std::fs::File::create(&tmp_path)
        .map_err(|e| format!("Failed to create temp file: {e}"))?;
    let mut writer = zip::ZipWriter::new(tmp_file);

    let mut charts_updated = 0usize;
    let mut series_updated = 0usize;

    // Pre-collect chart names, then check external links in a separate pass
    let all_chart_names: Vec<String> = (0..reader.len())
        .filter_map(|i| {
            let entry = reader.by_index(i).ok()?;
            let name = entry.name().to_string();
            if name.starts_with("ppt/charts/chart") && name.ends_with(".xml") && !name.contains(".rels") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    let chart_names_with_ext: Vec<String> = all_chart_names.into_iter()
        .filter(|name| {
            let chart_filename = name.rsplit('/').next().unwrap_or(name);
            let rels_path = format!("ppt/charts/_rels/{chart_filename}.rels");
            has_external_link(&mut reader, &rels_path)
        })
        .collect();

    for i in 0..reader.len() {
        let mut entry = reader.by_index(i).map_err(|e| format!("ZIP entry error: {e}"))?;
        let name = entry.name().to_string();
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(entry.compression());

        if chart_names_with_ext.contains(&name) {
            // Read chart XML and rewrite numCache values
            let mut xml_data = Vec::new();
            entry.read_to_end(&mut xml_data).map_err(|e| format!("Failed to read {name}: {e}"))?;

            match rewrite_chart_cache(&xml_data, range_values) {
                Ok((modified_xml, count)) => {
                    writer.start_file(&name, options).map_err(|e| format!("ZIP write error: {e}"))?;
                    writer.write_all(&modified_xml).map_err(|e| format!("ZIP write error: {e}"))?;
                    if count > 0 {
                        charts_updated += 1;
                        series_updated += count;
                    }
                }
                Err(_) => {
                    // Failed to rewrite — keep original
                    writer.start_file(&name, options).map_err(|e| format!("ZIP write error: {e}"))?;
                    writer.write_all(&xml_data).map_err(|e| format!("ZIP write error: {e}"))?;
                }
            }
        } else {
            writer.raw_copy_file(entry).map_err(|e| format!("ZIP copy error: {e}"))?;
        }
    }

    writer.finish().map_err(|e| format!("Failed to finalize ZIP: {e}"))?;

    std::fs::rename(&tmp_path, pptx_path).map_err(|e| {
        let _ = std::fs::remove_file(&tmp_path);
        format!("Failed to replace PPTX: {e}")
    })?;

    Ok(ChartDataResult { charts_updated, series_updated })
}

/// Write `<c:pt idx="i"><c:v>val</c:v></c:pt>` for every `Some` value, ascending idx.
/// `None` values produce no element (blank cell → no point). Per-point `formatCode`
/// attributes captured from the old cache are re-attached to the same idx.
fn write_points(
    writer: &mut quick_xml::writer::Writer<Vec<u8>>,
    vals: &[ChartValue],
    format_codes: &HashMap<usize, String>,
) -> Result<(), String> {
    use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

    for (idx, val) in vals.iter().enumerate() {
        let Some(v) = val else { continue };

        let mut pt_start = BytesStart::new("c:pt");
        pt_start.push_attribute(("idx", idx.to_string().as_str()));
        if let Some(fc) = format_codes.get(&idx) {
            pt_start.push_attribute(("formatCode", fc.as_str()));
        }
        writer.write_event(Event::Start(pt_start)).map_err(|e| e.to_string())?;
        writer.write_event(Event::Start(BytesStart::new("c:v"))).map_err(|e| e.to_string())?;
        writer.write_event(Event::Text(BytesText::new(&format!("{v}")))).map_err(|e| e.to_string())?;
        writer.write_event(Event::End(BytesEnd::new("c:v"))).map_err(|e| e.to_string())?;
        writer.write_event(Event::End(BytesEnd::new("c:pt"))).map_err(|e| e.to_string())?;
    }
    Ok(())
}

/// Rewrite `<c:numCache>` values in chart XML using streaming quick-xml.
///
/// For each `<c:ser>/<c:val>/<c:numRef>` whose `<c:f>` range is in `range_values`,
/// the cache is REBUILT: existing `<c:pt>` elements are dropped, `<c:ptCount>` is set
/// to the Excel value count, and a fresh `<c:pt>` is emitted for every `Some` value in
/// ascending idx order (none for `None`). This fills holes at any position and removes
/// stale points (GOTCHA #43, superseding the trailing-only logic of #36/#37).
/// Series whose range is not in the map pass through byte-for-byte.
///
/// Returns (modified_xml, number_of_series_rebuilt).
fn rewrite_chart_cache(
    xml: &[u8],
    range_values: &HashMap<String, Vec<ChartValue>>,
) -> Result<(Vec<u8>, usize), String> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;
    use quick_xml::writer::Writer;

    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::new());

    // State machine for tracking position in XML hierarchy
    let mut in_ser = false;
    let mut in_val = false;      // inside <c:val> (NOT <c:cat> — GOTCHA #23)
    let mut in_num_ref = false;
    let mut in_num_cache = false;
    let mut in_f = false;        // inside <c:f> (formula/range ref)
    let mut in_pt = false;       // inside <c:pt>

    let mut current_range_ref = String::new();
    let mut current_values: Option<&Vec<ChartValue>> = None;
    let mut combined_values_buf: Option<Vec<ChartValue>> = None; // Buffer for non-contiguous ranges
    let mut series_updated = 0usize;

    // Rebuild state for the numCache currently being rewritten
    let mut rebuild = false;          // true while inside a numCache we own
    let mut flushed = false;          // new <c:pt>s already emitted for this cache
    let mut pt_format_codes: HashMap<usize, String> = HashMap::new();

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,

            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"ser" => { in_ser = true; }
                    b"val" if in_ser => { in_val = true; }
                    b"numRef" if in_val => { in_num_ref = true; }
                    b"f" if in_num_ref => { in_f = true; }
                    b"numCache" if in_num_ref => {
                        in_num_cache = true;
                        flushed = false;
                        pt_format_codes.clear();
                        // Look up values for current range ref.
                        // For non-contiguous ranges (GOTCHA #20), split on commas
                        // and concatenate values from each sub-range.
                        let normalized = normalize_range_ref(&current_range_ref);
                        current_values = range_values.get(&normalized);
                        if current_values.is_none() && normalized.contains(',') {
                            let mut combined = Vec::new();
                            let mut all_found = true;
                            for sub in normalized.split(',') {
                                let sub = sub.trim();
                                if let Some(vals) = range_values.get(sub) {
                                    combined.extend(vals.iter().copied());
                                } else {
                                    all_found = false;
                                    break;
                                }
                            }
                            if all_found && !combined.is_empty() {
                                combined_values_buf = Some(combined);
                                current_values = combined_values_buf.as_ref();
                            }
                        }
                        rebuild = current_values.is_some();
                    }
                    b"pt" if in_num_cache => {
                        in_pt = true;
                        if rebuild {
                            // Swallow the old point; remember its formatCode (if any) by idx
                            let idx = e.try_get_attribute("idx")
                                .ok().flatten()
                                .and_then(|a| String::from_utf8_lossy(a.value.as_ref()).parse::<usize>().ok());
                            if let (Some(idx), Ok(Some(fc))) = (idx, e.try_get_attribute("formatCode")) {
                                pt_format_codes.insert(idx, String::from_utf8_lossy(fc.value.as_ref()).to_string());
                            }
                        }
                    }
                    _ if in_num_cache && rebuild && !in_pt && !flushed => {
                        // A non-pt child after the points (e.g. <c:extLst>): emit the new
                        // points first so they keep their schema position.
                        if let Some(vals) = current_values {
                            write_points(&mut writer, vals, &pt_format_codes)?;
                        }
                        flushed = true;
                    }
                    _ => {}
                }
                if !(rebuild && in_pt) {
                    writer.write_event(Event::Start(e.clone())).map_err(|e| e.to_string())?;
                }
            }

            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                let mut skip_write = false;
                match local.as_ref() {
                    b"ser" => {
                        in_ser = false;
                        in_val = false;
                        in_num_ref = false;
                        in_num_cache = false;
                        current_range_ref.clear();
                        current_values = None;
                        drop(combined_values_buf.take());
                    }
                    b"val" => { in_val = false; in_num_ref = false; in_num_cache = false; }
                    b"numRef" => { in_num_ref = false; in_num_cache = false; }
                    b"numCache" => {
                        if rebuild {
                            if !flushed
                                && let Some(vals) = current_values {
                                    write_points(&mut writer, vals, &pt_format_codes)?;
                                }
                            flushed = true;
                            series_updated += 1;
                        }
                        rebuild = false;
                        in_num_cache = false;
                    }
                    b"f" => { in_f = false; }
                    b"pt" => {
                        skip_write = rebuild && in_pt;
                        in_pt = false;
                    }
                    b"v" => { skip_write = rebuild && in_pt; }
                    _ => {}
                }
                if !skip_write {
                    writer.write_event(Event::End(e.clone())).map_err(|e| e.to_string())?;
                }
            }

            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                if in_num_cache && rebuild {
                    if local.as_ref() == b"ptCount" {
                        // ptCount = total categories, blanks included
                        if let Some(vals) = current_values {
                            let mut elem = e.clone();
                            elem.clear_attributes();
                            elem.push_attribute(("val", vals.len().to_string().as_str()));
                            writer.write_event(Event::Empty(elem)).map_err(|e| e.to_string())?;
                            continue;
                        }
                    } else if local.as_ref() == b"pt" {
                        // Degenerate self-closing point — drop it, we rebuild all points
                        continue;
                    }
                }
                writer.write_event(Event::Empty(e.clone())).map_err(|e| e.to_string())?;
            }

            Ok(Event::Text(ref t)) => {
                if in_f && in_num_ref && in_val {
                    // Capture the range reference
                    current_range_ref = String::from_utf8_lossy(t.as_ref()).to_string();
                    writer.write_event(Event::Text(t.clone())).map_err(|e| e.to_string())?;
                } else if rebuild && in_pt {
                    // Old point value — swallowed, rebuilt from Excel
                    continue;
                } else {
                    writer.write_event(Event::Text(t.clone())).map_err(|e| e.to_string())?;
                }
            }

            Ok(event) => {
                writer.write_event(event).map_err(|e| e.to_string())?;
            }

            Err(e) => return Err(format!("XML parse error: {e}")),
        }
    }

    Ok((writer.into_inner(), series_updated))
}

/// Extract value-axis range references from chart XML.
/// Only extracts `<c:ser>/<c:val>/<c:numRef>/<c:f>` — NOT `<c:cat>` (GOTCHA #23).
fn extract_val_refs(xml: &str) -> Vec<String> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(xml.as_bytes());
    let mut refs = Vec::new();
    let mut in_ser = false;
    let mut in_val = false;
    let mut in_num_ref = false;
    let mut found_for_series = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"ser" => { in_ser = true; found_for_series = false; }
                    b"val" if in_ser => { in_val = true; }
                    b"numRef" if in_val => { in_num_ref = true; }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.local_name().as_ref() {
                    b"ser" => { in_ser = false; in_val = false; in_num_ref = false; }
                    b"val" => { in_val = false; in_num_ref = false; }
                    b"numRef" => { in_num_ref = false; }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_ser && in_val && in_num_ref && !found_for_series {
                    let text = String::from_utf8_lossy(t.as_ref()).to_string();
                    if !text.trim().is_empty() {
                        refs.push(text.trim().to_string());
                        found_for_series = true;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    refs
}

/// Normalize a range reference for HashMap lookup.
/// Strips `$` signs and outer parentheses.
fn normalize_range_ref(range_ref: &str) -> String {
    range_ref
        .trim()
        .trim_start_matches('(')
        .trim_end_matches(')')
        .replace('$', "")
}

/// Check if a chart's .rels has an external link.
fn has_external_link(archive: &mut zip::ZipArchive<std::io::Cursor<&Vec<u8>>>, rels_path: &str) -> bool {
    match read_entry(archive, rels_path) {
        Some(data) => data.contains("TargetMode=\"External\""),
        None => false,
    }
}

/// Read a ZIP entry as a string.
fn read_entry(archive: &mut zip::ZipArchive<std::io::Cursor<&Vec<u8>>>, name: &str) -> Option<String> {
    let mut entry = archive.by_name(name).ok()?;
    let mut data = String::new();
    entry.read_to_string(&mut data).ok()?;
    Some(data)
}

/// Collect all unique range references from chart scan results.
/// Normalizes refs (strip $, parens) and splits non-contiguous ranges (GOTCHA #20).
pub fn collect_unique_ranges(chart_ranges: &HashMap<String, Vec<String>>) -> Vec<String> {
    let mut unique = std::collections::HashSet::new();
    for refs in chart_ranges.values() {
        for range_ref in refs {
            let normalized = normalize_range_ref(range_ref);
            // Split non-contiguous ranges (GOTCHA #20)
            for sub in normalized.split(',') {
                let sub = sub.trim();
                if !sub.is_empty() {
                    unique.insert(sub.to_string());
                }
            }
        }
    }
    unique.into_iter().collect()
}

/// Extract cached numCache values from chart XML per series.
///
/// Returns Vec of (range_ref, Vec<ChartValue>) for each series — the PPT-side cache.
/// Only extracts from `<c:val>` (GOTCHA #23).
///
/// Shape of the returned values (GOTCHA #43):
/// - No `<c:pt>` at all (empty cache, GOTCHA #36) → empty Vec, so callers can still
///   detect "unverifiable" caches with `is_empty()`.
/// - Otherwise length = max(`ptCount`, highest idx + 1); present points are `Some`,
///   holes anywhere are `None` (never padded with 0.0).
pub fn extract_cached_values(xml: &str) -> ChartSeriesData {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(xml.as_bytes());
    let mut series_data: ChartSeriesData = Vec::new();

    let mut in_ser = false;
    let mut in_val = false;
    let mut in_num_ref = false;
    let mut in_num_cache = false;
    let mut in_f = false;
    let mut in_pt = false;
    let mut in_v = false;

    let mut current_ref = String::new();
    let mut current_values: Vec<ChartValue> = Vec::new();
    let mut current_pt_idx: usize = 0;
    let mut pt_count: usize = 0;
    let mut seen_pt = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"ser" => {
                        in_ser = true;
                        current_ref.clear();
                        current_values.clear();
                        pt_count = 0;
                        seen_pt = false;
                    }
                    b"val" if in_ser => { in_val = true; }
                    b"numRef" if in_val => { in_num_ref = true; }
                    b"f" if in_num_ref => { in_f = true; }
                    b"numCache" if in_num_ref => { in_num_cache = true; }
                    b"pt" if in_num_cache => {
                        in_pt = true;
                        seen_pt = true;
                        current_pt_idx = e.try_get_attribute("idx")
                            .ok().flatten()
                            .and_then(|a| String::from_utf8_lossy(a.value.as_ref()).parse::<usize>().ok())
                            .unwrap_or(0);
                        // Extend with holes (None) up to this index
                        while current_values.len() <= current_pt_idx {
                            current_values.push(None);
                        }
                    }
                    b"v" if in_pt => { in_v = true; }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_num_cache && e.local_name().as_ref() == b"ptCount" {
                    pt_count = e.try_get_attribute("val")
                        .ok().flatten()
                        .and_then(|a| String::from_utf8_lossy(a.value.as_ref()).parse::<usize>().ok())
                        .unwrap_or(0);
                }
            }
            Ok(Event::End(ref e)) => {
                match e.local_name().as_ref() {
                    b"ser" => {
                        if !current_ref.is_empty() {
                            let values = if seen_pt {
                                // Pad trailing holes so the length reflects the category count
                                while current_values.len() < pt_count {
                                    current_values.push(None);
                                }
                                current_values.clone()
                            } else {
                                Vec::new() // empty cache — unverifiable (GOTCHA #36)
                            };
                            series_data.push((current_ref.clone(), values));
                        }
                        in_ser = false; in_val = false; in_num_ref = false;
                        in_num_cache = false; current_ref.clear(); current_values.clear();
                    }
                    b"val" => { in_val = false; in_num_ref = false; in_num_cache = false; }
                    b"numRef" => { in_num_ref = false; in_num_cache = false; }
                    b"numCache" => { in_num_cache = false; }
                    b"f" => { in_f = false; }
                    b"pt" => { in_pt = false; }
                    b"v" => { in_v = false; }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_f && in_num_ref && in_val {
                    current_ref = String::from_utf8_lossy(t.as_ref()).trim().to_string();
                } else if in_v && in_pt && in_num_cache
                    && let Ok(val) = String::from_utf8_lossy(t.as_ref()).trim().parse::<f64>()
                        && current_pt_idx < current_values.len() {
                            current_values[current_pt_idx] = Some(val);
                        }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    series_data
}

/// Read all chart cached values from a PPTX ZIP.
///
/// Returns: HashMap<chart_xml_path → Vec<(range_ref, cached_values)>> for every chart
/// with an external link. (`oa check` uses its own slide-position keyed traversal.)
pub fn read_all_chart_cache(pptx_path: &std::path::Path) -> Result<HashMap<String, ChartSeriesData>, String> {
    let data = std::fs::read(pptx_path).map_err(|e| format!("Failed to read PPTX: {e}"))?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&data))
        .map_err(|e| format!("Failed to open ZIP: {e}"))?;

    let mut result: HashMap<String, ChartSeriesData> = HashMap::new();

    let chart_names: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let entry = archive.by_index(i).ok()?;
            let name = entry.name().to_string();
            if name.starts_with("ppt/charts/chart") && name.ends_with(".xml") && !name.contains(".rels") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for chart_name in &chart_names {
        let chart_filename = chart_name.rsplit('/').next().unwrap_or(chart_name);
        let rels_path = format!("ppt/charts/_rels/{chart_filename}.rels");
        if !has_external_link(&mut archive, &rels_path) {
            continue;
        }

        if let Some(xml) = read_entry(&mut archive, chart_name) {
            let cached = extract_cached_values(&xml);
            if !cached.is_empty() {
                result.insert(chart_name.clone(), cached);
            }
        }
    }

    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_normalize_range_ref() {
        assert_eq!(normalize_range_ref("Tables!$B$388:$B$390"), "Tables!B388:B390");
        assert_eq!(normalize_range_ref("(Tables!$C$810,Tables!$F$810)"), "Tables!C810,Tables!F810");
    }

    #[test]
    fn test_collect_unique_ranges() {
        let mut chart_ranges = HashMap::new();
        chart_ranges.insert("chart1.xml".to_string(), vec![
            "Tables!$B$388:$B$390".to_string(),
            "(Tables!$C$810,Tables!$F$810)".to_string(),
        ]);
        let unique = collect_unique_ranges(&chart_ranges);
        assert!(unique.contains(&"Tables!B388:B390".to_string()));
        assert!(unique.contains(&"Tables!C810".to_string()));
        assert!(unique.contains(&"Tables!F810".to_string()));
    }

    #[test]
    fn test_extract_val_refs() {
        let xml = r#"<?xml version="1.0"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:barChart>
        <c:ser><c:val><c:numRef><c:f>Tables!$B$1:$B$3</c:f>
        <c:numCache><c:ptCount val="3"/>
        <c:pt idx="0"><c:v>1.0</c:v></c:pt>
        </c:numCache></c:numRef></c:val></c:ser>
        <c:ser><c:cat><c:strRef><c:f>Tables!$A$1:$A$3</c:f>
        </c:strRef></c:cat>
        <c:val><c:numRef><c:f>Tables!$C$1:$C$3</c:f>
        <c:numCache><c:ptCount val="3"/>
        </c:numCache></c:numRef></c:val></c:ser>
        </c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

        let refs = extract_val_refs(xml);
        assert_eq!(refs.len(), 2);
        assert_eq!(refs[0], "Tables!$B$1:$B$3");
        assert_eq!(refs[1], "Tables!$C$1:$C$3");
        // Category ref (Tables!$A$1:$A$3) should NOT be included (GOTCHA #23)
    }

    #[test]
    fn test_rewrite_chart_cache() {
        let xml = br#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart><c:plotArea><c:barChart>
<c:ser><c:val><c:numRef><c:f>Tables!$B$1:$B$3</c:f>
<c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="3"/>
<c:pt idx="0"><c:v>0.1</c:v></c:pt>
<c:pt idx="1"><c:v>0.2</c:v></c:pt>
<c:pt idx="2"><c:v>0.3</c:v></c:pt>
</c:numCache></c:numRef></c:val></c:ser>
</c:barChart></c:plotArea></c:chart></c:chartSpace>"#;

        let mut values = HashMap::new();
        values.insert("Tables!B1:B3".to_string(), vec![Some(0.5), Some(0.6), Some(0.7)]);

        let (output, count) = rewrite_chart_cache(xml, &values).unwrap();
        let output_str = String::from_utf8(output).unwrap();
        assert_eq!(count, 1);
        assert!(output_str.contains("0.5"));
        assert!(output_str.contains("0.6"));
        assert!(output_str.contains("0.7"));
        assert!(!output_str.contains("0.1"));
        assert!(!output_str.contains("0.2"));
        // formatCode child preserved, points in order
        assert!(output_str.contains("<c:formatCode>0%</c:formatCode>"));
        assert_eq!(pts_of(&output_str), vec![(0, "0.5".into()), (1, "0.6".into()), (2, "0.7".into())]);
    }

    // ── GOTCHA #43 helpers ──────────────────────────────────

    /// Wrap a numCache body in a minimal one-series chart referencing Tables!$B$1:$B$N.
    fn chart_with_cache(cache_body: &str) -> String {
        format!(r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart><c:plotArea><c:barChart>
<c:ser><c:cat><c:strRef><c:f>Tables!$A$1:$A$4</c:f><c:strCache><c:ptCount val="4"/>
<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt>
<c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Tables!$B$1:$B$4</c:f>
<c:numCache>{cache_body}</c:numCache></c:numRef></c:val></c:ser>
</c:barChart></c:plotArea></c:chart></c:chartSpace>"#)
    }

    /// (idx, value-text) of every <c:pt> inside the FIRST <c:val> numCache.
    fn pts_of(xml: &str) -> Vec<(usize, String)> {
        let val_start = xml.find("<c:val>").expect("no <c:val>");
        let val = &xml[val_start..xml.find("</c:val>").expect("no </c:val>")];
        let mut out = Vec::new();
        let mut rest = val;
        while let Some(p) = rest.find("<c:pt idx=\"") {
            let after = &rest[p + 11..];
            let idx_end = after.find('"').unwrap();
            let idx: usize = after[..idx_end].parse().unwrap();
            let v_start = after.find("<c:v>").unwrap() + 5;
            let v_end = after.find("</c:v>").unwrap();
            out.push((idx, after[v_start..v_end].to_string()));
            rest = &after[v_end..];
        }
        out
    }

    fn pt_count_of(xml: &str) -> usize {
        let val_start = xml.find("<c:val>").unwrap();
        let val = &xml[val_start..];
        let p = val.find("<c:ptCount val=\"").unwrap() + 16;
        val[p..p + val[p..].find('"').unwrap()].parse().unwrap()
    }

    fn vals(range: &str, v: Vec<ChartValue>) -> HashMap<String, Vec<ChartValue>> {
        let mut m = HashMap::new();
        m.insert(range.to_string(), v);
        m
    }

    // ── rewrite: holes anywhere ─────────────────────────────

    #[test]
    fn test_rewrite_fills_middle_gap() {
        // Template cache has pts 0,1,3 (idx 2 was a blank cell) — the Chart 33 case
        let xml = chart_with_cache(r#"<c:formatCode>0%</c:formatCode><c:ptCount val="4"/>
<c:pt idx="0"><c:v>0.05</c:v></c:pt><c:pt idx="1"><c:v>0.02</c:v></c:pt><c:pt idx="3"><c:v>0.13</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.08), Some(0.1), Some(0.18), Some(0.3)]);
        let (out, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(count, 1);
        assert_eq!(pt_count_of(&s), 4);
        assert_eq!(pts_of(&s), vec![
            (0, "0.08".into()), (1, "0.1".into()), (2, "0.18".into()), (3, "0.3".into()),
        ]);
        // Category cache untouched (GOTCHA #23)
        assert!(s.contains("<c:v>C</c:v>"));
    }

    #[test]
    fn test_rewrite_fills_leading_gap() {
        // Template cache has pts 1,2 (idx 0 blank) — the Chart 39 case
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>
<c:pt idx="1"><c:v>0.07</c:v></c:pt><c:pt idx="2"><c:v>0.12</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.25), Some(0.45), Some(0.22)]);
        let (out, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(count, 1);
        assert_eq!(pt_count_of(&s), 3);
        assert_eq!(pts_of(&s), vec![(0, "0.25".into()), (1, "0.45".into()), (2, "0.22".into())]);
    }

    #[test]
    fn test_rewrite_fills_trailing_gap_gotcha_37() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>
<c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.2</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.4), Some(0.5), Some(0.6)]);
        let (out, _) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(pts_of(&s), vec![(0, "0.4".into()), (1, "0.5".into()), (2, "0.6".into())]);
    }

    #[test]
    fn test_rewrite_fills_empty_cache_gotcha_36() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.4), Some(0.5), Some(0.6)]);
        let (out, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(count, 1);
        assert_eq!(pts_of(&s), vec![(0, "0.4".into()), (1, "0.5".into()), (2, "0.6".into())]);
    }

    // ── rewrite: blanks remove points, zero keeps them ─────

    #[test]
    fn test_rewrite_blank_removes_point_keeps_ptcount() {
        // Full template cache; new Excel has a blank in the middle (France → Indonesia)
        let xml = chart_with_cache(r#"<c:ptCount val="4"/>
<c:pt idx="0"><c:v>0.08</c:v></c:pt><c:pt idx="1"><c:v>0.1</c:v></c:pt>
<c:pt idx="2"><c:v>0.18</c:v></c:pt><c:pt idx="3"><c:v>0.3</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.05), Some(0.02), None, Some(0.13)]);
        let (out, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(count, 1);
        assert_eq!(pt_count_of(&s), 4, "ptCount stays the category count");
        assert_eq!(pts_of(&s), vec![(0, "0.05".into()), (1, "0.02".into()), (3, "0.13".into())]);
        assert!(!s.contains("<c:v>0</c:v>"), "a blank must not become a zero point");
    }

    #[test]
    fn test_rewrite_leading_blank_removes_point() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>
<c:pt idx="0"><c:v>0.25</c:v></c:pt><c:pt idx="1"><c:v>0.45</c:v></c:pt><c:pt idx="2"><c:v>0.22</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![None, Some(0.07), Some(0.12)]);
        let (out, _) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(pts_of(&s), vec![(1, "0.07".into()), (2, "0.12".into())]);
    }

    #[test]
    fn test_rewrite_real_zero_is_a_point() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>
<c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.2</c:v></c:pt><c:pt idx="2"><c:v>0.3</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.1), Some(0.0), Some(0.3)]);
        let (out, _) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(pts_of(&s), vec![(0, "0.1".into()), (1, "0".into()), (2, "0.3".into())]);
    }

    #[test]
    fn test_rewrite_all_blank_leaves_no_points() {
        let xml = chart_with_cache(r#"<c:ptCount val="2"/>
<c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.2</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![None, None]);
        let (out, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(count, 1);
        assert_eq!(pt_count_of(&s), 2);
        assert!(pts_of(&s).is_empty());
    }

    // ── rewrite: structure preservation ────────────────────

    #[test]
    fn test_rewrite_preserves_per_point_format_code() {
        let xml = chart_with_cache(r#"<c:ptCount val="2"/>
<c:pt idx="0" formatCode="0.0%"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.2</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.3), Some(0.4)]);
        let (out, _) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"<c:pt idx="0" formatCode="0.0%"><c:v>0.3</c:v></c:pt>"#));
        assert!(s.contains(r#"<c:pt idx="1"><c:v>0.4</c:v></c:pt>"#));
    }

    #[test]
    fn test_rewrite_keeps_extlst_after_points() {
        let xml = chart_with_cache(r#"<c:ptCount val="2"/>
<c:pt idx="0"><c:v>0.1</c:v></c:pt><c:extLst><c:ext uri="x"/></c:extLst>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.3), Some(0.4)]);
        let (out, _) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        let last_pt = s.rfind("</c:pt>").unwrap();
        let ext = s.find("<c:extLst>").unwrap();
        assert!(last_pt < ext, "injected points must precede <c:extLst>");
        assert_eq!(pts_of(&s), vec![(0, "0.3".into()), (1, "0.4".into())]);
        assert_eq!(s.matches("<c:extLst>").count(), 1);
    }

    #[test]
    fn test_rewrite_unknown_range_passes_through_unchanged() {
        let xml = chart_with_cache(r#"<c:ptCount val="4"/>
<c:pt idx="0"><c:v>0.05</c:v></c:pt><c:pt idx="3"><c:v>0.13</c:v></c:pt>"#);
        let m = vals("Tables!Z1:Z4", vec![Some(1.0)]);
        let (out, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        assert_eq!(count, 0);
        assert_eq!(String::from_utf8(out).unwrap(), xml);
    }

    #[test]
    fn test_rewrite_non_contiguous_with_blank_gotcha_20() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart><c:plotArea><c:barChart>
<c:ser><c:val><c:numRef><c:f>(Tables!$C$10,Tables!$F$10)</c:f>
<c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.2</c:v></c:pt></c:numCache>
</c:numRef></c:val></c:ser>
</c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let mut m = HashMap::new();
        m.insert("Tables!C10".to_string(), vec![Some(0.7)]);
        m.insert("Tables!F10".to_string(), vec![None]);
        let (out, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(count, 1);
        assert_eq!(pt_count_of(&s), 2);
        assert_eq!(pts_of(&s), vec![(0, "0.7".into())]);
    }

    #[test]
    fn test_rewrite_counts_one_per_series() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart><c:plotArea><c:barChart>
<c:ser><c:val><c:numRef><c:f>Tables!$B$1:$B$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
<c:ser><c:val><c:numRef><c:f>Tables!$C$1:$C$2</c:f><c:numCache><c:ptCount val="2"/></c:numCache></c:numRef></c:val></c:ser>
<c:ser><c:val><c:numRef><c:f>Tables!$D$1:$D$2</c:f><c:numCache><c:ptCount val="2"/></c:numCache></c:numRef></c:val></c:ser>
</c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let mut m = HashMap::new();
        m.insert("Tables!B1:B2".to_string(), vec![Some(1.0), Some(2.0)]);
        m.insert("Tables!C1:C2".to_string(), vec![Some(3.0), None]);
        // D not in map → untouched
        let (_, count) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        assert_eq!(count, 2);
    }

    // ── extract_cached_values ──────────────────────────────

    #[test]
    fn test_extract_full_cache() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>
<c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.2</c:v></c:pt><c:pt idx="2"><c:v>0.3</c:v></c:pt>"#);
        let got = extract_cached_values(&xml);
        assert_eq!(got.len(), 1);
        assert_eq!(got[0].0, "Tables!$B$1:$B$4");
        assert_eq!(got[0].1, vec![Some(0.1), Some(0.2), Some(0.3)]);
    }

    #[test]
    fn test_extract_middle_hole_is_none() {
        let xml = chart_with_cache(r#"<c:ptCount val="4"/>
<c:pt idx="0"><c:v>0.05</c:v></c:pt><c:pt idx="1"><c:v>0.02</c:v></c:pt><c:pt idx="3"><c:v>0.13</c:v></c:pt>"#);
        let got = extract_cached_values(&xml);
        assert_eq!(got[0].1, vec![Some(0.05), Some(0.02), None, Some(0.13)]);
    }

    #[test]
    fn test_extract_trailing_hole_padded_from_ptcount() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>
<c:pt idx="0"><c:v>0.1</c:v></c:pt><c:pt idx="1"><c:v>0.2</c:v></c:pt>"#);
        let got = extract_cached_values(&xml);
        assert_eq!(got[0].1, vec![Some(0.1), Some(0.2), None]);
    }

    #[test]
    fn test_extract_leading_hole_is_none() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>
<c:pt idx="1"><c:v>0.07</c:v></c:pt><c:pt idx="2"><c:v>0.12</c:v></c:pt>"#);
        let got = extract_cached_values(&xml);
        assert_eq!(got[0].1, vec![None, Some(0.07), Some(0.12)]);
    }

    #[test]
    fn test_extract_empty_cache_is_empty_vec() {
        let xml = chart_with_cache(r#"<c:ptCount val="3"/>"#);
        let got = extract_cached_values(&xml);
        assert_eq!(got.len(), 1);
        assert!(got[0].1.is_empty(), "no <c:pt> → empty (GOTCHA #36 semantics kept)");
    }

    #[test]
    fn test_extract_ignores_category_cache() {
        let xml = chart_with_cache(r#"<c:ptCount val="1"/><c:pt idx="0"><c:v>0.5</c:v></c:pt>"#);
        let got = extract_cached_values(&xml);
        assert_eq!(got.len(), 1);
        assert_eq!(got[0].1, vec![Some(0.5)]);
    }

    #[test]
    fn test_roundtrip_rewrite_then_extract() {
        let xml = chart_with_cache(r#"<c:ptCount val="4"/><c:pt idx="0"><c:v>0.9</c:v></c:pt>"#);
        let excel = vec![Some(0.08), None, Some(0.0), Some(0.3)];
        let m = vals("Tables!B1:B4", excel.clone());
        let (out, _) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let got = extract_cached_values(&String::from_utf8(out).unwrap());
        assert_eq!(got[0].1, excel, "what we write is exactly what we read back");
    }
}
