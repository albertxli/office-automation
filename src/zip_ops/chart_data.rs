//! ZIP-level chart data pre-update.
//!
//! Rewrites `<c:numCache>` values in chart XML files directly in the PPTX ZIP,
//! using fresh values read from Excel. This bypasses the extremely slow
//! `LinkFormat.Update()` COM call (~25ms/chart local, ~4s/chart network).
//!
//! GOTCHA #23: Only update `<c:val>` (value axis), NOT `<c:cat>` (category axis).
//! GOTCHA #20: Handle non-contiguous ranges (comma-separated in `<c:f>`).

use std::collections::HashMap;
use std::io::{Read, Write};
use std::path::Path;

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
/// `range_values` maps normalized range ref (e.g., "Tables!C388:C390") → Vec<f64>.
/// The PPTX is modified in-place via temp file + rename.
///
/// Returns the count of charts and series updated.
pub fn update_chart_data(
    pptx_path: &Path,
    range_values: &HashMap<String, Vec<f64>>,
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

/// Rewrite `<c:numCache>` values in chart XML using streaming quick-xml.
///
/// For each `<c:ser>/<c:val>/<c:numRef>`:
/// 1. Read `<c:f>` to get the range reference
/// 2. Look up values in `range_values` map
/// 3. Rewrite `<c:pt idx="N"><c:v>VALUE</c:v></c:pt>` elements
///
/// Returns (modified_xml, series_count_updated).
fn rewrite_chart_cache(
    xml: &[u8],
    range_values: &HashMap<String, Vec<f64>>,
) -> Result<(Vec<u8>, usize), String> {
    use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
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
    let mut in_v = false;        // inside <c:v>

    let mut current_range_ref = String::new();
    let mut current_values: Option<&Vec<f64>> = None;
    let mut combined_values_buf: Option<Vec<f64>> = None; // Buffer for non-contiguous ranges
    let mut current_pt_idx: usize = 0;
    let mut series_updated = 0usize;
    let mut seen_pt_in_cache = false; // Track if any <c:pt> exists in current numCache
    let mut max_pt_idx_seen: usize = 0; // Track highest pt idx written (for partial cache)

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
                        seen_pt_in_cache = false;
                        max_pt_idx_seen = 0;
                        // Look up values for current range ref.
                        // For non-contiguous ranges (GOTCHA #20), split on commas
                        // and concatenate values from each sub-range.
                        let normalized = normalize_range_ref(&current_range_ref);
                        current_values = range_values.get(&normalized);
                        if current_values.is_none() && normalized.contains(',') {
                            // Non-contiguous: build concatenated values
                            let mut combined = Vec::new();
                            let mut all_found = true;
                            for sub in normalized.split(',') {
                                let sub = sub.trim();
                                if let Some(vals) = range_values.get(sub) {
                                    combined.extend(vals);
                                } else {
                                    all_found = false;
                                    break;
                                }
                            }
                            if all_found && !combined.is_empty() {
                                // Store in a side buffer so we can reference it
                                combined_values_buf = Some(combined);
                                current_values = combined_values_buf.as_ref();
                            }
                        }
                    }
                    b"pt" if in_num_cache => {
                        in_pt = true;
                        seen_pt_in_cache = true;
                        // Get idx attribute (also track max for partial cache detection)
                        current_pt_idx = e.try_get_attribute("idx")
                            .ok().flatten()
                            .and_then(|a| String::from_utf8_lossy(a.value.as_ref()).parse::<usize>().ok())
                            .unwrap_or(0);
                        if current_pt_idx + 1 > max_pt_idx_seen {
                            max_pt_idx_seen = current_pt_idx + 1;
                        }
                    }
                    b"v" if in_pt => { in_v = true; }
                    _ => {}
                }
                writer.write_event(Event::Start(e.clone())).map_err(|e| e.to_string())?;
            }

            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"ser" => {
                        in_ser = false;
                        in_val = false;
                        in_num_ref = false;
                        in_num_cache = false;
                        current_range_ref.clear();
                        current_values = None;
                        combined_values_buf = None;
                    }
                    b"val" => { in_val = false; in_num_ref = false; in_num_cache = false; }
                    b"numRef" => { in_num_ref = false; in_num_cache = false; }
                    b"numCache" => {
                        // Inject missing <c:pt> elements:
                        // - Empty cache (!seen_pt_in_cache): inject ALL values
                        // - Partial cache (seen some pt but fewer than values): inject remaining
                        if let Some(vals) = current_values {
                            let start_idx = if !seen_pt_in_cache { 0 } else { max_pt_idx_seen };
                            if start_idx < vals.len() {
                                for idx in start_idx..vals.len() {
                                    let mut pt_start = BytesStart::new("c:pt");
                                    pt_start.push_attribute(("idx", idx.to_string().as_str()));
                                    writer.write_event(Event::Start(pt_start)).map_err(|e| e.to_string())?;

                                    writer.write_event(Event::Start(BytesStart::new("c:v"))).map_err(|e| e.to_string())?;
                                    writer.write_event(Event::Text(BytesText::new(&format!("{}", vals[idx])))).map_err(|e| e.to_string())?;
                                    writer.write_event(Event::End(BytesEnd::new("c:v"))).map_err(|e| e.to_string())?;

                                    writer.write_event(Event::End(BytesEnd::new("c:pt"))).map_err(|e| e.to_string())?;
                                }
                                series_updated += 1;
                            }
                        }
                        in_num_cache = false;
                    }
                    b"f" => { in_f = false; }
                    b"pt" => { in_pt = false; }
                    b"v" => { in_v = false; }
                    _ => {}
                }
                writer.write_event(Event::End(e.clone())).map_err(|e| e.to_string())?;
            }

            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"ptCount" && in_num_cache {
                    // Update ptCount if we have values
                    if let Some(vals) = current_values {
                        let mut elem = e.clone();
                        elem.clear_attributes();
                        elem.push_attribute(("val", vals.len().to_string().as_str()));
                        writer.write_event(Event::Empty(elem)).map_err(|e| e.to_string())?;
                        series_updated += 1;
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
                } else if in_v && in_pt && in_num_cache {
                    // Replace the value if we have data
                    if let Some(vals) = current_values {
                        if current_pt_idx < vals.len() {
                            let new_val = format!("{}", vals[current_pt_idx]);
                            let text = BytesText::new(&new_val);
                            writer.write_event(Event::Text(text)).map_err(|e| e.to_string())?;
                            continue;
                        }
                    }
                    // No replacement — keep original
                    writer.write_event(Event::Text(t.clone())).map_err(|e| e.to_string())?;
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
/// Returns Vec of (range_ref, Vec<f64>) for each series — the PPT-side cached values.
/// Only extracts from `<c:val>` (GOTCHA #23).
pub fn extract_cached_values(xml: &str) -> Vec<(String, Vec<f64>)> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(xml.as_bytes());
    let mut series_data: Vec<(String, Vec<f64>)> = Vec::new();

    let mut in_ser = false;
    let mut in_val = false;
    let mut in_num_ref = false;
    let mut in_num_cache = false;
    let mut in_f = false;
    let mut in_pt = false;
    let mut in_v = false;

    let mut current_ref = String::new();
    let mut current_values: Vec<f64> = Vec::new();
    let mut current_pt_idx: usize = 0;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"ser" => { in_ser = true; current_ref.clear(); current_values.clear(); }
                    b"val" if in_ser => { in_val = true; }
                    b"numRef" if in_val => { in_num_ref = true; }
                    b"f" if in_num_ref => { in_f = true; }
                    b"numCache" if in_num_ref => { in_num_cache = true; }
                    b"pt" if in_num_cache => {
                        in_pt = true;
                        current_pt_idx = e.try_get_attribute("idx")
                            .ok().flatten()
                            .and_then(|a| String::from_utf8_lossy(a.value.as_ref()).parse::<usize>().ok())
                            .unwrap_or(0);
                        // Extend values vec to fit this index
                        while current_values.len() <= current_pt_idx {
                            current_values.push(0.0);
                        }
                    }
                    b"v" if in_pt => { in_v = true; }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                match e.local_name().as_ref() {
                    b"ser" => {
                        if !current_ref.is_empty() {
                            series_data.push((current_ref.clone(), current_values.clone()));
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
                } else if in_v && in_pt && in_num_cache {
                    if let Ok(val) = String::from_utf8_lossy(t.as_ref()).trim().parse::<f64>() {
                        if current_pt_idx < current_values.len() {
                            current_values[current_pt_idx] = val;
                        }
                    }
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
/// Returns: HashMap<(slide_num, chart_position) → Vec<(range_ref, cached_values)>>
/// This matches the same key scheme as `build_chart_ref_map` in check.rs.
pub fn read_all_chart_cache(pptx_path: &std::path::Path) -> Result<HashMap<String, Vec<(String, Vec<f64>)>>, String> {
    let data = std::fs::read(pptx_path).map_err(|e| format!("Failed to read PPTX: {e}"))?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&data))
        .map_err(|e| format!("Failed to open ZIP: {e}"))?;

    let mut result: HashMap<String, Vec<(String, Vec<f64>)>> = HashMap::new();

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
        values.insert("Tables!B1:B3".to_string(), vec![0.5, 0.6, 0.7]);

        let (output, count) = rewrite_chart_cache(xml, &values).unwrap();
        let output_str = String::from_utf8(output).unwrap();
        assert_eq!(count, 1);
        assert!(output_str.contains("0.5"));
        assert!(output_str.contains("0.6"));
        assert!(output_str.contains("0.7"));
        assert!(!output_str.contains("0.1"));
        assert!(!output_str.contains("0.2"));
    }
}
