//! ZIP-level chart data pre-update.
//!
//! Rewrites `<c:numCache>` values in chart XML files directly in the PPTX ZIP,
//! using fresh values read from Excel. This bypasses the extremely slow
//! `LinkFormat.Update()` COM call (~25ms/chart local, ~4s/chart network).
//!
//! GOTCHA #48: EVERY series reference is rebuilt — values (`val`), category labels (`cat`),
//! series names (`tx`), scatter/bubble data (`xVal`/`yVal`/`bubbleSize`) — so the result
//! matches PowerPoint's own "refresh link". `strRef` caches are rebuilt from display text,
//! `numRef` caches from numbers. Multi-level categories and data-labels-from-cells are
//! left to PowerPoint (`needs_com`).
//! GOTCHA #23: a category range shared by every series is read once, never double-counted.
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

/// A chart whose `<c:f>` formulas carried a `[workbook]` qualifier that was removed (GOTCHA #44).
#[derive(Debug, Clone, PartialEq)]
pub struct FixedChart {
    /// ZIP part name, e.g. `ppt/charts/chart99.xml`.
    pub part: String,
    /// The workbook name that was stripped, e.g. `rpm_2025_Indonesia_v5.xlsx`.
    pub book: String,
    /// Number of `<c:f>` elements rewritten in this chart.
    pub formulas: usize,
}

/// Result of chart data pre-update.
#[derive(Debug, Default)]
pub struct ChartDataResult {
    pub charts_updated: usize,
    /// Value caches (`<c:val>`) rebuilt.
    pub series_updated: usize,
    /// Label caches rebuilt: categories, series names, xVal/yVal/bubbleSize (GOTCHA #48).
    pub labels_updated: usize,
    /// Charts whose formulas were normalised (GOTCHA #44) — reported as warnings.
    pub fixed: Vec<FixedChart>,
    /// Chart parts the ZIP rewrite could not fully rebuild (multi-level categories, data
    /// labels from cells) — must be refreshed by PowerPoint's `LinkFormat.Update()`.
    pub needs_refresh: Vec<String>,
    /// False when at least one range could not be read from Excel; the caller then
    /// keeps the COM `Update()` fallback for all charts.
    pub all_ranges_ok: bool,
}

/// Per-chart statistics from `rewrite_chart_cache`.
#[derive(Debug, Default, PartialEq)]
pub struct RewriteStats {
    /// `<c:val>` caches rebuilt.
    pub series_updated: usize,
    /// Other caches rebuilt (tx / cat / xVal / yVal / bubbleSize).
    pub labels_updated: usize,
    pub formulas_fixed: usize,
    /// Workbook name stripped from `<c:f>` (first one seen), if any.
    pub book: Option<String>,
    /// Chart holds a construct left to PowerPoint (multiLvlStrRef, datalabelsRange).
    pub needs_com: bool,
}

/// Which cache a reference feeds: numbers (`numRef`/`numCache`) or text (`strRef`/`strCache`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RefKind {
    Num,
    Str,
}

/// The series element a reference belongs to (GOTCHA #48).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SeriesElem {
    Tx,
    Cat,
    Val,
    XVal,
    YVal,
    BubbleSize,
}

impl SeriesElem {
    fn from_local(name: &[u8]) -> Option<Self> {
        match name {
            b"tx" => Some(Self::Tx),
            b"cat" => Some(Self::Cat),
            b"val" => Some(Self::Val),
            b"xVal" => Some(Self::XVal),
            b"yVal" => Some(Self::YVal),
            b"bubbleSize" => Some(Self::BubbleSize),
            _ => None,
        }
    }
}

/// One cell reference inside a `<c:ser>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SeriesRef {
    pub elem: SeriesElem,
    /// Raw formula text as found in `<c:f>` (may still carry `$` and `[workbook]`).
    pub formula: String,
    pub kind: RefKind,
}

/// Normalised range + kind: the key of the Excel read map.
pub type RangeKey = (String, RefKind);

/// Values read from Excel for one reference, shaped by its kind.
/// `None` entries are blank cells and produce no `<c:pt>`.
#[derive(Debug, Clone, PartialEq)]
pub enum CacheValues {
    Num(Vec<ChartValue>),
    Str(Vec<Option<String>>),
}

impl CacheValues {
    pub fn len(&self) -> usize {
        match self {
            CacheValues::Num(v) => v.len(),
            CacheValues::Str(v) => v.len(),
        }
    }
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
    fn empty_of(kind: RefKind) -> Self {
        match kind {
            RefKind::Num => CacheValues::Num(Vec::new()),
            RefKind::Str => CacheValues::Str(Vec::new()),
        }
    }
    fn extend_from(&mut self, other: &CacheValues) {
        match (self, other) {
            (CacheValues::Num(a), CacheValues::Num(b)) => a.extend(b.iter().copied()),
            (CacheValues::Str(a), CacheValues::Str(b)) => a.extend(b.iter().cloned()),
            _ => {}
        }
    }
}

/// Result of scanning the chart parts of a PPTX.
#[derive(Debug, Default)]
pub struct ChartScan {
    /// chart part → every series reference (tx / cat / val / xVal / yVal / bubbleSize)
    pub refs: HashMap<String, Vec<SeriesRef>>,
    /// chart parts with constructs the ZIP rewrite cannot rebuild (multi-level categories,
    /// data labels from cells) — refreshed by PowerPoint instead
    pub needs_com: Vec<String>,
}

/// Scan every externally linked chart part and collect ALL series references
/// (tx / cat / val / xVal / yVal / bubbleSize) with their kind (GOTCHA #48).
///
/// Charts containing constructs the ZIP rewrite cannot rebuild (multi-level categories,
/// data labels from cells) are listed in `needs_com` and left for `LinkFormat.Update()`.
pub fn scan_chart_ranges(pptx_path: &Path) -> Result<ChartScan, String> {
    let data = std::fs::read(pptx_path).map_err(|e| format!("Failed to read PPTX: {e}"))?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&data))
        .map_err(|e| format!("Failed to open ZIP: {e}"))?;

    let mut scan = ChartScan::default();

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

        let xml = match read_entry(&mut archive, chart_name) {
            Some(data) => data,
            None => continue,
        };

        let (refs, needs_com) = extract_series_refs_all(&xml);
        if needs_com {
            scan.needs_com.push(chart_name.clone());
        }
        if !refs.is_empty() {
            scan.refs.insert(chart_name.clone(), refs);
        }
    }

    Ok(scan)
}

/// Update every series cache (values, categories, series names, …) in the PPTX ZIP.
///
/// `range_values` maps `(normalized range, kind)` (e.g. `("Tables!C388:C390", Num)`) →
/// the Excel cells (`None` = blank = no point). The PPTX is modified in-place via temp
/// file + rename.
pub fn update_chart_data(
    pptx_path: &Path,
    range_values: &HashMap<RangeKey, CacheValues>,
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
    let mut labels_updated = 0usize;
    let mut fixed: Vec<FixedChart> = Vec::new();
    let mut needs_refresh: Vec<String> = Vec::new();

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
                Ok((modified_xml, stats)) => {
                    writer.start_file(&name, options).map_err(|e| format!("ZIP write error: {e}"))?;
                    writer.write_all(&modified_xml).map_err(|e| format!("ZIP write error: {e}"))?;
                    if stats.series_updated > 0 || stats.labels_updated > 0 {
                        charts_updated += 1;
                        series_updated += stats.series_updated;
                        labels_updated += stats.labels_updated;
                    }
                    if stats.needs_com {
                        needs_refresh.push(name.clone());
                    }
                    if stats.formulas_fixed > 0 {
                        fixed.push(FixedChart {
                            part: name.clone(),
                            book: stats.book.unwrap_or_default(),
                            formulas: stats.formulas_fixed,
                        });
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

    Ok(ChartDataResult {
        charts_updated, series_updated, labels_updated, fixed, needs_refresh, all_ranges_ok: true,
    })
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

/// Write `<c:pt idx="i"><c:v>text</c:v></c:pt>` for every `Some` string (strCache).
/// Blank cells (`None`) produce no element, exactly like numeric blanks.
fn write_str_points(
    writer: &mut quick_xml::writer::Writer<Vec<u8>>,
    vals: &[Option<String>],
) -> Result<(), String> {
    use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};

    for (idx, val) in vals.iter().enumerate() {
        let Some(s) = val else { continue };
        let mut pt_start = BytesStart::new("c:pt");
        pt_start.push_attribute(("idx", idx.to_string().as_str()));
        writer.write_event(Event::Start(pt_start)).map_err(|e| e.to_string())?;
        writer.write_event(Event::Start(BytesStart::new("c:v"))).map_err(|e| e.to_string())?;
        writer.write_event(Event::Text(BytesText::new(s))).map_err(|e| e.to_string())?;
        writer.write_event(Event::End(BytesEnd::new("c:v"))).map_err(|e| e.to_string())?;
        writer.write_event(Event::End(BytesEnd::new("c:pt"))).map_err(|e| e.to_string())?;
    }
    Ok(())
}

/// Values for `raw_ref` of `kind`, concatenating comma-separated sub-ranges (GOTCHA #20).
fn lookup_values(
    range_values: &HashMap<RangeKey, CacheValues>,
    raw_ref: &str,
    kind: RefKind,
) -> Option<CacheValues> {
    let normalized = normalize_range_ref(raw_ref);
    if let Some(v) = range_values.get(&(normalized.clone(), kind)) {
        return Some(v.clone());
    }
    if normalized.contains(',') {
        let mut combined = CacheValues::empty_of(kind);
        for sub in normalized.split(',') {
            let part = range_values.get(&(sub.trim().to_string(), kind))?;
            combined.extend_from(part);
        }
        if !combined.is_empty() {
            return Some(combined);
        }
    }
    None
}

/// Rewrite every series cache in a chart part using streaming quick-xml.
///
/// For each `<c:ser>` element (`tx`, `cat`, `val`, `xVal`, `yVal`, `bubbleSize`) whose
/// reference is in `range_values`, the cache is REBUILT: existing `<c:pt>` elements are
/// dropped, `<c:ptCount>` is set to the Excel cell count, and a fresh `<c:pt>` is emitted
/// for every `Some` value in ascending idx order (none for `None`). Numeric caches keep
/// per-point `formatCode` attributes (GOTCHA #43); string caches get plain text points
/// (GOTCHA #48). Series whose reference is not in the map pass through byte-for-byte.
/// `multiLvlStrRef` and `c15:datalabelsRange` are left untouched and flagged in
/// `stats.needs_com` so the chart is refreshed by PowerPoint instead.
///
/// GOTCHA #44: every `<c:f>` that carries a `[workbook]` qualifier is rewritten without it.
///
/// Returns (modified_xml, stats).
fn rewrite_chart_cache(
    xml: &[u8],
    range_values: &HashMap<RangeKey, CacheValues>,
) -> Result<(Vec<u8>, RewriteStats), String> {
    use quick_xml::events::{BytesText, Event};
    use quick_xml::reader::Reader;
    use quick_xml::writer::Writer;

    let mut reader = Reader::from_reader(xml);
    let mut writer = Writer::new(Vec::new());

    // State machine for tracking position in XML hierarchy
    let mut in_ser = false;
    let mut cur_elem: Option<SeriesElem> = None; // tx / cat / val / xVal / yVal / bubbleSize
    let mut cur_kind: Option<RefKind> = None;    // inside numRef / strRef of cur_elem
    let mut in_cache = false;                    // inside numCache / strCache of that ref
    let mut in_f = false;                        // inside the <c:f> of a numRef/strRef
    let mut in_any_f = false;                    // inside ANY <c:f> (prefix stripping, GOTCHA #44)
    let mut in_pt = false;                       // inside <c:pt>

    let mut current_range_ref = String::new();
    let mut current_values: Option<CacheValues> = None;
    let mut stats = RewriteStats::default();

    // Rebuild state for the cache currently being rewritten
    let mut rebuild = false;          // true while inside a cache we own
    let mut flushed = false;          // new <c:pt>s already emitted for this cache
    let mut pt_format_codes: HashMap<usize, String> = HashMap::new();

    macro_rules! flush_points {
        () => {
            match current_values.as_ref() {
                Some(CacheValues::Num(v)) => write_points(&mut writer, v, &pt_format_codes)?,
                Some(CacheValues::Str(v)) => write_str_points(&mut writer, v)?,
                None => {}
            }
        };
    }

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,

            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"ser" => { in_ser = true; }
                    b"numRef" if cur_elem.is_some() => { cur_kind = Some(RefKind::Num); }
                    b"strRef" if cur_elem.is_some() => { cur_kind = Some(RefKind::Str); }
                    b"multiLvlStrRef" if cur_elem.is_some() => { stats.needs_com = true; }
                    b"datalabelsRange" => { stats.needs_com = true; }
                    b"f" => {
                        in_any_f = true;
                        if cur_kind.is_some() { in_f = true; }
                    }
                    b"numCache" | b"strCache" if cur_kind.is_some() => {
                        in_cache = true;
                        flushed = false;
                        pt_format_codes.clear();
                        current_values = lookup_values(range_values, &current_range_ref, cur_kind.unwrap_or(RefKind::Num));
                        rebuild = current_values.is_some();
                    }
                    b"pt" if in_cache => {
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
                    b"extLst" if in_cache && rebuild && !in_pt && !flushed => {
                        // <c:extLst> follows the points in the schema: emit the new points first
                        flush_points!();
                        flushed = true;
                    }
                    other => {
                        if in_ser && cur_elem.is_none()
                            && let Some(el) = SeriesElem::from_local(other)
                        {
                            cur_elem = Some(el);
                            current_range_ref.clear();
                        }
                    }
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
                        cur_elem = None;
                        cur_kind = None;
                        in_cache = false;
                        current_range_ref.clear();
                        current_values = None;
                    }
                    b"numRef" | b"strRef" => { cur_kind = None; in_cache = false; }
                    b"numCache" | b"strCache" => {
                        if rebuild {
                            if !flushed {
                                flush_points!();
                            }
                            flushed = true;
                            match cur_elem {
                                Some(SeriesElem::Val) => stats.series_updated += 1,
                                _ => stats.labels_updated += 1,
                            }
                        }
                        rebuild = false;
                        in_cache = false;
                    }
                    b"f" => { in_f = false; in_any_f = false; }
                    b"pt" => {
                        skip_write = rebuild && in_pt;
                        in_pt = false;
                    }
                    b"v" => { skip_write = rebuild && in_pt; }
                    other => {
                        if let Some(el) = SeriesElem::from_local(other)
                            && cur_elem == Some(el)
                        {
                            cur_elem = None;
                            cur_kind = None;
                            in_cache = false;
                            current_range_ref.clear();
                            current_values = None;
                        }
                    }
                }
                if !skip_write {
                    writer.write_event(Event::End(e.clone())).map_err(|e| e.to_string())?;
                }
            }

            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                if in_cache && rebuild {
                    if local.as_ref() == b"ptCount" {
                        // ptCount = total cells, blanks included
                        if let Some(vals) = current_values.as_ref() {
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
                } else if local.as_ref() == b"datalabelsRange" {
                    stats.needs_com = true;
                }
                writer.write_event(Event::Empty(e.clone())).map_err(|e| e.to_string())?;
            }

            Ok(Event::Text(ref t)) => {
                if in_any_f {
                    // GOTCHA #44: drop any `[workbook]` qualifier from the formula text.
                    let raw = String::from_utf8_lossy(t.as_ref()).to_string();
                    let (book, cleaned) = strip_formula_prefixes(&raw);
                    if in_f {
                        // Capture the (cleaned) reference for the cache rebuild
                        current_range_ref = cleaned.clone();
                    }
                    if let Some(book) = book {
                        stats.formulas_fixed += 1;
                        stats.book.get_or_insert(book);
                        writer.write_event(Event::Text(BytesText::new(&cleaned))).map_err(|e| e.to_string())?;
                    } else {
                        writer.write_event(Event::Text(t.clone())).map_err(|e| e.to_string())?;
                    }
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

    Ok((writer.into_inner(), stats))
}

/// Every cell reference inside `<c:ser>` elements: `tx`, `cat`, `val`, `xVal`, `yVal`,
/// `bubbleSize`, with the kind taken from the enclosing `numRef` / `strRef`.
/// The bool is true when the chart holds a `multiLvlStrRef` or a `c15:datalabelsRange`
/// (constructs the ZIP rewrite leaves to PowerPoint).
pub fn extract_series_refs_all(xml: &str) -> (Vec<SeriesRef>, bool) {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(xml.as_bytes());
    let mut refs = Vec::new();
    let mut needs_com = false;
    let mut in_ser = false;
    let mut cur_elem: Option<SeriesElem> = None;
    let mut cur_kind: Option<RefKind> = None;
    let mut in_f = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"ser" => { in_ser = true; }
                    b"numRef" if cur_elem.is_some() => { cur_kind = Some(RefKind::Num); }
                    b"strRef" if cur_elem.is_some() => { cur_kind = Some(RefKind::Str); }
                    b"multiLvlStrRef" if cur_elem.is_some() => { needs_com = true; }
                    b"datalabelsRange" => { needs_com = true; }
                    b"f" if cur_kind.is_some() => { in_f = true; }
                    other => {
                        if in_ser && cur_elem.is_none()
                            && let Some(el) = SeriesElem::from_local(other)
                        {
                            cur_elem = Some(el);
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"datalabelsRange" {
                    needs_com = true;
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"ser" => { in_ser = false; cur_elem = None; cur_kind = None; in_f = false; }
                    b"numRef" | b"strRef" => { cur_kind = None; in_f = false; }
                    b"f" => { in_f = false; }
                    other => {
                        if let Some(el) = SeriesElem::from_local(other)
                            && cur_elem == Some(el)
                        {
                            cur_elem = None;
                            cur_kind = None;
                        }
                    }
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_f && let (Some(elem), Some(kind)) = (cur_elem, cur_kind) {
                    let text = String::from_utf8_lossy(t.as_ref()).trim().to_string();
                    if !text.is_empty() {
                        refs.push(SeriesRef { elem, formula: text, kind });
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    (refs, needs_com)
}

#[cfg_attr(not(test), allow(dead_code))]
/// Value-axis references only (`<c:val>`), in series order — kept for callers that
/// key series by their value range.
pub fn extract_val_refs(xml: &str) -> Vec<String> {
    extract_series_refs_all(xml).0.into_iter()
        .filter(|r| r.elem == SeriesElem::Val)
        .map(|r| r.formula)
        .collect()
}

/// Normalize a range reference for HashMap lookup and for reading from Excel.
/// Strips `$` signs, outer parentheses, and any `[workbook]` qualifier (GOTCHA #44)
/// from each comma-separated sub-range.
pub fn normalize_range_ref(range_ref: &str) -> String {
    let base = range_ref
        .trim()
        .trim_start_matches('(')
        .trim_end_matches(')')
        .replace('$', "");
    if !base.contains('[') {
        return base;
    }
    base.split(',')
        .map(|sub| strip_workbook_prefix(sub.trim()).1)
        .collect::<Vec<_>>()
        .join(",")
}

/// Strip a `[workbook]` qualifier from the sheet part of ONE range ref (GOTCHA #44).
///
/// `[b.xlsx]Tables!$A$1` → `Tables!$A$1`; `'[b.xlsx]My Sheet'!$A$1` → `'My Sheet'!$A$1`.
/// Returns the removed workbook name and the cleaned ref. Refs without `!` or without a
/// bracket pair before the `!` are returned unchanged.
pub fn strip_workbook_prefix(range_ref: &str) -> (Option<String>, String) {
    let Some(bang) = range_ref.find('!') else {
        return (None, range_ref.to_string());
    };
    let sheet_part = &range_ref[..bang];
    let (Some(open), Some(close)) = (sheet_part.find('['), sheet_part.find(']')) else {
        return (None, range_ref.to_string());
    };
    if close < open {
        return (None, range_ref.to_string());
    }
    let book = sheet_part[open + 1..close].to_string();
    let cleaned = format!("{}{}{}", &sheet_part[..open], &sheet_part[close + 1..], &range_ref[bang..]);
    (Some(book), cleaned)
}

/// Strip `[workbook]` qualifiers from a raw `<c:f>` formula, preserving `$` signs and
/// the parentheses of a multi-area formula (GOTCHA #20). Returns the first workbook name
/// removed (if any) and the cleaned formula.
fn strip_formula_prefixes(formula: &str) -> (Option<String>, String) {
    if !formula.contains('[') {
        return (None, formula.to_string());
    }
    let mut book: Option<String> = None;
    let cleaned = formula
        .split(',')
        .map(|sub| {
            let (b, s) = strip_workbook_prefix(sub);
            if let Some(b) = b {
                book.get_or_insert(b);
            }
            s
        })
        .collect::<Vec<_>>()
        .join(",");
    (book, cleaned)
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

/// Collect all unique `(normalised range, kind)` keys from a scan.
/// Strips `$`, parens and `[workbook]` (GOTCHA #44) and splits non-contiguous ranges (#20).
/// Sorted for deterministic COM read order.
pub fn collect_unique_ranges(chart_refs: &HashMap<String, Vec<SeriesRef>>) -> Vec<RangeKey> {
    let mut unique = std::collections::HashSet::new();
    for refs in chart_refs.values() {
        for r in refs {
            let normalized = normalize_range_ref(&r.formula);
            for sub in normalized.split(',') {
                let sub = sub.trim();
                if !sub.is_empty() {
                    unique.insert((sub.to_string(), r.kind));
                }
            }
        }
    }
    let mut keys: Vec<RangeKey> = unique.into_iter().collect();
    keys.sort_by(|a, b| a.0.cmp(&b.0).then_with(|| (a.1 as u8).cmp(&(b.1 as u8))));
    keys
}

/// Cached category labels and series name of one series, as PowerPoint stores them
/// (for `oa check`, GOTCHA #48). Numeric categories are reported as `cat_ref = None`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SeriesLabels {
    pub cat_ref: Option<String>,
    pub cat: Vec<Option<String>>,
    pub tx_ref: Option<String>,
    /// Series name cache — one point per referenced cell (a multi-cell name such as
    /// `Tables!$L$1067:$L$1069` is stored as three points, exactly as PowerPoint does).
    pub tx: Vec<Option<String>>,
    pub multi_level: bool,
}

/// Extract cached `tx` / `cat` strings per series (order = series order).
pub fn extract_cached_labels(xml: &str) -> Vec<SeriesLabels> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(xml.as_bytes());
    let mut out: Vec<SeriesLabels> = Vec::new();
    let mut cur: SeriesLabels = SeriesLabels::default();

    let mut in_ser = false;
    let mut cur_elem: Option<SeriesElem> = None;
    let mut in_str_ref = false;
    let mut in_f = false;
    let mut in_cache = false;
    let mut in_pt = false;
    let mut in_v = false;
    let mut pt_idx = 0usize;
    let mut pt_count = 0usize;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"ser" => { in_ser = true; cur = SeriesLabels::default(); }
                    b"strRef" if matches!(cur_elem, Some(SeriesElem::Tx) | Some(SeriesElem::Cat)) => { in_str_ref = true; }
                    b"multiLvlStrRef" if cur_elem == Some(SeriesElem::Cat) => { cur.multi_level = true; }
                    b"f" if in_str_ref => { in_f = true; }
                    b"strCache" if in_str_ref => { in_cache = true; pt_count = 0; }
                    b"pt" if in_cache => {
                        in_pt = true;
                        pt_idx = e.try_get_attribute("idx").ok().flatten()
                            .and_then(|a| String::from_utf8_lossy(a.value.as_ref()).parse::<usize>().ok())
                            .unwrap_or(0);
                        match cur_elem {
                            Some(SeriesElem::Cat) => { while cur.cat.len() <= pt_idx { cur.cat.push(None); } }
                            Some(SeriesElem::Tx) => { while cur.tx.len() <= pt_idx { cur.tx.push(None); } }
                            _ => {}
                        }
                    }
                    b"v" if in_pt => { in_v = true; }
                    other => {
                        if in_ser && cur_elem.is_none()
                            && let Some(el) = SeriesElem::from_local(other)
                        {
                            cur_elem = Some(el);
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_cache && e.local_name().as_ref() == b"ptCount" {
                    pt_count = e.try_get_attribute("val").ok().flatten()
                        .and_then(|a| String::from_utf8_lossy(a.value.as_ref()).parse::<usize>().ok())
                        .unwrap_or(0);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"ser" => {
                        in_ser = false; cur_elem = None; in_str_ref = false; in_cache = false;
                        out.push(std::mem::take(&mut cur));
                    }
                    b"strRef" => { in_str_ref = false; in_cache = false; }
                    b"strCache" => {
                        match cur_elem {
                            Some(SeriesElem::Cat) => { while cur.cat.len() < pt_count { cur.cat.push(None); } }
                            Some(SeriesElem::Tx) => { while cur.tx.len() < pt_count { cur.tx.push(None); } }
                            _ => {}
                        }
                        in_cache = false;
                    }
                    b"f" => { in_f = false; }
                    b"pt" => { in_pt = false; }
                    b"v" => { in_v = false; }
                    other => {
                        if let Some(el) = SeriesElem::from_local(other)
                            && cur_elem == Some(el)
                        {
                            cur_elem = None;
                            in_str_ref = false;
                            in_cache = false;
                        }
                    }
                }
            }
            Ok(Event::Text(ref t)) => {
                let text = String::from_utf8_lossy(t.as_ref()).to_string();
                if in_f {
                    match cur_elem {
                        Some(SeriesElem::Cat) => cur.cat_ref = Some(text.trim().to_string()),
                        Some(SeriesElem::Tx) => cur.tx_ref = Some(text.trim().to_string()),
                        _ => {}
                    }
                } else if in_v && in_pt && in_cache {
                    let slot = match cur_elem {
                        Some(SeriesElem::Cat) => cur.cat.get_mut(pt_idx),
                        Some(SeriesElem::Tx) => cur.tx.get_mut(pt_idx),
                        _ => None,
                    };
                    if let Some(slot) = slot {
                        slot.get_or_insert_with(String::new).push_str(&text);
                    }
                }
            }
            Ok(Event::GeneralRef(ref r)) => {
                // `&amp;` etc. arrive as separate events (quick-xml >= 0.37)
                if in_v && in_pt && in_cache {
                    let piece = if let Ok(Some(c)) = r.resolve_char_ref() {
                        c.to_string()
                    } else {
                        match r.decode().unwrap_or_default().as_ref() {
                            "amp" => "&".to_string(),
                            "lt" => "<".to_string(),
                            "gt" => ">".to_string(),
                            "quot" => "\"".to_string(),
                            "apos" => "'".to_string(),
                            other => format!("&{other};"),
                        }
                    };
                    let slot = match cur_elem {
                        Some(SeriesElem::Cat) => cur.cat.get_mut(pt_idx),
                        Some(SeriesElem::Tx) => cur.tx.get_mut(pt_idx),
                        _ => None,
                    };
                    if let Some(slot) = slot {
                        slot.get_or_insert_with(String::new).push_str(&piece);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    out
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

    fn sref(elem: SeriesElem, formula: &str, kind: RefKind) -> SeriesRef {
        SeriesRef { elem, formula: formula.to_string(), kind }
    }

    #[test]
    fn test_collect_unique_ranges() {
        let mut chart_ranges = HashMap::new();
        chart_ranges.insert("chart1.xml".to_string(), vec![
            sref(SeriesElem::Val, "Tables!$B$388:$B$390", RefKind::Num),
            sref(SeriesElem::Val, "(Tables!$C$810,Tables!$F$810)", RefKind::Num),
            sref(SeriesElem::Cat, "Tables!$A$388:$A$390", RefKind::Str),
            sref(SeriesElem::Cat, "Tables!$A$388:$A$390", RefKind::Str), // shared by a 2nd series → once
        ]);
        let unique = collect_unique_ranges(&chart_ranges);
        assert!(unique.contains(&("Tables!B388:B390".to_string(), RefKind::Num)));
        assert!(unique.contains(&("Tables!C810".to_string(), RefKind::Num)));
        assert!(unique.contains(&("Tables!F810".to_string(), RefKind::Num)));
        assert!(unique.contains(&("Tables!A388:A390".to_string(), RefKind::Str)));
        assert_eq!(unique.len(), 4, "shared category range counted once (GOTCHA #23)");
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

        let values = vals("Tables!B1:B3", vec![Some(0.5), Some(0.6), Some(0.7)]);

        let (output, st) = rewrite_chart_cache(xml, &values).unwrap();
        let output_str = String::from_utf8(output).unwrap();
        assert_eq!(st.series_updated, 1);
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

    fn vals(range: &str, v: Vec<ChartValue>) -> HashMap<RangeKey, CacheValues> {
        let mut m = HashMap::new();
        m.insert((range.to_string(), RefKind::Num), CacheValues::Num(v));
        m
    }

    fn strs(range: &str, v: &[Option<&str>]) -> (RangeKey, CacheValues) {
        ((range.to_string(), RefKind::Str), CacheValues::Str(v.iter().map(|s| s.map(|x| x.to_string())).collect()))
    }

    /// (idx, value-text) of every <c:pt> inside the FIRST `<c:{elem}>` block.
    fn pts_in(xml: &str, elem: &str) -> Vec<(usize, String)> {
        let open = format!("<c:{elem}>");
        let close = format!("</c:{elem}>");
        let start = xml.find(&open).unwrap_or_else(|| panic!("no {open}"));
        let block = &xml[start..xml[start..].find(&close).map(|p| start + p).expect("no close")];
        let mut out = Vec::new();
        let mut rest = block;
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

    // ── rewrite: holes anywhere ─────────────────────────────

    #[test]
    fn test_rewrite_fills_middle_gap() {
        // Template cache has pts 0,1,3 (idx 2 was a blank cell) — the Chart 33 case
        let xml = chart_with_cache(r#"<c:formatCode>0%</c:formatCode><c:ptCount val="4"/>
<c:pt idx="0"><c:v>0.05</c:v></c:pt><c:pt idx="1"><c:v>0.02</c:v></c:pt><c:pt idx="3"><c:v>0.13</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.08), Some(0.1), Some(0.18), Some(0.3)]);
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.series_updated, 1);
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
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.series_updated, 1);
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
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.series_updated, 1);
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
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.series_updated, 1);
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
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.series_updated, 1);
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
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        assert_eq!(st.series_updated, 0);
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
        m.insert(("Tables!C10".to_string(), RefKind::Num), CacheValues::Num(vec![Some(0.7)]));
        m.insert(("Tables!F10".to_string(), RefKind::Num), CacheValues::Num(vec![None]));
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.series_updated, 1);
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
        m.insert(("Tables!B1:B2".to_string(), RefKind::Num), CacheValues::Num(vec![Some(1.0), Some(2.0)]));
        m.insert(("Tables!C1:C2".to_string(), RefKind::Num), CacheValues::Num(vec![Some(3.0), None]));
        // D not in map → untouched
        let (_, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        assert_eq!(st.series_updated, 2);
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

    // ── GOTCHA #44: [workbook] qualified formulas ──────────

    #[test]
    fn test_strip_workbook_prefix() {
        assert_eq!(
            strip_workbook_prefix("[rpm_2025_Indonesia_v5.xlsx]Tables!$V$778:$Y$778"),
            (Some("rpm_2025_Indonesia_v5.xlsx".into()), "Tables!$V$778:$Y$778".into())
        );
        assert_eq!(
            strip_workbook_prefix("'[b.xlsx]My Sheet'!$A$1"),
            (Some("b.xlsx".into()), "'My Sheet'!$A$1".into())
        );
        assert_eq!(strip_workbook_prefix("Tables!$A$1"), (None, "Tables!$A$1".into()));
        assert_eq!(strip_workbook_prefix("A1:B2"), (None, "A1:B2".into()));
        // Bracket after the '!' is not a workbook qualifier
        assert_eq!(strip_workbook_prefix("Tables!A[1]"), (None, "Tables!A[1]".into()));
    }

    #[test]
    fn test_normalize_strips_workbook_prefix() {
        assert_eq!(
            normalize_range_ref("[rpm_2025_Indonesia_v5.xlsx]Tables!$V$778:$Y$778"),
            "Tables!V778:Y778"
        );
        assert_eq!(
            normalize_range_ref("([b.xlsx]Tables!$C$10,[b.xlsx]Tables!$F$10)"),
            "Tables!C10,Tables!F10"
        );
        assert_eq!(normalize_range_ref("'[b.xlsx]My Sheet'!$A$1"), "'My Sheet'!A1");
    }

    /// Chart 31 shape: tx, cat and val formulas all name the workbook.
    fn qualified_chart() -> &'static str {
        r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart><c:plotArea><c:barChart>
<c:ser><c:tx><c:strRef><c:f>[rpm_2025_Indonesia_v5.xlsx]Tables!$V$776</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Market</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:cat><c:strRef><c:f>[rpm_2025_Indonesia_v5.xlsx]Tables!$V$777:$Y$777</c:f><c:strCache><c:ptCount val="4"/>
<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>[rpm_2025_Indonesia_v5.xlsx]Tables!$V$778:$Y$778</c:f>
<c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="4"/>
<c:pt idx="0"><c:v>0.12</c:v></c:pt><c:pt idx="1"><c:v>0.27</c:v></c:pt><c:pt idx="3"><c:v>0.05</c:v></c:pt>
</c:numCache></c:numRef></c:val></c:ser>
</c:barChart></c:plotArea></c:chart></c:chartSpace>"#
    }

    #[test]
    fn test_rewrite_strips_prefix_from_all_formulas_and_rebuilds_values() {
        // Excel map is keyed by the PLAIN ref (as collect_unique_ranges produces it)
        let m = vals("Tables!V778:Y778", vec![Some(0.1), Some(0.26), Some(0.2), Some(0.09)]);
        let (out, st) = rewrite_chart_cache(qualified_chart().as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();

        assert_eq!(st.series_updated, 1);
        assert_eq!(st.formulas_fixed, 3, "tx + cat + val formulas rewritten");
        assert_eq!(st.book.as_deref(), Some("rpm_2025_Indonesia_v5.xlsx"));

        assert!(!s.contains('['), "no workbook qualifier may survive: {s}");
        assert!(s.contains("<c:f>Tables!$V$776</c:f>"));
        assert!(s.contains("<c:f>Tables!$V$777:$Y$777</c:f>"));
        assert!(s.contains("<c:f>Tables!$V$778:$Y$778</c:f>"));

        // Values rebuilt from France, blank at idx 2 filled (GOTCHA #43)
        assert_eq!(pt_count_of(&s), 4);
        assert_eq!(pts_of(&s), vec![
            (0, "0.1".into()), (1, "0.26".into()), (2, "0.2".into()), (3, "0.09".into()),
        ]);
        // Category cache untouched (GOTCHA #23)
        assert!(s.contains("<c:v>C</c:v>"));
    }

    #[test]
    fn test_collect_unique_ranges_strips_prefix() {
        let mut chart_ranges = HashMap::new();
        chart_ranges.insert("chart99.xml".to_string(), vec![
            sref(SeriesElem::Val, "[rpm_2025_Indonesia_v5.xlsx]Tables!$V$778:$Y$778", RefKind::Num),
        ]);
        let unique = collect_unique_ranges(&chart_ranges);
        assert_eq!(unique, vec![("Tables!V778:Y778".to_string(), RefKind::Num)]);
    }

    // ── GOTCHA #48: categories, series names, other refs ───

    /// Chart 11 shape: tx + cat (strRef) + val, Indonesia labels cached.
    fn labelled_chart() -> &'static str {
        r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart><c:plotArea><c:barChart>
<c:ser><c:idx val="0"/>
<c:tx><c:strRef><c:f>Tables!$L$728</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Market  </c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:cat><c:strRef><c:f>Tables!$A$729:$A$733</c:f><c:strCache><c:ptCount val="5"/>
<c:pt idx="0"><c:v>Outdoor Ads</c:v></c:pt><c:pt idx="1"><c:v>In-person: Saw Someone Using</c:v></c:pt>
<c:pt idx="2"><c:v>Ad / Promotion</c:v></c:pt><c:pt idx="3"><c:v>Tv / Streaming (ads Or Content)</c:v></c:pt>
<c:pt idx="4"><c:v>E-cigarettes / Vapes (general)</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Tables!$L$729:$L$733</c:f><c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="5"/>
<c:pt idx="0"><c:v>0.24</c:v></c:pt><c:pt idx="1"><c:v>0.17</c:v></c:pt><c:pt idx="2"><c:v>0.12</c:v></c:pt>
<c:pt idx="3"><c:v>0.11</c:v></c:pt><c:pt idx="4"><c:v>0.1</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
    }

    #[test]
    fn test_rewrite_rebuilds_categories_series_name_and_values() {
        let mut m = vals("Tables!L729:L733", vec![Some(0.29), Some(0.07), Some(0.05), Some(0.04), Some(0.03)]);
        let (k, v) = strs("Tables!A729:A733", &[
            Some("Health Risks / Safety Concerns"), Some("Cigarettes & Vapes"),
            Some("Harm Reduction / 'safer' <alt>"), Some("Negative (general)"), Some("E-cigarettes / Vapes (general)"),
        ]);
        m.insert(k, v);
        let (k, v) = strs("Tables!L728", &[Some("France  ")]);
        m.insert(k, v);

        let (out, st) = rewrite_chart_cache(labelled_chart().as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.series_updated, 1);
        assert_eq!(st.labels_updated, 2, "cat + tx");
        assert!(!st.needs_com);

        assert_eq!(pts_in(&s, "tx"), vec![(0, "France  ".into())]);
        assert_eq!(pts_in(&s, "cat"), vec![
            (0, "Health Risks / Safety Concerns".into()),
            (1, "Cigarettes &amp; Vapes".into()),
            (2, "Harm Reduction / &apos;safer&apos; &lt;alt&gt;".into()),
            (3, "Negative (general)".into()),
            (4, "E-cigarettes / Vapes (general)".into()),
        ]);
        assert_eq!(pts_in(&s, "val"), vec![
            (0, "0.29".into()), (1, "0.07".into()), (2, "0.05".into()), (3, "0.04".into()), (4, "0.03".into()),
        ]);
        assert!(!s.contains("Outdoor Ads"), "stale label must be gone");
        // formatCode / ptCount still precede the first point (schema order)
        let v = &s[s.find("<c:val>").unwrap()..];
        assert!(v.find("<c:formatCode>").unwrap() < v.find("<c:ptCount").unwrap());
        assert!(v.find("<c:ptCount").unwrap() < v.find("<c:pt idx").unwrap());
    }

    #[test]
    fn test_rewrite_labels_only_when_values_missing_from_map() {
        // Only the category range is known → cat rebuilt, val passes through unchanged
        let mut m = HashMap::new();
        let (k, v) = strs("Tables!A729:A733", &[Some("a"), None, Some("c"), Some("d"), Some("e")]);
        m.insert(k, v);
        let (out, st) = rewrite_chart_cache(labelled_chart().as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!((st.series_updated, st.labels_updated), (0, 1));
        assert_eq!(pts_in(&s, "cat"), vec![(0, "a".into()), (2, "c".into()), (3, "d".into()), (4, "e".into())],
            "blank label cell → no point, idx kept");
        assert!(s.contains(r#"<c:cat><c:strRef><c:f>Tables!$A$729:$A$733</c:f><c:strCache><c:ptCount val="5"/>"#));
        assert!(s.contains("<c:v>0.24</c:v>"), "values untouched");
        assert!(s.contains("<c:v>Market  </c:v>"), "series name untouched");
    }

    #[test]
    fn test_rewrite_numeric_categories_use_num_cache() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart>
<c:ser><c:cat><c:numRef><c:f>Tables!$A$1:$A$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
<c:pt idx="0"><c:v>2022</c:v></c:pt><c:pt idx="1"><c:v>2023</c:v></c:pt><c:pt idx="2"><c:v>2024</c:v></c:pt></c:numCache></c:numRef></c:cat>
<c:val><c:numRef><c:f>Tables!$B$1:$B$3</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>"#;
        let mut m = vals("Tables!B1:B3", vec![Some(4.0), Some(5.0), Some(6.0)]);
        m.insert(("Tables!A1:A3".to_string(), RefKind::Num), CacheValues::Num(vec![Some(2023.0), Some(2024.0), Some(2025.0)]));
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!((st.series_updated, st.labels_updated), (1, 1));
        assert_eq!(pts_in(&s, "cat"), vec![(0, "2023".into()), (1, "2024".into()), (2, "2025".into())]);
        assert!(s.contains("<c:formatCode>General</c:formatCode>"));
    }

    #[test]
    fn test_rewrite_scatter_xval_yval_and_bubble() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:bubbleChart>
<c:ser><c:xVal><c:numRef><c:f>Tables!$A$1:$A$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:xVal>
<c:yVal><c:numRef><c:f>Tables!$B$1:$B$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:yVal>
<c:bubbleSize><c:numRef><c:f>Tables!$C$1:$C$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:bubbleSize>
</c:ser></c:bubbleChart></c:plotArea></c:chart></c:chartSpace>"#;
        let mut m = HashMap::new();
        for (r, a, b) in [("Tables!A1:A2", 10.0, 11.0), ("Tables!B1:B2", 20.0, 21.0), ("Tables!C1:C2", 30.0, 31.0)] {
            m.insert((r.to_string(), RefKind::Num), CacheValues::Num(vec![Some(a), Some(b)]));
        }
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!((st.series_updated, st.labels_updated), (0, 3));
        assert_eq!(pts_in(&s, "xVal"), vec![(0, "10".into()), (1, "11".into())]);
        assert_eq!(pts_in(&s, "yVal"), vec![(0, "20".into()), (1, "21".into())]);
        assert_eq!(pts_in(&s, "bubbleSize"), vec![(0, "30".into()), (1, "31".into())]);
    }

    #[test]
    fn test_rewrite_multilevel_categories_left_to_powerpoint() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart>
<c:ser><c:cat><c:multiLvlStrRef><c:f>Tables!$A$1:$B$2</c:f><c:multiLvlStrCache><c:ptCount val="2"/>
<c:lvl><c:pt idx="0"><c:v>x</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
<c:val><c:numRef><c:f>Tables!$C$1:$C$2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let m = vals("Tables!C1:C2", vec![Some(7.0), Some(8.0)]);
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert!(st.needs_com, "multi-level categories need PowerPoint's refresh");
        assert_eq!(st.series_updated, 1, "values are still rebuilt");
        assert!(s.contains("<c:lvl><c:pt idx=\"0\"><c:v>x</c:v></c:pt></c:lvl>"), "multi-level cache untouched");
        let (refs, needs_com) = extract_series_refs_all(xml);
        assert!(needs_com);
        assert_eq!(refs, vec![sref(SeriesElem::Val, "Tables!$C$1:$C$2", RefKind::Num)]);
    }

    #[test]
    fn test_extract_series_refs_all_kinds_and_labels() {
        let (refs, needs_com) = extract_series_refs_all(labelled_chart());
        assert!(!needs_com);
        assert_eq!(refs, vec![
            sref(SeriesElem::Tx, "Tables!$L$728", RefKind::Str),
            sref(SeriesElem::Cat, "Tables!$A$729:$A$733", RefKind::Str),
            sref(SeriesElem::Val, "Tables!$L$729:$L$733", RefKind::Num),
        ]);
        assert_eq!(extract_val_refs(labelled_chart()), vec!["Tables!$L$729:$L$733".to_string()]);

        let labels = extract_cached_labels(labelled_chart());
        assert_eq!(labels.len(), 1);
        assert_eq!(labels[0].tx_ref.as_deref(), Some("Tables!$L$728"));
        assert_eq!(labels[0].tx, vec![Some("Market  ".to_string())]);
        assert_eq!(labels[0].cat_ref.as_deref(), Some("Tables!$A$729:$A$733"));
        assert_eq!(labels[0].cat, vec![
            Some("Outdoor Ads".to_string()), Some("In-person: Saw Someone Using".to_string()),
            Some("Ad / Promotion".to_string()), Some("Tv / Streaming (ads Or Content)".to_string()),
            Some("E-cigarettes / Vapes (general)".to_string()),
        ]);
        assert!(!labels[0].multi_level);
    }

    #[test]
    fn test_extract_cached_labels_entities_and_holes() {
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart>
<c:ser><c:cat><c:strRef><c:f>Tables!$A$1:$A$3</c:f><c:strCache><c:ptCount val="3"/>
<c:pt idx="0"><c:v>R&amp;D</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Tables!$B$1:$B$3</c:f><c:numCache><c:ptCount val="3"/></c:numCache></c:numRef></c:val>
</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let labels = extract_cached_labels(xml);
        assert_eq!(labels[0].cat, vec![Some("R&D".to_string()), None, Some("c".to_string())]);
        assert!(labels[0].tx.is_empty());
    }

    #[test]
    fn test_multi_cell_series_name_is_one_point_per_cell() {
        // PowerPoint stores `tx` = Tables!$L$1067:$L$1069 as three points, not one joined string
        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart>
<c:ser><c:tx><c:strRef><c:f>Tables!$L$1067:$L$1069</c:f><c:strCache><c:ptCount val="3"/>
<c:pt idx="0"><c:v>2025</c:v></c:pt><c:pt idx="1"><c:v>Gen Pop</c:v></c:pt><c:pt idx="2"><c:v>Market  </c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:val><c:numRef><c:f>Tables!$L$1070:$L$1071</c:f><c:numCache><c:ptCount val="2"/></c:numCache></c:numRef></c:val>
</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let labels = extract_cached_labels(xml);
        assert_eq!(labels[0].tx, vec![Some("2025".to_string()), Some("Gen Pop".to_string()), Some("Market  ".to_string())]);

        let mut m = HashMap::new();
        let (k, v) = strs("Tables!L1067:L1069", &[Some("2025"), Some("Gen Pop"), Some("France  ")]);
        m.insert(k, v);
        let (out, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        let s = String::from_utf8(out).unwrap();
        assert_eq!(st.labels_updated, 1);
        assert_eq!(pts_in(&s, "tx"), vec![(0, "2025".into()), (1, "Gen Pop".into()), (2, "France  ".into())]);
        assert!(s.contains(r#"<c:tx><c:strRef><c:f>Tables!$L$1067:$L$1069</c:f><c:strCache><c:ptCount val="3"/>"#));
    }

    #[test]
    fn test_roundtrip_labels_rewrite_then_extract() {
        let mut m = HashMap::new();
        let (k, v) = strs("Tables!A729:A733", &[Some("Ä & Ö"), Some("b"), None, Some("d"), Some("e")]);
        m.insert(k, v);
        let (out, _) = rewrite_chart_cache(labelled_chart().as_bytes(), &m).unwrap();
        let labels = extract_cached_labels(&String::from_utf8(out).unwrap());
        assert_eq!(labels[0].cat, vec![Some("Ä & Ö".to_string()), Some("b".to_string()), None, Some("d".to_string()), Some("e".to_string())]);
    }

    #[test]
    fn test_rewrite_plain_chart_reports_no_fix() {
        let xml = chart_with_cache(r#"<c:ptCount val="1"/><c:pt idx="0"><c:v>0.5</c:v></c:pt>"#);
        let m = vals("Tables!B1:B4", vec![Some(0.7)]);
        let (_, st) = rewrite_chart_cache(xml.as_bytes(), &m).unwrap();
        assert_eq!(st.formulas_fixed, 0);
        assert_eq!(st.book, None);
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
