//! Slide-order and chart-ownership helpers over the PPTX ZIP — no COM needed.
//!
//! GOTCHA #19: chart parts (`ppt/charts/chartN.xml`) are not numbered in slide order;
//! resolve them through `presentation.xml` → slide `.rels` → slide XML.
//! GOTCHA #44: warnings about a chart part must name the slide and shape the user sees.

use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::Path;

/// Read a ZIP entry as a string.
pub fn read_zip_entry<R: Read + Seek>(archive: &mut zip::ZipArchive<R>, name: &str) -> Option<String> {
    let mut entry = archive.by_name(name).ok()?;
    let mut data = String::new();
    entry.read_to_string(&mut data).ok()?;
    Some(data)
}

/// Read a `.rels` part and return its rId → Target map.
pub fn read_rels_map<R: Read + Seek>(archive: &mut zip::ZipArchive<R>, path: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let Some(xml) = read_zip_entry(archive, path) else { return map };
    let mut reader = quick_xml::Reader::from_reader(xml.as_bytes());
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e)) | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let id = attr(e, "Id");
                    let target = attr(e, "Target");
                    if let (Some(id), Some(target)) = (id, target) {
                        map.insert(id, target);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

/// Slide parts in presentation order, e.g. `["ppt/slides/slide1.xml", ...]`.
/// Index + 1 is the slide number PowerPoint (and the COM inventory) uses.
pub fn get_slide_order<R: Read + Seek>(archive: &mut zip::ZipArchive<R>) -> Result<Vec<String>, String> {
    let pres_xml = read_zip_entry(archive, "ppt/presentation.xml")
        .ok_or("Missing presentation.xml")?;
    let rid_map = read_rels_map(archive, "ppt/_rels/presentation.xml.rels");
    if rid_map.is_empty() {
        return Err("Missing presentation.xml.rels".into());
    }

    let mut slides = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(pres_xml.as_bytes());
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e)) | Ok(quick_xml::events::Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"sldId"
                    && let Some(rid) = rid_attr(e)
                    && let Some(target) = rid_map.get(&rid)
                {
                    slides.push(format!("ppt/{target}"));
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(slides)
}

/// Map every chart part to the slide number and shape name that owns it:
/// `"ppt/charts/chart99.xml" → (34, "Chart 31")`.
pub fn chart_part_owners(pptx_path: &Path) -> Result<HashMap<String, (i32, String)>, String> {
    let file = std::fs::File::open(pptx_path).map_err(|e| format!("Failed to open PPTX: {e}"))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| format!("Failed to open ZIP: {e}"))?;
    let slide_order = get_slide_order(&mut archive)?;

    let mut owners = HashMap::new();
    for (i, slide_file) in slide_order.iter().enumerate() {
        let slide_num = (i + 1) as i32;
        let Some(slide_xml) = read_zip_entry(&mut archive, slide_file) else { continue };
        let slide_name = slide_file.rsplit('/').next().unwrap_or(slide_file);
        let rels = read_rels_map(&mut archive, &format!("ppt/slides/_rels/{slide_name}.rels"));

        for (rid, shape_name) in chart_rids_in_slide(&slide_xml) {
            if let Some(target) = rels.get(&rid) {
                let part = if target.starts_with("ppt/") {
                    target.clone()
                } else {
                    format!("ppt/{}", target.trim_start_matches("../"))
                };
                owners.insert(part, (slide_num, shape_name));
            }
        }
    }
    Ok(owners)
}

/// `(r:id, shape name)` for every `<c:chart>` in a slide, in document order.
/// The shape name is the last `<p:cNvPr name="…">` seen before the chart element
/// (the graphicFrame's own non-visual properties).
fn chart_rids_in_slide(slide_xml: &str) -> Vec<(String, String)> {
    let mut out = Vec::new();
    let mut reader = quick_xml::Reader::from_reader(slide_xml.as_bytes());
    let mut buf = Vec::new();
    let mut last_name = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(ref e)) | Ok(quick_xml::events::Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"cNvPr" => {
                        if let Some(n) = attr(e, "name") {
                            last_name = n;
                        }
                    }
                    b"chart" => {
                        if let Some(rid) = rid_attr(e) {
                            out.push((rid, last_name.clone()));
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn attr(e: &quick_xml::events::BytesStart<'_>, key: &str) -> Option<String> {
    e.try_get_attribute(key).ok().flatten()
        .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string())
}

/// `r:id` attribute, tolerant of any namespace prefix.
fn rid_attr(e: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    e.attributes().filter_map(|a| a.ok()).find(|a| {
        let key = String::from_utf8_lossy(a.key.as_ref());
        key == "r:id" || key.ends_with(":id")
    }).map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_chart_rids_in_slide_pairs_name_with_rid() {
        let xml = r#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:c="c" xmlns:r="r"><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/></p:nvSpPr></p:sp>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 28"/></p:nvGraphicFramePr>
<a:graphic><a:graphicData><c:chart r:id="rId10"/></a:graphicData></a:graphic></p:graphicFrame>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Chart 31"/></p:nvGraphicFramePr>
<a:graphic><a:graphicData><c:chart r:id="rId11"/></a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>"#;
        assert_eq!(chart_rids_in_slide(xml), vec![
            ("rId10".to_string(), "Chart 28".to_string()),
            ("rId11".to_string(), "Chart 31".to_string()),
        ]);
    }
}
