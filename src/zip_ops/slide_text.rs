//! Read paragraph text out of a PPTX at ZIP level — no COM, no PowerPoint (GOTCHA #47).
//!
//! Covers slides, speaker notes, slide layouts and slide masters. Every `<a:t>` inside one
//! `<a:p>` is joined before it is handed out, because PowerPoint splits a sentence across
//! runs at arbitrary points (`positive ` + ` change`). The shape a paragraph belongs to is
//! the nearest preceding `<p:cNvPr name="…">`, so grouped shapes report the child's name and
//! table cells report the table's name.
//!
//! Notes parts are numbered independently of slides (`notesSlide7.xml` may belong to slide 3),
//! so they are resolved through each slide's `.rels`.

use std::fmt;
use std::io::{Read, Seek};
use std::path::Path;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use super::slide_map::{get_slide_order, read_rels_map, read_zip_entry};

/// Where a paragraph lives. `Display` gives the row label used by `oa find`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Location {
    Slide(i32),
    Notes(i32),
    Layout(String),
    Master(String),
}

impl fmt::Display for Location {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Location::Slide(n) => write!(f, "Slide {n:>2}"),
            Location::Notes(n) => write!(f, "Notes {n:>2}"),
            Location::Layout(name) => write!(f, "Layout {name}"),
            Location::Master(name) => write!(f, "Master {name}"),
        }
    }
}

/// One paragraph of text and where it came from.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextBlock {
    pub location: Location,
    pub shape: String,
    pub text: String,
}

/// How many parts of each kind were scanned.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct DeckStats {
    pub slides: usize,
    pub layouts: usize,
    pub masters: usize,
    pub notes: usize,
}

/// Extract every paragraph from slides (in presentation order, each followed by its notes),
/// then all layouts, then all masters.
pub fn extract_text_blocks(pptx_path: &Path) -> Result<(Vec<TextBlock>, DeckStats), String> {
    let file = std::fs::File::open(pptx_path).map_err(|e| format!("Failed to open PPTX: {e}"))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| format!("Failed to open ZIP: {e}"))?;

    let mut blocks = Vec::new();
    let mut stats = DeckStats::default();

    // --- Slides + their notes ---
    for (i, slide_part) in get_slide_order(&mut archive)?.iter().enumerate() {
        let n = (i + 1) as i32;
        if let Some(xml) = read_zip_entry(&mut archive, slide_part) {
            stats.slides += 1;
            push_blocks(&mut blocks, &Location::Slide(n), &xml);
        }
        let slide_name = slide_part.rsplit('/').next().unwrap_or(slide_part);
        let rels = read_rels_map(&mut archive, &format!("ppt/slides/_rels/{slide_name}.rels"));
        if let Some(target) = rels.values().find(|t| t.contains("notesSlides/"))
            && let Some(xml) = read_zip_entry(&mut archive, &normalize_part(target))
        {
            stats.notes += 1;
            push_blocks(&mut blocks, &Location::Notes(n), &xml);
        }
    }

    // --- Layouts, then masters (numeric part order) ---
    for part in sorted_parts(&mut archive, "ppt/slideLayouts/slideLayout") {
        let Some(xml) = read_zip_entry(&mut archive, &part) else { continue };
        stats.layouts += 1;
        let name = csld_name(&xml).unwrap_or_else(|| part_stem(&part));
        push_blocks(&mut blocks, &Location::Layout(name), &xml);
    }
    for part in sorted_parts(&mut archive, "ppt/slideMasters/slideMaster") {
        let Some(xml) = read_zip_entry(&mut archive, &part) else { continue };
        stats.masters += 1;
        let name = csld_name(&xml).unwrap_or_else(|| part_number(&part).to_string());
        push_blocks(&mut blocks, &Location::Master(name), &xml);
    }

    Ok((blocks, stats))
}

fn push_blocks(blocks: &mut Vec<TextBlock>, location: &Location, xml: &str) {
    for (shape, text) in paragraphs_from_xml(xml) {
        blocks.push(TextBlock { location: location.clone(), shape, text });
    }
}

/// `(shape name, paragraph text)` for every non-blank `<a:p>` in a slide-family XML part.
pub fn paragraphs_from_xml(xml: &str) -> Vec<(String, String)> {
    let mut reader = Reader::from_str(xml);
    let mut out = Vec::new();
    let mut shape_name = String::new();
    let mut para: Option<String> = None;
    let mut in_t = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"cNvPr" => {
                    if let Some(n) = attr(e, "name") {
                        shape_name = n;
                    }
                }
                b"p" => para = Some(String::new()),
                b"t" => in_t = para.is_some(),
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"cNvPr" => {
                    if let Some(n) = attr(e, "name") {
                        shape_name = n;
                    }
                }
                b"br" => {
                    if let Some(p) = para.as_mut() {
                        p.push('\n');
                    }
                }
                _ => {}
            },
            Ok(Event::Text(ref t)) => {
                if in_t && let Some(p) = para.as_mut() {
                    p.push_str(&t.xml_content().unwrap_or_default());
                }
            }
            // quick-xml ≥ 0.37 reports `&amp;` / `&#x2019;` as separate events
            Ok(Event::GeneralRef(ref r)) => {
                if in_t && let Some(p) = para.as_mut() {
                    if let Ok(Some(c)) = r.resolve_char_ref() {
                        p.push(c);
                    } else {
                        let name = r.decode().unwrap_or_default();
                        match name.as_ref() {
                            "amp" => p.push('&'),
                            "lt" => p.push('<'),
                            "gt" => p.push('>'),
                            "quot" => p.push('"'),
                            "apos" => p.push('\''),
                            other => { p.push('&'); p.push_str(other); p.push(';'); }
                        }
                    }
                }
            }
            Ok(Event::CData(ref c)) => {
                if in_t && let Some(p) = para.as_mut() {
                    p.push_str(&String::from_utf8_lossy(c.as_ref()));
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"t" => in_t = false,
                b"p" => {
                    if let Some(text) = para.take()
                        && !text.trim().is_empty()
                    {
                        out.push((shape_name.clone(), text));
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    out
}

/// `<p:cSld name="Title Slide">` — layouts carry a name here, masters usually do not.
fn csld_name(xml: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"cSld" {
                    return attr(e, "name").filter(|n| !n.trim().is_empty());
                }
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
    }
}

/// All `ppt/<prefix>N.xml` parts (not `.rels`), sorted by N.
fn sorted_parts<R: Read + Seek>(archive: &mut zip::ZipArchive<R>, prefix: &str) -> Vec<String> {
    let mut parts: Vec<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|e| e.name().to_string()))
        .filter(|n| n.starts_with(prefix) && n.ends_with(".xml") && !n.contains("/_rels/"))
        .collect();
    parts.sort_by_key(|p| part_number(p));
    parts
}

/// `ppt/slideLayouts/slideLayout12.xml` → `12` (0 when no digits).
fn part_number(part: &str) -> u32 {
    part_stem(part)
        .chars()
        .filter(|c| c.is_ascii_digit())
        .collect::<String>()
        .parse()
        .unwrap_or(0)
}

/// `ppt/slideLayouts/slideLayout12.xml` → `slideLayout12`.
fn part_stem(part: &str) -> String {
    part.rsplit('/').next().unwrap_or(part).trim_end_matches(".xml").to_string()
}

/// `../notesSlides/notesSlide3.xml` → `ppt/notesSlides/notesSlide3.xml`.
fn normalize_part(target: &str) -> String {
    if target.starts_with("ppt/") {
        target.to_string()
    } else {
        format!("ppt/{}", target.trim_start_matches("../").trim_start_matches('/'))
    }
}

fn attr(e: &BytesStart<'_>, key: &str) -> Option<String> {
    e.try_get_attribute(key).ok().flatten()
        .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#;

    #[test]
    fn test_runs_joined_per_paragraph() {
        let xml = format!(r#"<p:sld {NS}><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox 4"/></p:nvSpPr><p:txBody>
<a:p><a:r><a:rPr lang="en-US"/><a:t>positive </a:t></a:r><a:r><a:rPr b="1"/><a:t> change</a:t></a:r></a:p>
<a:p><a:r><a:t>second para</a:t></a:r></a:p>
<a:p><a:endParaRPr/></a:p>
</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#);
        assert_eq!(paragraphs_from_xml(&xml), vec![
            ("TextBox 4".to_string(), "positive  change".to_string()),
            ("TextBox 4".to_string(), "second para".to_string()),
        ]);
    }

    #[test]
    fn test_group_child_name_wins_and_table_cells_inherit_table_name() {
        let xml = format!(r#"<p:sld {NS}><p:cSld><p:spTree>
<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 9"/></p:nvGrpSpPr>
  <p:sp><p:nvSpPr><p:cNvPr id="11" name="Inner Box"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>grouped [country]</a:t></a:r></a:p></p:txBody></p:sp>
</p:grpSp>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="ntbl_Object_edu"/></p:nvGraphicFramePr>
  <a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>74%</a:t></a:r></a:p></a:txBody></a:tc>
  <a:tc><a:txBody><a:p><a:r><a:t>Base: [country]</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic>
</p:graphicFrame></p:spTree></p:cSld></p:sld>"#);
        assert_eq!(paragraphs_from_xml(&xml), vec![
            ("Inner Box".to_string(), "grouped [country]".to_string()),
            ("ntbl_Object_edu".to_string(), "74%".to_string()),
            ("ntbl_Object_edu".to_string(), "Base: [country]".to_string()),
        ]);
    }

    #[test]
    fn test_field_text_line_break_and_entities() {
        let xml = format!(r#"<p:sld {NS}><p:cSld><p:spTree>
<p:sp><p:nvSpPr><p:cNvPr id="5" name="Footer 5"/></p:nvSpPr><p:txBody>
<a:p><a:fld id="x" type="slidenum"><a:t>3</a:t></a:fld><a:r><a:t> of </a:t></a:r><a:br/><a:r><a:t>Q&amp;A</a:t></a:r></a:p>
</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#);
        assert_eq!(paragraphs_from_xml(&xml), vec![("Footer 5".to_string(), "3 of \nQ&A".to_string())]);
    }

    #[test]
    fn test_csld_name_and_helpers() {
        let layout = format!(r#"<p:sldLayout {NS}><p:cSld name="Title Slide"><p:spTree/></p:cSld></p:sldLayout>"#);
        assert_eq!(csld_name(&layout).as_deref(), Some("Title Slide"));
        let master = format!(r#"<p:sldMaster {NS}><p:cSld><p:spTree/></p:cSld></p:sldMaster>"#);
        assert_eq!(csld_name(&master), None);
        assert_eq!(part_number("ppt/slideLayouts/slideLayout12.xml"), 12);
        assert_eq!(part_stem("ppt/slideLayouts/slideLayout12.xml"), "slideLayout12");
        assert_eq!(normalize_part("../notesSlides/notesSlide3.xml"), "ppt/notesSlides/notesSlide3.xml");
        assert_eq!(normalize_part("ppt/notesSlides/notesSlide3.xml"), "ppt/notesSlides/notesSlide3.xml");
    }

    #[test]
    fn test_location_display() {
        assert_eq!(Location::Slide(3).to_string(), "Slide  3");
        assert_eq!(Location::Notes(12).to_string(), "Notes 12");
        assert_eq!(Location::Layout("Title Slide".into()).to_string(), "Layout Title Slide");
        assert_eq!(Location::Master("1".into()).to_string(), "Master 1");
    }
}
