//! Streaming XML transform helpers using quick-xml.
//!
//! Provides utilities for efficient single-pass XML modifications in PPTX files.

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;

/// Rewrite attributes in XML elements matching a predicate.
///
/// Reads XML from `input`, writes to output. For each empty or start element,
/// calls `rewriter` which can modify attributes. Returns the count of modified elements.
pub fn rewrite_xml_attributes<F>(input: &[u8], mut rewriter: F) -> Result<(Vec<u8>, usize), String>
where
    F: FnMut(&mut BytesStart) -> bool,
{
    let mut reader = Reader::from_reader(input);
    let mut writer = Writer::new(Vec::new());
    let mut count = 0;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Empty(ref e)) => {
                let mut elem = e.clone();
                if rewriter(&mut elem) {
                    count += 1;
                }
                writer.write_event(Event::Empty(elem)).map_err(|e| e.to_string())?;
            }
            Ok(Event::Start(ref e)) => {
                let mut elem = e.clone();
                if rewriter(&mut elem) {
                    count += 1;
                }
                writer.write_event(Event::Start(elem)).map_err(|e| e.to_string())?;
            }
            Ok(event) => {
                writer.write_event(event).map_err(|e| e.to_string())?;
            }
            Err(e) => return Err(format!("XML parse error: {e}")),
        }
    }

    Ok((writer.into_inner(), count))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_identity_transform() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<Root><Child attr="value"/></Root>"#;

        let (output, count) = rewrite_xml_attributes(xml, |_| false).unwrap();
        assert_eq!(count, 0);
        // Output should be valid XML (may differ in formatting)
        let output_str = String::from_utf8(output).unwrap();
        assert!(output_str.contains("Child"));
        assert!(output_str.contains(r#"attr="value""#));
    }

    #[test]
    fn test_attribute_rewrite() {
        let xml = br#"<?xml version="1.0"?>
<Root><Item name="old"/><Item name="keep"/></Root>"#;

        let (output, count) = rewrite_xml_attributes(xml, |elem| {
            if elem.name().as_ref() == b"Item" {
                let name = elem.try_get_attribute("name")
                    .ok()
                    .flatten()
                    .map(|a| String::from_utf8_lossy(a.value.as_ref()).to_string());
                if name.as_deref() == Some("old") {
                    elem.clear_attributes();
                    elem.push_attribute(("name", "new"));
                    return true;
                }
            }
            false
        }).unwrap();

        assert_eq!(count, 1);
        let output_str = String::from_utf8(output).unwrap();
        assert!(output_str.contains(r#"name="new""#));
        assert!(output_str.contains(r#"name="keep""#));
    }
}
