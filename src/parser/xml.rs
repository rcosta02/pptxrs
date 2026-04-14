use std::io::{Cursor, Read};
use zip::ZipArchive;
use quick_xml::events::Event;
use quick_xml::Reader;
use base64::Engine;
use base64::engine::general_purpose::STANDARD as B64;

use crate::model::{
    elements::{
        ChartData, ChartOptions, ChartType, CoordVal, HorizAlign, ImageOptions, Position,
        ShapeOptions, SlideElement, TableCell, TableOptions, TextContent, TextOptions,
        TextRun, VertAlign,
    },
    presentation::{Layout, Presentation, PresentationMeta},
    slide::{Slide, SlideBackground},
    master::SlideMaster,
};

pub fn parse_pptx(data: &[u8]) -> Result<Presentation, String> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let mut pres = Presentation::new();

    // Parse presentation.xml for slide list and dimensions
    let slide_count = parse_presentation_xml(&mut archive, &mut pres)?;

    // Parse each slide
    for idx in 0..slide_count {
        let slide_name = format!("ppt/slides/slide{}.xml", idx + 1);
        let rels_name = format!("ppt/slides/_rels/slide{}.xml.rels", idx + 1);

        // Build a map: rId -> target path (for images)
        let rel_map = parse_rels(&mut archive, &rels_name);

        let slide_xml = read_zip_entry(&mut archive, &slide_name)?;
        let slide = parse_slide_xml(&slide_xml, &rel_map, &mut archive, idx)?;
        pres.slides.push(slide);
    }

    Ok(pres)
}

fn read_zip_entry(archive: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Result<String, String> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| format!("entry '{}': {}", name, e))?;
    let mut s = String::new();
    entry.read_to_string(&mut s).map_err(|e| e.to_string())?;
    Ok(s)
}

fn read_zip_entry_bytes(archive: &mut ZipArchive<Cursor<&[u8]>>, name: &str) -> Option<Vec<u8>> {
    let mut entry = archive.by_name(name).ok()?;
    let mut buf = Vec::new();
    entry.read_to_end(&mut buf).ok()?;
    Some(buf)
}

/// Returns number of slides found
fn parse_presentation_xml(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    pres: &mut Presentation,
) -> Result<usize, String> {
    let xml = read_zip_entry(archive, "ppt/presentation.xml")?;
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut slide_count = 0usize;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"p:sldSz" => {
                        // cx / cy in EMU
                        let cx = attr_i64(e, b"cx");
                        let cy = attr_i64(e, b"cy");
                        if let (Some(cx), Some(cy)) = (cx, cy) {
                            // Detect layout from dimensions
                            pres.meta.layout = match (cx, cy) {
                                (9_144_000, 6_858_000) => Layout::Layout4x3,
                                (12_192_000, 6_858_000) => Layout::LayoutWide,
                                _ => Layout::Layout16x9,
                            };
                        }
                    }
                    b"p:sldId" => {
                        slide_count += 1;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.to_string()),
            _ => {}
        }
        buf.clear();
    }

    Ok(slide_count)
}

/// Parse _rels file into a HashMap<rId, target_path>
fn parse_rels(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    rels_path: &str,
) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let Ok(xml) = read_zip_entry(archive, rels_path) else {
        return map;
    };

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) if e.name().as_ref() == b"Relationship" => {
                let id = attr_str(e, b"Id").unwrap_or_default();
                let target = attr_str(e, b"Target").unwrap_or_default();
                map.insert(id, target);
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    map
}

fn parse_slide_xml(
    xml: &str,
    rel_map: &std::collections::HashMap<String, String>,
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    _idx: usize,
) -> Result<Slide, String> {
    let mut slide = Slide::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    // We'll do a simplified parse: extract text boxes (p:sp with text) and images (p:pic)
    // A full parse would recursively walk shape trees; this covers the common cases.
    let mut in_sp = false;
    let mut in_pic = false;
    let mut in_txBody = false;
    let mut current_text = String::new();
    let mut sp_x = 0i64;
    let mut sp_y = 0i64;
    let mut sp_w = 0i64;
    let mut sp_h = 0i64;
    let mut pic_rid = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.name().as_ref() {
                b"p:sp" => { in_sp = true; }
                b"p:pic" => { in_pic = true; }
                b"p:txBody" => { in_txBody = true; }
                b"a:xfrm" if in_sp || in_pic => {
                    // position comes from child a:off / a:ext
                }
                b"a:off" if in_sp || in_pic => {
                    sp_x = attr_i64(e, b"x").unwrap_or(0);
                    sp_y = attr_i64(e, b"y").unwrap_or(0);
                }
                b"a:ext" if in_sp || in_pic => {
                    sp_w = attr_i64(e, b"cx").unwrap_or(0);
                    sp_h = attr_i64(e, b"cy").unwrap_or(0);
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.name().as_ref() {
                b"a:off" if in_sp || in_pic => {
                    sp_x = attr_i64(e, b"x").unwrap_or(0);
                    sp_y = attr_i64(e, b"y").unwrap_or(0);
                }
                b"a:ext" if in_sp || in_pic => {
                    sp_w = attr_i64(e, b"cx").unwrap_or(0);
                    sp_h = attr_i64(e, b"cy").unwrap_or(0);
                }
                b"r:embed" if in_pic => {
                    pic_rid = attr_str(e, b"r:id").unwrap_or_default();
                }
                _ => {}
            },
            Ok(Event::Text(ref e)) if in_txBody => {
                let t = e.unescape().unwrap_or_default();
                if !t.is_empty() {
                    if !current_text.is_empty() {
                        current_text.push(' ');
                    }
                    current_text.push_str(&t);
                }
            }
            Ok(Event::End(ref e)) => match e.name().as_ref() {
                b"p:sp" => {
                    if in_sp && !current_text.is_empty() {
                        slide.elements.push(SlideElement::Text {
                            text: TextContent::Plain(current_text.clone()),
                            options: TextOptions {
                                pos: emu_to_position(sp_x, sp_y, sp_w, sp_h),
                                ..Default::default()
                            },
                        });
                    }
                    in_sp = false;
                    in_txBody = false;
                    current_text.clear();
                    sp_x = 0; sp_y = 0; sp_w = 0; sp_h = 0;
                }
                b"p:pic" => {
                    if in_pic {
                        // Resolve image from relationship map
                        let image_data = if !pic_rid.is_empty() {
                            if let Some(target) = rel_map.get(&pic_rid) {
                                // target is relative like "../media/image1.png"
                                let full = format!("ppt/slides/{}", target)
                                    .replace("/./", "/");
                                // Normalize ../
                                let full = normalize_path(&full);
                                read_zip_entry_bytes(archive, &full)
                                    .map(|b| B64.encode(&b))
                            } else {
                                None
                            }
                        } else {
                            None
                        };

                        slide.elements.push(SlideElement::Image {
                            options: ImageOptions {
                                pos: emu_to_position(sp_x, sp_y, sp_w, sp_h),
                                data: image_data,
                                ..Default::default()
                            },
                        });
                    }
                    in_pic = false;
                    pic_rid.clear();
                    sp_x = 0; sp_y = 0; sp_w = 0; sp_h = 0;
                }
                b"p:txBody" => { in_txBody = false; }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(e.to_string()),
            _ => {}
        }
        buf.clear();
    }

    Ok(slide)
}

// ── Helpers ───────────────────────────────────────────────────────────────────

fn emu_to_position(x: i64, y: i64, w: i64, h: i64) -> Position {
    // 1 EMU = 1/914 400 inch = 1/9 525 px (at 96 DPI)
    Position {
        x: CoordVal::Pixels(x as f64 / 9_525.0),
        y: CoordVal::Pixels(y as f64 / 9_525.0),
        w: CoordVal::Pixels(w as f64 / 9_525.0),
        h: CoordVal::Pixels(h as f64 / 9_525.0),
    }
}

fn attr_i64(e: &quick_xml::events::BytesStart, name: &[u8]) -> Option<i64> {
    e.attributes()
        .filter_map(|a| a.ok())
        .find(|a| a.key.as_ref() == name)
        .and_then(|a| std::str::from_utf8(&a.value).ok()?.parse().ok())
}

fn attr_str(e: &quick_xml::events::BytesStart, name: &[u8]) -> Option<String> {
    e.attributes()
        .filter_map(|a| a.ok())
        .find(|a| a.key.as_ref() == name)
        .and_then(|a| String::from_utf8(a.value.into_owned()).ok())
}

/// Resolve `../` components in a path string
fn normalize_path(path: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for segment in path.split('/') {
        match segment {
            ".." => { parts.pop(); }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}
