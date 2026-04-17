use std::io::{Cursor, Read};
use zip::ZipArchive;
use quick_xml::events::Event;
use quick_xml::Reader;
use base64::Engine;
use base64::engine::general_purpose::STANDARD as B64;

use crate::model::{
    elements::{
        ChartData, ChartOptions, ChartType, CoordVal, HorizAlign, ImageOptions, LineOptions,
        Position, ShapeOptions, SlideElement, TableCell, TableOptions, TextContent, TextOptions,
        TextRun, TextRunOptions, VertAlign,
    },
    presentation::{Layout, Presentation, PresentationMeta},
    slide::{Slide, SlideBackground},
    master::SlideMaster,
};

// ── Fill context tracker (keeps solidFill handling unambiguous) ───────────────

#[derive(PartialEq, Clone, Copy)]
enum FillCtx {
    None,
    Bg,
    SpPr,
    Line,
    RunPr,
}

// ── Run accumulator ───────────────────────────────────────────────────────────

#[derive(Default)]
struct RunAccum {
    text: String,
    font_size: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    color: Option<String>,
    lang: Option<String>,
}

// ── Public API ────────────────────────────────────────────────────────────────

pub fn parse_pptx(data: &[u8]) -> Result<Presentation, String> {
    let cursor = Cursor::new(data);
    let mut archive = ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let mut pres = Presentation::new();

    // Store original ZIP bytes for passthrough write support
    pres.source_zip = Some(data.to_vec());

    // Determine the next chart ID by scanning existing chart files
    {
        let n = archive.len();
        let mut max_chart_id = 0u32;
        for i in 0..n {
            if let Ok(f) = archive.by_index(i) {
                let name = f.name().to_string();
                if let Some(rest) = name.strip_prefix("ppt/charts/chart") {
                    if let Some(stem) = rest.strip_suffix(".xml") {
                        if let Ok(id) = stem.parse::<u32>() {
                            max_chart_id = max_chart_id.max(id);
                        }
                    }
                }
            }
        }
        pres.next_chart_id = max_chart_id + 1;
    }

    // Parse presentation.xml for slide list and dimensions
    let slide_count = parse_presentation_xml(&mut archive, &mut pres)?;
    pres.original_slide_count = slide_count;

    // Parse each slide
    for idx in 0..slide_count {
        let slide_name = format!("ppt/slides/slide{}.xml", idx + 1);
        let rels_name = format!("ppt/slides/_rels/slide{}.xml.rels", idx + 1);

        // Build a map: rId -> target path (for images / charts)
        let rel_map = parse_rels(&mut archive, &rels_name);
        let raw_rels = read_zip_entry(&mut archive, &rels_name).unwrap_or_default();

        let slide_xml = read_zip_entry(&mut archive, &slide_name)?;
        let mut slide = parse_slide_xml(&slide_xml, &rel_map, &mut archive, idx)?;

        // Capture raw <p:graphicFrame> XML for tables and charts (preserves formatting)
        let frames = extract_graphic_frames(&slide_xml);
        associate_raw_frames(&frames, &mut slide);

        slide.raw_xml = Some(slide_xml);
        slide.raw_rels = Some(raw_rels);
        slide.original_element_count = slide.elements.len();

        pres.slides.push(slide);
    }

    Ok(pres)
}

// ── Internal helpers ──────────────────────────────────────────────────────────

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
                        let cx = attr_i64(e, b"cx");
                        let cy = attr_i64(e, b"cy");
                        if let (Some(cx), Some(cy)) = (cx, cy) {
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

// ── Slide parser ──────────────────────────────────────────────────────────────

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

    // ── Top-level element flags (mutually exclusive) ──────────────────────────
    let mut in_bg    = false;
    let mut in_sp    = false;
    let mut in_pic   = false;
    let mut in_gframe = false;

    // ── Fill context (one at a time) ──────────────────────────────────────────
    let mut fill_ctx = FillCtx::None;

    // ── Background ────────────────────────────────────────────────────────────
    let mut bg_color_hex: Option<String> = None;

    // ── Shape accumulator ─────────────────────────────────────────────────────
    let mut sp_is_txbox    = false;
    let mut sp_prst        = String::from("rect");
    let mut sp_x = 0i64; let mut sp_y = 0i64;
    let mut sp_w = 0i64; let mut sp_h = 0i64;
    let mut sp_fill_color: Option<String> = None;
    let mut sp_no_fill     = false;
    let mut sp_line_w:     Option<i64>   = None;
    let mut sp_line_color: Option<String> = None;
    let mut sp_wrap:   Option<bool>     = None;
    let mut sp_valign: Option<VertAlign> = None;

    // collected paragraphs: (align, runs)
    let mut sp_paragraphs: Vec<(Option<HorizAlign>, Vec<RunAccum>)> = Vec::new();
    let mut sp_cur_align:  Option<HorizAlign> = None;
    let mut sp_cur_runs:   Vec<RunAccum>      = Vec::new();
    let mut cur_run:       Option<RunAccum>   = None;

    // Shape sub-context flags
    let mut in_sp_pr    = false;
    let mut in_xfrm     = false;  // a:xfrm inside spPr (sp or pic)
    let mut in_ln       = false;  // a:ln inside spPr
    let mut in_tx_body  = false;  // p:txBody inside p:sp
    let mut in_paragraph = false; // a:p inside tx_body
    let mut in_run      = false;  // a:r inside paragraph
    let mut in_run_pr   = false;  // a:rPr inside run
    let mut in_run_text = false;  // a:t inside run

    // Picture accumulator
    let mut pic_x = 0i64; let mut pic_y = 0i64;
    let mut pic_w = 0i64; let mut pic_h = 0i64;
    let mut pic_rid    = String::new();
    let mut in_pic_sp_pr = false;

    // Table / graphicFrame accumulator
    let mut tbl_x = 0i64; let mut tbl_y = 0i64;
    let mut tbl_w = 0i64; let mut tbl_h = 0i64;
    let mut in_gframe_xfrm  = false;
    let mut in_tbl          = false;
    let mut in_tbl_grid     = false;
    let mut in_tr           = false;
    let mut in_tc           = false;
    let mut in_tc_body      = false;
    let mut in_tc_para      = false;
    let mut in_tc_run       = false;
    let mut in_tc_run_text  = false;
    let mut tbl_col_widths: Vec<i64>       = Vec::new();
    let mut tbl_rows:       Vec<Vec<TableCell>> = Vec::new();
    let mut tbl_cur_row:    Vec<TableCell> = Vec::new();
    let mut tbl_cur_cell:   String         = String::new();

    // Chart detection inside graphicFrame
    let mut gframe_idx:        usize         = 0;
    let mut gframe_chart_path: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {

            // ──────────────────────────────────────────────────────────────────
            Ok(Event::Start(ref e)) => {
                let local = e.name();
                let local = local.as_ref();

                match local {

                    // ── Background ────────────────────────────────────────────
                    b"p:bg" => { in_bg = true; }

                    b"a:solidFill" if in_bg && !in_sp && !in_pic && !in_gframe => {
                        fill_ctx = FillCtx::Bg;
                    }

                    // ── Shape (p:sp) ──────────────────────────────────────────
                    b"p:sp" => {
                        in_sp = true;
                        // reset all shape state
                        sp_is_txbox = false;
                        sp_prst = "rect".into();
                        sp_x = 0; sp_y = 0; sp_w = 0; sp_h = 0;
                        sp_fill_color = None;
                        sp_no_fill = false;
                        sp_line_w = None;
                        sp_line_color = None;
                        sp_wrap = None;
                        sp_valign = None;
                        sp_paragraphs.clear();
                        sp_cur_align = None;
                        sp_cur_runs.clear();
                        cur_run = None;
                        in_sp_pr = false;
                        in_xfrm = false;
                        in_ln = false;
                        in_tx_body = false;
                        in_paragraph = false;
                        in_run = false;
                        in_run_pr = false;
                        in_run_text = false;
                        fill_ctx = FillCtx::None;
                    }

                    b"p:cNvSpPr" if in_sp => {
                        if attr_str(e, b"txBox").as_deref() == Some("1") {
                            sp_is_txbox = true;
                        }
                    }

                    b"p:spPr" if in_sp  => { in_sp_pr    = true; }
                    b"p:spPr" if in_pic => { in_pic_sp_pr = true; }

                    b"a:xfrm" if in_sp_pr || in_pic_sp_pr => { in_xfrm = true; }

                    b"a:prstGeom" if in_sp_pr => {
                        if let Some(p) = attr_str(e, b"prst") { sp_prst = p; }
                    }

                    b"a:solidFill" if in_sp_pr && !in_ln => { fill_ctx = FillCtx::SpPr; }

                    b"a:ln" if in_sp_pr => {
                        in_ln = true;
                        sp_line_w = attr_i64(e, b"w");
                    }
                    b"a:solidFill" if in_ln => { fill_ctx = FillCtx::Line; }

                    // a:srgbClr in a solidFill context (Start variant — has child a:alpha etc.)
                    b"a:srgbClr" => {
                        let val = attr_str(e, b"val");
                        match fill_ctx {
                            FillCtx::Bg    => bg_color_hex  = val,
                            FillCtx::SpPr  => sp_fill_color = val,
                            FillCtx::Line  => sp_line_color = val,
                            FillCtx::RunPr => {
                                if let Some(ref mut r) = cur_run { r.color = val; }
                            }
                            FillCtx::None  => {}
                        }
                    }

                    // ── Text body ─────────────────────────────────────────────
                    b"p:txBody" if in_sp => { in_tx_body = true; }

                    b"a:bodyPr" if in_tx_body => {
                        sp_wrap = attr_str(e, b"wrap").map(|w| w != "none");
                        sp_valign = attr_str(e, b"anchor").map(|a| match a.as_str() {
                            "ctr" => VertAlign::Middle,
                            "b"   => VertAlign::Bottom,
                            _     => VertAlign::Top,
                        });
                    }

                    b"a:p" if in_tx_body => {
                        in_paragraph = true;
                        sp_cur_align = None;
                        sp_cur_runs.clear();
                    }

                    b"a:pPr" if in_paragraph => {
                        sp_cur_align = attr_str(e, b"algn").map(|a| match a.as_str() {
                            "ctr"  => HorizAlign::Center,
                            "r"    => HorizAlign::Right,
                            "just" => HorizAlign::Justify,
                            _      => HorizAlign::Left,
                        });
                    }

                    b"a:r" if in_paragraph => {
                        in_run = true;
                        cur_run = Some(RunAccum::default());
                    }

                    b"a:rPr" if in_run => {
                        in_run_pr = true;
                        if let Some(ref mut r) = cur_run {
                            if let Some(sz) = attr_i64(e, b"sz") {
                                r.font_size = Some(sz as f64 / 100.0);
                            }
                            r.bold   = attr_str(e, b"b").map(|v| v == "1");
                            r.italic = attr_str(e, b"i").map(|v| v == "1");
                            r.lang   = attr_str(e, b"lang");
                        }
                    }

                    b"a:solidFill" if in_run_pr => { fill_ctx = FillCtx::RunPr; }

                    b"a:t" if in_run     => { in_run_text    = true; }
                    b"a:t" if in_tc_run  => { in_tc_run_text = true; }

                    // ── Picture (p:pic) ───────────────────────────────────────
                    b"p:pic" => {
                        in_pic = true;
                        pic_x = 0; pic_y = 0; pic_w = 0; pic_h = 0;
                        pic_rid.clear();
                        in_pic_sp_pr = false;
                    }

                    b"a:blip" if in_pic => {
                        // r:embed — try both prefixed and full-ns key
                        pic_rid = e.attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| a.key.as_ref().ends_with(b"embed"))
                            .and_then(|a| String::from_utf8(a.value.into_owned()).ok())
                            .unwrap_or_default();
                    }

                    // ── GraphicFrame (table or chart) ─────────────────────────
                    b"p:graphicFrame" => {
                        in_gframe = true;
                        tbl_x = 0; tbl_y = 0; tbl_w = 0; tbl_h = 0;
                        tbl_col_widths.clear();
                        tbl_rows.clear();
                        tbl_cur_row.clear();
                        tbl_cur_cell.clear();
                        in_gframe_xfrm = false;
                        in_tbl = false; in_tbl_grid = false;
                        in_tr = false; in_tc = false;
                        in_tc_body = false; in_tc_para = false;
                        in_tc_run = false; in_tc_run_text = false;
                        gframe_chart_path = None;
                    }

                    b"p:xfrm" if in_gframe => { in_gframe_xfrm = true; }

                    b"a:tbl"    if in_gframe           => { in_tbl      = true; }
                    b"a:tblGrid" if in_tbl             => { in_tbl_grid = true; }

                    b"a:tr" if in_tbl => {
                        in_tr = true;
                        tbl_cur_row.clear();
                    }
                    b"a:tc" if in_tr => {
                        in_tc = true;
                        tbl_cur_cell.clear();
                    }
                    b"a:txBody" if in_tc   => { in_tc_body = true; }
                    b"a:p"      if in_tc_body => { in_tc_para = true; }
                    b"a:r"      if in_tc_para => { in_tc_run = true; }

                    // a:off / a:ext under any active xfrm
                    b"a:off" => {
                        let x = attr_i64(e, b"x").unwrap_or(0);
                        let y = attr_i64(e, b"y").unwrap_or(0);
                        if      in_xfrm        && in_sp  { sp_x = x; sp_y = y; }
                        else if in_xfrm        && in_pic { pic_x = x; pic_y = y; }
                        else if in_gframe_xfrm           { tbl_x = x; tbl_y = y; }
                    }
                    b"a:ext" => {
                        let cx = attr_i64(e, b"cx").unwrap_or(0);
                        let cy = attr_i64(e, b"cy").unwrap_or(0);
                        if      in_xfrm        && in_sp  { sp_w = cx; sp_h = cy; }
                        else if in_xfrm        && in_pic { pic_w = cx; pic_h = cy; }
                        else if in_gframe_xfrm           { tbl_w = cx; tbl_h = cy; }
                    }

                    _ => {}
                }
            }

            // ──────────────────────────────────────────────────────────────────
            Ok(Event::Empty(ref e)) => {
                let local = e.name();
                let local = local.as_ref();

                match local {
                    // coordinates (self-closing form)
                    b"a:off" => {
                        let x = attr_i64(e, b"x").unwrap_or(0);
                        let y = attr_i64(e, b"y").unwrap_or(0);
                        if      in_xfrm        && in_sp  { sp_x = x; sp_y = y; }
                        else if in_xfrm        && in_pic { pic_x = x; pic_y = y; }
                        else if in_gframe_xfrm           { tbl_x = x; tbl_y = y; }
                    }
                    b"a:ext" => {
                        let cx = attr_i64(e, b"cx").unwrap_or(0);
                        let cy = attr_i64(e, b"cy").unwrap_or(0);
                        if      in_xfrm        && in_sp  { sp_w = cx; sp_h = cy; }
                        else if in_xfrm        && in_pic { pic_w = cx; pic_h = cy; }
                        else if in_gframe_xfrm           { tbl_w = cx; tbl_h = cy; }
                    }

                    // shape detection
                    b"p:cNvSpPr" if in_sp => {
                        if attr_str(e, b"txBox").as_deref() == Some("1") {
                            sp_is_txbox = true;
                        }
                    }
                    b"a:noFill" if in_sp_pr => { sp_no_fill = true; }

                    b"a:prstGeom" if in_sp_pr => {
                        if let Some(p) = attr_str(e, b"prst") { sp_prst = p; }
                    }
                    b"a:ln" if in_sp_pr => {
                        // self-closing line — capture width only (no fill)
                        sp_line_w = attr_i64(e, b"w");
                    }

                    // srgbClr (self-closing)
                    b"a:srgbClr" => {
                        let val = attr_str(e, b"val");
                        match fill_ctx {
                            FillCtx::Bg    => bg_color_hex  = val,
                            FillCtx::SpPr  => sp_fill_color = val,
                            FillCtx::Line  => sp_line_color = val,
                            FillCtx::RunPr => {
                                if let Some(ref mut r) = cur_run { r.color = val; }
                            }
                            FillCtx::None  => {}
                        }
                    }

                    // run properties (self-closing rPr)
                    b"a:rPr" if in_run => {
                        if let Some(ref mut r) = cur_run {
                            if let Some(sz) = attr_i64(e, b"sz") {
                                r.font_size = Some(sz as f64 / 100.0);
                            }
                            r.bold   = attr_str(e, b"b").map(|v| v == "1");
                            r.italic = attr_str(e, b"i").map(|v| v == "1");
                            r.lang   = attr_str(e, b"lang");
                        }
                    }

                    // paragraph props (self-closing pPr)
                    b"a:pPr" if in_paragraph => {
                        sp_cur_align = attr_str(e, b"algn").map(|a| match a.as_str() {
                            "ctr"  => HorizAlign::Center,
                            "r"    => HorizAlign::Right,
                            "just" => HorizAlign::Justify,
                            _      => HorizAlign::Left,
                        });
                    }

                    // body props (self-closing bodyPr)
                    b"a:bodyPr" if in_tx_body => {
                        sp_wrap = attr_str(e, b"wrap").map(|w| w != "none");
                        sp_valign = attr_str(e, b"anchor").map(|a| match a.as_str() {
                            "ctr" => VertAlign::Middle,
                            "b"   => VertAlign::Bottom,
                            _     => VertAlign::Top,
                        });
                    }

                    // blip (self-closing)
                    b"a:blip" if in_pic => {
                        pic_rid = e.attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| a.key.as_ref().ends_with(b"embed"))
                            .and_then(|a| String::from_utf8(a.value.into_owned()).ok())
                            .unwrap_or_default();
                    }

                    // chart reference inside a graphicFrame (self-closing)
                    b"c:chart" if in_gframe => {
                        if let Some(rid) = e.attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| {
                                let k = a.key.as_ref();
                                k == b"r:id" || k.ends_with(b":id")
                            })
                            .and_then(|a| String::from_utf8(a.value.into_owned()).ok())
                        {
                            if let Some(target) = rel_map.get(&rid) {
                                // target may be relative ("../charts/chart1.xml")
                                // or absolute ("/ppt/charts/chart1.xml")
                                let chart_path = resolve_rel_target("ppt/slides", target);
                                gframe_chart_path = Some(chart_path);
                            }
                        }
                    }

                    // table grid column
                    b"a:gridCol" if in_tbl_grid => {
                        if let Some(w) = attr_i64(e, b"w") {
                            tbl_col_widths.push(w);
                        }
                    }

                    _ => {}
                }
            }

            // ──────────────────────────────────────────────────────────────────
            Ok(Event::Text(ref e)) => {
                let t = e.unescape().unwrap_or_default();
                if in_run_text {
                    if let Some(ref mut r) = cur_run { r.text.push_str(&t); }
                } else if in_tc_run_text {
                    tbl_cur_cell.push_str(&t);
                }
            }

            // ──────────────────────────────────────────────────────────────────
            Ok(Event::End(ref e)) => {
                let local = e.name();
                let local = local.as_ref();

                match local {
                    // ── fill / sub-context closes ─────────────────────────────
                    b"a:solidFill" => { fill_ctx = FillCtx::None; }
                    b"a:ln"        => { in_ln = false; }
                    b"a:xfrm"      => { in_xfrm = false; }
                    b"a:rPr"       => { in_run_pr = false; }
                    b"a:t"         => { in_run_text = false; in_tc_run_text = false; }

                    // ── text body ─────────────────────────────────────────────
                    b"a:r" => {
                        if in_run {
                            if let Some(run) = cur_run.take() {
                                sp_cur_runs.push(run);
                            }
                            in_run = false;
                            in_run_text = false;
                        } else if in_tc_run {
                            in_tc_run = false;
                            in_tc_run_text = false;
                        }
                    }

                    b"a:p" => {
                        if in_paragraph {
                            sp_paragraphs.push((
                                sp_cur_align.take(),
                                sp_cur_runs.drain(..).collect(),
                            ));
                            in_paragraph = false;
                        } else if in_tc_para {
                            in_tc_para = false;
                        }
                    }

                    b"p:txBody" if in_sp     => { in_tx_body  = false; }
                    b"a:txBody" if in_tc      => { in_tc_body  = false; }
                    b"p:spPr"   if in_sp      => { in_sp_pr    = false; }
                    b"p:spPr"   if in_pic     => { in_pic_sp_pr = false; }
                    b"p:xfrm"   if in_gframe  => { in_gframe_xfrm = false; }
                    b"a:tblGrid"              => { in_tbl_grid = false; }

                    // ── table cell / row ──────────────────────────────────────
                    b"a:tc" if in_tr => {
                        tbl_cur_row.push(TableCell::Text(tbl_cur_cell.clone()));
                        tbl_cur_cell.clear();
                        in_tc = false;
                        in_tc_body = false;
                    }
                    b"a:tr" if in_tbl => {
                        tbl_rows.push(tbl_cur_row.drain(..).collect());
                        in_tr = false;
                    }
                    b"a:tbl" => { in_tbl = false; }

                    // ── background close ──────────────────────────────────────
                    b"p:bg" => {
                        if let Some(color) = bg_color_hex.take() {
                            slide.background.color = Some(color);
                        }
                        in_bg = false;
                        fill_ctx = FillCtx::None;
                    }

                    // ── shape close ───────────────────────────────────────────
                    b"p:sp" if in_sp => {
                        let has_text = sp_paragraphs.iter()
                            .any(|(_, runs)| runs.iter().any(|r| !r.text.is_empty()));

                        if sp_is_txbox || (has_text && sp_no_fill) {
                            // ── Text element ──────────────────────────────────
                            slide.elements.push(build_text_element(
                                &sp_paragraphs,
                                emu_to_position(sp_x, sp_y, sp_w, sp_h),
                                sp_wrap,
                                sp_valign.take(),
                            ));
                        } else {
                            // ── Shape element ─────────────────────────────────
                            let fill = sp_fill_color.as_ref().map(|c| {
                                crate::model::elements::ShapeFill {
                                    color: Some(c.clone()),
                                    transparency: None,
                                }
                            });
                            let line = if sp_line_w.is_some() || sp_line_color.is_some() {
                                Some(LineOptions {
                                    color: sp_line_color.clone(),
                                    width: sp_line_w.map(|w| w as f64 / 12_700.0),
                                    ..Default::default()
                                })
                            } else {
                                None
                            };
                            let text = has_text.then(|| flatten_para_text(&sp_paragraphs));

                            slide.elements.push(SlideElement::Shape {
                                shape_type: sp_prst.clone(),
                                options: ShapeOptions {
                                    pos: emu_to_position(sp_x, sp_y, sp_w, sp_h),
                                    fill,
                                    line,
                                    text,
                                    ..Default::default()
                                },
                            });
                        }

                        in_sp = false;
                        fill_ctx = FillCtx::None;
                    }

                    // ── picture close ─────────────────────────────────────────
                    b"p:pic" if in_pic => {
                        let image_data = if !pic_rid.is_empty() {
                            if let Some(target) = rel_map.get(&pic_rid) {
                                let full = resolve_rel_target("ppt/slides", target);
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
                                pos: emu_to_position(pic_x, pic_y, pic_w, pic_h),
                                data: image_data,
                                ..Default::default()
                            },
                        });

                        in_pic = false;
                        in_pic_sp_pr = false;
                        fill_ctx = FillCtx::None;
                    }

                    // ── graphicFrame close ────────────────────────────────────
                    b"p:graphicFrame" if in_gframe => {
                        if !tbl_rows.is_empty() {
                            // ── Table ──────────────────────────────────────────
                            let col_w = if !tbl_col_widths.is_empty() {
                                let px: Vec<f64> = tbl_col_widths
                                    .iter()
                                    .map(|&w| w as f64 / 9_525.0)
                                    .collect();
                                Some(crate::model::elements::ColRowSizes::PerColumn(px))
                            } else {
                                None
                            };

                            slide.elements.push(SlideElement::Table {
                                data: tbl_rows.drain(..).collect(),
                                options: TableOptions {
                                    x: Some(CoordVal::Pixels(tbl_x as f64 / 9_525.0)),
                                    y: Some(CoordVal::Pixels(tbl_y as f64 / 9_525.0)),
                                    w: Some(CoordVal::Pixels(tbl_w as f64 / 9_525.0)),
                                    h: Some(CoordVal::Pixels(tbl_h as f64 / 9_525.0)),
                                    col_w,
                                    ..Default::default()
                                },
                                frame_index: Some(gframe_idx),
                                raw_frame_xml: None, // filled in by associate_raw_frames()
                                modified: false,
                            });
                            gframe_idx += 1;
                        } else if let Some(chart_path) = gframe_chart_path.take() {
                            // ── Chart ──────────────────────────────────────────
                            let (chart_type, chart_data) =
                                parse_chart_xml_from_archive(archive, &chart_path)
                                    .unwrap_or_else(|| (ChartType::Bar, vec![]));

                            slide.elements.push(SlideElement::Chart {
                                chart_type,
                                data: chart_data,
                                combo_types: vec![],
                                options: ChartOptions {
                                    pos: emu_to_position(tbl_x, tbl_y, tbl_w, tbl_h),
                                    ..Default::default()
                                },
                                source_chart_path: Some(chart_path),
                                frame_index: Some(gframe_idx),
                                modified: false,
                            });
                            gframe_idx += 1;
                        }

                        in_gframe = false;
                        fill_ctx = FillCtx::None;
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

    Ok(slide)
}

// ── Element builders ──────────────────────────────────────────────────────────

/// Turn the collected paragraphs into a `SlideElement::Text`.
///
/// - Single paragraph, single run → `TextContent::Plain`
/// - Multiple paragraphs or runs  → `TextContent::Runs` (with `\n` between paragraphs)
fn build_text_element(
    paragraphs: &[(Option<HorizAlign>, Vec<RunAccum>)],
    pos: Position,
    wrap: Option<bool>,
    valign: Option<VertAlign>,
) -> SlideElement {
    // Derive global properties from the first run
    let first = paragraphs.iter().flat_map(|(_, rs)| rs.iter()).next();
    let font_size = first.and_then(|r| r.font_size);
    let bold      = first.and_then(|r| r.bold);
    let italic    = first.and_then(|r| r.italic);
    let color     = first.and_then(|r| r.color.clone());
    let align     = paragraphs.first().and_then(|(a, _)| a.clone());

    let total_runs: usize = paragraphs.iter().map(|(_, rs)| rs.len()).sum();
    let total_paras = paragraphs.len();

    let (text_content, effective_align) = if total_paras <= 1 && total_runs <= 1 {
        // Simple: plain string
        let text = paragraphs.iter()
            .flat_map(|(_, rs)| rs.iter())
            .map(|r| r.text.as_str())
            .collect::<Vec<_>>()
            .join("");
        (TextContent::Plain(text), align)
    } else {
        // Rich: runs with paragraph breaks
        let mut runs: Vec<TextRun> = Vec::new();
        for (para_idx, (_para_align, para_runs)) in paragraphs.iter().enumerate() {
            for run in para_runs {
                runs.push(TextRun {
                    text: run.text.clone(),
                    options: Some(TextRunOptions {
                        font_size: run.font_size,
                        bold:      run.bold,
                        italic:    run.italic,
                        color:     run.color.clone(),
                        lang:      run.lang.clone(),
                        ..Default::default()
                    }),
                });
            }
            // paragraph separator (except after the last one)
            if para_idx + 1 < total_paras {
                runs.push(TextRun { text: "\n".into(), options: None });
            }
        }
        (TextContent::Runs(runs), align)
    };

    SlideElement::Text {
        text: text_content,
        options: TextOptions {
            pos,
            font_size,
            bold,
            italic,
            color,
            align: effective_align,
            valign,
            wrap,
            ..Default::default()
        },
    }
}

/// Join all paragraph text into a single `TextContent::Plain` (used for shape text).
fn flatten_para_text(paragraphs: &[(Option<HorizAlign>, Vec<RunAccum>)]) -> TextContent {
    let text = paragraphs.iter()
        .map(|(_, rs)| rs.iter().map(|r| r.text.as_str()).collect::<Vec<_>>().join(""))
        .collect::<Vec<_>>()
        .join("\n");
    TextContent::Plain(text)
}

// ── Low-level helpers ─────────────────────────────────────────────────────────

fn emu_to_position(x: i64, y: i64, w: i64, h: i64) -> Position {
    // 1 px (96 DPI) = 9 525 EMU
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

/// Resolve `../` and `.` components in a path string
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

/// Resolve a relationship target against a base directory.
///
/// Handles both:
/// - relative targets: `../media/image1.png`  → resolved relative to `base_dir`
/// - absolute targets: `/ppt/charts/chart1.xml` → strip leading `/`, use directly
fn resolve_rel_target(base_dir: &str, target: &str) -> String {
    if target.starts_with('/') {
        // Absolute pack URI — strip leading slash, use as-is
        target.trim_start_matches('/').to_string()
    } else {
        normalize_path(&format!("{}/{}", base_dir, target))
    }
}

// ── Chart XML parser ──────────────────────────────────────────────────────────

/// Parse a chart XML file from the archive and return (chart_type, series_data).
fn parse_chart_xml_from_archive(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    chart_path: &str,
) -> Option<(ChartType, Vec<ChartData>)> {
    let xml = read_zip_entry(archive, chart_path).ok()?;
    parse_chart_xml_str(&xml)
}

fn parse_chart_xml_str(xml: &str) -> Option<(ChartType, Vec<ChartData>)> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut chart_type: Option<ChartType> = None;
    let mut series: Vec<ChartData> = Vec::new();

    // Current series accumulator
    let mut cur_series: Option<ChartData> = None;

    // State flags
    let mut in_ser     = false;
    let mut in_tx      = false;  // series name
    let mut in_cat     = false;  // category labels
    let mut in_val     = false;  // values
    let mut in_str_cache = false;
    let mut in_num_cache = false;
    let mut in_pt      = false;
    let mut in_v       = false;
    let mut cur_pt_idx: u32 = 0;

    // Temporary accumulation for current string/numeric cache
    let mut cat_labels: Vec<(u32, String)> = Vec::new();
    let mut val_pts:    Vec<(u32, f64)>    = Vec::new();
    let mut tx_name:    String             = String::new();
    let mut _in_tx_v   = false;
    let mut cur_text   = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let local = e.name();
                let local = local.as_ref();
                match local {
                    // Detect chart type
                    b"c:barChart" | b"c:bar3DChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Bar); }
                    }
                    b"c:lineChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Line); }
                    }
                    b"c:pieChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Pie); }
                    }
                    b"c:areaChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Area); }
                    }
                    b"c:doughnutChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Doughnut); }
                    }
                    b"c:radarChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Radar); }
                    }
                    b"c:scatterChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Scatter); }
                    }
                    b"c:bubbleChart" => {
                        if chart_type.is_none() { chart_type = Some(ChartType::Bubble); }
                    }

                    b"c:ser" => {
                        in_ser = true;
                        cur_series = Some(ChartData::default());
                        cat_labels.clear();
                        val_pts.clear();
                        tx_name.clear();
                    }
                    b"c:tx"  if in_ser => { in_tx  = true; _in_tx_v = false; }
                    b"c:cat" if in_ser => { in_cat = true; }
                    b"c:val" if in_ser => { in_val = true; }

                    b"c:strCache"  if in_tx || in_cat => { in_str_cache = true; }
                    b"c:numCache"  if in_val           => { in_num_cache = true; }

                    b"c:pt" if in_str_cache || in_num_cache => {
                        in_pt = true;
                        cur_pt_idx = attr_str(e, b"idx")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                    }
                    b"c:v" if in_pt => { in_v = true; cur_text.clear(); }

                    _ => {}
                }
            }

            Ok(Event::Text(ref e)) => {
                if in_v {
                    let t = e.unescape().unwrap_or_default();
                    cur_text.push_str(&t);
                }
            }

            Ok(Event::End(ref e)) => {
                let local = e.name();
                let local = local.as_ref();
                match local {
                    b"c:v" => {
                        if in_v {
                            if in_str_cache {
                                if in_tx {
                                    tx_name = cur_text.clone();
                                } else if in_cat {
                                    cat_labels.push((cur_pt_idx, cur_text.clone()));
                                }
                            } else if in_num_cache {
                                if let Ok(v) = cur_text.parse::<f64>() {
                                    val_pts.push((cur_pt_idx, v));
                                }
                            }
                            in_v = false;
                            cur_text.clear();
                        }
                    }
                    b"c:pt"       => { in_pt = false; }
                    b"c:strCache" => { in_str_cache = false; }
                    b"c:numCache" => { in_num_cache = false; }
                    b"c:tx"  => { in_tx = false; _in_tx_v = false; }
                    b"c:cat" => { in_cat = false; }
                    b"c:val" => { in_val = false; }

                    b"c:ser" => {
                        if let Some(mut s) = cur_series.take() {
                            s.name = if tx_name.is_empty() { None } else { Some(tx_name.clone()) };

                            // Sort by index and collect labels
                            cat_labels.sort_by_key(|(i, _)| *i);
                            val_pts.sort_by_key(|(i, _)| *i);

                            if !cat_labels.is_empty() {
                                s.labels = Some(cat_labels.iter().map(|(_, l)| l.clone()).collect());
                            }
                            s.values = val_pts.iter().map(|(_, v)| *v).collect();
                            series.push(s);
                        }
                        in_ser = false;
                        cat_labels.clear();
                        val_pts.clear();
                        tx_name.clear();
                    }
                    _ => {}
                }
            }

            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    chart_type.map(|ct| (ct, series))
}

// ── Graphic frame raw XML extraction ─────────────────────────────────────────

/// Extract all `<p:graphicFrame>…</p:graphicFrame>` substrings from slide XML in order.
fn extract_graphic_frames(xml: &str) -> Vec<String> {
    let mut frames = Vec::new();
    let mut pos = 0;
    while pos < xml.len() {
        let search = &xml[pos..];
        if let Some(start_off) = search.find("<p:graphicFrame") {
            let abs_start = pos + start_off;
            let from = &xml[abs_start..];
            if let Some(end_off) = from.find("</p:graphicFrame>") {
                let abs_end = abs_start + end_off + "</p:graphicFrame>".len();
                frames.push(xml[abs_start..abs_end].to_string());
                pos = abs_end;
            } else {
                break;
            }
        } else {
            break;
        }
    }
    frames
}

/// Associate the raw frame XML with each Table/Chart element that has a `frame_index`.
fn associate_raw_frames(frames: &[String], slide: &mut Slide) {
    for el in &mut slide.elements {
        match el {
            SlideElement::Table { frame_index: Some(i), raw_frame_xml, .. } => {
                if let Some(frame) = frames.get(*i) {
                    *raw_frame_xml = Some(frame.clone());
                }
            }
            // Charts don't need raw_frame_xml — they use source_chart_path
            _ => {}
        }
    }
}
