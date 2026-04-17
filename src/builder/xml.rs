use std::collections::HashMap;
use std::io::{Cursor, Read as IoRead, Write as IoWrite};
use zip::{ZipArchive, ZipWriter, write::FileOptions};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as B64;

use crate::model::{
    elements::{
        BarDir, BarGrouping, ChartData, ChartType, HorizAlign, SlideElement,
        TextContent, TextOptions,
    },
    presentation::Presentation,
    slide::Slide,
};


pub fn build_pptx(pres: &Presentation) -> Result<Vec<u8>, String> {
    // If this presentation was loaded from a ZIP, use the passthrough builder
    // which preserves all original content and only replaces what changed.
    if let Some(ref zip_bytes) = pres.source_zip {
        return build_pptx_passthrough(pres, zip_bytes);
    }

    build_pptx_fresh(pres)
}

fn build_pptx_fresh(pres: &Presentation) -> Result<Vec<u8>, String> {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip = ZipWriter::new(cursor);

    let options: FileOptions<()> = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Track media files: (filename, bytes)
    let mut media: Vec<(String, Vec<u8>)> = Vec::new();
    // Track chart files: (zip_path, xml_content)
    let mut chart_files: Vec<(String, String)> = Vec::new();

    // Collect slide XMLs and their relationship files
    let slide_count = pres.slides.len();
    let mut slide_xmls: Vec<String> = Vec::new();
    let mut slide_rels: Vec<String> = Vec::new();

    let mut media_counter = 1u32;
    let mut chart_counter = 1u32;
    for (idx, slide) in pres.slides.iter().enumerate() {
        let (xml, rels, new_media, new_charts) =
            build_slide_xml(slide, idx, pres, &mut media_counter, &mut chart_counter);
        slide_xmls.push(xml);
        slide_rels.push(rels);
        media.extend(new_media);
        chart_files.extend(new_charts);
    }

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).map_err(|e| e.to_string())?;
    zip.write_all(build_content_types(slide_count, &media, &chart_files).as_bytes())
        .map_err(|e| e.to_string())?;

    // _rels/.rels
    zip.start_file("_rels/.rels", options).map_err(|e| e.to_string())?;
    zip.write_all(ROOT_RELS.as_bytes()).map_err(|e| e.to_string())?;

    // ppt/_rels/presentation.xml.rels
    zip.start_file("ppt/_rels/presentation.xml.rels", options)
        .map_err(|e| e.to_string())?;
    zip.write_all(build_pres_rels(slide_count).as_bytes())
        .map_err(|e| e.to_string())?;

    // ppt/presentation.xml
    zip.start_file("ppt/presentation.xml", options).map_err(|e| e.to_string())?;
    zip.write_all(build_presentation_xml(pres).as_bytes())
        .map_err(|e| e.to_string())?;

    // ppt/theme/theme1.xml (minimal)
    zip.start_file("ppt/theme/theme1.xml", options).map_err(|e| e.to_string())?;
    zip.write_all(MINIMAL_THEME.as_bytes()).map_err(|e| e.to_string())?;

    // ppt/slideLayouts/slideLayout1.xml (minimal)
    zip.start_file("ppt/slideLayouts/slideLayout1.xml", options)
        .map_err(|e| e.to_string())?;
    zip.write_all(MINIMAL_SLIDE_LAYOUT.as_bytes()).map_err(|e| e.to_string())?;

    zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", options)
        .map_err(|e| e.to_string())?;
    zip.write_all(SLIDE_LAYOUT_RELS.as_bytes()).map_err(|e| e.to_string())?;

    // ppt/slideMasters/slideMaster1.xml (minimal)
    zip.start_file("ppt/slideMasters/slideMaster1.xml", options)
        .map_err(|e| e.to_string())?;
    zip.write_all(MINIMAL_SLIDE_MASTER.as_bytes()).map_err(|e| e.to_string())?;

    zip.start_file("ppt/slideMasters/_rels/slideMaster1.xml.rels", options)
        .map_err(|e| e.to_string())?;
    zip.write_all(SLIDE_MASTER_RELS.as_bytes()).map_err(|e| e.to_string())?;

    // Slides
    for (idx, (xml, rels)) in slide_xmls.iter().zip(slide_rels.iter()).enumerate() {
        let slide_path = format!("ppt/slides/slide{}.xml", idx + 1);
        let rels_path = format!("ppt/slides/_rels/slide{}.xml.rels", idx + 1);
        zip.start_file(&slide_path, options).map_err(|e| e.to_string())?;
        zip.write_all(xml.as_bytes()).map_err(|e| e.to_string())?;
        zip.start_file(&rels_path, options).map_err(|e| e.to_string())?;
        zip.write_all(rels.as_bytes()).map_err(|e| e.to_string())?;
    }

    // Media
    for (name, bytes) in &media {
        zip.start_file(format!("ppt/media/{}", name), options)
            .map_err(|e| e.to_string())?;
        zip.write_all(bytes).map_err(|e| e.to_string())?;
    }

    // Charts
    for (zip_path, xml) in &chart_files {
        zip.start_file(zip_path, options).map_err(|e| e.to_string())?;
        zip.write_all(xml.as_bytes()).map_err(|e| e.to_string())?;
        // Empty rels for each chart
        let rels_path = chart_rels_path(zip_path);
        zip.start_file(&rels_path, options).map_err(|e| e.to_string())?;
        zip.write_all(EMPTY_RELS.as_bytes()).map_err(|e| e.to_string())?;
    }

    let inner = zip.finish().map_err(|e| e.to_string())?;
    Ok(inner.into_inner())
}

// ── Per-slide XML ─────────────────────────────────────────────────────────────

/// Returns (slide_xml, rels_xml, media_files, chart_files)
/// `chart_files` — `(zip_path, xml_content)` pairs, e.g. `("ppt/charts/chart1.xml", "…")`
fn build_slide_xml(
    slide: &Slide,
    _idx: usize,
    pres: &Presentation,
    media_counter: &mut u32,
    chart_counter: &mut u32,
) -> (String, String, Vec<(String, Vec<u8>)>, Vec<(String, String)>) {
    let (w_emu, h_emu) = pres.meta.layout.dimensions_emu();
    let mut shapes       = String::new();
    let mut rels_entries = String::new();
    let mut media:       Vec<(String, Vec<u8>)>  = Vec::new();
    let mut chart_files: Vec<(String, String)>   = Vec::new();
    let mut sp_id  = 2u32; // 1 is reserved for background
    let mut rel_id = 1u32;

    for element in &slide.elements {
        match element {
            SlideElement::Text { text, options } => {
                shapes.push_str(&build_text_shape(text, options, sp_id, w_emu, h_emu));
                sp_id += 1;
            }
            SlideElement::Image { options } => {
                let rid = format!("rId{}", rel_id);
                if let Some(data_b64) = &options.data {
                    if let Ok(bytes) = B64.decode(data_b64) {
                        let ext = detect_image_ext(&bytes);
                        let fname = format!("image{}.{}", *media_counter, ext);
                        rels_entries.push_str(&format!(
                            r#"<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{fname}"/>"#
                        ));
                        media.push((fname, bytes));
                        shapes.push_str(&build_pic_shape(options, sp_id, &rid, w_emu, h_emu));
                        sp_id += 1;
                        rel_id += 1;
                        *media_counter += 1;
                    }
                }
            }
            SlideElement::Shape { shape_type, options } => {
                shapes.push_str(&build_shape(shape_type, options, sp_id, w_emu, h_emu));
                sp_id += 1;
            }
            SlideElement::Table { data, options, raw_frame_xml, .. } => {
                // If we have a preserved raw frame, use it directly (style-preserving path)
                if let Some(raw) = raw_frame_xml {
                    shapes.push_str(raw);
                } else {
                    shapes.push_str(&build_table(data, options, sp_id, w_emu, h_emu));
                }
                sp_id += 1;
            }
            SlideElement::Chart { chart_type, data, options, .. } => {
                let rid        = format!("rId{}", rel_id);
                let chart_name = format!("chart{}.xml", *chart_counter);
                let zip_path   = format!("ppt/charts/{}", chart_name);
                let rel_target = format!("../charts/{}", chart_name);

                rels_entries.push_str(&format!(
                    r#"<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{rel_target}"/>"#
                ));
                shapes.push_str(&build_chart_frame(options, sp_id, &rid, w_emu, h_emu));

                let chart_xml = build_chart_xml_from_data(chart_type, data, options);
                chart_files.push((zip_path, chart_xml));

                sp_id += 1;
                rel_id += 1;
                *chart_counter += 1;
            }
            SlideElement::Notes { .. } => {
                // Notes are stored at slide level; skipped here
            }
        }
    }

    // Background fill
    let bg_color = slide
        .background
        .color
        .as_deref()
        .unwrap_or("FFFFFF");
    let bg_xml = format!(
        r#"<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"#
    );

    let slide_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>{bg_xml}<p:spTree>
    <p:nvGrpSpPr>
      <p:cNvPr id="1" name=""/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="{w_emu}" cy="{h_emu}"/>
        <a:chOff x="0" y="0"/><a:chExt cx="{w_emu}" cy="{h_emu}"/></a:xfrm>
    </p:grpSpPr>
    {shapes}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"#
    );

    // Base relationship to slide layout
    let rels_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  {rels_entries}
</Relationships>"#
    );

    (slide_xml, rels_xml, media, chart_files)
}

// ── Shape builders ────────────────────────────────────────────────────────────

fn build_text_shape(
    text: &TextContent,
    opts: &TextOptions,
    id: u32,
    slide_w: i64,
    slide_h: i64,
) -> String {
    let (x, y, cx, mut cy) = pos_emu(&opts.pos, slide_w, slide_h);
    if cy == 0 {
        // h was omitted — estimate one line at 1.5× the font size.
        // font_pt / 72 inches × 914 400 EMU/inch × 1.5 line factor
        let font_pt = opts.font_size.unwrap_or(18.0);
        cy = (font_pt / 72.0 * 914_400.0 * 1.5) as i64;
    }
    let font_size_hp = opts.font_size.unwrap_or(18.0) as i64 * 100; // half-points × 2 = hundredths of a point
    let bold = if opts.bold.unwrap_or(false) { "1" } else { "0" };
    let italic = if opts.italic.unwrap_or(false) { "1" } else { "0" };
    let color = opts.color.as_deref().unwrap_or("000000");
    let align = match opts.align.as_ref().unwrap_or(&HorizAlign::Left) {
        HorizAlign::Left    => "l",
        HorizAlign::Center  => "ctr",
        HorizAlign::Right   => "r",
        HorizAlign::Justify => "just",
    };

    let _text_str = match text {
        TextContent::Plain(s) => s.as_str(),
        TextContent::Runs(_)  => "", // runs handled separately
    };

    // Build paragraph content
    let para_content = match text {
        TextContent::Plain(s) => format!(
            r#"<a:r><a:rPr lang="en-US" sz="{font_size_hp}" b="{bold}" i="{italic}" dirty="0">
                 <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
               </a:rPr><a:t>{}</a:t></a:r>"#,
            xml_escape(s)
        ),
        TextContent::Runs(runs) => {
            let mut out = String::new();
            for run in runs {
                let rs = run.options.as_ref();
                let sz = rs.and_then(|o| o.font_size).unwrap_or(opts.font_size.unwrap_or(18.0)) as i64 * 100;
                let b = if rs.and_then(|o| o.bold).unwrap_or(opts.bold.unwrap_or(false)) { "1" } else { "0" };
                let i = if rs.and_then(|o| o.italic).unwrap_or(opts.italic.unwrap_or(false)) { "1" } else { "0" };
                let c = rs.and_then(|o| o.color.as_deref())
                    .or(opts.color.as_deref())
                    .unwrap_or("000000");
                out.push_str(&format!(
                    r#"<a:r><a:rPr lang="en-US" sz="{sz}" b="{b}" i="{i}" dirty="0">
                         <a:solidFill><a:srgbClr val="{c}"/></a:solidFill>
                       </a:rPr><a:t>{}</a:t></a:r>"#,
                    xml_escape(&run.text)
                ));
            }
            out
        }
    };

    let wrap = if opts.wrap.unwrap_or(true) { "square" } else { "none" };

    format!(
        r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{id}" name="TextBox {id}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="{wrap}" rtlCol="0">
      <a:normAutofit/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="{align}"/>
      {para_content}
    </a:p>
  </p:txBody>
</p:sp>"#
    )
}

fn build_pic_shape(
    opts: &crate::model::elements::ImageOptions,
    id: u32,
    rid: &str,
    slide_w: i64,
    slide_h: i64,
) -> String {
    let (x, y, cx, cy) = pos_emu(&opts.pos, slide_w, slide_h);
    format!(
        r#"<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="{id}" name="Image {id}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rid}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"#
    )
}

fn build_shape(
    shape_type: &str,
    opts: &crate::model::elements::ShapeOptions,
    id: u32,
    slide_w: i64,
    slide_h: i64,
) -> String {
    let (x, y, cx, cy) = pos_emu(&opts.pos, slide_w, slide_h);
    let prst = pptx_shape_name(shape_type);
    let fill_xml = match &opts.fill {
        Some(f) => {
            let color = f.color.as_deref().unwrap_or("4472C4");
            format!(r#"<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>"#)
        }
        None => "<a:noFill/>".into(),
    };
    let line_xml = match &opts.line {
        Some(l) => {
            let color = l.color.as_deref().unwrap_or("000000");
            let w = ((l.width.unwrap_or(1.0)) * 12_700.0) as i64;
            format!(r#"<a:ln w="{w}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>"#)
        }
        None => String::new(),
    };

    format!(
        r#"<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{id}" name="Shape {id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>
    {fill_xml}
    {line_xml}
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>"#
    )
}

fn build_table(
    data: &[Vec<crate::model::elements::TableCell>],
    opts: &crate::model::elements::TableOptions,
    id: u32,
    slide_w: i64,
    slide_h: i64,
) -> String {
    let x = opts.x.as_ref().map(|v| v.to_emu(slide_w)).unwrap_or(914_400);
    let y = opts.y.as_ref().map(|v| v.to_emu(slide_h)).unwrap_or(914_400);
    let w = opts.w.as_ref().map(|v| v.to_emu(slide_w)).unwrap_or(7_315_200);
    let h = opts.h.as_ref().map(|v| v.to_emu(slide_h)).unwrap_or(914_400);

    let _row_count = data.len();
    let col_count = data.first().map(|r| r.len()).unwrap_or(0);

    // Column widths — evenly distribute if not specified
    // colW values are in pixels (96 DPI); 1 px = 9 525 EMU
    let col_w_emu: Vec<i64> = match &opts.col_w {
        Some(crate::model::elements::ColRowSizes::Uniform(v)) => {
            vec![(*v * 9_525.0) as i64; col_count]
        }
        Some(crate::model::elements::ColRowSizes::PerColumn(vs)) => {
            vs.iter().map(|v| (*v * 9_525.0) as i64).collect()
        }
        None => vec![w / col_count.max(1) as i64; col_count],
    };

    let cols_xml: String = col_w_emu
        .iter()
        .map(|cw| format!(r#"<a:gridCol w="{}"/>"#, cw))
        .collect();

    let mut rows_xml = String::new();
    for row in data {
        let mut cells_xml = String::new();
        for cell in row {
            let text = match cell {
                crate::model::elements::TableCell::Text(s) => s.clone(),
                crate::model::elements::TableCell::Rich(r) => match &r.text {
                    TextContent::Plain(s) => s.clone(),
                    TextContent::Runs(runs) => runs.iter().map(|r| r.text.as_str()).collect::<Vec<_>>().join(""),
                },
            };
            cells_xml.push_str(&format!(
                r#"<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>
                     <a:p><a:r><a:t>{}</a:t></a:r></a:p>
                   </a:txBody><a:tcPr/></a:tc>"#,
                xml_escape(&text)
            ));
        }
        rows_xml.push_str(&format!("<a:tr h=\"457200\">{cells_xml}</a:tr>"));
    }

    format!(
        r#"<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="{id}" name="Table {id}"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1">
          <a:tableStyleId>{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}</a:tableStyleId>
        </a:tblPr>
        <a:tblGrid>{cols_xml}</a:tblGrid>
        {rows_xml}
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#
    )
}

fn build_chart_frame(
    opts: &crate::model::elements::ChartOptions,
    id: u32,
    rid: &str,
    slide_w: i64,
    slide_h: i64,
) -> String {
    let (x, y, cx, cy) = pos_emu(&opts.pos, slide_w, slide_h);
    format!(
        r#"<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="{id}" name="Chart {id}"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
               r:id="{rid}"/>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>"#
    )
}

// ── Presentation-level XML ────────────────────────────────────────────────────

fn build_presentation_xml(pres: &Presentation) -> String {
    let (w, h) = pres.meta.layout.dimensions_emu();
    let slide_refs: String = (0..pres.slides.len())
        .map(|i| format!(r#"<p:sldId id="{}" r:id="rId{}"/>"#, 256 + i, i + 2))
        .collect();

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>{slide_refs}</p:sldIdLst>
  <p:sldSz cx="{w}" cy="{h}" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#
    )
}

fn build_pres_rels(slide_count: usize) -> String {
    let mut rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>"#.to_string();

    for i in 0..slide_count {
        rels.push_str(&format!(
            r#"
  <Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
            i + 2,
            i + 1
        ));
    }
    rels.push_str("\n</Relationships>");
    rels
}

fn build_content_types(
    slide_count: usize,
    _media: &[(String, Vec<u8>)],
    chart_files: &[(String, String)],
) -> String {
    let mut xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Default Extension="jpg"  ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif"  ContentType="image/gif"/>
  <Default Extension="svg"  ContentType="image/svg+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>"#.to_string();

    for i in 0..slide_count {
        xml.push_str(&format!(
            r#"
  <Override PartName="/ppt/slides/slide{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
            n = i + 1
        ));
    }

    for (zip_path, _) in chart_files {
        xml.push_str(&format!(
            r#"
  <Override PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
            zip_path
        ));
    }

    xml.push_str("\n</Types>");
    xml
}

// ── Static boilerplate XML ────────────────────────────────────────────────────

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#;

const MINIMAL_THEME: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1>
      <a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    </a:fillStyleLst>
    <a:lnStyleLst>
      <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
    </a:lnStyleLst>
    <a:effectStyleLst>
      <a:effectStyle><a:effectLst/></a:effectStyle>
      <a:effectStyle><a:effectLst/></a:effectStyle>
      <a:effectStyle><a:effectLst/></a:effectStyle>
    </a:effectStyleLst>
    <a:bgFillStyleLst>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
    </a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>"#;

const MINIMAL_SLIDE_LAYOUT: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"#;

const SLIDE_LAYOUT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

const MINIMAL_SLIDE_MASTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"
            accent2="accent2" accent3="accent3" accent4="accent4"
            accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle>
    <p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>"#;

const SLIDE_MASTER_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;

// ── Utility ───────────────────────────────────────────────────────────────────

fn pos_emu(pos: &crate::model::elements::Position, w: i64, h: i64) -> (i64, i64, i64, i64) {
    (
        pos.x.to_emu(w),
        pos.y.to_emu(h),
        pos.w.to_emu(w),
        pos.h.to_emu(h),
    )
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
     .replace('"', "&quot;")
     .replace('\'', "&apos;")
}

fn detect_image_ext(bytes: &[u8]) -> &'static str {
    if bytes.starts_with(&[0x89, b'P', b'N', b'G']) { return "png"; }
    if bytes.starts_with(&[0xFF, 0xD8]) { return "jpg"; }
    if bytes.starts_with(b"GIF") { return "gif"; }
    if bytes.starts_with(b"<svg") || bytes.starts_with(b"<?xml") { return "svg"; }
    "png"
}

fn image_content_type(ext: &str) -> &'static str {
    match ext {
        "jpg" | "jpeg" => "image/jpeg",
        "gif"          => "image/gif",
        "svg"          => "image/svg+xml",
        _              => "image/png",
    }
}

/// Map pptxrs shape names to OOXML preset geometry names
fn pptx_shape_name(name: &str) -> &str {
    match name {
        "rect"          => "rect",
        "roundRect"     => "roundRect",
        "ellipse"       => "ellipse",
        "triangle"      => "triangle",
        "rightTriangle" => "rtTriangle",
        "diamond"       => "diamond",
        "pentagon"      => "pentagon",
        "hexagon"       => "hexagon",
        "star4"         => "star4",
        "star5"         => "star5",
        "line"          => "line",
        "rightArrow"    => "rightArrow",
        "leftArrow"     => "leftArrow",
        "upArrow"       => "upArrow",
        "downArrow"     => "downArrow",
        "heart"         => "heart",
        "cloud"         => "cloud",
        other           => other,
    }
}

/// Convert `ppt/charts/chart1.xml` → `ppt/charts/_rels/chart1.xml.rels`
fn chart_rels_path(zip_path: &str) -> String {
    // e.g. "ppt/charts/chart1.xml" → "ppt/charts/_rels/chart1.xml.rels"
    if let Some(slash) = zip_path.rfind('/') {
        let dir  = &zip_path[..slash];
        let file = &zip_path[slash + 1..];
        format!("{}/_rels/{}.rels", dir, file)
    } else {
        format!("_rels/{}.rels", zip_path)
    }
}

const EMPTY_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;

// ── Chart XML generator ───────────────────────────────────────────────────────

/// Generate a complete chart XML file from structured data.
pub fn build_chart_xml_from_data(
    chart_type: &ChartType,
    data: &[ChartData],
    options: &crate::model::elements::ChartOptions,
) -> String {
    let (chart_element, axes_xml) = match chart_type {
        ChartType::Bar | ChartType::Bar3d => {
            let bar_dir = match options.bar_dir.as_ref() {
                Some(BarDir::Bar) => "bar",
                _ => "col",
            };
            let grouping = match options.bar_grouping.as_ref() {
                Some(BarGrouping::Stacked)        => "stacked",
                Some(BarGrouping::PercentStacked) => "percentStacked",
                _                                 => "clustered",
            };
            let tag = if *chart_type == ChartType::Bar3d { "c:bar3DChart" } else { "c:barChart" };
            let series: String = data.iter().enumerate()
                .map(|(i, d)| build_chart_series(i, d))
                .collect();
            let elem = format!(
                "<{tag}><c:barDir val=\"{bar_dir}\"/><c:grouping val=\"{grouping}\"/>\
                 {series}<c:axId val=\"1\"/><c:axId val=\"2\"/></{tag}>"
            );
            (elem, CAT_VAL_AXES.to_string())
        }

        ChartType::Line => {
            let series: String = data.iter().enumerate()
                .map(|(i, d)| build_chart_series(i, d))
                .collect();
            let elem = format!(
                "<c:lineChart><c:grouping val=\"standard\"/>\
                 {series}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>"
            );
            (elem, CAT_VAL_AXES.to_string())
        }

        ChartType::Area => {
            let series: String = data.iter().enumerate()
                .map(|(i, d)| build_chart_series(i, d))
                .collect();
            let elem = format!(
                "<c:areaChart><c:grouping val=\"standard\"/>\
                 {series}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:areaChart>"
            );
            (elem, CAT_VAL_AXES.to_string())
        }

        ChartType::Pie | ChartType::Doughnut => {
            let tag = if *chart_type == ChartType::Pie { "c:pieChart" } else { "c:doughnutChart" };
            let series: String = data.iter().enumerate()
                .map(|(i, d)| build_chart_series(i, d))
                .collect();
            let elem = format!("<{tag}><c:varyColors val=\"1\"/>{series}</{tag}>");
            (elem, String::new()) // pie/doughnut have no axes
        }

        ChartType::Radar => {
            let series: String = data.iter().enumerate()
                .map(|(i, d)| build_chart_series(i, d))
                .collect();
            let elem = format!(
                "<c:radarChart><c:radarStyle val=\"marker\"/>\
                 {series}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:radarChart>"
            );
            (elem, CAT_VAL_AXES.to_string())
        }

        ChartType::Scatter | ChartType::Bubble | ChartType::Bubble3d => {
            let tag = if matches!(chart_type, ChartType::Bubble | ChartType::Bubble3d) {
                "c:bubbleChart"
            } else {
                "c:scatterChart"
            };
            let series: String = data.iter().enumerate()
                .map(|(i, d)| build_scatter_series(i, d))
                .collect();
            let elem = format!(
                "<{tag}><c:scatterStyle val=\"marker\"/>\
                 {series}<c:axId val=\"1\"/><c:axId val=\"2\"/></{tag}>"
            );
            (elem, SCATTER_AXES.to_string())
        }
    };

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      {chart_element}
      {axes_xml}
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>"#
    )
}

fn build_chart_series(idx: usize, data: &ChartData) -> String {
    let name = data.name.as_deref().unwrap_or("");
    let val_count = data.values.len();

    let cat_xml = if let Some(labels) = &data.labels {
        let pts: String = labels.iter().enumerate()
            .map(|(i, l)| format!("<c:pt idx=\"{i}\"><c:v>{}</c:v></c:pt>", xml_escape(l)))
            .collect();
        format!(
            "<c:cat><c:strRef><c:strCache>\
             <c:ptCount val=\"{}\"/>{pts}\
             </c:strCache></c:strRef></c:cat>",
            labels.len()
        )
    } else {
        String::new()
    };

    let val_pts: String = data.values.iter().enumerate()
        .map(|(i, v)| format!("<c:pt idx=\"{i}\"><c:v>{v}</c:v></c:pt>"))
        .collect();

    format!(
        "<c:ser>\
         <c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>\
         <c:tx><c:strRef><c:strCache>\
         <c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{name}</c:v></c:pt>\
         </c:strCache></c:strRef></c:tx>\
         {cat_xml}\
         <c:val><c:numRef><c:numCache>\
         <c:formatCode>General</c:formatCode>\
         <c:ptCount val=\"{val_count}\"/>{val_pts}\
         </c:numCache></c:numRef></c:val>\
         </c:ser>",
        name = xml_escape(name),
    )
}

/// Scatter / bubble series use xVal + yVal instead of cat + val
fn build_scatter_series(idx: usize, data: &ChartData) -> String {
    let name = data.name.as_deref().unwrap_or("");
    let count = data.values.len();

    // x values come from labels (if numeric strings), otherwise 1..N
    let x_pts: String = (0..count)
        .map(|i| {
            let x = data.labels.as_ref()
                .and_then(|l| l.get(i))
                .and_then(|s| s.parse::<f64>().ok())
                .unwrap_or((i + 1) as f64);
            format!("<c:pt idx=\"{i}\"><c:v>{x}</c:v></c:pt>")
        })
        .collect();

    let y_pts: String = data.values.iter().enumerate()
        .map(|(i, v)| format!("<c:pt idx=\"{i}\"><c:v>{v}</c:v></c:pt>"))
        .collect();

    format!(
        "<c:ser>\
         <c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>\
         <c:tx><c:strRef><c:strCache>\
         <c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{name}</c:v></c:pt>\
         </c:strCache></c:strRef></c:tx>\
         <c:xVal><c:numRef><c:numCache>\
         <c:formatCode>General</c:formatCode>\
         <c:ptCount val=\"{count}\"/>{x_pts}\
         </c:numCache></c:numRef></c:xVal>\
         <c:yVal><c:numRef><c:numCache>\
         <c:formatCode>General</c:formatCode>\
         <c:ptCount val=\"{count}\"/>{y_pts}\
         </c:numCache></c:numRef></c:yVal>\
         </c:ser>",
        name = xml_escape(name),
    )
}

const CAT_VAL_AXES: &str =
    r#"<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>
    <c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>
    <c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>"#;

const SCATTER_AXES: &str =
    r#"<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>
    <c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>
    <c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>"#;

// ── Passthrough builder ───────────────────────────────────────────────────────

/// Build a PPTX using the original ZIP as a base, replacing only changed content.
fn build_pptx_passthrough(pres: &Presentation, zip_bytes: &[u8]) -> Result<Vec<u8>, String> {
    // Pre-build all overrides (dirty slides, chart patches, new charts)
    let mut overrides: HashMap<String, Vec<u8>> = HashMap::new();

    let mut media_counter = {
        // Determine first unused media ID from existing ZIP entries
        let mut src = ZipArchive::new(Cursor::new(zip_bytes)).map_err(|e| e.to_string())?;
        let mut max_id = 0u32;
        for i in 0..src.len() {
            if let Ok(f) = src.by_index(i) {
                let name = f.name().to_string();
                if let Some(rest) = name.strip_prefix("ppt/media/image") {
                    if let Some(stem) = rest.split('.').next() {
                        if let Ok(n) = stem.parse::<u32>() { max_id = max_id.max(n); }
                    }
                }
            }
        }
        max_id + 1
    };
    let mut chart_counter = pres.next_chart_id;

    // If original_slide_count wasn't set (e.g. came from fromJson), derive it
    // from the ZIP itself by counting ppt/slides/slideN.xml entries.
    let original_slide_count = if pres.original_slide_count > 0 {
        pres.original_slide_count
    } else {
        let mut src = ZipArchive::new(Cursor::new(zip_bytes)).map_err(|e| e.to_string())?;
        let mut count = 0usize;
        for i in 0..src.len() {
            if let Ok(f) = src.by_index(i) {
                let name = f.name();
                if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
                    count += 1;
                }
            }
        }
        count
    };

    for (idx, slide) in pres.slides.iter().enumerate() {
        let slide_path = format!("ppt/slides/slide{}.xml", idx + 1);
        let rels_path  = format!("ppt/slides/_rels/slide{}.xml.rels", idx + 1);

        if idx < original_slide_count {
            // ── Existing slide ────────────────────────────────────────────────
            if !slide.dirty {
                continue; // copy verbatim from source
            }

            // Handle charts whose data was modified (updateChart was called)
            for el in &slide.elements[..slide.original_element_count.min(slide.elements.len())] {
                if let SlideElement::Chart {
                    chart_type, data, options,
                    source_chart_path: Some(chart_path),
                    modified: true, ..
                } = el {
                    let chart_xml = build_chart_xml_from_data(chart_type, data, options);
                    overrides.insert(chart_path.clone(), chart_xml.into_bytes());
                }
            }

            // Handle tables whose data was modified (updateTable was called)
            // Rebuild the slide XML with patched table frames
            let (slide_xml, rels_xml, new_media, new_charts) =
                build_slide_modified(slide, idx, pres, &mut media_counter, &mut chart_counter);

            overrides.insert(slide_path, slide_xml.into_bytes());
            overrides.insert(rels_path, rels_xml.into_bytes());

            for (name, bytes) in new_media {
                overrides.insert(format!("ppt/media/{}", name), bytes);
            }
            for (path, xml) in new_charts {
                overrides.insert(path.clone(), xml.into_bytes());
                let rels = chart_rels_path(&path);
                overrides.entry(rels).or_insert_with(|| EMPTY_RELS.as_bytes().to_vec());
            }
        } else {
            // ── Newly added slide (not in original ZIP) ───────────────────────
            let (xml, rels, new_media, new_charts) =
                build_slide_xml(slide, idx, pres, &mut media_counter, &mut chart_counter);

            overrides.insert(slide_path, xml.into_bytes());
            overrides.insert(rels_path, rels.into_bytes());

            for (name, bytes) in new_media {
                overrides.insert(format!("ppt/media/{}", name), bytes);
            }
            for (path, xml) in new_charts {
                overrides.insert(path.clone(), xml.into_bytes());
                let rels = chart_rels_path(&path);
                overrides.entry(rels).or_insert_with(|| EMPTY_RELS.as_bytes().to_vec());
            }
        }
    }

    // Regenerate presentation.xml.rels if slide count changed
    if pres.slides.len() != original_slide_count {
        overrides.insert(
            "ppt/_rels/presentation.xml.rels".to_string(),
            build_pres_rels(pres.slides.len()).into_bytes(),
        );
        // Also regenerate ppt/presentation.xml to update sldIdLst
        overrides.insert(
            "ppt/presentation.xml".to_string(),
            build_presentation_xml(pres).into_bytes(),
        );
    }

    // Read all source file names first
    let mut source = ZipArchive::new(Cursor::new(zip_bytes)).map_err(|e| e.to_string())?;
    let source_len = source.len();
    let mut source_names: Vec<String> = Vec::with_capacity(source_len);
    for i in 0..source_len {
        if let Ok(f) = source.by_index(i) {
            source_names.push(f.name().to_string());
        }
    }

    // Build output ZIP
    let out_buf = Vec::new();
    let mut zip = ZipWriter::new(Cursor::new(out_buf));
    let options: FileOptions<()> =
        FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // Copy/replace source files
    for i in 0..source_len {
        let (name, raw_bytes) = {
            let mut entry = source.by_index(i).map_err(|e| e.to_string())?;
            let name = entry.name().to_string();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).map_err(|e| e.to_string())?;
            (name, bytes)
        };

        zip.start_file(&name, options).map_err(|e| e.to_string())?;
        if let Some(override_bytes) = overrides.get(&name) {
            zip.write_all(override_bytes).map_err(|e| e.to_string())?;
        } else {
            zip.write_all(&raw_bytes).map_err(|e| e.to_string())?;
        }
    }

    // Write new files (new slides, media, charts not in source)
    let source_set: std::collections::HashSet<&str> =
        source_names.iter().map(|s| s.as_str()).collect();

    for (path, bytes) in &overrides {
        if !source_set.contains(path.as_str()) {
            zip.start_file(path, options).map_err(|e| e.to_string())?;
            zip.write_all(bytes).map_err(|e| e.to_string())?;
        }
    }

    let inner = zip.finish().map_err(|e| e.to_string())?;
    Ok(inner.into_inner())
}

/// Build slide XML for a dirty slide that has a raw_xml base.
/// Applies table patches and injects new elements surgically.
fn build_slide_modified(
    slide: &Slide,
    _idx: usize,
    pres: &Presentation,
    media_counter: &mut u32,
    chart_counter: &mut u32,
) -> (String, String, Vec<(String, Vec<u8>)>, Vec<(String, String)>) {
    let raw_xml  = slide.raw_xml.as_deref().unwrap_or("");
    let raw_rels = slide.raw_rels.as_deref().unwrap_or("");

    let (w_emu, h_emu) = pres.meta.layout.dimensions_emu();
    let orig_count = slide.original_element_count;

    // Step 1: patch the raw slide XML for any modified original elements
    let mut patched_xml = raw_xml.to_string();
    for el in &slide.elements[..orig_count.min(slide.elements.len())] {
        if let SlideElement::Table { data, raw_frame_xml: Some(old_frame), modified: true, .. } = el {
            let new_frame = patch_table_frame_xml(old_frame, data);
            patched_xml = patched_xml.replacen(old_frame.as_str(), &new_frame, 1);
        }
    }

    // Step 2: build XML for new elements (those past orig_count)
    let new_elements = &slide.elements[orig_count..];
    if new_elements.is_empty() {
        return (patched_xml, raw_rels.to_string(), vec![], vec![]);
    }

    let max_sp_id  = find_max_id_in_xml(raw_xml, r#" id=""#);
    let max_rel_id = find_max_rel_id(raw_rels);
    let mut sp_id  = max_sp_id + 1;
    let mut rel_id = max_rel_id + 1;

    let mut new_shapes = String::new();
    let mut new_rels   = String::new();
    let mut new_media: Vec<(String, Vec<u8>)>  = Vec::new();
    let mut new_charts: Vec<(String, String)>  = Vec::new();

    for el in new_elements {
        match el {
            SlideElement::Text { text, options } => {
                new_shapes.push_str(&build_text_shape(text, options, sp_id, w_emu, h_emu));
                sp_id += 1;
            }
            SlideElement::Image { options } => {
                let rid = format!("rId{}", rel_id);
                if let Some(data_b64) = &options.data {
                    if let Ok(bytes) = B64.decode(data_b64) {
                        let ext   = detect_image_ext(&bytes);
                        let fname = format!("image{}.{}", *media_counter, ext);
                        new_rels.push_str(&format!(
                            r#"<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{fname}"/>"#
                        ));
                        new_media.push((fname, bytes));
                        new_shapes.push_str(&build_pic_shape(options, sp_id, &rid, w_emu, h_emu));
                        sp_id  += 1;
                        rel_id += 1;
                        *media_counter += 1;
                    }
                }
            }
            SlideElement::Shape { shape_type, options } => {
                new_shapes.push_str(&build_shape(shape_type, options, sp_id, w_emu, h_emu));
                sp_id += 1;
            }
            SlideElement::Table { data, options, .. } => {
                new_shapes.push_str(&build_table(data, options, sp_id, w_emu, h_emu));
                sp_id += 1;
            }
            SlideElement::Chart { chart_type, data, options, .. } => {
                let rid        = format!("rId{}", rel_id);
                let chart_name = format!("chart{}.xml", *chart_counter);
                let zip_path   = format!("ppt/charts/{}", chart_name);
                let rel_target = format!("../charts/{}", chart_name);
                new_rels.push_str(&format!(
                    r#"<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{rel_target}"/>"#
                ));
                new_shapes.push_str(&build_chart_frame(options, sp_id, &rid, w_emu, h_emu));
                let chart_xml = build_chart_xml_from_data(chart_type, data, options);
                new_charts.push((zip_path, chart_xml));
                sp_id  += 1;
                rel_id += 1;
                *chart_counter += 1;
            }
            SlideElement::Notes { .. } => {}
        }
    }

    // Inject new shapes before </p:spTree>
    let final_xml = if let Some(pos) = patched_xml.rfind("</p:spTree>") {
        format!("{}{}{}", &patched_xml[..pos], new_shapes, &patched_xml[pos..])
    } else {
        patched_xml
    };

    // Inject new rels before </Relationships>
    let final_rels = if !new_rels.is_empty() {
        if let Some(pos) = raw_rels.rfind("</Relationships>") {
            format!("{}  {}{}", &raw_rels[..pos], new_rels, &raw_rels[pos..])
        } else {
            raw_rels.to_string()
        }
    } else {
        raw_rels.to_string()
    };

    (final_xml, final_rels, new_media, new_charts)
}

/// Patch the text content of table cells in a raw `<p:graphicFrame>` XML block.
/// Preserves all formatting (borders, shading, fonts) — only replaces `<a:t>` text.
fn patch_table_frame_xml(
    frame_xml: &str,
    data: &[Vec<crate::model::elements::TableCell>],
) -> String {
    let flat: Vec<&crate::model::elements::TableCell> =
        data.iter().flat_map(|r| r.iter()).collect();
    let mut result = frame_xml.to_string();
    let mut cell_idx = 0usize;
    let mut search_pos = 0usize;

    while search_pos < result.len() {
        // Find next <a:tc>
        let Some(tc_start) = result[search_pos..].find("<a:tc>").map(|o| o + search_pos) else { break };
        let Some(tc_end_off) = result[tc_start..].find("</a:tc>") else { break };
        let tc_end = tc_start + tc_end_off + "</a:tc>".len();

        if let Some(cell) = flat.get(cell_idx) {
            let new_text = match cell {
                crate::model::elements::TableCell::Text(s) => s.clone(),
                crate::model::elements::TableCell::Rich(r) => match &r.text {
                    TextContent::Plain(s) => s.clone(),
                    TextContent::Runs(runs) => runs.iter().map(|r| r.text.as_str()).collect::<Vec<_>>().join(""),
                },
            };

            let tc_xml = result[tc_start..tc_end].to_string();
            let patched = replace_at_text(&tc_xml, &new_text);
            let new_len = patched.len();
            result.replace_range(tc_start..tc_end, &patched);
            search_pos = tc_start + new_len;
        } else {
            search_pos = tc_end;
        }

        cell_idx += 1;
    }

    result
}

/// Replace the text inside the first `<a:t>…</a:t>` pair within a `<a:tc>` block.
fn replace_at_text(tc_xml: &str, new_text: &str) -> String {
    if let Some(start) = tc_xml.find("<a:t>") {
        if let Some(end) = tc_xml.find("</a:t>") {
            let before = &tc_xml[..start + "<a:t>".len()];
            let after  = &tc_xml[end..];
            return format!("{}{}{}", before, xml_escape(new_text), after);
        }
    }
    tc_xml.to_string()
}

// ── XML scanning helpers ──────────────────────────────────────────────────────

/// Find the highest `id="N"` value in slide XML (for shape ID allocation).
fn find_max_id_in_xml(xml: &str, attr_prefix: &str) -> u32 {
    let mut max = 1u32;
    let mut s = xml;
    while let Some(pos) = s.find(attr_prefix) {
        let rest = &s[pos + attr_prefix.len()..];
        if let Some(end) = rest.find('"') {
            if let Ok(n) = rest[..end].parse::<u32>() {
                max = max.max(n);
            }
        }
        s = &s[pos + attr_prefix.len()..];
    }
    max
}

/// Find the highest `rIdN` number in a rels XML (for relationship ID allocation).
fn find_max_rel_id(rels_xml: &str) -> u32 {
    let mut max = 0u32;
    let mut s = rels_xml;
    while let Some(pos) = s.find("Id=\"rId") {
        let rest = &s[pos + 7..];
        if let Some(end) = rest.find('"') {
            if let Ok(n) = rest[..end].parse::<u32>() {
                max = max.max(n);
            }
        }
        s = &s[pos + 7..];
    }
    max
}
