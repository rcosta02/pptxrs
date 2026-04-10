use std::io::{Cursor, Write as IoWrite};
use zip::{ZipWriter, write::FileOptions};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as B64;

use crate::model::{
    elements::{
        HorizAlign, SlideElement, TextContent, TextOptions,
    },
    presentation::Presentation,
    slide::Slide,
};


pub fn build_pptx(pres: &Presentation) -> Result<Vec<u8>, String> {
    let buf = Vec::new();
    let cursor = Cursor::new(buf);
    let mut zip = ZipWriter::new(cursor);

    let options: FileOptions<()> = FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Track media files: (filename, bytes)
    let mut media: Vec<(String, Vec<u8>)> = Vec::new();

    // Collect slide XMLs and their relationship files
    let slide_count = pres.slides.len();
    let mut slide_xmls: Vec<String> = Vec::new();
    let mut slide_rels: Vec<String> = Vec::new();

    for (idx, slide) in pres.slides.iter().enumerate() {
        let (xml, rels, new_media) = build_slide_xml(slide, idx, pres);
        slide_xmls.push(xml);
        slide_rels.push(rels);
        media.extend(new_media);
    }

    // [Content_Types].xml
    zip.start_file("[Content_Types].xml", options).map_err(|e| e.to_string())?;
    zip.write_all(build_content_types(slide_count, &media).as_bytes())
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

    let inner = zip.finish().map_err(|e| e.to_string())?;
    Ok(inner.into_inner())
}

// ── Per-slide XML ─────────────────────────────────────────────────────────────

/// Returns (slide_xml, rels_xml, media_files)
fn build_slide_xml(
    slide: &Slide,
    _idx: usize,
    pres: &Presentation,
) -> (String, String, Vec<(String, Vec<u8>)>) {
    let (w_emu, h_emu) = pres.meta.layout.dimensions_emu();
    let mut shapes = String::new();
    let mut rels_entries = String::new();
    let mut media: Vec<(String, Vec<u8>)> = Vec::new();
    let mut sp_id = 2u32; // 1 is reserved for background
    let mut rel_id = 1u32;
    let mut chart_id = 1u32;

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
                        let fname = format!("image{}.{}", rel_id, ext);
                        let _content_type = image_content_type(&ext);
                        rels_entries.push_str(&format!(
                            r#"<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{fname}"/>"#
                        ));
                        media.push((fname, bytes));
                        shapes.push_str(&build_pic_shape(options, sp_id, &rid, w_emu, h_emu));
                        sp_id += 1;
                        rel_id += 1;
                    }
                }
            }
            SlideElement::Shape { shape_type, options } => {
                shapes.push_str(&build_shape(shape_type, options, sp_id, w_emu, h_emu));
                sp_id += 1;
            }
            SlideElement::Table { data, options } => {
                shapes.push_str(&build_table(data, options, sp_id, w_emu, h_emu));
                sp_id += 1;
            }
            SlideElement::Chart { chart_type, data, combo_types, options } => {
                let rid = format!("rId{}", rel_id);
                let chart_fname = format!("chart{}.xml", chart_id);
                let chart_path = format!("../charts/{}", chart_fname);
                rels_entries.push_str(&format!(
                    r#"<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="{chart_path}"/>"#
                ));
                shapes.push_str(&build_chart_frame(options, sp_id, &rid, w_emu, h_emu));
                // Chart XML would be written separately; for now embed a placeholder note
                sp_id += 1;
                rel_id += 1;
                chart_id += 1;
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

    (slide_xml, rels_xml, media)
}

// ── Shape builders ────────────────────────────────────────────────────────────

fn build_text_shape(
    text: &TextContent,
    opts: &TextOptions,
    id: u32,
    slide_w: i64,
    slide_h: i64,
) -> String {
    let (x, y, cx, cy) = pos_emu(&opts.pos, slide_w, slide_h);
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
    let col_w_emu: Vec<i64> = match &opts.col_w {
        Some(crate::model::elements::ColRowSizes::Uniform(v)) => {
            vec![(*v * 914_400.0) as i64; col_count]
        }
        Some(crate::model::elements::ColRowSizes::PerColumn(vs)) => {
            vs.iter().map(|v| (*v * 914_400.0) as i64).collect()
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

fn build_content_types(slide_count: usize, _media: &[(String, Vec<u8>)]) -> String {
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

/// Map pptrs shape names to OOXML preset geometry names
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
