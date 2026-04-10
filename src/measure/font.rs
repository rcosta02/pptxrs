use std::collections::HashMap;
use rustybuzz::{Face, UnicodeBuffer};
use unicode_linebreak::{linebreaks, BreakOpportunity};

#[derive(Debug, Default)]
pub struct FontRegistry {
    fonts: HashMap<String, Vec<u8>>,
}

impl FontRegistry {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn register(&mut self, name: impl Into<String>, data: Vec<u8>) {
        self.fonts.insert(name.into(), data);
    }

    pub fn get(&self, name: &str) -> Option<&Vec<u8>> {
        self.fonts.get(name)
    }
}

#[derive(Debug)]
pub struct MeasureOptions {
    /// Font name — must be registered via `FontRegistry::register`
    pub font: String,
    /// Font size in points
    pub font_size: f64,
    pub bold: bool,
    pub italic: bool,
    /// Additional character spacing in points
    pub char_spacing: f64,
    /// Line spacing multiplier (1.0 = normal)
    pub line_spacing_multiple: f64,
    /// Text box width in inches. When `Some`, enables word-wrap.
    pub width_inches: Option<f64>,
}

impl Default for MeasureOptions {
    fn default() -> Self {
        Self {
            font: String::new(),
            font_size: 18.0,
            bold: false,
            italic: false,
            char_spacing: 0.0,
            line_spacing_multiple: 1.0,
            width_inches: None,
        }
    }
}

#[derive(Debug, serde::Serialize)]
pub struct TextMetrics {
    /// Total height of the text block in points
    pub height: f64,
    /// Width of the longest rendered line in points
    pub width: f64,
    /// Number of rendered lines
    pub lines: u32,
    /// Per-line height in points (ascender + descender + line gap)
    #[serde(rename = "lineHeight")]
    pub line_height: f64,
}

/// Measure text using HarfBuzz shaping via rustybuzz.
///
/// Returns an error string if the font is not registered or the font data is invalid.
pub fn measure_text(
    text: &str,
    opts: &MeasureOptions,
    registry: &FontRegistry,
) -> Result<TextMetrics, String> {
    let font_data = registry
        .get(&opts.font)
        .ok_or_else(|| format!("font '{}' not registered", opts.font))?;

    let face = Face::from_slice(font_data, 0)
        .ok_or_else(|| "failed to parse font face".to_string())?;

    let units_per_em = face.units_per_em() as f64;
    let ascender = face.ascender() as f64;
    let descender = face.descender() as f64; // negative
    let line_gap = face.line_gap() as f64;

    // Line height in points
    let line_height_raw = (ascender - descender + line_gap) / units_per_em * opts.font_size;
    let line_height = line_height_raw * opts.line_spacing_multiple;

    // Width constraint in font units
    let max_width_pts: Option<f64> = opts.width_inches.map(|w| w * 72.0); // inches → points

    // Shape each line using rustybuzz
    let logical_lines: Vec<&str> = text.split('\n').collect();
    let mut all_lines: Vec<f64> = Vec::new(); // widths in points

    for logical_line in logical_lines {
        if logical_line.is_empty() {
            all_lines.push(0.0);
            continue;
        }

        // Shape the whole logical line
        let mut ubuf = UnicodeBuffer::new();
        ubuf.push_str(logical_line);
        let shaped = rustybuzz::shape(&face, &[], ubuf);

        let positions = shaped.glyph_positions();
        let _infos = shaped.glyph_infos();

        // Build per-char advance widths (by cluster → char index mapping)
        // For simplicity we accumulate advance per glyph cluster
        let glyph_advances: Vec<f64> = positions
            .iter()
            .map(|p| {
                let advance_pts = p.x_advance as f64 / units_per_em * opts.font_size;
                advance_pts + opts.char_spacing
            })
            .collect();

        let total_pts: f64 = glyph_advances.iter().sum();

        if let Some(max_w) = max_width_pts {
            // Word-wrap: use UAX#14 break opportunities
            let breaks: Vec<(usize, BreakOpportunity)> = linebreaks(logical_line).collect();

            let mut line_start_pts = 0.0f64;

            // Map byte positions to cumulative glyph widths
            // (Simplified: assume 1 glyph per cluster — sufficient for Latin text)
            let char_advances: Vec<f64> = {
                let chars: Vec<char> = logical_line.chars().collect();
                if chars.len() == glyph_advances.len() {
                    glyph_advances.clone()
                } else {
                    // Redistribute evenly as fallback
                    let per = total_pts / chars.len().max(1) as f64;
                    vec![per; chars.len()]
                }
            };

            let mut byte_to_pts: Vec<(usize, f64)> = Vec::new();
            let mut cumulative = 0.0f64;
            for (i, c) in logical_line.char_indices() {
                let adv = char_advances.get(i).copied().unwrap_or(0.0);
                cumulative += adv;
                byte_to_pts.push((i + c.len_utf8(), cumulative));
            }

            let pts_at_byte = |byte: usize| -> f64 {
                byte_to_pts
                    .iter()
                    .find(|(b, _)| *b >= byte)
                    .map(|(_, p)| *p)
                    .unwrap_or(cumulative)
            };

            let mut prev_break = 0usize;

            for (byte_pos, opp) in &breaks {
                let chunk_end_pts = pts_at_byte(*byte_pos);
                let chunk_width = chunk_end_pts - line_start_pts;

                if chunk_width > max_w && prev_break > 0 {
                    // Wrap before this segment
                    let line_w = pts_at_byte(prev_break) - line_start_pts;
                    all_lines.push(line_w);
                    line_start_pts = pts_at_byte(prev_break);
                }

                if matches!(opp, BreakOpportunity::Mandatory) {
                    let line_w = pts_at_byte(*byte_pos) - line_start_pts;
                    all_lines.push(line_w);
                    line_start_pts = pts_at_byte(*byte_pos);
                }

                prev_break = *byte_pos;
            }

            // Last line
            let last_w = total_pts - line_start_pts;
            if last_w > 0.0 {
                all_lines.push(last_w);
            }
        } else {
            all_lines.push(total_pts);
        }
    }

    let line_count = all_lines.len().max(1) as u32;
    let max_width = all_lines.iter().cloned().fold(0.0_f64, f64::max);
    let total_height = line_height * line_count as f64;

    Ok(TextMetrics {
        height: total_height,
        width: max_width,
        lines: line_count,
        line_height,
    })
}
