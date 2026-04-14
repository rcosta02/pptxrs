use serde::{Deserialize, Serialize};

// ── Shared ────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Position {
    pub x: CoordVal,
    pub y: CoordVal,
    pub w: CoordVal,
    /// Optional — text elements may omit `h` and have it estimated from font size.
    #[serde(default)]
    pub h: CoordVal,
}

/// A coordinate value — either pixels (f64, at 96 DPI) or a percentage string like "50%"
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum CoordVal {
    Pixels(f64),
    Pct(String),
}

impl Default for CoordVal {
    fn default() -> Self {
        CoordVal::Pixels(0.0)
    }
}

impl CoordVal {
    /// Convert to EMU (English Metric Units). Percentage requires slide dimension context.
    /// 1 px (96 DPI) = 9 525 EMU  (914 400 EMU/inch ÷ 96 px/inch)
    pub fn to_emu(&self, slide_dim_emu: i64) -> i64 {
        match self {
            CoordVal::Pixels(v) => (*v * 9_525.0) as i64,
            CoordVal::Pct(s) => {
                let pct: f64 = s.trim_end_matches('%').parse().unwrap_or(0.0);
                (slide_dim_emu as f64 * pct / 100.0) as i64
            }
        }
    }

    /// Return the value in pixels (96 DPI).
    /// `slide_dim_emu` is required to resolve percentage values; pass 0 if the
    /// value is guaranteed to be pixel-based.
    pub fn to_pixels(&self, slide_dim_emu: i64) -> f64 {
        const EMU_PER_PX: f64 = 9_525.0;
        match self {
            CoordVal::Pixels(v) => *v,
            CoordVal::Pct(s) => {
                let pct: f64 = s.trim_end_matches('%').parse().unwrap_or(0.0);
                slide_dim_emu as f64 * pct / 100.0 / EMU_PER_PX
            }
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ShadowOptions {
    #[serde(rename = "type")]
    pub kind: ShadowKind,
    pub angle: Option<f64>,
    pub blur: Option<f64>,
    pub color: Option<String>,
    pub offset: Option<f64>,
    pub opacity: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum ShadowKind {
    #[default]
    Outer,
    Inner,
    None,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct HyperlinkOptions {
    pub url: Option<String>,
    pub slide: Option<u32>,
    pub tooltip: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct LineOptions {
    pub color: Option<String>,
    pub width: Option<f64>,
    #[serde(rename = "dashType")]
    pub dash_type: Option<String>,
    #[serde(rename = "beginArrowType")]
    pub begin_arrow_type: Option<String>,
    #[serde(rename = "endArrowType")]
    pub end_arrow_type: Option<String>,
}

// ── Text ──────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TextOptions {
    #[serde(flatten)]
    pub pos: Position,
    #[serde(rename = "fontSize")]
    pub font_size: Option<f64>,
    #[serde(rename = "fontFace")]
    pub font_face: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<serde_json::Value>, // bool or string style
    pub strike: Option<String>,
    pub color: Option<String>,
    pub fill: Option<String>,
    pub highlight: Option<String>,
    pub subscript: Option<bool>,
    pub superscript: Option<bool>,
    pub align: Option<HorizAlign>,
    pub valign: Option<VertAlign>,
    pub vert: Option<String>,
    #[serde(rename = "rtlMode")]
    pub rtl_mode: Option<bool>,
    #[serde(rename = "autoFit")]
    pub auto_fit: Option<bool>,
    pub fit: Option<TextFit>,
    pub wrap: Option<bool>,
    #[serde(rename = "breakLine")]
    pub break_line: Option<bool>,
    #[serde(rename = "softBreakBefore")]
    pub soft_break_before: Option<bool>,
    #[serde(rename = "lineSpacing")]
    pub line_spacing: Option<f64>,
    #[serde(rename = "lineSpacingMultiple")]
    pub line_spacing_multiple: Option<f64>,
    #[serde(rename = "charSpacing")]
    pub char_spacing: Option<f64>,
    pub baseline: Option<f64>,
    #[serde(rename = "paraSpaceBefore")]
    pub para_space_before: Option<f64>,
    #[serde(rename = "paraSpaceAfter")]
    pub para_space_after: Option<f64>,
    #[serde(rename = "indentLevel")]
    pub indent_level: Option<u8>,
    pub bullet: Option<BulletValue>,
    pub inset: Option<f64>,
    pub margin: Option<MarginValue>,
    #[serde(rename = "rectRadius")]
    pub rect_radius: Option<f64>,
    pub rotate: Option<f64>,
    #[serde(rename = "isTextBox")]
    pub is_text_box: Option<bool>,
    pub hyperlink: Option<HyperlinkOptions>,
    pub lang: Option<String>,
    pub line: Option<LineOptions>,
    pub transparency: Option<f64>,
    pub glow: Option<GlowOptions>,
    pub outline: Option<OutlineOptions>,
    pub shadow: Option<ShadowOptions>,
    pub placeholder: Option<PlaceholderKind>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum BulletValue {
    Simple(bool),
    Options(BulletOptions),
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct BulletOptions {
    #[serde(rename = "type")]
    pub kind: Option<BulletKind>,
    pub code: Option<String>,
    pub font: Option<String>,
    pub color: Option<String>,
    pub size: Option<f64>,
    pub indent: Option<f64>,
    #[serde(rename = "numberStartAt")]
    pub number_start_at: Option<u32>,
    pub style: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum BulletKind {
    Bullet,
    Number,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum MarginValue {
    Uniform(f64),
    Sides([f64; 4]),
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct GlowOptions {
    pub size: f64,
    pub opacity: f64,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct OutlineOptions {
    pub color: String,
    pub size: f64,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum HorizAlign {
    #[default]
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "lowercase")]
pub enum VertAlign {
    #[default]
    Top,
    Middle,
    Bottom,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum TextFit {
    None,
    Shrink,
    Resize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum PlaceholderKind {
    Title,
    Body,
}

/// A text run within a paragraph (inline style variation)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextRun {
    pub text: String,
    pub options: Option<TextRunOptions>,
}

/// TextOptions minus position fields — used for inline runs
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TextRunOptions {
    #[serde(rename = "fontSize")]
    pub font_size: Option<f64>,
    #[serde(rename = "fontFace")]
    pub font_face: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<serde_json::Value>,
    pub color: Option<String>,
    pub highlight: Option<String>,
    pub subscript: Option<bool>,
    pub superscript: Option<bool>,
    pub hyperlink: Option<HyperlinkOptions>,
    pub lang: Option<String>,
    pub char_spacing: Option<f64>,
    pub break_line: Option<bool>,
    pub soft_break_before: Option<bool>,
    pub bullet: Option<BulletValue>,
    pub indent_level: Option<u8>,
}

/// The content of a text element — either a plain string or an array of runs
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum TextContent {
    Plain(String),
    Runs(Vec<TextRun>),
}

// ── Image ─────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ImageOptions {
    #[serde(flatten)]
    pub pos: Position,
    /// Base64-encoded image bytes (PNG/JPG/GIF/SVG/WEBP)
    pub data: Option<String>,
    /// Node.js filesystem path or URL
    pub path: Option<String>,
    pub rotate: Option<f64>,
    #[serde(rename = "flipH")]
    pub flip_h: Option<bool>,
    #[serde(rename = "flipV")]
    pub flip_v: Option<bool>,
    pub rounding: Option<bool>,
    pub transparency: Option<f64>,
    #[serde(rename = "altText")]
    pub alt_text: Option<String>,
    pub hyperlink: Option<HyperlinkOptions>,
    pub sizing: Option<ImageSizing>,
    pub shadow: Option<ShadowOptions>,
    pub placeholder: Option<PlaceholderKind>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ImageSizing {
    #[serde(rename = "type")]
    pub kind: ImageSizingType,
    pub w: Option<CoordVal>,
    pub h: Option<CoordVal>,
    pub x: Option<CoordVal>,
    pub y: Option<CoordVal>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum ImageSizingType {
    Contain,
    Cover,
    Crop,
}

// ── Shape ─────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ShapeFill {
    pub color: Option<String>,
    pub transparency: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ShapeOptions {
    #[serde(flatten)]
    pub pos: Position,
    pub fill: Option<ShapeFill>,
    pub line: Option<LineOptions>,
    #[serde(rename = "flipH")]
    pub flip_h: Option<bool>,
    #[serde(rename = "flipV")]
    pub flip_v: Option<bool>,
    pub rotate: Option<f64>,
    #[serde(rename = "rectRadius")]
    pub rect_radius: Option<f64>,
    #[serde(rename = "shapeName")]
    pub shape_name: Option<String>,
    pub hyperlink: Option<HyperlinkOptions>,
    pub shadow: Option<ShadowOptions>,
    // optional text inside the shape
    pub text: Option<TextContent>,
    #[serde(rename = "fontSize")]
    pub font_size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<String>,
    pub align: Option<HorizAlign>,
    pub valign: Option<VertAlign>,
}

// ── Table ─────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum TableCell {
    Text(String),
    Rich(RichTableCell),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RichTableCell {
    pub text: TextContent,
    pub options: Option<TableCellOptions>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableCellOptions {
    pub align: Option<HorizAlign>,
    pub valign: Option<VertAlign>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    #[serde(rename = "fontSize")]
    pub font_size: Option<f64>,
    #[serde(rename = "fontFace")]
    pub font_face: Option<String>,
    pub color: Option<String>,
    pub fill: Option<String>,
    pub margin: Option<MarginValue>,
    pub border: Option<BorderValue>,
    pub colspan: Option<u32>,
    pub rowspan: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum BorderValue {
    Single(BorderOptions),
    Sides(Vec<BorderOptions>),
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct BorderOptions {
    #[serde(rename = "type")]
    pub kind: Option<String>,
    pub pt: Option<f64>,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableOptions {
    pub x: Option<CoordVal>,
    pub y: Option<CoordVal>,
    pub w: Option<CoordVal>,
    pub h: Option<CoordVal>,
    #[serde(rename = "colW")]
    pub col_w: Option<ColRowSizes>,
    #[serde(rename = "rowH")]
    pub row_h: Option<ColRowSizes>,
    pub align: Option<HorizAlign>,
    pub valign: Option<VertAlign>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    #[serde(rename = "fontSize")]
    pub font_size: Option<f64>,
    #[serde(rename = "fontFace")]
    pub font_face: Option<String>,
    pub color: Option<String>,
    pub fill: Option<String>,
    pub border: Option<BorderValue>,
    pub margin: Option<MarginValue>,
    #[serde(rename = "autoPage")]
    pub auto_page: Option<bool>,
    #[serde(rename = "autoPageCharWeight")]
    pub auto_page_char_weight: Option<f64>,
    #[serde(rename = "autoPageLineWeight")]
    pub auto_page_line_weight: Option<f64>,
    #[serde(rename = "autoPageRepeatHeader")]
    pub auto_page_repeat_header: Option<bool>,
    #[serde(rename = "autoPageHeaderRows")]
    pub auto_page_header_rows: Option<u32>,
    #[serde(rename = "newSlideStartY")]
    pub new_slide_start_y: Option<CoordVal>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ColRowSizes {
    Uniform(f64),
    PerColumn(Vec<f64>),
}

// ── Chart ─────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum ChartType {
    Area,
    Bar,
    Bar3d,
    Bubble,
    Bubble3d,
    Doughnut,
    Line,
    Pie,
    Radar,
    Scatter,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ChartData {
    pub name: Option<String>,
    pub labels: Option<Vec<String>>,
    pub values: Vec<f64>,
    pub sizes: Option<Vec<f64>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ChartOptions {
    #[serde(flatten)]
    pub pos: Position,

    // Title
    #[serde(rename = "showTitle")]
    pub show_title: Option<bool>,
    pub title: Option<String>,
    #[serde(rename = "titleAlign")]
    pub title_align: Option<String>,
    #[serde(rename = "titleColor")]
    pub title_color: Option<String>,
    #[serde(rename = "titleFontFace")]
    pub title_font_face: Option<String>,
    #[serde(rename = "titleFontSize")]
    pub title_font_size: Option<f64>,
    #[serde(rename = "titlePos")]
    pub title_pos: Option<String>,
    #[serde(rename = "titleRotate")]
    pub title_rotate: Option<f64>,

    // Legend
    #[serde(rename = "showLegend")]
    pub show_legend: Option<bool>,
    #[serde(rename = "legendPos")]
    pub legend_pos: Option<LegendPos>,
    #[serde(rename = "legendFontFace")]
    pub legend_font_face: Option<String>,
    #[serde(rename = "legendFontSize")]
    pub legend_font_size: Option<f64>,
    #[serde(rename = "legendColor")]
    pub legend_color: Option<String>,

    // Data labels
    #[serde(rename = "showLabel")]
    pub show_label: Option<bool>,
    #[serde(rename = "showValue")]
    pub show_value: Option<bool>,
    #[serde(rename = "showPercent")]
    pub show_percent: Option<bool>,
    #[serde(rename = "showDataTable")]
    pub show_data_table: Option<bool>,
    #[serde(rename = "showDataTableKeys")]
    pub show_data_table_keys: Option<bool>,
    #[serde(rename = "showDataTableHorzBorder")]
    pub show_data_table_horz_border: Option<bool>,
    #[serde(rename = "showDataTableVertBorder")]
    pub show_data_table_vert_border: Option<bool>,
    #[serde(rename = "showDataTableOutline")]
    pub show_data_table_outline: Option<bool>,
    #[serde(rename = "dataLabelPosition")]
    pub data_label_position: Option<String>,
    #[serde(rename = "dataLabelFontSize")]
    pub data_label_font_size: Option<f64>,
    #[serde(rename = "dataLabelFontFace")]
    pub data_label_font_face: Option<String>,
    #[serde(rename = "dataLabelFontBold")]
    pub data_label_font_bold: Option<bool>,
    #[serde(rename = "dataLabelColor")]
    pub data_label_color: Option<String>,
    #[serde(rename = "dataLabelFormatCode")]
    pub data_label_format_code: Option<String>,

    // Colors
    #[serde(rename = "chartColors")]
    pub chart_colors: Option<Vec<String>>,
    #[serde(rename = "chartColorsOpacity")]
    pub chart_colors_opacity: Option<f64>,

    // Category axis
    #[serde(rename = "catAxisTitle")]
    pub cat_axis_title: Option<String>,
    #[serde(rename = "catAxisLabelPos")]
    pub cat_axis_label_pos: Option<String>,
    #[serde(rename = "catAxisLabelRotate")]
    pub cat_axis_label_rotate: Option<f64>,
    #[serde(rename = "catAxisLineStyle")]
    pub cat_axis_line_style: Option<String>,
    #[serde(rename = "catAxisMajorUnit")]
    pub cat_axis_major_unit: Option<f64>,
    #[serde(rename = "catAxisOrientation")]
    pub cat_axis_orientation: Option<String>,

    // Value axis
    #[serde(rename = "valAxisTitle")]
    pub val_axis_title: Option<String>,
    #[serde(rename = "valAxisMaxVal")]
    pub val_axis_max_val: Option<f64>,
    #[serde(rename = "valAxisMinVal")]
    pub val_axis_min_val: Option<f64>,
    #[serde(rename = "valAxisMajorUnit")]
    pub val_axis_major_unit: Option<f64>,
    #[serde(rename = "valAxisLogScaleBase")]
    pub val_axis_log_scale_base: Option<f64>,
    #[serde(rename = "valAxisDisplayUnit")]
    pub val_axis_display_unit: Option<String>,
    #[serde(rename = "valAxisLabelFormatCode")]
    pub val_axis_label_format_code: Option<String>,

    // Bar-specific
    #[serde(rename = "barDir")]
    pub bar_dir: Option<BarDir>,
    #[serde(rename = "barGrouping")]
    pub bar_grouping: Option<BarGrouping>,
    #[serde(rename = "barGapWidthPct")]
    pub bar_gap_width_pct: Option<f64>,
    #[serde(rename = "barOverlapPct")]
    pub bar_overlap_pct: Option<f64>,

    // Line-specific
    #[serde(rename = "lineSize")]
    pub line_size: Option<f64>,
    #[serde(rename = "lineSmooth")]
    pub line_smooth: Option<bool>,
    #[serde(rename = "lineDash")]
    pub line_dash: Option<String>,
    #[serde(rename = "lineDataSymbol")]
    pub line_data_symbol: Option<String>,
    #[serde(rename = "lineDataSymbolSize")]
    pub line_data_symbol_size: Option<f64>,

    // 3D
    #[serde(rename = "bar3DShape")]
    pub bar3d_shape: Option<Bar3dShape>,
    #[serde(rename = "v3DRotX")]
    pub v3d_rot_x: Option<f64>,
    #[serde(rename = "v3DRotY")]
    pub v3d_rot_y: Option<f64>,
    #[serde(rename = "v3DPerspective")]
    pub v3d_perspective: Option<f64>,
    #[serde(rename = "v3DRAngAx")]
    pub v3d_r_ang_ax: Option<bool>,

    // Combo-chart multi-axis
    #[serde(rename = "secondaryValAxis")]
    pub secondary_val_axis: Option<bool>,
    #[serde(rename = "secondaryCatAxis")]
    pub secondary_cat_axis: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum LegendPos {
    #[serde(rename = "b")]  Bottom,
    #[serde(rename = "tr")] TopRight,
    #[serde(rename = "l")]  Left,
    #[serde(rename = "r")]  Right,
    #[serde(rename = "t")]  Top,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum BarDir {
    Col,
    Bar,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum BarGrouping {
    Clustered,
    Stacked,
    PercentStacked,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Bar3dShape {
    Box,
    Cylinder,
    ConeToMax,
    Pyramid,
    PyramidToMax,
}

// ── Top-level element discriminated union ─────────────────────────────────────

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "lowercase")]
pub enum SlideElement {
    Text {
        text: TextContent,
        options: TextOptions,
    },
    Image {
        options: ImageOptions,
    },
    Shape {
        #[serde(rename = "shapeType")]
        shape_type: String,
        options: ShapeOptions,
    },
    Table {
        data: Vec<Vec<TableCell>>,
        options: TableOptions,
    },
    Chart {
        #[serde(rename = "chartType")]
        chart_type: ChartType,
        /// For combo charts this holds all series; chart_type = primary type
        data: Vec<ChartData>,
        /// Additional chart types for combo charts (empty for single-type charts)
        #[serde(rename = "comboTypes", default)]
        combo_types: Vec<ChartType>,
        options: ChartOptions,
    },
    Notes {
        text: String,
    },
}
