use wasm_bindgen::prelude::*;
use crate::model::elements::{CoordVal, SlideElement};

/// JS-facing wrapper around a single slide element.
///
/// Obtained from `slide.getElements()`.  Provides typed dimension accessors
/// so you don't have to pick through the raw options object yourself.
///
/// All pixel values use **96 DPI** (CSS standard: 1 inch = 96 px).
#[wasm_bindgen]
pub struct JsSlideElement {
    pub(crate) inner: SlideElement,
    slide_width_emu: i64,
    slide_height_emu: i64,
}

#[wasm_bindgen]
impl JsSlideElement {
    // ── Type ──────────────────────────────────────────────────────────────────

    /// The element kind: `"text"` | `"image"` | `"shape"` | `"table"` |
    /// `"chart"` | `"notes"`.
    #[wasm_bindgen(getter, js_name = elementType)]
    pub fn element_type(&self) -> String {
        match &self.inner {
            SlideElement::Text { .. }  => "text".into(),
            SlideElement::Image { .. } => "image".into(),
            SlideElement::Shape { .. } => "shape".into(),
            SlideElement::Table { .. } => "table".into(),
            SlideElement::Chart { .. } => "chart".into(),
            SlideElement::Notes { .. } => "notes".into(),
        }
    }

    // ── Width ─────────────────────────────────────────────────────────────────

    /// Element width in **pixels** (96 DPI).
    #[wasm_bindgen(js_name = getWidth)]
    pub fn get_width(&self) -> f64 {
        self.w_coord()
            .map(|c| c.to_pixels(self.slide_width_emu))
            .unwrap_or(0.0)
    }

    /// Element width in **inches**.
    #[wasm_bindgen(js_name = getWidthInches)]
    pub fn get_width_inches(&self) -> f64 {
        self.w_coord()
            .map(|c| coord_to_inches(c, self.slide_width_emu))
            .unwrap_or(0.0)
    }

    // ── Height ────────────────────────────────────────────────────────────────

    /// Element height in **pixels** (96 DPI).
    ///
    /// For text elements where `h` was omitted, the height is estimated as
    /// `fontSize / 72 * 96 * 1.5` (one line at 1.5× the font size).
    #[wasm_bindgen(js_name = getHeight)]
    pub fn get_height(&self) -> f64 {
        if let SlideElement::Text { options, .. } = &self.inner {
            let h = options.pos.h.to_pixels(self.slide_height_emu);
            if h == 0.0 {
                return text_height_px(options.font_size);
            }
            return h;
        }
        self.h_coord()
            .map(|c| c.to_pixels(self.slide_height_emu))
            .unwrap_or(0.0)
    }

    /// Element height in **inches**.
    ///
    /// For text elements where `h` was omitted, the height is estimated from font size.
    #[wasm_bindgen(js_name = getHeightInches)]
    pub fn get_height_inches(&self) -> f64 {
        if let SlideElement::Text { options, .. } = &self.inner {
            let h = coord_to_inches(&options.pos.h, self.slide_height_emu);
            if h == 0.0 {
                return text_height_px(options.font_size) / 96.0;
            }
            return h;
        }
        self.h_coord()
            .map(|c| coord_to_inches(c, self.slide_height_emu))
            .unwrap_or(0.0)
    }

    // ── X / Y ─────────────────────────────────────────────────────────────────

    /// Element X position in **pixels** (96 DPI).
    #[wasm_bindgen(js_name = getX)]
    pub fn get_x(&self) -> f64 {
        self.x_coord()
            .map(|c| c.to_pixels(self.slide_width_emu))
            .unwrap_or(0.0)
    }

    /// Element X position in **inches**.
    #[wasm_bindgen(js_name = getXInches)]
    pub fn get_x_inches(&self) -> f64 {
        self.x_coord()
            .map(|c| coord_to_inches(c, self.slide_width_emu))
            .unwrap_or(0.0)
    }

    /// Element Y position in **pixels** (96 DPI).
    #[wasm_bindgen(js_name = getY)]
    pub fn get_y(&self) -> f64 {
        self.y_coord()
            .map(|c| c.to_pixels(self.slide_height_emu))
            .unwrap_or(0.0)
    }

    /// Element Y position in **inches**.
    #[wasm_bindgen(js_name = getYInches)]
    pub fn get_y_inches(&self) -> f64 {
        self.y_coord()
            .map(|c| coord_to_inches(c, self.slide_height_emu))
            .unwrap_or(0.0)
    }

    // ── Full data ─────────────────────────────────────────────────────────────

    /// Return the full element data as a plain JS object (same shape as the
    /// old `getElements()` array entries).  Useful for accessing text content,
    /// style options, chart data, etc.
    #[wasm_bindgen(js_name = toJson)]
    pub fn to_json(&self) -> Result<JsValue, JsValue> {
        serde_wasm_bindgen::to_value(&self.inner)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }
}

// ── Private coordinate extractors ─────────────────────────────────────────────

impl JsSlideElement {
    fn x_coord(&self) -> Option<&CoordVal> {
        let zero = &CoordVal::Pixels(0.0);
        match &self.inner {
            SlideElement::Text  { options, .. } => Some(&options.pos.x),
            SlideElement::Image { options }      => Some(&options.pos.x),
            SlideElement::Shape { options, .. }  => Some(&options.pos.x),
            SlideElement::Chart { options, .. }  => Some(&options.pos.x),
            SlideElement::Table { options, .. }  => Some(options.x.as_ref().unwrap_or(zero)),
            SlideElement::Notes { .. }           => None,
        }
    }

    fn y_coord(&self) -> Option<&CoordVal> {
        let zero = &CoordVal::Pixels(0.0);
        match &self.inner {
            SlideElement::Text  { options, .. } => Some(&options.pos.y),
            SlideElement::Image { options }      => Some(&options.pos.y),
            SlideElement::Shape { options, .. }  => Some(&options.pos.y),
            SlideElement::Chart { options, .. }  => Some(&options.pos.y),
            SlideElement::Table { options, .. }  => Some(options.y.as_ref().unwrap_or(zero)),
            SlideElement::Notes { .. }           => None,
        }
    }

    fn w_coord(&self) -> Option<&CoordVal> {
        let zero = &CoordVal::Pixels(0.0);
        match &self.inner {
            SlideElement::Text  { options, .. } => Some(&options.pos.w),
            SlideElement::Image { options }      => Some(&options.pos.w),
            SlideElement::Shape { options, .. }  => Some(&options.pos.w),
            SlideElement::Chart { options, .. }  => Some(&options.pos.w),
            SlideElement::Table { options, .. }  => Some(options.w.as_ref().unwrap_or(zero)),
            SlideElement::Notes { .. }           => None,
        }
    }

    fn h_coord(&self) -> Option<&CoordVal> {
        let zero = &CoordVal::Pixels(0.0);
        match &self.inner {
            SlideElement::Text  { options, .. } => Some(&options.pos.h),
            SlideElement::Image { options }      => Some(&options.pos.h),
            SlideElement::Shape { options, .. }  => Some(&options.pos.h),
            SlideElement::Chart { options, .. }  => Some(&options.pos.h),
            SlideElement::Table { options, .. }  => Some(options.h.as_ref().unwrap_or(zero)),
            SlideElement::Notes { .. }           => None,
        }
    }

    pub fn new(inner: SlideElement, slide_width_emu: i64, slide_height_emu: i64) -> Self {
        Self { inner, slide_width_emu, slide_height_emu }
    }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/// Estimate the rendered height of a single-line text box in pixels.
/// Mirrors the formula used in the PPTX builder: font_pt / 72 * 96 * 1.5.
fn text_height_px(font_size: Option<f64>) -> f64 {
    let font_pt = font_size.unwrap_or(18.0);
    font_pt / 72.0 * 96.0 * 1.5
}

fn coord_to_inches(c: &CoordVal, slide_dim_emu: i64) -> f64 {
    const PX_PER_INCH: f64 = 96.0;
    const EMU_PER_PX: f64 = 9_525.0;
    match c {
        CoordVal::Pixels(v) => v / PX_PER_INCH,
        CoordVal::Pct(s) => {
            let pct: f64 = s.trim_end_matches('%').parse().unwrap_or(0.0);
            // pct of slide_dim_emu → pixels → inches
            slide_dim_emu as f64 * pct / 100.0 / EMU_PER_PX / PX_PER_INCH
        }
    }
}
