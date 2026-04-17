use wasm_bindgen::prelude::*;
use crate::bindings::element::JsSlideElement;
use crate::model::{
    elements::{
        ChartData, ChartOptions, ChartType, CoordVal, ImageOptions, ShapeOptions, SlideElement,
        TableCell, TableOptions, TextContent, TextOptions,
    },
    slide::Slide,
};

/// JS-facing wrapper around a single slide.
///
/// Obtain via `Presentation.addSlide()` or `Presentation.getSlides()`.
#[wasm_bindgen]
pub struct JsSlide {
    pub(crate) inner: Slide,
    /// Width of the parent presentation slide in EMU (needed to resolve `%` coords).
    slide_width_emu: i64,
    /// Height of the parent presentation slide in EMU.
    slide_height_emu: i64,
}

#[wasm_bindgen]
impl JsSlide {
    // ── Text ──────────────────────────────────────────────────────────────────

    /// Add a text element to the slide.
    ///
    /// `text` — plain string or JSON-encoded `TextRun[]`
    /// `options` — `TextOptions` object (x, y, w, h required)
    #[wasm_bindgen(js_name = addText)]
    pub fn add_text(&mut self, text: JsValue, options: JsValue) -> Result<(), JsValue> {
        let opts: TextOptions = serde_wasm_bindgen::from_value(options)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;

        let content: TextContent = if text.is_string() {
            TextContent::Plain(text.as_string().unwrap())
        } else {
            serde_wasm_bindgen::from_value::<TextContent>(text)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        };

        self.inner.elements.push(SlideElement::Text {
            text: content,
            options: opts,
        });
        self.inner.dirty = true;
        Ok(())
    }

    // ── Image ─────────────────────────────────────────────────────────────────

    /// Add an image to the slide.
    ///
    /// `options.data` — base64-encoded image bytes
    /// `options.path` — Node.js filesystem path (resolved at call time)
    #[wasm_bindgen(js_name = addImage)]
    pub fn add_image(&mut self, options: JsValue) -> Result<(), JsValue> {
        let mut opts: ImageOptions = serde_wasm_bindgen::from_value(options)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;

        // Resolve Node.js path to base64 if `data` is absent
        if opts.data.is_none() {
            if let Some(path) = &opts.path {
                // In WASM/Node context we read via js-sys / node:fs.
                // We delegate to a JS helper injected at init time.
                // For now, emit a clear error so the caller can pre-encode.
                return Err(JsValue::from_str(
                    "addImage: use `options.data` (base64) instead of `options.path` in WASM context",
                ));
            }
        }

        self.inner.elements.push(SlideElement::Image { options: opts });
        self.inner.dirty = true;
        Ok(())
    }

    // ── Shape ─────────────────────────────────────────────────────────────────

    /// Add a preset shape.
    ///
    /// `shapeType` — e.g. `"rect"`, `"ellipse"`, `"rightArrow"` …
    #[wasm_bindgen(js_name = addShape)]
    pub fn add_shape(&mut self, shape_type: &str, options: JsValue) -> Result<(), JsValue> {
        let opts: ShapeOptions = serde_wasm_bindgen::from_value(options)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        self.inner.elements.push(SlideElement::Shape {
            shape_type: shape_type.to_string(),
            options: opts,
        });
        self.inner.dirty = true;
        Ok(())
    }

    // ── Table ─────────────────────────────────────────────────────────────────

    /// Add a table.
    ///
    /// `data` — 2-D array: `string[][]` or `TableCell[][]`
    /// `options` — `TableOptions`
    #[wasm_bindgen(js_name = addTable)]
    pub fn add_table(&mut self, data: JsValue, options: JsValue) -> Result<(), JsValue> {
        let rows: Vec<Vec<TableCell>> = serde_wasm_bindgen::from_value(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let opts: TableOptions = serde_wasm_bindgen::from_value(options)
            .unwrap_or_default();
        self.inner.elements.push(SlideElement::Table {
            data: rows,
            options: opts,
            frame_index: None,
            raw_frame_xml: None,
            modified: false,
        });
        self.inner.dirty = true;
        Ok(())
    }

    // ── Chart ─────────────────────────────────────────────────────────────────

    /// Add a chart.
    ///
    /// `chartType` — `"bar"` | `"line"` | `"pie"` | …
    /// `data` — `ChartData[]`
    /// `options` — `ChartOptions`
    #[wasm_bindgen(js_name = addChart)]
    pub fn add_chart(
        &mut self,
        chart_type: &str,
        data: JsValue,
        options: JsValue,
    ) -> Result<(), JsValue> {
        let ct: ChartType = serde_json::from_str(&format!("\"{}\"", chart_type))
            .map_err(|e| JsValue::from_str(&format!("unknown chartType '{}': {}", chart_type, e)))?;
        let series: Vec<ChartData> = serde_wasm_bindgen::from_value(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let opts: ChartOptions = serde_wasm_bindgen::from_value(options)
            .unwrap_or_default();
        self.inner.elements.push(SlideElement::Chart {
            chart_type: ct,
            data: series,
            combo_types: vec![],
            options: opts,
            source_chart_path: None,
            frame_index: None,
            modified: false,
        });
        self.inner.dirty = true;
        Ok(())
    }

    /// Add a combo chart (multiple chart types on one frame).
    ///
    /// `chartTypes` — array of chart type strings (first = primary)
    /// `data` — parallel array of `ChartData[]` per type
    #[wasm_bindgen(js_name = addComboChart)]
    pub fn add_combo_chart(
        &mut self,
        chart_types: JsValue,
        data: JsValue,
        options: JsValue,
    ) -> Result<(), JsValue> {
        let type_strs: Vec<String> = serde_wasm_bindgen::from_value(chart_types)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let mut types: Vec<ChartType> = Vec::new();
        for t in &type_strs {
            let ct: ChartType = serde_json::from_str(&format!("\"{}\"", t))
                .map_err(|e| JsValue::from_str(&format!("unknown chartType '{}': {}", t, e)))?;
            types.push(ct);
        }
        let primary = types.remove(0);
        let series_of_series: Vec<Vec<ChartData>> = serde_wasm_bindgen::from_value(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        let flat: Vec<ChartData> = series_of_series.into_iter().flatten().collect();
        let opts: ChartOptions = serde_wasm_bindgen::from_value(options)
            .unwrap_or_default();
        self.inner.elements.push(SlideElement::Chart {
            chart_type: primary,
            data: flat,
            combo_types: types,
            options: opts,
            source_chart_path: None,
            frame_index: None,
            modified: false,
        });
        self.inner.dirty = true;
        Ok(())
    }

    // ── Notes ─────────────────────────────────────────────────────────────────

    /// Set speaker notes for this slide (plain text).
    #[wasm_bindgen(js_name = addNotes)]
    pub fn add_notes(&mut self, text: &str) {
        self.inner.notes = Some(text.to_string());
        // Also push as a Notes element for JSON round-trip fidelity
        self.inner.elements.push(SlideElement::Notes {
            text: text.to_string(),
        });
        self.inner.dirty = true;
    }

    // ── Background ────────────────────────────────────────────────────────────

    /// Set the slide background color (hex string, e.g. `"FF0000"`).
    #[wasm_bindgen(js_name = setBackground)]
    pub fn set_background(&mut self, color: &str) {
        self.inner.background.color = Some(color.to_string());
        self.inner.dirty = true;
    }

    // ── Update existing elements ──────────────────────────────────────────────

    /// Replace the data for a chart element (identified by its index in `getElements()`).
    ///
    /// Marks the slide dirty so the chart XML is regenerated on `write()`.
    /// Works for both charts loaded from an existing file and charts created fresh.
    ///
    /// ```js
    /// const slides = pres.getSlides();
    /// // slide index 0, element index 1 is a chart
    /// slides[0].updateChart(1, [{ name: 'Sales', labels: ['Q1','Q2'], values: [10, 20] }]);
    /// pres.syncSlide(0, slides[0]);
    /// ```
    #[wasm_bindgen(js_name = updateChart)]
    pub fn update_chart(&mut self, index: usize, data: JsValue) -> Result<(), JsValue> {
        let series: Vec<ChartData> = serde_wasm_bindgen::from_value(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;

        let el = self.inner.elements.get_mut(index).ok_or_else(|| {
            JsValue::from_str(&format!("updateChart: no element at index {}", index))
        })?;

        if let SlideElement::Chart { data, modified, .. } = el {
            *data = series;
            *modified = true;
            self.inner.dirty = true;
            Ok(())
        } else {
            Err(JsValue::from_str(&format!(
                "updateChart: element at index {} is not a chart",
                index
            )))
        }
    }

    /// Replace the cell data for a table element (identified by its index in `getElements()`).
    ///
    /// Preserves all original formatting (borders, fonts, colors) — only cell text is changed.
    /// Marks the slide dirty so the table XML is patched on `write()`.
    ///
    /// ```js
    /// const slides = pres.getSlides();
    /// slides[0].updateTable(0, [['R1C1', 'R1C2'], ['R2C1', 'R2C2']]);
    /// pres.syncSlide(0, slides[0]);
    /// ```
    #[wasm_bindgen(js_name = updateTable)]
    pub fn update_table(&mut self, index: usize, data: JsValue) -> Result<(), JsValue> {
        let rows: Vec<Vec<TableCell>> = serde_wasm_bindgen::from_value(data)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;

        let el = self.inner.elements.get_mut(index).ok_or_else(|| {
            JsValue::from_str(&format!("updateTable: no element at index {}", index))
        })?;

        if let SlideElement::Table { data, modified, .. } = el {
            *data = rows;
            *modified = true;
            self.inner.dirty = true;
            Ok(())
        } else {
            Err(JsValue::from_str(&format!(
                "updateTable: element at index {} is not a table",
                index
            )))
        }
    }

    // ── Introspection ─────────────────────────────────────────────────────────

    /// Return all elements on this slide as an array of `JsSlideElement`.
    ///
    /// Each element exposes `getWidth()`, `getHeight()`, `getX()`, `getY()`
    /// (pixels at 96 DPI) plus `getWidthInches()` / `getHeightInches()` /
    /// `getXInches()` / `getYInches()` for inch values.  Call `toJson()` on
    /// any element to access the full options/data payload.
    #[wasm_bindgen(js_name = getElements)]
    pub fn get_elements(&self) -> js_sys::Array {
        let arr = js_sys::Array::new();
        for el in &self.inner.elements {
            let js_el = JsSlideElement::new(
                el.clone(),
                self.slide_width_emu,
                self.slide_height_emu,
            );
            arr.push(&JsValue::from(js_el));
        }
        arr
    }

    /// Return the pixel-perfect bounding box of the element at `index`.
    ///
    /// All values are in **CSS pixels at 96 DPI**, matching what browsers and
    /// most rendering pipelines use as "1 inch = 96 px".
    ///
    /// Returns an object `{ x, y, w, h }`.  Throws if `index` is out of range
    /// or the element is a Notes element (which has no spatial bounds).
    ///
    /// Works for elements added to a brand-new `Presentation`, loaded from a
    /// `.pptx` file, or reconstructed from JSON — as long as the `JsSlide` was
    /// obtained from a `JsPresentation` (which supplies the slide dimensions
    /// needed to resolve percentage-based coordinates).
    #[wasm_bindgen(js_name = getElementBounds)]
    pub fn get_element_bounds(&self, index: usize) -> Result<JsValue, JsValue> {
        let el = self.inner.elements.get(index).ok_or_else(|| {
            JsValue::from_str(&format!(
                "getElementBounds: index {} out of range (slide has {} elements)",
                index,
                self.inner.elements.len()
            ))
        })?;

        let bounds = element_bounds_px(el, self.slide_width_emu, self.slide_height_emu)
            .ok_or_else(|| {
                JsValue::from_str(
                    "getElementBounds: Notes elements do not have spatial bounds",
                )
            })?;

        let obj = js_sys::Object::new();
        js_sys::Reflect::set(&obj, &"x".into(), &bounds[0].into())?;
        js_sys::Reflect::set(&obj, &"y".into(), &bounds[1].into())?;
        js_sys::Reflect::set(&obj, &"w".into(), &bounds[2].into())?;
        js_sys::Reflect::set(&obj, &"h".into(), &bounds[3].into())?;
        Ok(obj.into())
    }
}

// ── Internal helpers ──────────────────────────────────────────────────────────

/// Extract `[x_px, y_px, w_px, h_px]` from any spatial element.
/// Returns `None` for Notes elements, which carry no position.
fn element_bounds_px(
    el: &SlideElement,
    slide_w_emu: i64,
    slide_h_emu: i64,
) -> Option<[f64; 4]> {
    let coord = |c: &CoordVal, dim: i64| c.to_pixels(dim);
    let zero = CoordVal::Pixels(0.0);

    match el {
        SlideElement::Text { options, .. } => Some([
            coord(&options.pos.x, slide_w_emu),
            coord(&options.pos.y, slide_h_emu),
            coord(&options.pos.w, slide_w_emu),
            coord(&options.pos.h, slide_h_emu),
        ]),
        SlideElement::Image { options } => Some([
            coord(&options.pos.x, slide_w_emu),
            coord(&options.pos.y, slide_h_emu),
            coord(&options.pos.w, slide_w_emu),
            coord(&options.pos.h, slide_h_emu),
        ]),
        SlideElement::Shape { options, .. } => Some([
            coord(&options.pos.x, slide_w_emu),
            coord(&options.pos.y, slide_h_emu),
            coord(&options.pos.w, slide_w_emu),
            coord(&options.pos.h, slide_h_emu),
        ]),
        SlideElement::Chart { options, .. } => Some([
            coord(&options.pos.x, slide_w_emu),
            coord(&options.pos.y, slide_h_emu),
            coord(&options.pos.w, slide_w_emu),
            coord(&options.pos.h, slide_h_emu),
        ]),
        SlideElement::Table { options, .. } => Some([
            coord(options.x.as_ref().unwrap_or(&zero), slide_w_emu),
            coord(options.y.as_ref().unwrap_or(&zero), slide_h_emu),
            coord(options.w.as_ref().unwrap_or(&zero), slide_w_emu),
            coord(options.h.as_ref().unwrap_or(&zero), slide_h_emu),
        ]),
        SlideElement::Notes { .. } => None,
    }
}

impl JsSlide {
    /// Create a `JsSlide` with full slide-dimension context (preferred).
    ///
    /// `slide_width_emu` / `slide_height_emu` — the parent presentation's slide
    /// dimensions in EMU, used to resolve percentage-based coordinates in
    /// `getElementBounds`.
    pub fn from_slide_with_dims(slide: Slide, slide_width_emu: i64, slide_height_emu: i64) -> Self {
        Self { inner: slide, slide_width_emu, slide_height_emu }
    }

    /// Convenience constructor that assumes LAYOUT_16x9 (9 144 000 × 5 143 500 EMU).
    /// Use `from_slide_with_dims` when the actual layout is known.
    pub fn from_slide(slide: Slide) -> Self {
        Self::from_slide_with_dims(slide, 9_144_000, 5_143_500)
    }

    pub fn into_slide(self) -> Slide {
        self.inner
    }
}
