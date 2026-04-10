use wasm_bindgen::prelude::*;
use crate::model::{
    elements::{
        ChartData, ChartOptions, ChartType, ImageOptions, ShapeOptions, SlideElement,
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
        });
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
        });
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
        });
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
    }

    // ── Background ────────────────────────────────────────────────────────────

    /// Set the slide background color (hex string, e.g. `"FF0000"`).
    #[wasm_bindgen(js_name = setBackground)]
    pub fn set_background(&mut self, color: &str) {
        self.inner.background.color = Some(color.to_string());
    }

    // ── Introspection ─────────────────────────────────────────────────────────

    /// Return all elements on this slide as a JSON-serialized array.
    ///
    /// Each element is a `SlideElement` discriminated union tagged with `"type"`.
    #[wasm_bindgen(js_name = getElements)]
    pub fn get_elements(&self) -> Result<JsValue, JsValue> {
        serde_wasm_bindgen::to_value(&self.inner.elements)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }
}

impl JsSlide {
    pub fn from_slide(slide: Slide) -> Self {
        Self { inner: slide }
    }

    pub fn into_slide(self) -> Slide {
        self.inner
    }
}
