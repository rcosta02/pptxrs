use wasm_bindgen::prelude::*;
use js_sys::Uint8Array;
use serde_json;

use crate::builder::build_pptx;
use crate::measure::{measure_text, FontRegistry, MeasureOptions};
use crate::model::{
    master::SlideMaster,
    presentation::{Layout, Presentation, PresentationMeta},
    slide::{Slide, SlideBackground},
};
use crate::parser::parse_pptx;

use super::slide::JsSlide;

/// ```js
/// import { Presentation } from 'pptxrs';
/// const pres = new Presentation();
/// const slide = pres.addSlide();
/// slide.addText('Hello world', { x: 1, y: 1, w: 8, h: 1, fontSize: 36 });
/// const buf = await pres.write('nodebuffer');
/// ```
#[wasm_bindgen]
pub struct JsPresentation {
    inner: Presentation,
    fonts: FontRegistry,
}

#[wasm_bindgen]
impl JsPresentation {
    // ── Construction ──────────────────────────────────────────────────────────

    /// Create a new empty presentation.
    ///
    /// `options` — `PresentationOptions` (all fields optional)
    #[wasm_bindgen(constructor)]
    pub fn new(options: JsValue) -> Self {
        let mut meta = PresentationMeta::default();
        if let Ok(opts) = serde_wasm_bindgen::from_value::<serde_json::Value>(options) {
            if let Some(v) = opts.get("layout").and_then(|v| v.as_str()) {
                meta.layout = match v {
                    "LAYOUT_4x3"  => Layout::Layout4x3,
                    "LAYOUT_WIDE" => Layout::LayoutWide,
                    _             => Layout::Layout16x9,
                };
            }
            if let Some(v) = opts.get("title").and_then(|v| v.as_str()) {
                meta.title = Some(v.to_string());
            }
            if let Some(v) = opts.get("author").and_then(|v| v.as_str()) {
                meta.author = Some(v.to_string());
            }
            if let Some(v) = opts.get("company").and_then(|v| v.as_str()) {
                meta.company = Some(v.to_string());
            }
        }
        Self {
            inner: Presentation { meta, ..Default::default() },
            fonts: FontRegistry::new(),
        }
    }

    /// Import an existing `.pptx` file from a `Uint8Array`.
    ///
    /// ```js
    /// const buf = fs.readFileSync('deck.pptx');
    /// const pres = Presentation.fromBuffer(buf);
    /// ```
    #[wasm_bindgen(js_name = fromBuffer)]
    pub fn from_buffer(data: &[u8]) -> Result<JsPresentation, JsValue> {
        let inner = parse_pptx(data).map_err(|e| JsValue::from_str(&e))?;
        Ok(Self {
            inner,
            fonts: FontRegistry::new(),
        })
    }

    /// Reconstruct a presentation from a `PresentationJson` object (output of `toJson()`).
    ///
    /// ```js
    /// const json = JSON.parse(fs.readFileSync('deck.json', 'utf8'));
    /// const pres = Presentation.fromJson(json);
    /// ```
    #[wasm_bindgen(js_name = fromJson)]
    pub fn from_json(json: JsValue) -> Result<JsPresentation, JsValue> {
        let inner: Presentation = serde_wasm_bindgen::from_value(json)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        Ok(Self {
            inner,
            fonts: FontRegistry::new(),
        })
    }

    // ── Font registration ─────────────────────────────────────────────────────

    /// Register a TTF/OTF font for use in `measureText()`.
    ///
    /// `name` — the font name used to reference it (e.g. `"Calibri"`)
    /// `data` — raw font file bytes (`Uint8Array` / `Buffer`)
    #[wasm_bindgen(js_name = registerFont)]
    pub fn register_font(&mut self, name: &str, data: &[u8]) {
        self.fonts.register(name.to_string(), data.to_vec());
    }

    // ── Text measurement ──────────────────────────────────────────────────────

    /// Measure text dimensions before adding it to a slide.
    ///
    /// `font` must be registered via `registerFont()` first.
    ///
    /// Returns `{ height, width, lines, lineHeight }` — all values in points.
    ///
    /// ```js
    /// pres.registerFont('Calibri', fs.readFileSync('Calibri.ttf'));
    /// const m = pres.measureText('Hello world', {
    ///   font: 'Calibri', fontSize: 24, width: 5
    /// });
    /// console.log(m.height, m.width, m.lines);
    /// ```
    #[wasm_bindgen(js_name = measureText)]
    pub fn measure_text(&self, text: &str, opts: JsValue) -> Result<JsValue, JsValue> {
        let opts_json: serde_json::Value = serde_wasm_bindgen::from_value(opts)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;

        let mo = MeasureOptions {
            font: opts_json["font"].as_str().unwrap_or("").to_string(),
            font_size: opts_json["fontSize"].as_f64().unwrap_or(18.0),
            bold: opts_json["bold"].as_bool().unwrap_or(false),
            italic: opts_json["italic"].as_bool().unwrap_or(false),
            char_spacing: opts_json["charSpacing"].as_f64().unwrap_or(0.0),
            line_spacing_multiple: opts_json["lineSpacingMultiple"].as_f64().unwrap_or(1.0),
            width_inches: opts_json["width"].as_f64(),
        };

        let metrics = measure_text(text, &mo, &self.fonts)
            .map_err(|e| JsValue::from_str(&e))?;

        serde_wasm_bindgen::to_value(&metrics)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }

    // ── Slide masters ─────────────────────────────────────────────────────────

    /// Define a slide master for use as a background template.
    ///
    /// `options` — `SlideMasterOptions` (must include `title`)
    #[wasm_bindgen(js_name = defineSlideMaster)]
    pub fn define_slide_master(&mut self, options: JsValue) -> Result<(), JsValue> {
        let master: SlideMaster = serde_wasm_bindgen::from_value(options)
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        self.inner.masters.push(master);
        Ok(())
    }

    // ── Slide management ──────────────────────────────────────────────────────

    /// Add a new blank slide and return it.
    ///
    /// `masterName` — optional name of a previously defined slide master.
    #[wasm_bindgen(js_name = addSlide)]
    pub fn add_slide(&mut self, master_name: Option<String>) -> JsSlide {
        let slide = if let Some(m) = master_name {
            Slide::with_master(m)
        } else {
            Slide::new()
        };
        self.inner.slides.push(slide);
        // We return a copy — the slide is already pushed into inner
        // Caller mutates it and then calls `syncSlide` or we use a different pattern.
        // For the binding pattern, we give them a JsSlide backed by a clone; on
        // export, we re-serialize all slides (they were mutated via the JsSlide
        // reference tracking approach below).
        //
        // Simpler: we use a push-then-pop-last approach, give them a JsSlide
        // wrapper that holds the slide data, and on export we collect via `getSlides`.
        let slide_data = self.inner.slides.pop().unwrap();
        JsSlide::from_slide(slide_data)
    }

    /// Return all slides in the presentation as `JsSlide` objects.
    ///
    /// Note: modifying the returned slides does NOT update the presentation.
    /// Use `syncSlide(index, slide)` to push changes back.
    #[wasm_bindgen(js_name = getSlides)]
    pub fn get_slides(&self) -> js_sys::Array {
        let arr = js_sys::Array::new();
        for slide in &self.inner.slides {
            let js = JsSlide::from_slide(slide.clone());
            arr.push(&JsValue::from(js));
        }
        arr
    }

    /// Push a (possibly modified) slide back into the presentation at the given index.
    /// If `index` equals the current slide count the slide is appended.
    #[wasm_bindgen(js_name = syncSlide)]
    pub fn sync_slide(&mut self, index: usize, slide: JsSlide) -> Result<(), JsValue> {
        let s = slide.into_slide();
        if index < self.inner.slides.len() {
            self.inner.slides[index] = s;
        } else if index == self.inner.slides.len() {
            self.inner.slides.push(s);
        } else {
            return Err(JsValue::from_str(&format!(
                "syncSlide: index {} out of range ({})",
                index,
                self.inner.slides.len()
            )));
        }
        Ok(())
    }

    /// Remove the slide at `index`.
    #[wasm_bindgen(js_name = removeSlide)]
    pub fn remove_slide(&mut self, index: usize) -> Result<(), JsValue> {
        if index >= self.inner.slides.len() {
            return Err(JsValue::from_str(&format!(
                "removeSlide: index {} out of range",
                index
            )));
        }
        self.inner.slides.remove(index);
        Ok(())
    }

    // ── Metadata setters ──────────────────────────────────────────────────────

    #[wasm_bindgen(setter)]
    pub fn set_layout(&mut self, layout: &str) {
        self.inner.meta.layout = match layout {
            "LAYOUT_4x3"  => Layout::Layout4x3,
            "LAYOUT_WIDE" => Layout::LayoutWide,
            _             => Layout::Layout16x9,
        };
    }

    #[wasm_bindgen(getter)]
    pub fn layout(&self) -> String {
        match self.inner.meta.layout {
            Layout::Layout4x3   => "LAYOUT_4x3".into(),
            Layout::LayoutWide  => "LAYOUT_WIDE".into(),
            Layout::Layout16x9  => "LAYOUT_16x9".into(),
            Layout::LayoutUser  => "LAYOUT_USER".into(),
        }
    }

    #[wasm_bindgen(setter)]
    pub fn set_title(&mut self, v: &str) { self.inner.meta.title = Some(v.to_string()); }
    #[wasm_bindgen(getter)]
    pub fn title(&self) -> Option<String> { self.inner.meta.title.clone() }

    #[wasm_bindgen(setter)]
    pub fn set_author(&mut self, v: &str) { self.inner.meta.author = Some(v.to_string()); }
    #[wasm_bindgen(getter)]
    pub fn author(&self) -> Option<String> { self.inner.meta.author.clone() }

    #[wasm_bindgen(setter)]
    pub fn set_company(&mut self, v: &str) { self.inner.meta.company = Some(v.to_string()); }
    #[wasm_bindgen(getter)]
    pub fn company(&self) -> Option<String> { self.inner.meta.company.clone() }

    // ── Export ────────────────────────────────────────────────────────────────

    /// Export the presentation.
    ///
    /// `outputType`:
    /// - `"nodebuffer"` — Node.js `Buffer`
    /// - `"uint8array"` — `Uint8Array`
    /// - `"base64"` — base64 string
    ///
    /// Note: the slides must have been pushed back via `syncSlide()` before
    /// calling `write()` if they were obtained via `addSlide()`.
    #[wasm_bindgen]
    pub fn write(&self, output_type: &str) -> Result<JsValue, JsValue> {
        let bytes = build_pptx(&self.inner)
            .map_err(|e| JsValue::from_str(&e))?;

        match output_type {
            "base64" => {
                use base64::Engine;
                let s = base64::engine::general_purpose::STANDARD.encode(&bytes);
                Ok(JsValue::from_str(&s))
            }
            _ => {
                // Both "nodebuffer" and "uint8array" return a Uint8Array from WASM;
                // Node.js Buffer is a subclass of Uint8Array and is compatible.
                let arr = Uint8Array::new_with_length(bytes.len() as u32);
                arr.copy_from(&bytes);
                Ok(arr.into())
            }
        }
    }

    // ── JSON interchange ──────────────────────────────────────────────────────

    /// Return the full presentation as a JSON-serializable JS object.
    ///
    /// The result can be stored, versioned, or passed to `Presentation.fromJson()`.
    #[wasm_bindgen(js_name = toJson)]
    pub fn to_json(&self) -> Result<JsValue, JsValue> {
        serde_wasm_bindgen::to_value(&self.inner)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Return the full presentation serialized as a JSON string.
    #[wasm_bindgen(js_name = toJsonString)]
    pub fn to_json_string(&self) -> Result<String, JsValue> {
        serde_json::to_string(&self.inner)
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }
}
