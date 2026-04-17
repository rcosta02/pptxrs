use serde::{Deserialize, Serialize};
use crate::model::{Slide, SlideMaster};

/// Serde module: serialize `Option<Vec<u8>>` as an optional base64 string.
///
/// The field is omitted from JSON when `None` (so new empty presentations have no bloat).
/// On deserialisation, a base64 string is decoded back to bytes.
mod base64_bytes_opt {
    use serde::{Deserializer, Serializer, Deserialize};
    use base64::{Engine as _, engine::general_purpose::STANDARD};

    pub fn serialize<S>(value: &Option<Vec<u8>>, s: S) -> Result<S::Ok, S::Error>
    where S: Serializer
    {
        match value {
            Some(bytes) => s.serialize_some(&STANDARD.encode(bytes)),
            None        => s.serialize_none(),
        }
    }

    pub fn deserialize<'de, D>(d: D) -> Result<Option<Vec<u8>>, D::Error>
    where D: Deserializer<'de>
    {
        let opt: Option<String> = Option::deserialize(d)?;
        match opt {
            Some(s) => STANDARD.decode(&s)
                .map(Some)
                .map_err(serde::de::Error::custom),
            None => Ok(None),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq, Eq)]
pub enum Layout {
    #[default]
    #[serde(rename = "LAYOUT_16x9")]
    Layout16x9,
    #[serde(rename = "LAYOUT_4x3")]
    Layout4x3,
    #[serde(rename = "LAYOUT_WIDE")]
    LayoutWide,
    #[serde(rename = "LAYOUT_USER")]
    LayoutUser,
}

impl Layout {
    /// Returns (width_emu, height_emu)
    pub fn dimensions_emu(&self) -> (i64, i64) {
        match self {
            Layout::Layout16x9  => (9_144_000, 5_143_500),
            Layout::Layout4x3   => (9_144_000, 6_858_000),
            Layout::LayoutWide  => (12_192_000, 6_858_000),
            Layout::LayoutUser  => (9_144_000, 5_143_500), // fallback to 16x9
        }
    }

    /// Returns (width_inches, height_inches)
    pub fn dimensions_inches(&self) -> (f64, f64) {
        let (w, h) = self.dimensions_emu();
        (w as f64 / 914_400.0, h as f64 / 914_400.0)
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct PresentationMeta {
    pub title: Option<String>,
    pub author: Option<String>,
    pub company: Option<String>,
    pub layout: Layout,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Presentation {
    #[serde(default)]
    pub meta: PresentationMeta,
    #[serde(default)]
    pub masters: Vec<SlideMaster>,
    #[serde(default)]
    pub slides: Vec<Slide>,
    // ── Passthrough fields ────────────────────────────────────────────────────
    /// Original ZIP bytes. Serialised as a base64 string so `fromJson(toJson())`
    /// round-trips the presentation faithfully (slide masters, themes, chart
    /// formatting, etc. are all preserved).  Omitted when `None`.
    #[serde(
        rename = "sourceZipB64",
        default,
        skip_serializing_if = "Option::is_none",
        with = "base64_bytes_opt",
    )]
    pub source_zip: Option<Vec<u8>>,
    /// Number of slides present in the source ZIP (derived — not serialised).
    #[serde(skip)] pub original_slide_count: usize,
    /// Next chart ID to allocate when new charts are added (derived — not serialised).
    #[serde(skip)] pub next_chart_id: u32,
}

impl Presentation {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_slide(&mut self, master: Option<String>) -> &mut Slide {
        self.slides.push(Slide::new());
        if let Some(m) = master {
            self.slides.last_mut().unwrap().master = Some(m);
        }
        self.slides.last_mut().unwrap()
    }

    pub fn slide_width_emu(&self) -> i64 {
        self.meta.layout.dimensions_emu().0
    }

    pub fn slide_height_emu(&self) -> i64 {
        self.meta.layout.dimensions_emu().1
    }
}
