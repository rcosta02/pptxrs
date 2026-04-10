use serde::{Deserialize, Serialize};
use crate::model::{Slide, SlideMaster};

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
    pub meta: PresentationMeta,
    pub masters: Vec<SlideMaster>,
    pub slides: Vec<Slide>,
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
