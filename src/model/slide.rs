use serde::{Deserialize, Serialize};
use crate::model::elements::SlideElement;

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct SlideBackground {
    pub color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Slide {
    #[serde(default)]
    pub background: SlideBackground,
    #[serde(default)]
    pub elements: Vec<SlideElement>,
    /// Name of the slide master to use (if any)
    pub master: Option<String>,
    /// Speaker notes
    pub notes: Option<String>,
    // ── Passthrough fields — never serialised ─────────────────────────────────
    /// Original raw slide XML (from ZIP); `None` for newly-created slides.
    #[serde(skip)] pub raw_xml: Option<String>,
    /// Original raw rels XML (from ZIP); `None` for newly-created slides.
    #[serde(skip)] pub raw_rels: Option<String>,
    /// `true` when the slide has been mutated since it was loaded.
    /// Triggers surgical rebuild in the passthrough writer.
    #[serde(skip)] pub dirty: bool,
    /// Number of elements after initial parse.  New elements start at this index.
    #[serde(skip)] pub original_element_count: usize,
}

impl Slide {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_master(master: impl Into<String>) -> Self {
        Self {
            master: Some(master.into()),
            ..Default::default()
        }
    }
}
