use serde::{Deserialize, Serialize};
use crate::model::elements::SlideElement;

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct SlideBackground {
    pub color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Slide {
    pub background: SlideBackground,
    pub elements: Vec<SlideElement>,
    /// Name of the slide master to use (if any)
    pub master: Option<String>,
    /// Speaker notes
    pub notes: Option<String>,
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
