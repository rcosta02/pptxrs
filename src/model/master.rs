use serde::{Deserialize, Serialize};
use crate::model::elements::{ImageOptions, LineOptions, ShapeOptions, TextContent, TextOptions};

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct SlideMasterBackground {
    pub color: Option<String>,
    pub transparency: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "lowercase")]
pub enum MasterObject {
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
    Line {
        x: f64,
        y: f64,
        x2: f64,
        y2: f64,
        options: Option<LineOptions>,
    },
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct SlideNumberOptions {
    pub x: Option<f64>,
    pub y: Option<f64>,
    pub w: Option<f64>,
    pub h: Option<f64>,
    pub align: Option<String>,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideMaster {
    /// Unique name referenced via `addSlide(masterName)`
    pub title: String,
    pub background: Option<SlideMasterBackground>,
    pub margin: Option<serde_json::Value>, // f64 or [f64;4]
    pub objects: Vec<MasterObject>,
    pub slide_number: Option<SlideNumberOptions>,
}
