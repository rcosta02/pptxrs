mod bindings;
mod builder;
mod measure;
mod model;
mod parser;

pub use bindings::{JsPresentation, JsSlide};

use wasm_bindgen::prelude::*;

/// Called automatically by the WASM runtime on module load.
#[wasm_bindgen(start)]
pub fn init() {}
