/**
 * pptxrs — Create, read, modify, and export .pptx files.
 *
 * Node.js wrapper around the Rust/WASM core.
 */

"use strict";

const core = require("./pptxrs.js");
const fs = require("fs");
const path = require("path");

/**
 * @example
 * const { Presentation } = require('pptxrs');
 *
 * // ── Create from scratch ────────────────────────────────────────────────────
 * const pres = new Presentation({ layout: 'LAYOUT_16x9', title: 'My Deck' });
 *
 * pres.addSlide(null, slide => {
 *   slide.addText('Hello world', { x: 96, y: 96, w: 768, h: 96, fontSize: 36 });
 * });
 * await pres.writeFile('deck.pptx');
 *
 * // ── Measure text before layout ─────────────────────────────────────────────
 * pres.registerFont('Calibri', fs.readFileSync('Calibri.ttf'));
 * const m = pres.measureText('Hello world', { font: 'Calibri', fontSize: 24 });
 * console.log(m.height, m.width);  // in points
 *
 * // ── Import and inspect an existing file ───────────────────────────────────
 * const pres2 = Presentation.fromBuffer(fs.readFileSync('existing.pptx'));
 * for (const slide of pres2.getSlides()) {
 *   for (const el of slide.getElements()) {
 *     const data = el.toJson();          // full options — fontSize, color, fill, …
 *     console.log(el.elementType, el.getWidth(), el.getHeight()); // pixels
 *   }
 * }
 *
 * // ── Modify and re-export ───────────────────────────────────────────────────
 * const slides = pres2.getSlides();
 * slides.forEach((slide, i) => {
 *   slide.addText('DRAFT', { x: 192, y: 192, w: 576, h: 192, fontSize: 72,
 *                             color: 'FF0000', rotate: 45, bold: true });
 *   pres2.syncSlide(i, slide);
 * });
 * await pres2.writeFile('modified.pptx');
 */
class Presentation {
  constructor(options = {}) {
    this._inner = new core.JsPresentation(options);
  }

  // ── Static factory methods ──────────────────────────────────────────────────

  /**
   * Import an existing .pptx file.
   *
   * Extracts all element types with their full options:
   *  - **Text**: x/y/w/h, fontSize, bold, italic, color, align, valign, wrap
   *  - **Shape**: x/y/w/h, shape type, fill color, line width/color
   *  - **Image**: x/y/w/h, image data (base64)
   *  - **Table**: x/y/w/h, column widths, all cell text
   *  - **Slide**: background fill color
   *
   * @param {Uint8Array|Buffer} buffer
   * @returns {Presentation}
   */
  static fromBuffer(buffer) {
    const p = Object.create(Presentation.prototype);
    p._inner = core.JsPresentation.fromBuffer(
      buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer),
    );
    return p;
  }

  /**
   * Reconstruct a presentation from a PresentationJson object (from `toJson()`).
   * @param {object} json
   * @returns {Presentation}
   */
  static fromJson(json) {
    const p = Object.create(Presentation.prototype);
    p._inner = core.JsPresentation.fromJson(json);
    return p;
  }

  // ── Font registration ───────────────────────────────────────────────────────

  /**
   * Register a TTF/OTF font for use in `measureText()`.
   * @param {string} name
   * @param {Uint8Array|Buffer} buffer  Raw font file bytes
   * @returns {this}
   */
  registerFont(name, buffer) {
    this._inner.registerFont(
      name,
      buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer),
    );
    return this;
  }

  // ── Text measurement ────────────────────────────────────────────────────────

  /**
   * Measure the rendered dimensions of a text string.
   *
   * The font must be registered via `registerFont()` first.
   *
   * @param {string} text
   * @param {{ font: string, fontSize: number, bold?: boolean, italic?: boolean,
   *            charSpacing?: number, lineSpacingMultiple?: number, width?: number }} options
   *   `width` is in **pixels** (96 DPI) and enables word-wrap when set.
   * @returns {{ height: number, width: number, lines: number, lineHeight: number }}
   *   All values in **points**.
   */
  measureText(text, options) {
    return this._inner.measureText(text, options);
  }

  // ── Slide masters ───────────────────────────────────────────────────────────

  /**
   * Define a slide master template.
   * @param {import('./index.d.ts').SlideMasterOptions} options  Must include a unique `title`.
   * @returns {this}
   */
  defineSlideMaster(options) {
    this._inner.defineSlideMaster(options);
    return this;
  }

  // ── Slide management ────────────────────────────────────────────────────────

  /**
   * Add a new blank slide to the presentation.
   *
   * **Callback form (recommended)** — auto-syncs the slide:
   * ```js
   * pres.addSlide(null, slide => {
   *   slide.addText('Hello', { x: 96, y: 96, w: 768, h: 96 });
   * });
   * ```
   *
   * **Manual form** — you must call `syncSlide` when done:
   * ```js
   * const slide = pres.addSlide();
   * slide.addText('Hello', { x: 96, y: 96, w: 768, h: 96 });
   * pres.syncSlide(0, slide);
   * ```
   *
   * @param {string|null} [masterName]  Optional slide master name.
   * @param {(slide: Slide) => void} [fn]  Optional setup callback (auto-syncs).
   * @returns {Slide}
   */
  addSlide(masterName, fn) {
    const slide = new Slide(this._inner.addSlide(masterName ?? undefined));
    if (fn) {
      fn(slide);
      this._inner.syncSlide(this._inner.getSlides().length, slide._inner);
    }
    return slide;
  }

  /**
   * Get all slides in the presentation as `Slide` instances.
   * @returns {Slide[]}
   */
  getSlides() {
    return this._inner.getSlides().map((s) => new Slide(s));
  }

  /**
   * Push a modified slide back into the presentation at `index`.
   *
   * Required after modifying a slide returned by `addSlide()` or `getSlides()`
   * unless the callback form of `addSlide()` was used.
   *
   * @param {number} index
   * @param {Slide} slide
   * @returns {this}
   */
  syncSlide(index, slide) {
    this._inner.syncSlide(index, slide._inner);
    return this;
  }

  /**
   * Remove the slide at `index`.
   * @param {number} index
   * @returns {this}
   */
  removeSlide(index) {
    this._inner.removeSlide(index);
    return this;
  }

  // ── Metadata ────────────────────────────────────────────────────────────────

  get layout() { return this._inner.layout; }
  set layout(v) { this._inner.layout = v; }

  get title()  { return this._inner.title; }
  set title(v) { this._inner.title = v; }

  get author()  { return this._inner.author; }
  set author(v) { this._inner.author = v; }

  get company()  { return this._inner.company; }
  set company(v) { this._inner.company = v; }

  // ── Export ──────────────────────────────────────────────────────────────────

  /**
   * Export the presentation bytes.
   *
   * Note: slides added via `addSlide()` must be pushed back with `syncSlide()`
   * before calling `write()` (the callback form of `addSlide()` does this automatically).
   *
   * @param {'nodebuffer'|'uint8array'|'base64'} [outputType='nodebuffer']
   * @returns {Buffer|Uint8Array|string}
   */
  write(outputType = "nodebuffer") {
    return this._inner.write(outputType);
  }

  /**
   * Write the presentation to a file on disk (Node.js only).
   * @param {string} filePath  Destination file path.
   * @returns {Promise<void>}
   */
  async writeFile(filePath) {
    const buf = this._inner.write("nodebuffer");
    return fs.promises.writeFile(filePath, buf);
  }

  // ── JSON interchange ────────────────────────────────────────────────────────

  /**
   * Serialize the presentation to a plain JS object.
   *
   * All element options — including position fields `x`, `y`, `w`, `h`,
   * and styling fields like `fontSize`, `color`, `fill`, etc. — are included.
   *
   * The result can be passed directly to `Presentation.fromJson()`.
   *
   * @returns {import('./index.d.ts').PresentationJson}
   */
  toJson() {
    return this._inner.toJson();
  }

  /**
   * Serialize the presentation to a JSON string.
   * Equivalent to `JSON.stringify(pres.toJson())` but faster.
   * @returns {string}
   */
  toJsonString() {
    return this._inner.toJsonString();
  }
}

/**
 * Slide — wraps a single slide in the presentation.
 */
class Slide {
  /** @param {import('./pptxrs.js').JsSlide} inner */
  constructor(inner) {
    this._inner = inner;
  }

  /**
   * Add text to the slide.
   *
   * `h` is optional — when omitted, height is estimated from `fontSize`
   * (`fontSize / 72 * 96 * 1.5` px).
   *
   * @param {string|import('./index.d.ts').TextRun[]} text
   * @param {import('./index.d.ts').TextOptions} options
   * @returns {this}
   */
  addText(text, options) {
    this._inner.addText(text, options);
    return this;
  }

  /**
   * Add an image.
   *
   * `options.path` is resolved from the filesystem automatically in Node.js.
   *
   * @param {import('./index.d.ts').ImageOptions} options
   * @returns {this}
   */
  addImage(options) {
    if (options.path && !options.data) {
      const resolved = path.resolve(options.path);
      const bytes = fs.readFileSync(resolved);
      options = { ...options, data: bytes.toString("base64"), path: undefined };
    }
    this._inner.addImage(options);
    return this;
  }

  /**
   * Add a preset shape.
   * @param {import('./index.d.ts').ShapeType} shapeType
   * @param {import('./index.d.ts').ShapeOptions} options
   * @returns {this}
   */
  addShape(shapeType, options) {
    this._inner.addShape(shapeType, options);
    return this;
  }

  /**
   * Add a table.
   * @param {import('./index.d.ts').TableCell[][]} data
   * @param {import('./index.d.ts').TableOptions} [options]
   * @returns {this}
   */
  addTable(data, options = {}) {
    this._inner.addTable(data, options);
    return this;
  }

  /**
   * Add a chart.
   * @param {import('./index.d.ts').ChartType} chartType
   * @param {import('./index.d.ts').ChartData[]} data
   * @param {import('./index.d.ts').ChartOptions} [options]
   * @returns {this}
   */
  addChart(chartType, data, options = {}) {
    this._inner.addChart(chartType, data, options);
    return this;
  }

  /**
   * Add a combo chart (multiple chart types on one axis).
   * @param {import('./index.d.ts').ChartType[]} chartTypes
   * @param {import('./index.d.ts').ChartData[][]} data
   * @param {import('./index.d.ts').ChartOptions} [options]
   * @returns {this}
   */
  addComboChart(chartTypes, data, options = {}) {
    this._inner.addComboChart(chartTypes, data, options);
    return this;
  }

  /**
   * Set speaker notes.
   * @param {string} text
   * @returns {this}
   */
  addNotes(text) {
    this._inner.addNotes(text);
    return this;
  }

  /**
   * Set background color.
   * @param {string} hexColor  Hex color without `#`, e.g. `"FF0000"`.
   * @returns {this}
   */
  setBackground(color) {
    this._inner.setBackground(color);
    return this;
  }

  /**
   * Update the data for an existing chart element on this slide.
   *
   * Works for charts parsed from an imported `.pptx` and for freshly-created charts.
   * Preserves all chart formatting (colors, axes, labels) when the slide was imported.
   *
   * After updating, push the slide back with `pres.syncSlide(index, slide)`.
   *
   * @param {number} elementIndex  Index of the chart in `getElements()`.
   * @param {import('./index.d.ts').ChartData[]} data  New series data.
   * @returns {this}
   */
  updateChart(elementIndex, data) {
    this._inner.updateChart(elementIndex, data);
    return this;
  }

  /**
   * Update the cell data for an existing table element on this slide.
   *
   * When the table was parsed from an imported `.pptx`, all original formatting
   * (borders, shading, fonts, colors) is preserved — only the text content changes.
   *
   * After updating, push the slide back with `pres.syncSlide(index, slide)`.
   *
   * @param {number} elementIndex  Index of the table in `getElements()`.
   * @param {import('./index.d.ts').TableCell[][]} data  New row/cell data.
   * @returns {this}
   */
  updateTable(elementIndex, data) {
    this._inner.updateTable(elementIndex, data);
    return this;
  }

  /**
   * Get all elements on this slide as `SlideElementObject` instances.
   *
   * Each object has:
   * - `elementType` — `"text"` | `"image"` | `"shape"` | `"table"` | `"chart"` | `"notes"`
   * - `getWidth()` / `getHeight()` — dimensions in **pixels** (96 DPI)
   * - `getX()` / `getY()` — position in **pixels** (96 DPI)
   * - `getWidthInches()` / `getHeightInches()` / `getXInches()` / `getYInches()` — in inches
   * - `toJson()` — full element data including all styling options
   *
   * Works on slides from `new Presentation()`, `fromBuffer()`, or `fromJson()`.
   *
   * @returns {import('./index.d.ts').SlideElementObject[]}
   */
  getElements() {
    return this._inner.getElements();
  }
}

module.exports = { Presentation, Slide };
