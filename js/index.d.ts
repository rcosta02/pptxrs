// pptxrs — TypeScript declarations
// Create, read, modify, and export .pptx files (Rust/WASM, Node.js)

// ── Shared ────────────────────────────────────────────────────────────────────

/** A coordinate value — pixels (96 DPI number) or a percentage string like `"50%"`. */
export type CoordVal = number | string;

export interface ShadowOptions {
  type: "outer" | "inner" | "none";
  angle?: number;
  blur?: number;
  color?: string;
  offset?: number;
  opacity?: number;
}

export interface HyperlinkOptions {
  url?: string;
  slide?: number;
  tooltip?: string;
}

export interface LineOptions {
  color?: string;
  width?: number;
  dashType?: string;
  beginArrowType?: string;
  endArrowType?: string;
}

// ── Text ──────────────────────────────────────────────────────────────────────

export interface BulletOptions {
  type?: "bullet" | "number";
  code?: string;
  font?: string;
  color?: string;
  size?: number;
  indent?: number;
  numberStartAt?: number;
  style?: string;
}

export interface TextOptions {
  x: CoordVal;
  y: CoordVal;
  w: CoordVal;
  /**
   * Box height in pixels (96 DPI).
   * Optional for text — when omitted, `getHeight()` estimates the height as
   * `fontSize / 72 * 96 * 1.5` (one line at 1.5× the font size).
   */
  h?: CoordVal;
  /** Font size in points. */
  fontSize?: number;
  fontFace?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  strike?: string;
  /** Hex color without `#`, e.g. `"FF0000"`. */
  color?: string;
  /** Background fill hex color, e.g. `"FFFF00"`. */
  fill?: string;
  highlight?: string;
  subscript?: boolean;
  superscript?: boolean;
  align?: "left" | "center" | "right" | "justify";
  valign?: "top" | "middle" | "bottom";
  vert?: string;
  rtlMode?: boolean;
  autoFit?: boolean;
  fit?: "none" | "shrink" | "resize";
  wrap?: boolean;
  breakLine?: boolean;
  softBreakBefore?: boolean;
  lineSpacing?: number;
  lineSpacingMultiple?: number;
  charSpacing?: number;
  baseline?: number;
  paraSpaceBefore?: number;
  paraSpaceAfter?: number;
  indentLevel?: number;
  bullet?: boolean | BulletOptions;
  inset?: number;
  margin?: number | [number, number, number, number];
  rectRadius?: number;
  rotate?: number;
  isTextBox?: boolean;
  hyperlink?: HyperlinkOptions;
  lang?: string;
  line?: LineOptions;
  transparency?: number;
  glow?: { size: number; opacity: number; color?: string };
  outline?: { color: string; size: number };
  shadow?: ShadowOptions;
  placeholder?: "title" | "body";
}

export interface TextRunOptions extends Omit<TextOptions, "x" | "y" | "w" | "h"> {}

export interface TextRun {
  text: string;
  options?: TextRunOptions;
}

// ── Image ─────────────────────────────────────────────────────────────────────

export interface ImageOptions {
  x: CoordVal;
  y: CoordVal;
  w: CoordVal;
  h: CoordVal;
  /** Base64-encoded image bytes (PNG / JPG / GIF / SVG / WEBP). */
  data?: string;
  /** Node.js filesystem path (resolved automatically by the JS wrapper). */
  path?: string;
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
  rounding?: boolean;
  transparency?: number;
  altText?: string;
  hyperlink?: HyperlinkOptions;
  sizing?: {
    type: "contain" | "cover" | "crop";
    w?: CoordVal;
    h?: CoordVal;
    x?: CoordVal;
    y?: CoordVal;
  };
  shadow?: ShadowOptions;
  placeholder?: "title" | "body";
}

// ── Shape ─────────────────────────────────────────────────────────────────────

export type ShapeType =
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "rightTriangle"
  | "diamond"
  | "parallelogram"
  | "trapezoid"
  | "pentagon"
  | "hexagon"
  | "heptagon"
  | "octagon"
  | "star4"
  | "star5"
  | "star6"
  | "star7"
  | "star8"
  | "star10"
  | "star12"
  | "star16"
  | "star24"
  | "star32"
  | "ribbon"
  | "ribbon2"
  | "ellipseRibbon"
  | "ellipseRibbon2"
  | "callout1"
  | "callout2"
  | "callout3"
  | "rightArrow"
  | "leftArrow"
  | "upArrow"
  | "downArrow"
  | "leftRightArrow"
  | "upDownArrow"
  | "bentArrow"
  | "uturnArrow"
  | "curvedRightArrow"
  | "curvedLeftArrow"
  | "heart"
  | "cloud"
  | "sun"
  | "moon"
  | "lightningBolt"
  | "smileyFace"
  | "line"
  | "arc"
  | "donut"
  | "pie"
  | "blockArc"
  | "mathPlus"
  | "mathMinus"
  | "mathMultiply"
  | "mathDivide"
  | "mathEqual"
  | "mathNotEqual"
  | string; // extensible — any OOXML preset geometry name

export interface ShapeOptions {
  x: CoordVal;
  y: CoordVal;
  w: CoordVal;
  h: CoordVal;
  fill?: { color?: string; transparency?: number };
  line?: LineOptions;
  flipH?: boolean;
  flipV?: boolean;
  rotate?: number;
  rectRadius?: number;
  shapeName?: string;
  hyperlink?: HyperlinkOptions;
  shadow?: ShadowOptions;
  /** Optional text rendered inside the shape. */
  text?: string | TextRun[];
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
}

// ── Table ─────────────────────────────────────────────────────────────────────

export interface BorderOptions {
  type?: string;
  pt?: number;
  color?: string;
}

export interface TableCellOptions {
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFace?: string;
  color?: string;
  fill?: string;
  margin?: number | number[];
  border?: BorderOptions | BorderOptions[];
  colspan?: number;
  rowspan?: number;
}

export type TableCell =
  | string
  | { text: string | TextRun[]; options?: TableCellOptions };

export interface TableOptions {
  x?: CoordVal;
  y?: CoordVal;
  w?: CoordVal;
  h?: CoordVal;
  /** Column widths in pixels. Uniform number or per-column array. */
  colW?: number | number[];
  /** Row heights in pixels. Uniform number or per-row array. */
  rowH?: number | number[];
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  fontFace?: string;
  color?: string;
  fill?: string;
  border?: BorderOptions | BorderOptions[];
  margin?: number | number[];
  autoPage?: boolean;
  autoPageCharWeight?: number;
  autoPageLineWeight?: number;
  autoPageRepeatHeader?: boolean;
  autoPageHeaderRows?: number;
  newSlideStartY?: CoordVal;
}

// ── Chart ─────────────────────────────────────────────────────────────────────

export type ChartType =
  | "area"
  | "bar"
  | "bar3d"
  | "bubble"
  | "bubble3d"
  | "doughnut"
  | "line"
  | "pie"
  | "radar"
  | "scatter";

export interface ChartData {
  name?: string;
  /** Category axis labels. */
  labels?: string[];
  values: number[];
  /** Bubble sizes (bubble chart only). */
  sizes?: number[];
}

export interface ChartOptions {
  x: CoordVal;
  y: CoordVal;
  w: CoordVal;
  h: CoordVal;
  // Title
  showTitle?: boolean;
  title?: string;
  titleAlign?: string;
  titleColor?: string;
  titleFontFace?: string;
  titleFontSize?: number;
  titlePos?: string;
  titleRotate?: number;
  // Legend
  showLegend?: boolean;
  legendPos?: "b" | "tr" | "l" | "r" | "t";
  legendFontFace?: string;
  legendFontSize?: number;
  legendColor?: string;
  // Data labels
  showLabel?: boolean;
  showValue?: boolean;
  showPercent?: boolean;
  showDataTable?: boolean;
  showDataTableKeys?: boolean;
  showDataTableHorzBorder?: boolean;
  showDataTableVertBorder?: boolean;
  showDataTableOutline?: boolean;
  dataLabelPosition?:
    | "bestFit"
    | "b"
    | "ctr"
    | "inBase"
    | "inEnd"
    | "l"
    | "outEnd"
    | "r"
    | "t";
  dataLabelFontSize?: number;
  dataLabelFontFace?: string;
  dataLabelFontBold?: boolean;
  dataLabelColor?: string;
  dataLabelFormatCode?: string;
  // Colors
  chartColors?: string[];
  chartColorsOpacity?: number;
  // Category axis
  catAxisTitle?: string;
  catAxisLabelPos?: string;
  catAxisLabelRotate?: number;
  catAxisLineStyle?: string;
  catAxisMajorUnit?: number;
  catAxisOrientation?: string;
  // Value axis
  valAxisTitle?: string;
  valAxisMaxVal?: number;
  valAxisMinVal?: number;
  valAxisMajorUnit?: number;
  valAxisLogScaleBase?: number;
  valAxisDisplayUnit?: string;
  valAxisLabelFormatCode?: string;
  // Bar
  barDir?: "col" | "bar";
  barGrouping?: "clustered" | "stacked" | "percentStacked";
  barGapWidthPct?: number;
  barOverlapPct?: number;
  // Line
  lineSize?: number;
  lineSmooth?: boolean;
  lineDash?: string;
  lineDataSymbol?: string;
  lineDataSymbolSize?: number;
  // 3D
  bar3DShape?: "box" | "cylinder" | "coneToMax" | "pyramid" | "pyramidToMax";
  v3DRotX?: number;
  v3DRotY?: number;
  v3DPerspective?: number;
  v3DRAngAx?: boolean;
  // Combo
  secondaryValAxis?: boolean;
  secondaryCatAxis?: boolean;
}

// ── Slide element union ───────────────────────────────────────────────────────

export type SlideElement =
  | { type: "text";  text: string | TextRun[];    options: TextOptions  }
  | { type: "image";                               options: ImageOptions }
  | { type: "shape"; shapeType: ShapeType;         options: ShapeOptions }
  | { type: "table"; data: TableCell[][];          options: TableOptions }
  | { type: "chart"; chartType: ChartType; data: ChartData[]; options: ChartOptions }
  | { type: "notes"; text: string };

// ── SlideElementObject ────────────────────────────────────────────────────────

/**
 * A live element handle returned by `slide.getElements()`.
 *
 * Provides pixel-perfect dimension accessors (96 DPI) and their inch equivalents
 * for every spatial element type. Call `toJson()` to access the complete
 * options payload — text content, font size, colors, fill, chart series, etc.
 *
 * Dimension methods return `0` for `"notes"` elements (no spatial bounds).
 *
 * Works identically whether the slide came from `new Presentation()`,
 * `Presentation.fromBuffer()`, or `Presentation.fromJson()`.
 */
export declare class SlideElementObject {
  /** Element kind. */
  readonly elementType: "text" | "image" | "shape" | "table" | "chart" | "notes";

  /** Width in **pixels** (96 DPI). */
  getWidth(): number;
  /** Width in **inches**. */
  getWidthInches(): number;

  /**
   * Height in **pixels** (96 DPI).
   *
   * For text elements where `h` was omitted, the height is estimated as
   * `fontSize / 72 * 96 * 1.5` (one line at 1.5× font size).
   */
  getHeight(): number;
  /** Height in **inches**. */
  getHeightInches(): number;

  /** X position in **pixels** (96 DPI). */
  getX(): number;
  /** X position in **inches**. */
  getXInches(): number;

  /** Y position in **pixels** (96 DPI). */
  getY(): number;
  /** Y position in **inches**. */
  getYInches(): number;

  /**
   * Full element data as a plain JS object.
   *
   * The returned shape matches the `SlideElement` discriminated union, including
   * all styling options (`x`, `y`, `w`, `h`, `fontSize`, `color`, `fill`, …).
   */
  toJson(): SlideElement;
}

// ── Slide Master ──────────────────────────────────────────────────────────────

export interface SlideMasterOptions {
  title: string;
  background?: { color?: string; transparency?: number };
  margin?: number | [number, number, number, number];
  objects?: Array<
    | { type: "text"; text: string | TextRun[]; options: TextOptions }
    | { type: "image"; options: ImageOptions }
    | { type: "shape"; shapeType: ShapeType; options: ShapeOptions }
    | { type: "line"; x: number; y: number; x2: number; y2: number; options?: LineOptions }
  >;
  slideNumber?: {
    x?: CoordVal;
    y?: CoordVal;
    w?: CoordVal;
    h?: CoordVal;
    align?: "left" | "center" | "right";
    color?: string;
  };
}

// ── Text Measurement ──────────────────────────────────────────────────────────

export interface MeasureOptions {
  /** Font name — must be registered via `registerFont()`. */
  font: string;
  /** Font size in points. */
  fontSize: number;
  bold?: boolean;
  italic?: boolean;
  /** Additional char spacing in points. */
  charSpacing?: number;
  /** Line spacing multiplier (default `1.0`). */
  lineSpacingMultiple?: number;
  /** Text box width in **pixels** (96 DPI). Enables word-wrap when set. */
  width?: number;
}

export interface TextMetrics {
  /** Total text block height in points. */
  height: number;
  /** Width of the longest rendered line in points. */
  width: number;
  /** Number of rendered lines. */
  lines: number;
  /** Per-line height in points. */
  lineHeight: number;
}

// ── Presentation Options ──────────────────────────────────────────────────────

export interface PresentationOptions {
  layout?: "LAYOUT_16x9" | "LAYOUT_4x3" | "LAYOUT_WIDE";
  title?: string;
  author?: string;
  company?: string;
}

// ── JSON Schema ───────────────────────────────────────────────────────────────

export interface PresentationJson {
  meta: {
    title?: string;
    author?: string;
    company?: string;
    layout: "LAYOUT_16x9" | "LAYOUT_4x3" | "LAYOUT_WIDE" | string;
  };
  slides: SlideJson[];
  masters?: SlideMasterOptions[];
}

export interface SlideJson {
  background?: { color?: string };
  elements: SlideElement[];
  master?: string;
}

// ── Slide class ───────────────────────────────────────────────────────────────

export declare class Slide {
  /**
   * Add text to the slide.
   *
   * `h` is optional — when omitted, height is estimated from `fontSize`
   * (`fontSize / 72 * 96 * 1.5` px).
   */
  addText(text: string | TextRun[], options: TextOptions): this;

  /**
   * Add an image.
   *
   * `options.path` is resolved from the filesystem automatically in Node.js,
   * so you can pass a relative or absolute path without reading the file yourself.
   */
  addImage(options: ImageOptions): this;

  /** Add a preset shape. */
  addShape(shapeType: ShapeType, options: ShapeOptions): this;

  /** Add a table. */
  addTable(data: TableCell[][], options?: TableOptions): this;

  /** Add a chart. */
  addChart(chartType: ChartType, data: ChartData[], options?: ChartOptions): this;

  /** Add a combo chart (multiple chart types on one set of axes). */
  addComboChart(chartTypes: ChartType[], data: ChartData[][], options?: ChartOptions): this;

  /** Set speaker notes for this slide. */
  addNotes(text: string): this;

  /** Set the slide background color (hex without `#`, e.g. `"002060"`). */
  setBackground(color: string): this;

  /**
   * Update the data for an existing chart element on this slide.
   *
   * Works for charts parsed from an imported `.pptx` and for freshly-created charts.
   * Preserves all chart formatting (colors, axes, labels) when the slide was imported.
   *
   * After updating, push the slide back with `pres.syncSlide(index, slide)`.
   *
   * @param elementIndex  Index of the chart in `getElements()`.
   * @param data          New series data.
   */
  updateChart(elementIndex: number, data: ChartData[]): this;

  /**
   * Update the cell data for an existing table element on this slide.
   *
   * When the table was parsed from an imported `.pptx`, all original formatting
   * (borders, shading, fonts, colors) is preserved — only the text content changes.
   *
   * After updating, push the slide back with `pres.syncSlide(index, slide)`.
   *
   * @param elementIndex  Index of the table in `getElements()`.
   * @param data          New row/cell data.
   */
  updateTable(elementIndex: number, data: TableCell[][]): this;

  /**
   * Get all elements on this slide as `SlideElementObject` instances.
   *
   * Each object exposes:
   * - `elementType` — the element kind
   * - `getWidth()` / `getHeight()` — dimensions in pixels (96 DPI)
   * - `getX()` / `getY()` — position in pixels (96 DPI)
   * - `getWidthInches()` etc. — inch equivalents
   * - `toJson()` — full element data including all styling options
   *
   * Works on slides from `new Presentation()`, `fromBuffer()`, or `fromJson()`.
   */
  getElements(): SlideElementObject[];
}

// ── Presentation class ────────────────────────────────────────────────────────

export declare class Presentation {
  constructor(options?: PresentationOptions);

  /**
   * Import an existing `.pptx` file.
   *
   * Extracts the following from each element type:
   * - **Text**: position, `fontSize`, `bold`, `italic`, `color`, `align`, `valign`, `wrap`, text content (plain or runs)
   * - **Shape**: position, shape type, fill color, line width & color, optional text
   * - **Image**: position, image data as base64
   * - **Table**: position, column widths, cell text
   * - **Slide**: background fill color
   *
   * Call `slide.getElements()` and `el.toJson()` to access the parsed data.
   */
  static fromBuffer(buffer: Uint8Array | Buffer): Presentation;

  /**
   * Reconstruct a presentation from a `PresentationJson` object (output of `toJson()`).
   *
   * ```js
   * const json = JSON.parse(fs.readFileSync('deck.json', 'utf8'));
   * const pres = Presentation.fromJson(json);
   * ```
   */
  static fromJson(json: PresentationJson): Presentation;

  /** Register a TTF/OTF font for use in `measureText()`. */
  registerFont(name: string, buffer: Uint8Array | Buffer): this;

  /**
   * Measure text dimensions before adding it to a slide.
   *
   * Requires the font to be registered first via `registerFont()`.
   * Returns values in **points**.
   */
  measureText(text: string, options: MeasureOptions): TextMetrics;

  /** Define a named slide master template. */
  defineSlideMaster(options: SlideMasterOptions): this;

  /**
   * Add a new blank slide.
   *
   * **Callback form (recommended)** — auto-syncs the slide back:
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
   */
  addSlide(masterName?: string | null, fn?: (slide: Slide) => void): Slide;

  /** Return all slides as `Slide` instances. */
  getSlides(): Slide[];

  /**
   * Push a modified slide back into the presentation at `index`.
   *
   * Required after modifying a slide obtained via `addSlide()` or `getSlides()`
   * (unless the callback form of `addSlide()` was used).
   */
  syncSlide(index: number, slide: Slide): this;

  /** Remove the slide at `index`. */
  removeSlide(index: number): this;

  layout: "LAYOUT_16x9" | "LAYOUT_4x3" | "LAYOUT_WIDE" | "LAYOUT_USER";
  title: string | undefined;
  author: string | undefined;
  company: string | undefined;

  /**
   * Export the presentation.
   *
   * Note: slides must have been pushed back via `syncSlide()` before calling
   * `write()`. The callback form of `addSlide()` does this automatically.
   */
  write(outputType?: "nodebuffer"): Buffer;
  write(outputType: "uint8array"): Uint8Array;
  write(outputType: "base64"): string;

  /** Write to a file on disk (Node.js only). */
  writeFile(filePath: string): Promise<void>;

  /**
   * Serialize the presentation to a plain JS object.
   *
   * All element options — including position (`x`, `y`, `w`, `h`) and styling
   * fields (`fontSize`, `color`, `fill`, etc.) — are included.
   *
   * The result can be passed directly to `Presentation.fromJson()`.
   */
  toJson(): PresentationJson;

  /**
   * Serialize the presentation to a JSON string.
   *
   * Equivalent to `JSON.stringify(pres.toJson())` but faster.
   */
  toJsonString(): string;
}
