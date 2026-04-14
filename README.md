# pptxrs

_Created by [Rafael Costa - @rcosta02](https://github.com/rcosta02)._

Performant Rust/WASM library for creating, reading, and modifying PowerPoint (.pptx) files in Node.js — with text measurement and JSON interchange.

**Why pptxrs over PptxGenJS?**

- **Import existing files** — open any `.pptx`, read its elements, modify them, re-export.
- **Text measurement** — get the exact rendered height and width of a text string (using real font metrics) before writing it to a slide, so you can position other elements relative to it.
- **JSON interchange** — serialize a full presentation to/from JSON for storage, diffing, or templating.
- **Performance** — the core runs in Rust compiled to WASM; no native dependencies.

---

## Installation

```bash
npm install pptxrs
```

Requires Node.js ≥ 16.

---

## Quick start

```js
const { Presentation } = require("pptxrs");
const fs = require("fs");

const pres = new Presentation({ layout: "LAYOUT_16x9", title: "My Deck" });

pres.addSlide(null, (slide) => {
  slide.addText("Hello, pptxrs!", {
    x: 96,
    y: 96,
    w: 768,
    h: 144,
    fontSize: 44,
    bold: true,
  });
});

await pres.writeFile("deck.pptx");
```

---

## Table of contents

1. [Coordinates](#coordinates)
2. [Creating a presentation](#creating-a-presentation)
3. [Adding slides](#adding-slides)
4. [Text](#text)
5. [Images](#images)
6. [Shapes](#shapes)
7. [Tables](#tables)
8. [Charts](#charts)
9. [Speaker notes](#speaker-notes)
10. [Slide masters](#slide-masters)
11. [Exporting](#exporting)
12. [Importing existing .pptx files](#importing-existing-pptx-files)
13. [Reading element dimensions](#reading-element-dimensions-1)
14. [Modifying imported slides](#modifying-imported-slides)
15. [Text measurement](#text-measurement)
16. [JSON interchange](#json-interchange)
17. [Full API reference](#full-api-reference)

---

## Coordinates

All position and size values (`x`, `y`, `w`, `h`) are in **pixels** (96 DPI). You can also pass a percentage string.

| Value    | Meaning                    |
| -------- | -------------------------- |
| `96`     | 96 px = 1 inch             |
| `"50%"`  | 50% of the slide dimension |

Standard slide dimensions at 96 DPI:

| Layout                  | Width (px) | Height (px) |
| ----------------------- | ---------- | ----------- |
| `LAYOUT_16x9` (default) | 960        | 540         |
| `LAYOUT_4x3`            | 960        | 720         |
| `LAYOUT_WIDE`           | 1280       | 720         |

Colors are **hex strings without `#`**, e.g. `"FF0000"` for red.

### Reading element dimensions

`slide.getElements()` returns `SlideElementObject` instances with dimension accessors:

```js
const elements = slide.getElements();
const el = elements[0];

el.getWidth()         // pixels (96 DPI)
el.getHeight()        // pixels (96 DPI)
el.getWidthInches()   // inches
el.getHeightInches()  // inches
el.getX()             // x position in pixels
el.getY()             // y position in pixels
el.elementType        // "text" | "image" | "shape" | "table" | "chart" | "notes"
el.toJson()           // full element data/options object
```

---

## Creating a presentation

```js
const { Presentation } = require("pptxrs");

const pres = new Presentation({
  layout: "LAYOUT_16x9", // 'LAYOUT_4x3' | 'LAYOUT_WIDE'
  title: "Q3 Results",
  author: "Jane Smith",
  company: "Acme Corp",
});

// Metadata can also be set as properties
pres.title = "Updated Title";
pres.author = "John Doe";
pres.company = "ACME";
pres.layout = "LAYOUT_4x3";
```

---

## Adding slides

Use the **callback form** (recommended) to add and configure a slide in one step — it auto-syncs:

```js
pres.addSlide(null, (slide) => {
  slide.addText("Slide 1", { x: 96, y: 96, w: 768, h: 96 });
  slide.setBackground("F0F4FF");
});
```

Or use the **manual form** and call `syncSlide` when done:

```js
const slide = pres.addSlide(); // not yet in the presentation
slide.addText("Hello", { x: 96, y: 96, w: 768, h: 96 });
pres.syncSlide(0, slide); // push it in at index 0
```

To use a named slide master:

```js
pres.addSlide("MASTER_BRAND", (slide) => {
  slide.addText("Content here", { x: 96, y: 192, w: 768, h: 288 });
});
```

---

## Text

### Plain text

```js
slide.addText("Hello world", {
  x: 96,
  y: 96,
  w: 768,
  h: 144,

  // Font
  fontSize: 24, // points
  fontFace: "Calibri",
  bold: true,
  italic: false,
  underline: true,
  strike: "sngStrike", // 'sngStrike' | 'dblStrike'
  color: "4472C4", // hex
  highlight: "FFFF00", // hex

  // Alignment
  align: "center", // 'left' | 'center' | 'right'
  valign: "middle", // 'top' | 'middle' | 'bottom'

  // Spacing
  lineSpacingMultiple: 1.5,
  charSpacing: 2,
  paraSpaceBefore: 6,
  paraSpaceAfter: 6,

  // Box behaviour
  wrap: true,
  autoFit: true, // shrink text to fit box
  fit: "shrink", // 'none' | 'shrink' | 'resize'
  margin: 0.1, // or [top, right, bottom, left] in inches

  // Effects
  shadow: { type: "outer", angle: 45, blur: 4, color: "000000", opacity: 0.5 },
  glow: { size: 8, opacity: 0.4, color: "4472C4" },

  // Hyperlink
  hyperlink: { url: "https://example.com", tooltip: "Visit site" },
  // or link to slide number:
  hyperlink: { slide: 3 },

  // RTL / language
  rtlMode: false,
  lang: "en-US",
});
```

### Mixed formatting (TextRun array)

Pass an array of runs to mix styles within one text box:

```js
slide.addText(
  [
    { text: "Normal text, " },
    { text: "bold red, ", options: { bold: true, color: "FF0000" } },
    { text: "italic blue.", options: { italic: true, color: "0070C0" } },
  ],
  { x: 96, y: 96, w: 768, h: 96, fontSize: 18 },
);
```

### Bullet lists

```js
slide.addText(
  [
    { text: "First item", options: { bullet: true } },
    { text: "Second item", options: { bullet: true } },
    {
      text: "Numbered",
      options: { bullet: { type: "number", style: "arabicPeriod" } },
    },
    {
      text: "Custom char",
      options: { bullet: { code: "2713", color: "70AD47" } },
    },
  ],
  { x: 96, y: 96, w: 768, h: 384, fontSize: 18 },
);
```

Bullet indent levels (1–32):

```js
slide.addText(
  [
    { text: "Level 1", options: { bullet: true, indentLevel: 1 } },
    { text: "Level 2", options: { bullet: true, indentLevel: 2 } },
  ],
  { x: 96, y: 96, w: 768, h: 288 },
);
```

### Superscript / subscript

```js
slide.addText(
  [
    { text: "E = mc" },
    { text: "2", options: { superscript: true, fontSize: 12 } },
  ],
  { x: 96, y: 96, w: 384, h: 96, fontSize: 24 },
);
```

---

## Images

Images can be loaded from the filesystem (auto-resolved) or passed as base64.

```js
// From file path (Node.js — resolved automatically)
slide.addImage({
  path: "./assets/logo.png",
  x: 48,
  y: 48,
  w: 192,
  h: 96,
});

// From base64
const imgData = fs.readFileSync("./photo.jpg").toString("base64");
slide.addImage({
  data: imgData,
  x: 288,
  y: 96,
  w: 384,
  h: 288,
});

// Sizing modes
slide.addImage({
  path: "./bg.jpg",
  x: 0,
  y: 0,
  w: 960,
  h: 540,
  sizing: { type: "cover" }, // 'contain' | 'cover' | 'crop'
});

// Effects
slide.addImage({
  path: "./photo.png",
  x: 96,
  y: 96,
  w: 288,
  h: 288,
  rotate: 15,
  flipH: true,
  rounding: true, // circular crop
  transparency: 20, // 0–100
  shadow: { type: "outer", angle: 45, blur: 6, color: "000000", opacity: 0.4 },
  hyperlink: { url: "https://example.com" },
  altText: "Company logo",
});
```

---

## Shapes

`shapeType` is a string matching PowerPoint preset geometry names.

```js
slide.addShape("rect", {
  x: 96,
  y: 96,
  w: 288,
  h: 192,
  fill: { color: "4472C4", transparency: 20 },
  line: { color: "002060", width: 2, dashType: "dash" },
  shadow: { type: "outer", angle: 45, blur: 4, color: "000000", opacity: 0.3 },
  rotate: 10,
  rectRadius: 0.1, // corner rounding (for roundRect)
});

// Shape with text inside
slide.addShape("roundRect", {
  x: 384,
  y: 96,
  w: 384,
  h: 192,
  fill: { color: "ED7D31" },
  text: "Click me",
  fontSize: 20,
  bold: true,
  color: "FFFFFF",
  align: "center",
  valign: "middle",
  hyperlink: { url: "https://example.com" },
});
```

**Common shape types:**

| Category | Values                                                                                                                                      |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| Basic    | `rect` `roundRect` `ellipse` `triangle` `rightTriangle` `diamond`                                                                           |
| Polygons | `pentagon` `hexagon` `heptagon` `octagon`                                                                                                   |
| Stars    | `star4` `star5` `star6` `star7` `star8` `star10` `star12` `star16` `star24` `star32`                                                        |
| Arrows   | `rightArrow` `leftArrow` `upArrow` `downArrow` `leftRightArrow` `upDownArrow` `bentArrow` `uturnArrow` `curvedRightArrow` `curvedLeftArrow` |
| Callouts | `callout1` `callout2` `callout3`                                                                                                            |
| Symbols  | `heart` `cloud` `sun` `moon` `lightningBolt` `smileyFace`                                                                                   |
| Lines    | `line` `arc`                                                                                                                                |
| Math     | `mathPlus` `mathMinus` `mathMultiply` `mathDivide` `mathEqual` `mathNotEqual`                                                               |
| Misc     | `donut` `pie` `blockArc` `ribbon` `ribbon2`                                                                                                 |

---

## Tables

```js
// Simple string data
slide.addTable(
  [
    ["Name", "Department", "Score"],
    ["Alice", "Engineering", "95"],
    ["Bob", "Design", "87"],
    ["Carol", "PM", "91"],
  ],
  {
    x: 96,
    y: 192,
    w: 768,
    h: 288,
    colW: [288, 288, 192], // column widths in pixels
    fontSize: 14,
    border: { pt: 1, color: "CCCCCC" },
  },
);

// Per-cell formatting
slide.addTable(
  [
    [
      {
        text: "Header",
        options: {
          bold: true,
          fill: "4472C4",
          color: "FFFFFF",
          align: "center",
        },
      },
      {
        text: "Value",
        options: {
          bold: true,
          fill: "4472C4",
          color: "FFFFFF",
          align: "center",
        },
      },
    ],
    ["Row 1", "100"],
    ["Row 2", "200"],
  ],
  { x: 96, y: 192, w: 576, h: 288 },
);

// Cell spanning
slide.addTable(
  [
    [
      {
        text: "Merged header",
        options: { colspan: 3, align: "center", bold: true },
      },
    ],
    ["Col A", "Col B", "Col C"],
  ],
  { x: 96, y: 96, w: 768, h: 192 },
);

// Auto-paging (table continues onto new slides)
slide.addTable(data, {
  x: 48,
  y: 96,
  w: 864,
  autoPage: true,
  autoPageRepeatHeader: true,
  autoPageHeaderRows: 1,
  newSlideStartY: 96,
});
```

---

## Charts

### Bar / column chart

```js
slide.addChart(
  "bar",
  [
    {
      name: "Revenue",
      labels: ["Q1", "Q2", "Q3", "Q4"],
      values: [120, 190, 160, 230],
    },
    {
      name: "Expenses",
      labels: ["Q1", "Q2", "Q3", "Q4"],
      values: [80, 110, 90, 130],
    },
  ],
  {
    x: 48,
    y: 48,
    w: 864,
    h: 480,

    barDir: "col", // 'col' (vertical) | 'bar' (horizontal)
    barGrouping: "clustered", // 'clustered' | 'stacked' | 'percentStacked'

    showTitle: true,
    title: "Quarterly Financials",
    titleFontSize: 14,

    showLegend: true,
    legendPos: "b", // 'b' | 't' | 'l' | 'r' | 'tr'

    showValue: true,
    dataLabelPosition: "outEnd",

    chartColors: ["4472C4", "ED7D31", "A9D18E"],
    catAxisTitle: "Quarter",
    valAxisTitle: "USD (thousands)",
    valAxisMaxVal: 300,
    valAxisMajorUnit: 50,
  },
);
```

### Line chart

```js
slide.addChart(
  "line",
  [
    {
      name: "Series A",
      labels: ["Jan", "Feb", "Mar", "Apr"],
      values: [10, 25, 18, 32],
    },
  ],
  {
    x: 48,
    y: 48,
    w: 864,
    h: 480,
    lineSmooth: true,
    lineSize: 2.5,
    lineDataSymbol: "circle",
    lineDataSymbolSize: 8,
    showTitle: true,
    title: "Monthly Trend",
  },
);
```

### Pie / doughnut chart

```js
slide.addChart(
  "pie",
  [
    {
      labels: ["North", "South", "East", "West"],
      values: [35, 25, 20, 20],
    },
  ],
  {
    x: 96,
    y: 48,
    w: 768,
    h: 480,
    showPercent: true,
    showLegend: true,
    legendPos: "r",
    chartColors: ["4472C4", "ED7D31", "A9D18E", "FFC000"],
  },
);
```

### Scatter / bubble chart

```js
slide.addChart(
  "scatter",
  [{ name: "Group A", labels: ["1", "2", "3"], values: [10, 20, 30] }],
  { x: 48, y: 48, w: 864, h: 480 },
);

slide.addChart(
  "bubble",
  [{ name: "Data", values: [10, 20, 30], sizes: [5, 10, 15] }],
  { x: 48, y: 48, w: 864, h: 480 },
);
```

### Combo chart (multiple types)

```js
slide.addComboChart(
  ["bar", "line"], // first type = primary
  [
    // Data for 'bar' series
    [
      {
        name: "Revenue",
        labels: ["Q1", "Q2", "Q3", "Q4"],
        values: [120, 190, 160, 230],
      },
    ],
    // Data for 'line' series
    [
      {
        name: "Margin %",
        labels: ["Q1", "Q2", "Q3", "Q4"],
        values: [30, 40, 35, 45],
      },
    ],
  ],
  {
    x: 48,
    y: 48,
    w: 864,
    h: 480,
    secondaryValAxis: true,
  },
);
```

### Chart types

`'area'` `'bar'` `'bar3d'` `'bubble'` `'bubble3d'` `'doughnut'` `'line'` `'pie'` `'radar'` `'scatter'`

---

## Speaker notes

```js
slide.addNotes(
  "Mention the 30% YoY growth. Pause for questions after this slide.",
);
```

---

## Slide masters

Define a master before adding slides that reference it:

```js
pres.defineSlideMaster({
  title: "MASTER_BRAND",
  background: { color: "002060" },
  objects: [
    // Logo in the top-right corner
    {
      type: "image",
      options: { path: "./logo.png", x: 816, y: 10, w: 115, h: 48 },
    },
    // Footer bar
    {
      type: "shape",
      shapeType: "rect",
      options: { x: 0, y: 509, w: 960, h: 31, fill: { color: "4472C4" } },
    },
    // Footer text
    {
      type: "text",
      text: "Confidential — Acme Corp",
      options: {
        x: 29,
        y: 514,
        w: 480,
        h: 24,
        fontSize: 10,
        color: "FFFFFF",
      },
    },
  ],
  slideNumber: { x: 864, y: 514, w: 58, color: "FFFFFF", align: "right" },
});

pres.addSlide("MASTER_BRAND", (slide) => {
  slide.addText("Branded slide", {
    x: 96,
    y: 192,
    w: 768,
    h: 192,
    fontSize: 36,
    color: "FFFFFF",
  });
});
```

---

## Exporting

### Write to file

```js
await pres.writeFile("output.pptx");
```

### Get a Buffer (Node.js)

```js
const buf = pres.write("nodebuffer"); // Buffer
fs.writeFileSync("output.pptx", buf);
```

### Get a Uint8Array

```js
const bytes = pres.write("uint8array");
```

### Get base64

```js
const b64 = pres.write("base64");
// e.g. send as HTTP response:
res.json({ file: b64 });
```

---

## Importing existing .pptx files

```js
const { Presentation } = require("pptxrs");
const fs = require("fs");

// From file
const pres = Presentation.fromBuffer(fs.readFileSync("existing.pptx"));

// From a Uint8Array (e.g. received over HTTP)
const pres2 = Presentation.fromBuffer(new Uint8Array(arrayBuffer));

console.log(pres.layout); // 'LAYOUT_16x9'
console.log(pres.title); // 'My Presentation'

const slides = pres.getSlides(); // Slide[]
console.log(slides.length);

for (const slide of slides) {
  const elements = slide.getElements();
  for (const el of elements) {
    console.log(el.elementType, el.getWidth(), el.getHeight()); // pixels
    const data = el.toJson();
    if (data.type === "text") {
      console.log(data.text, el.getX(), el.getY());
    }
    if (data.type === "image") {
      console.log("image at", el.getX(), el.getY());
    }
  }
}
```

---

## Reading element dimensions

`slide.getElements()` returns `SlideElementObject` instances. Call `getHeight()` / `getWidth()` to read dimensions in pixels, or the `*Inches()` variants for inches. Works identically whether the slide came from a file, JSON, or was just built from scratch.

```js
const { Presentation } = require("pptxrs");
const fs = require("fs");

// ── Example 1: element added to a brand-new slide ─────────────────────────────

const pres = new Presentation();
const slide = pres.addSlide();
slide.addText("Hello", { x: 96, y: 96, w: 480, h: 144, fontSize: 32 });
slide.addImage({ data: imgBase64, x: 96, y: 288, w: 192, h: 192 });
pres.syncSlide(0, slide);

const elements = pres.getSlides()[0].getElements();

const textEl = elements[0];
console.log(textEl.elementType);    // "text"
console.log(textEl.getHeight());    // 144   (pixels)
console.log(textEl.getWidth());     // 480   (pixels)
console.log(textEl.getHeightInches()); // 1.5 (inches)

const imgEl = elements[1];
console.log(imgEl.elementType);     // "image"
console.log(imgEl.getHeight());     // 192   (pixels)
console.log(imgEl.getWidth());      // 192   (pixels)

// ── Example 2: elements read from an existing .pptx ───────────────────────────

const imported = Presentation.fromBuffer(fs.readFileSync("deck.pptx"));

for (const slide of imported.getSlides()) {
  for (const el of slide.getElements()) {
    if (el.elementType === "notes") continue; // notes have no dimensions

    console.log(
      el.elementType,
      `${el.getWidth()}×${el.getHeight()} px`,
      `(${el.getWidthInches().toFixed(2)}×${el.getHeightInches().toFixed(2)} in)`,
    );
  }
}
```

---

## Modifying imported slides

After importing, get slides, mutate them, push them back with `syncSlide`:

```js
const pres = Presentation.fromBuffer(fs.readFileSync("deck.pptx"));
const slides = pres.getSlides();

// Add a watermark to every slide
slides.forEach((slide, i) => {
  slide.addText("DRAFT", {
    x: 192,
    y: 192,
    w: 576,
    h: 192,
    fontSize: 72,
    color: "FF0000",
    transparency: 70,
    rotate: 45,
    bold: true,
    align: "center",
    valign: "middle",
  });
  pres.syncSlide(i, slide);
});

await pres.writeFile("deck-draft.pptx");
```

Remove a slide:

```js
pres.removeSlide(0); // remove first slide
```

Reorder slides by rebuilding:

```js
const slides = pres.getSlides();
// swap slide 0 and slide 1
pres.syncSlide(0, slides[1]);
pres.syncSlide(1, slides[0]);
```

---

## Text measurement

Measure the exact rendered height and width of a string **before** placing it, so you can stack or position elements dynamically.

You must provide the raw font file so pptxrs can use real font metrics (via HarfBuzz shaping).

```js
const fs = require("fs");

// 1. Register the font
pres.registerFont("Calibri", fs.readFileSync("./fonts/Calibri.ttf"));

// 2. Measure
const metrics = pres.measureText("Hello, world!", {
  font: "Calibri",
  fontSize: 24, // points
  bold: false,
  italic: false,
});

// { height: 30.4, width: 148.2, lines: 1, lineHeight: 30.4 }
console.log(metrics.height); // total height in points
console.log(metrics.width); // longest line width in points
console.log(metrics.lines); // number of rendered lines
console.log(metrics.lineHeight); // per-line height in points
```

### Word-wrap: measure in a constrained box

Pass `width` in pixels to enable automatic word-wrap:

```js
const m = pres.measureText(longText, {
  font: "Calibri",
  fontSize: 18,
  width: 576, // text box width in pixels (= 6 inches × 96)
});

console.log(m.lines); // how many lines it wraps to
console.log(m.height); // total height; use to size the text box
```

### Dynamic layout: stack text boxes

```js
pres.registerFont("Calibri", fs.readFileSync("./fonts/Calibri.ttf"));

const title = "Section Title";
const body = "Here is a longer body paragraph that may wrap…";
const padding = 19; // pixels (~0.2 inch)

const titleMetrics = pres.measureText(title, {
  font: "Calibri",
  fontSize: 36,
  bold: true,
});
// measureText returns points; convert to pixels: points / 72 * 96
const titleH = (titleMetrics.height / 72) * 96;

const bodyMetrics = pres.measureText(body, {
  font: "Calibri",
  fontSize: 18,
  width: 768, // pixels
});
const bodyH = (bodyMetrics.height / 72) * 96;

pres.addSlide(null, (slide) => {
  slide.addText(title, {
    x: 96,
    y: 96,
    w: 768,
    h: titleH + padding,
    fontSize: 36,
    bold: true,
  });

  slide.addText(body, {
    x: 96,
    y: 96 + titleH + padding * 2,
    w: 768,
    h: bodyH + padding,
    fontSize: 18,
    wrap: true,
  });
});
```

---

## JSON interchange

Convert a presentation to a plain JSON object and back. The JSON is self-contained: images are stored as base64 inside the object.

### Serialize

```js
// As a JS object
const json = pres.toJson();

// As a JSON string
const jsonStr = pres.toJsonString();
fs.writeFileSync("deck.json", jsonStr);
```

### Deserialize

```js
// From a JS object
const pres2 = Presentation.fromJson(
  JSON.parse(fs.readFileSync("deck.json", "utf8")),
);
await pres2.writeFile("rebuilt.pptx");

// fromJson also accepts the live JS object directly
const pres3 = Presentation.fromJson(pres.toJson());
```

### Schema

```ts
interface PresentationJson {
  meta: {
    title?: string;
    author?: string;
    company?: string;
    layout: string; // 'LAYOUT_16x9' | 'LAYOUT_4x3' | 'LAYOUT_WIDE'
  };
  slides: {
    background?: { color?: string };
    master?: string;
    elements: SlideElement[]; // discriminated union on element.type
  }[];
}
```

Each element is a discriminated union:

```ts
type SlideElement =
  | { type: "text"; text: string | TextRun[]; options: TextOptions }
  | { type: "image"; options: ImageOptions }
  | { type: "shape"; shapeType: string; options: ShapeOptions }
  | { type: "table"; data: TableCell[][]; options: TableOptions }
  | {
      type: "chart";
      chartType: string;
      data: ChartData[];
      options: ChartOptions;
    }
  | { type: "notes"; text: string };
```

### Use cases

**Store in a database:**

```js
await db.query("INSERT INTO decks (id, data) VALUES ($1, $2)", [
  id,
  pres.toJsonString(),
]);

// Later:
const row = await db.query("SELECT data FROM decks WHERE id = $1", [id]);
const pres = Presentation.fromJson(JSON.parse(row.data));
```

**Generate from a template object:**

```js
function buildReport(data) {
  const json = {
    meta: { title: data.title, layout: "LAYOUT_16x9" },
    slides: data.sections.map((section) => ({
      elements: [
        {
          type: "text",
          text: section.heading,
          options: { x: 96, y: 96, w: 768, h: 96, fontSize: 32, bold: true },
        },
        {
          type: "text",
          text: section.body,
          options: { x: 96, y: 240, w: 768, h: 288, fontSize: 16 },
        },
      ],
    })),
  };
  return Presentation.fromJson(json);
}
```

---

## Full API reference

### `new Presentation(options?)`

| Option    | Type                                                 | Default         |
| --------- | ---------------------------------------------------- | --------------- |
| `layout`  | `'LAYOUT_16x9'` \| `'LAYOUT_4x3'` \| `'LAYOUT_WIDE'` | `'LAYOUT_16x9'` |
| `title`   | `string`                                             | —               |
| `author`  | `string`                                             | —               |
| `company` | `string`                                             | —               |

### `Presentation` static methods

| Method                         | Description                                           |
| ------------------------------ | ----------------------------------------------------- |
| `Presentation.fromBuffer(buf)` | Import a `.pptx` file from a `Buffer` or `Uint8Array` |
| `Presentation.fromJson(json)`  | Reconstruct from a `PresentationJson` object          |

### `Presentation` instance methods

| Method                       | Returns                              | Description                                             |
| ---------------------------- | ------------------------------------ | ------------------------------------------------------- |
| `registerFont(name, buf)`    | `this`                               | Register a TTF/OTF font for `measureText`               |
| `measureText(text, opts)`    | `TextMetrics`                        | Measure text dimensions in points                       |
| `defineSlideMaster(opts)`    | `this`                               | Define a named slide master                             |
| `addSlide(masterName?, fn?)` | `Slide`                              | Add a slide; optional callback auto-syncs               |
| `getSlides()`                | `Slide[]`                            | Get all slides                                          |
| `syncSlide(index, slide)`    | `this`                               | Push a modified slide back                              |
| `removeSlide(index)`         | `this`                               | Remove slide at index                                   |
| `write(outputType?)`         | `Buffer` \| `Uint8Array` \| `string` | Export (`'nodebuffer'` \| `'uint8array'` \| `'base64'`) |
| `writeFile(path)`            | `Promise<void>`                      | Write to disk (Node.js)                                 |
| `toJson()`                   | `PresentationJson`                   | Serialize to JS object                                  |
| `toJsonString()`             | `string`                             | Serialize to JSON string                                |

### `Slide` instance methods

| Method                              | Returns          | Description                                 |
| ----------------------------------- | ---------------- | ------------------------------------------- |
| `addText(text, opts)`               | `this`           | Add text or `TextRun[]`                     |
| `addImage(opts)`                    | `this`           | Add image (`path` auto-resolved in Node.js) |
| `addShape(type, opts)`              | `this`           | Add a preset shape                          |
| `addTable(data, opts?)`             | `this`           | Add a table                                 |
| `addChart(type, data, opts?)`       | `this`           | Add a chart                                 |
| `addComboChart(types, data, opts?)` | `this`           | Add a combo chart                           |
| `addNotes(text)`                    | `this`           | Set speaker notes                           |
| `setBackground(hexColor)`           | `this`           | Set background color                        |
| `getElements()`                     | `SlideElementObject[]` | Get all elements with dimension accessors   |

### `SlideElementObject` methods

| Method / Property  | Returns  | Description                                       |
| ------------------ | -------- | ------------------------------------------------- |
| `elementType`      | `string` | `"text"` \| `"image"` \| `"shape"` \| `"table"` \| `"chart"` \| `"notes"` |
| `getWidth()`       | `number` | Width in pixels (96 DPI)                          |
| `getHeight()`      | `number` | Height in pixels (96 DPI)                         |
| `getX()`           | `number` | X position in pixels (96 DPI)                     |
| `getY()`           | `number` | Y position in pixels (96 DPI)                     |
| `getWidthInches()` | `number` | Width in inches                                   |
| `getHeightInches()`| `number` | Height in inches                                  |
| `getXInches()`     | `number` | X position in inches                              |
| `getYInches()`     | `number` | Y position in inches                              |
| `toJson()`         | `SlideElement` | Full element data/options as a plain object  |

### `MeasureOptions`

| Field                 | Type      | Description                             |
| --------------------- | --------- | --------------------------------------- |
| `font`                | `string`  | Font name (must be registered)          |
| `fontSize`            | `number`  | Points                                  |
| `bold`                | `boolean` |                                         |
| `italic`              | `boolean` |                                         |
| `charSpacing`         | `number`  | Extra char spacing in points            |
| `lineSpacingMultiple` | `number`  | Multiplier, default `1.0`               |
| `width`               | `number`  | Box width in **pixels**; enables word-wrap |

### `TextMetrics`

| Field        | Type     | Description                       |
| ------------ | -------- | --------------------------------- |
| `height`     | `number` | Total text block height in points |
| `width`      | `number` | Longest line width in points      |
| `lines`      | `number` | Number of rendered lines          |
| `lineHeight` | `number` | Per-line height in points         |

---

## Building from source

```bash
# Prerequisites
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
cargo install wasm-pack

# Build
wasm-pack build --target nodejs --out-dir pkg

# Publish
cd pkg && npm publish --access public
```

---

## License

MIT
