"use strict";
/**
 * 05-text-measurement.js
 *
 * Demonstrates font-aware text measurement before layout:
 *
 *   A. Basic measurement — get height and width of a string
 *   B. Word-wrap measurement — how many lines and total height in a fixed-width box
 *   C. Dynamic layout — stack multiple text blocks without hardcoded positions
 *   D. Auto-size a text box to exactly fit its content
 *   E. Multi-column layout driven by measurements
 *
 * NOTE: measureText requires a real TTF/OTF font file.
 *       The example looks for fonts in common system locations.
 *       If none is found it prints instructions and exits gracefully.
 *
 * Run:  node examples/05-text-measurement.js
 * Out:  examples/out/05-text-measurement.pptx
 */

const { Presentation } = require("../pkg/index.js");
const fs = require("fs");
const path = require("path");

const OUT = path.join(__dirname, "out");
fs.mkdirSync(OUT, { recursive: true });

// ── Locate a system font ──────────────────────────────────────────────────────
function findFont(names) {
  const candidates = [
    // macOS
    "/Library/Fonts",
    "/System/Library/Fonts",
    `${process.env.HOME}/Library/Fonts`,
    // Windows
    "C:\\Windows\\Fonts",
    // Linux
    "/usr/share/fonts/truetype",
    "/usr/share/fonts/opentype",
    "/usr/local/share/fonts",
  ];

  for (const dir of candidates) {
    if (!fs.existsSync(dir)) continue;
    for (const name of names) {
      const f = path.join(dir, name);
      if (fs.existsSync(f)) return f;
      // Recurse one level
      try {
        for (const sub of fs.readdirSync(dir)) {
          const ff = path.join(dir, sub, name);
          if (fs.existsSync(ff)) return ff;
        }
      } catch {
        /* skip */
      }
    }
  }
  return null;
}

const FONT_PATH = findFont([
  "Arial.ttf",
  "arial.ttf",
  "Helvetica.ttc",
  "helvetica.ttf",
  "DejaVuSans.ttf",
  "LiberationSans-Regular.ttf",
  "NotoSans-Regular.ttf",
]);

if (!FONT_PATH) {
  console.error(`
No suitable font found.
Place a TTF/OTF font file accessible from one of these paths and re-run:
  /Library/Fonts/Arial.ttf  (macOS)
  C:\\Windows\\Fonts\\arial.ttf  (Windows)
  /usr/share/fonts/truetype/dejavu/DejaVuSans.ttf  (Linux)

Or pass the path directly by editing FONT_PATH in this file.
`);
  process.exit(0);
}

const FONT_NAME = path.basename(FONT_PATH, path.extname(FONT_PATH));
console.log(`Using font: ${FONT_PATH}`);

// Points → inches helper
const pts2in = (pts) => pts / 72;

async function main() {
  const pres = new Presentation({ title: "Text Measurement Demo" });
  pres.registerFont(FONT_NAME, fs.readFileSync(FONT_PATH));

  // ── A. Basic measurement ────────────────────────────────────────────────────
  {
    const m = pres.measureText("Hello, pptxrs!", {
      font: FONT_NAME,
      fontSize: 36,
    });
    console.log("\n── A. Basic measurement ────────────────────────────────");
    console.log(`  text    : "Hello, pptxrs!"`);
    console.log(`  fontSize: 36 pt`);
    console.log(
      `  height  : ${m.height.toFixed(1)} pt  (${pts2in(m.height).toFixed(3)} in)`,
    );
    console.log(
      `  width   : ${m.width.toFixed(1)} pt  (${pts2in(m.width).toFixed(3)} in)`,
    );
    console.log(`  lines   : ${m.lines}`);
    console.log(`  lineH   : ${m.lineHeight.toFixed(1)} pt`);
  }

  // ── B. Word-wrap measurement ─────────────────────────────────────────────────
  {
    const longText =
      "This is a longer sentence that will wrap when constrained to a narrow box width.";
    const boxWidthIn = 4;

    const mWrap = pres.measureText(longText, {
      font: FONT_NAME,
      fontSize: 16,
      width: boxWidthIn, // enable word-wrap at 4 inches
    });

    const mNoWrap = pres.measureText(longText, {
      font: FONT_NAME,
      fontSize: 16,
      // no width → single line
    });

    console.log("\n── B. Word-wrap measurement ────────────────────────────");
    console.log(`  text      : "${longText.slice(0, 50)}…"`);
    console.log(`  box width : ${boxWidthIn} in`);
    console.log(
      `  no wrap   : ${mNoWrap.lines} line,  width ${mNoWrap.width.toFixed(0)} pt`,
    );
    console.log(
      `  wrapped   : ${mWrap.lines} lines, height ${mWrap.height.toFixed(0)} pt`,
    );
  }

  // ── C. Dynamic layout — stack text blocks ─────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Dynamic Layout", {
      x: 0.5,
      y: 0.1,
      w: 9,
      h: 0.7,
      fontSize: 28,
      bold: true,
      color: "002060",
    });

    const blocks = [
      { text: "Section Heading", fontSize: 24, bold: true, color: "002060" },
      {
        text: "A subtitle line goes here.",
        fontSize: 16,
        bold: false,
        color: "666666",
      },
      {
        text: "Body paragraph: This text wraps automatically based on measured height, so the next element always starts exactly below it regardless of how long this content is.",
        fontSize: 14,
        bold: false,
        color: "333333",
      },
      { text: "Another heading", fontSize: 20, bold: true, color: "4472C4" },
      {
        text: "Final line of content.",
        fontSize: 14,
        bold: false,
        color: "555555",
      },
    ];

    const BOX_W = 9; // inches — width of all text boxes
    const LEFT = 0.5; // inches — left margin
    const TOP = 0.9; // inches — starting Y
    const GAP = 0.12; // inches — gap between blocks
    const PADV = 0.08; // inches — extra vertical padding per box

    let currentY = TOP;

    blocks.forEach(({ text, fontSize, bold, color }) => {
      const m = pres.measureText(text, {
        font: FONT_NAME,
        fontSize,
        bold,
        width: BOX_W, // wrap within the box
      });

      const boxH = pts2in(m.height) + PADV;

      slide.addText(text, {
        x: LEFT,
        y: currentY,
        w: BOX_W,
        h: boxH,
        fontSize,
        bold,
        color,
        wrap: true,
      });

      currentY += boxH + GAP;
    });

    console.log(`\n── C. Dynamic layout ───────────────────────────────────`);
    console.log(
      `  Stacked ${blocks.length} blocks. Final Y: ${currentY.toFixed(3)} in`,
    );
  });

  // ── D. Auto-size a text box ───────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Auto-sized Text Boxes", {
      x: 0.5,
      y: 0.1,
      w: 9,
      h: 0.7,
      fontSize: 28,
      bold: true,
      color: "002060",
    });

    const samples = [
      "Short.",
      "A medium-length sentence that takes more space.",
      "A much longer piece of text that will definitely require the box to be taller than a single line when constrained to a fixed width column.",
    ];

    samples.forEach((text, i) => {
      const BOX_W = 2.7;
      const X = 0.4 + i * 3.1;

      const m = pres.measureText(text, {
        font: FONT_NAME,
        fontSize: 14,
        width: BOX_W,
      });
      const boxH = pts2in(m.height) + 0.1;

      // Background rect sized to content
      slide.addShape("rect", {
        x: X - 0.05,
        y: 0.9 - 0.05,
        w: BOX_W + 0.1,
        h: boxH + 0.1,
        fill: { color: "EEF2FF" },
        line: { color: "4472C4", width: 1 },
      });

      slide.addText(text, {
        x: X,
        y: 0.9,
        w: BOX_W,
        h: boxH,
        fontSize: 14,
        wrap: true,
        color: "222222",
      });

      // Show measurement annotation
      slide.addText(
        `${m.lines} line${m.lines !== 1 ? "s" : ""} · ${pts2in(m.height).toFixed(2)}"`,
        {
          x: X,
          y: 0.9 + boxH + 0.05,
          w: BOX_W,
          h: 0.3,
          fontSize: 9,
          color: "888888",
          align: "center",
        },
      );
    });
  });

  // ── E. Multi-column layout ────────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Multi-column Measured Layout", {
      x: 0.5,
      y: 0.1,
      w: 9,
      h: 0.6,
      fontSize: 24,
      bold: true,
      color: "002060",
    });

    const COL_W = 4.1; // column width in inches
    const COL_GAP = 0.3;
    const TOP_Y = 0.85;
    const MAX_H = 4.5; // max column height before overflow

    const items = [
      {
        heading: "Performance",
        body: "Rust compiles to WebAssembly, running at near-native speed in Node.js with no native dependencies.",
      },
      {
        heading: "Compatibility",
        body: "Works in Node.js 16+. The output .pptx opens in PowerPoint, Google Slides, LibreOffice, and Keynote.",
      },
      {
        heading: "API Design",
        body: "Mirrors PptxGenJS so migration is straightforward. All coordinates are in inches.",
      },
      {
        heading: "Text Metrics",
        body: "Register any TTF/OTF font and get pixel-accurate height and width before export.",
      },
      {
        heading: "JSON Export",
        body: "Serialize a full presentation to JSON for storage, versioning, or template generation.",
      },
      {
        heading: "Import",
        body: "Open any .pptx file, inspect its elements, modify them, and re-export.",
      },
    ];

    let colIndex = 0;
    let colY = TOP_Y;

    const ITEM_GAP = 0.15;
    const HEAD_SZ = 16;
    const BODY_SZ = 12;

    items.forEach(({ heading, body }) => {
      const mHead = pres.measureText(heading, {
        font: FONT_NAME,
        fontSize: HEAD_SZ,
        bold: true,
        width: COL_W,
      });
      const mBody = pres.measureText(body, {
        font: FONT_NAME,
        fontSize: BODY_SZ,
        width: COL_W,
      });

      const headH = pts2in(mHead.height) + 0.05;
      const bodyH = pts2in(mBody.height) + 0.05;
      const totalH = headH + bodyH + ITEM_GAP;

      // Overflow to next column
      if (colY + totalH > TOP_Y + MAX_H && colIndex < 1) {
        colIndex++;
        colY = TOP_Y;
      }

      const X = 0.4 + colIndex * (COL_W + COL_GAP);

      slide.addText(heading, {
        x: X,
        y: colY,
        w: COL_W,
        h: headH,
        fontSize: HEAD_SZ,
        bold: true,
        color: "4472C4",
      });

      slide.addText(body, {
        x: X,
        y: colY + headH,
        w: COL_W,
        h: bodyH,
        fontSize: BODY_SZ,
        color: "444444",
        wrap: true,
      });

      colY += totalH + ITEM_GAP;
    });

    console.log("\n── E. Multi-column layout ──────────────────────────────");
    console.log(`  Distributed ${items.length} items across columns.`);
  });

  const out = path.join(OUT, "05-text-measurement.pptx");
  await pres.writeFile(out);
  console.log(`\nWritten: ${out}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
