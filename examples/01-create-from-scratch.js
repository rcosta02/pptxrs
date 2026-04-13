"use strict";
/**
 * 01-create-from-scratch.js
 *
 * Demonstrates creating a presentation from nothing:
 * text, images, shapes, tables, charts, speaker notes,
 * and all common styling options.
 *
 * Run:  node examples/01-create-from-scratch.js
 * Out:  examples/out/01-create-from-scratch.pptx
 */

const { Presentation } = require("../pkg/index.js");
const fs = require("fs");
const path = require("path");

const OUT = path.join(__dirname, "out");
fs.mkdirSync(OUT, { recursive: true });

async function main() {
  const pres = new Presentation({
    layout: "LAYOUT_16x9",
    title: "pptxrs demo",
    author: "Jane Smith",
    company: "Acme Corp",
  });

  // ── Slide 1 — Title slide ───────────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.setBackground("002060");

    slide.addText("pptxrs", {
      x: 1,
      y: 1.5,
      w: 8,
      h: 1.5,
      fontSize: 60,
      bold: true,
      color: "FFFFFF",
      align: "center",
    });

    slide.addText("Create · Read · Modify · Export .pptx files", {
      x: 1,
      y: 3.2,
      w: 8,
      h: 0.8,
      fontSize: 20,
      color: "C0D0F0",
      align: "center",
    });

    slide.addText("Powered by Rust + WebAssembly", {
      x: 1,
      y: 4.5,
      w: 8,
      h: 0.5,
      fontSize: 13,
      color: "8090B0",
      align: "center",
      italic: true,
    });

    slide.addNotes("Welcome slide. Mention the Rust/WASM architecture.");
  });

  // ── Slide 2 — Text styles ───────────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Text Styles", {
      x: 0.5,
      y: 0.2,
      w: 9,
      h: 0.7,
      fontSize: 28,
      bold: true,
      color: "002060",
    });

    // Plain text with every style option exercised
    slide.addText("Bold  Italic  Underline  Strike  Highlight", {
      x: 0.5,
      y: 1.1,
      w: 9,
      h: 0.6,
      fontSize: 18,
      bold: true,
      italic: true,
      underline: true,
      strike: "sngStrike",
      highlight: "FFFF00",
      color: "333333",
    });

    // Mixed-format runs
    slide.addText(
      [
        { text: "Normal, " },
        { text: "bold red, ", options: { bold: true, color: "C00000" } },
        { text: "italic blue, ", options: { italic: true, color: "0070C0" } },
        { text: "big orange", options: { fontSize: 28, color: "ED7D31" } },
        { text: " and superscript", options: { fontSize: 12 } },
        { text: "2", options: { superscript: true, fontSize: 10 } },
      ],
      { x: 0.5, y: 1.9, w: 9, h: 0.7, fontSize: 18 },
    );

    // Hyperlink
    slide.addText("Click here → pptxrs on GitHub", {
      x: 0.5,
      y: 2.8,
      w: 5,
      h: 0.5,
      fontSize: 16,
      color: "0563C1",
      underline: true,
      hyperlink: { url: "https://github.com", tooltip: "Open GitHub" },
    });

    // Bullet list
    slide.addText(
      [
        { text: "First bullet point", options: { bullet: true } },
        { text: "Second bullet point", options: { bullet: true } },
        {
          text: "Numbered item one",
          options: { bullet: { type: "number", style: "arabicPeriod" } },
        },
        {
          text: "Numbered item two",
          options: { bullet: { type: "number", style: "arabicPeriod" } },
        },
        {
          text: "Custom checkmark",
          options: {
            bullet: { code: "2713", color: "70AD47", font: "Wingdings 2" },
          },
        },
      ],
      { x: 0.5, y: 3.4, w: 4.5, h: 2, fontSize: 15 },
    );

    // Shadow + glow
    slide.addText("Shadow & Glow Effects", {
      x: 5,
      y: 3.4,
      w: 4.5,
      h: 1.2,
      fontSize: 22,
      bold: true,
      color: "4472C4",
      shadow: {
        type: "outer",
        angle: 45,
        blur: 6,
        color: "000000",
        opacity: 0.4,
      },
      glow: { size: 10, opacity: 0.3, color: "4472C4" },
    });

    slide.addNotes(
      "All major text formatting options are shown on this slide.",
    );
  });

  // ── Slide 3 — Shapes ────────────────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Shapes", {
      x: 0.5,
      y: 0.2,
      w: 9,
      h: 0.7,
      fontSize: 28,
      bold: true,
      color: "002060",
    });

    const shapes = [
      { type: "rect", x: 0.3, fill: "4472C4" },
      { type: "roundRect", x: 2.1, fill: "ED7D31", rectRadius: 0.15 },
      { type: "ellipse", x: 3.9, fill: "A9D18E" },
      { type: "triangle", x: 5.7, fill: "FFC000" },
      { type: "diamond", x: 7.5, fill: "FF0000" },
    ];

    shapes.forEach(({ type, x, fill, rectRadius }) => {
      slide.addShape(type, {
        x,
        y: 0.9,
        w: 1.6,
        h: 1.4,
        fill: { color: fill },
        line: { color: "000000", width: 1 },
        ...(rectRadius ? { rectRadius } : {}),
      });
    });

    // Arrow shapes
    ["rightArrow", "leftArrow", "upArrow", "downArrow"].forEach((type, i) => {
      slide.addShape(type, {
        x: 0.3 + i * 2.3,
        y: 2.6,
        w: 2,
        h: 1,
        fill: { color: "5B9BD5" },
        line: { color: "2E75B6", width: 1 },
      });
    });

    // Shape with text inside
    slide.addShape("roundRect", {
      x: 0.5,
      y: 4,
      w: 4,
      h: 1.4,
      fill: { color: "002060" },
      line: { color: "001040", width: 2 },
      text: "Shape with text",
      fontSize: 18,
      bold: true,
      color: "FFFFFF",
      align: "center",
      valign: "middle",
    });

    // Shadow on a shape
    slide.addShape("ellipse", {
      x: 5,
      y: 4,
      w: 4,
      h: 1.4,
      fill: { color: "ED7D31", transparency: 10 },
      shadow: {
        type: "outer",
        angle: 45,
        blur: 8,
        color: "000000",
        opacity: 0.5,
      },
    });
  });

  // ── Slide 4 — Table ─────────────────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Tables", {
      x: 0.5,
      y: 0.1,
      w: 9,
      h: 0.6,
      fontSize: 28,
      bold: true,
      color: "002060",
    });

    // Styled header row + data rows
    slide.addTable(
      [
        // Header row — per-cell options
        [
          {
            text: "Employee",
            options: {
              bold: true,
              fill: "002060",
              color: "FFFFFF",
              align: "center",
            },
          },
          {
            text: "Department",
            options: {
              bold: true,
              fill: "002060",
              color: "FFFFFF",
              align: "center",
            },
          },
          {
            text: "Score",
            options: {
              bold: true,
              fill: "002060",
              color: "FFFFFF",
              align: "center",
            },
          },
          {
            text: "Status",
            options: {
              bold: true,
              fill: "002060",
              color: "FFFFFF",
              align: "center",
            },
          },
        ],
        [
          "Alice Zhang",
          "Engineering",
          "97",
          { text: "Exceeds", options: { color: "00B050", bold: true } },
        ],
        [
          "Bob Marley",
          "Design",
          "85",
          { text: "Meets", options: { color: "0070C0" } },
        ],
        [
          "Carol White",
          "PM",
          "91",
          { text: "Exceeds", options: { color: "00B050", bold: true } },
        ],
        [
          "David Kim",
          "Engineering",
          "72",
          { text: "Below", options: { color: "FF0000" } },
        ],
        [
          "Eva Rossi",
          "Marketing",
          "88",
          { text: "Meets", options: { color: "0070C0" } },
        ],
      ],
      {
        x: 0.5,
        y: 0.85,
        w: 9,
        h: 3.8,
        colW: [2.5, 2.5, 1.5, 2.5],
        fontSize: 14,
        border: { pt: 1, color: "D0D0D0" },
        valign: "middle",
      },
    );

    slide.addNotes("Performance review table with conditional color coding.");
  });

  // ── Slide 5 — Charts ────────────────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Charts", {
      x: 0.5,
      y: 0.1,
      w: 9,
      h: 0.5,
      fontSize: 24,
      bold: true,
      color: "002060",
    });

    // Bar chart (left half)
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
        x: 0.2,
        y: 0.7,
        w: 4.6,
        h: 4.7,
        barDir: "col",
        barGrouping: "clustered",
        showTitle: true,
        title: "Quarterly Financials",
        showLegend: true,
        legendPos: "b",
        showValue: true,
        chartColors: ["4472C4", "ED7D31"],
        valAxisTitle: "USD (k)",
      },
    );

    // Pie chart (right half)
    slide.addChart(
      "pie",
      [
        {
          labels: ["North", "South", "East", "West"],
          values: [35, 25, 20, 20],
        },
      ],
      {
        x: 5.1,
        y: 0.7,
        w: 4.6,
        h: 4.7,
        showTitle: true,
        title: "Market Share",
        showPercent: true,
        showLegend: true,
        legendPos: "b",
        chartColors: ["4472C4", "ED7D31", "A9D18E", "FFC000"],
      },
    );
  });

  // ── Slide 6 — Line & scatter ────────────────────────────────────────────────
  pres.addSlide(null, (slide) => {
    slide.addText("Line & Scatter Charts", {
      x: 0.5,
      y: 0.1,
      w: 9,
      h: 0.5,
      fontSize: 24,
      bold: true,
      color: "002060",
    });

    slide.addChart(
      "line",
      [
        {
          name: "Product A",
          labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
          values: [42, 58, 51, 73, 68, 89],
        },
        {
          name: "Product B",
          labels: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
          values: [30, 35, 45, 40, 60, 55],
        },
      ],
      {
        x: 0.2,
        y: 0.7,
        w: 4.6,
        h: 4.7,
        lineSmooth: true,
        lineSize: 2.5,
        lineDataSymbol: "circle",
        lineDataSymbolSize: 7,
        showTitle: true,
        title: "Monthly Sales",
        showLegend: true,
        legendPos: "b",
        chartColors: ["4472C4", "ED7D31"],
        valAxisTitle: "Units",
      },
    );

    slide.addChart(
      "scatter",
      [
        {
          name: "Dataset 1",
          labels: ["A", "B", "C", "D", "E"],
          values: [10, 30, 20, 50, 40],
        },
      ],
      {
        x: 5.1,
        y: 0.7,
        w: 4.6,
        h: 4.7,
        showTitle: true,
        title: "Scatter Plot",
        showLegend: false,
        chartColors: ["4472C4"],
      },
    );
  });

  // ── Write ───────────────────────────────────────────────────────────────────
  const outPath = path.join(OUT, "01-create-from-scratch.pptx");
  await pres.writeFile(outPath);
  console.log(`Written: ${outPath}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
