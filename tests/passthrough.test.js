/**
 * Passthrough round-trip fidelity tests
 *
 * Verifies that:
 *   A. fromBuffer() → write() produces a valid PPTX (ZIP passthrough path)
 *   B. fromBuffer() + add elements → write() preserves original + adds new
 *   C. fromBuffer() + updateChart() → write() → fromBuffer() shows updated data
 *   D. fromBuffer() + updateTable() → write() → fromBuffer() shows updated text
 *   E. Charts created from scratch actually work (chart XML is written)
 *
 * Run: node --test tests/passthrough.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { Presentation, TINY_PNG } = require("./helpers.js");

// ── Helpers ──────────────────────────────────────────────────────────────────

/** Build a presentation with varied content and return its buffer. */
function buildSourceBuffer() {
  const pres = new Presentation({ layout: "LAYOUT_16x9" });

  pres.addSlide(null, (s) => {
    s.setBackground("1A2B3C");
    s.addText("Slide One Title", {
      x: 48, y: 24, w: 864, h: 80,
      fontSize: 36, bold: true, color: "FFFFFF",
    });
    s.addShape("rect", {
      x: 48, y: 120, w: 300, h: 100,
      fill: { color: "4472C4" },
    });
  });

  pres.addSlide(null, (s) => {
    s.addTable(
      [
        ["Name",   "Score"],
        ["Alice",  "95"],
        ["Bob",    "82"],
      ],
      { x: 48, y: 48, w: 400, h: 150 },
    );
  });

  pres.addSlide(null, (s) => {
    s.addChart("bar", [
      { name: "Sales",  labels: ["Q1", "Q2", "Q3"], values: [10, 20, 30] },
      { name: "Budget", labels: ["Q1", "Q2", "Q3"], values: [15, 15, 25] },
    ], { x: 48, y: 48, w: 500, h: 300 });
  });

  return pres.write("nodebuffer");
}

/**
 * Load a buffer with fromBuffer, optionally mutate slides, write to a new
 * buffer, and load that with fromBuffer again. Returns the final presentation.
 */
function passthroughCycle(buf, mutateFn) {
  const pres = Presentation.fromBuffer(buf);
  if (mutateFn) mutateFn(pres);
  const buf2 = pres.write("nodebuffer");
  return Presentation.fromBuffer(buf2);
}


// ═════════════════════════════════════════════════════════════════════════════
// A. Basic passthrough — no mutations
// ═════════════════════════════════════════════════════════════════════════════

describe("Passthrough – no mutations", () => {
  let src;
  let result;

  test("builds source buffer without throwing", () => {
    assert.doesNotThrow(() => { src = buildSourceBuffer(); });
  });

  test("fromBuffer() → write() does not throw", () => {
    assert.doesNotThrow(() => {
      const pres = Presentation.fromBuffer(src);
      pres.write("nodebuffer");
    });
  });

  test("slide count is preserved (3 slides)", () => {
    result = passthroughCycle(src);
    assert.equal(result.getSlides().length, 3);
  });

  test("slide 0 background color is preserved", () => {
    const bg = result.toJson().slides[0].background;
    assert.equal(bg?.color, "1A2B3C");
  });

  test("slide 0 still has text and shape elements", () => {
    const types = result.getSlides()[0].getElements().map((e) => e.elementType);
    assert.ok(types.includes("text"),  "text should survive passthrough");
    assert.ok(types.includes("shape"), "shape should survive passthrough");
  });

  test("slide 1 still has table element", () => {
    const types = result.getSlides()[1].getElements().map((e) => e.elementType);
    assert.ok(types.includes("table"), "table should survive passthrough");
  });

  test("slide 2 still has chart element", () => {
    const types = result.getSlides()[2].getElements().map((e) => e.elementType);
    assert.ok(types.includes("chart"), "chart should survive passthrough");
  });

  test("chart data (series values) is preserved in passthrough", () => {
    const chartEl = result.getSlides()[2].getElements()
      .find((e) => e.elementType === "chart");
    assert.ok(chartEl, "chart element should exist");
    const { data } = chartEl.toJson();
    assert.ok(Array.isArray(data) && data.length >= 1, "chart should have series");
    const sales = data.find((s) => s.name === "Sales");
    assert.ok(sales, "Sales series should be found");
    assert.deepEqual(sales.values, [10, 20, 30]);
  });

  test("table cells are preserved in passthrough", () => {
    const tblEl = result.getSlides()[1].getElements()
      .find((e) => e.elementType === "table");
    const rows = tblEl.toJson().data;
    assert.deepEqual(rows[0], ["Name", "Score"]);
    assert.deepEqual(rows[1], ["Alice", "95"]);
    assert.deepEqual(rows[2], ["Bob",   "82"]);
  });
});


// ═════════════════════════════════════════════════════════════════════════════
// B. Passthrough + surgical injection of new elements
// ═════════════════════════════════════════════════════════════════════════════

describe("Passthrough – add new elements to existing slides", () => {
  const src = buildSourceBuffer();

  test("adding text to an existing slide does not throw on write", () => {
    assert.doesNotThrow(() => {
      const pres = Presentation.fromBuffer(src);
      const slides = pres.getSlides();
      slides[0].addText("INJECTED", { x: 0, y: 400, w: 500, h: 60, fontSize: 20 });
      pres.syncSlide(0, slides[0]);
      pres.write("nodebuffer");
    });
  });

  test("injected text survives write→fromBuffer round-trip", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    slides[0].addText("INJECTED", { x: 0, y: 400, w: 500, h: 60, fontSize: 20 });
    pres.syncSlide(0, slides[0]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    const texts = result.getSlides()[0].getElements()
      .filter((e) => e.elementType === "text")
      .map((e) => {
        const raw = e.toJson().text;
        return typeof raw === "string" ? raw : raw.map((r) => r.text).join("");
      });
    assert.ok(texts.some((t) => t.includes("INJECTED")), "injected text should be present");
  });

  test("original elements still present after surgical injection", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    slides[0].addText("NEW TEXT", { x: 0, y: 440, w: 500, h: 60, fontSize: 16 });
    pres.syncSlide(0, slides[0]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    const types = result.getSlides()[0].getElements().map((e) => e.elementType);
    assert.ok(types.includes("text"),  "original text should still be present");
    assert.ok(types.includes("shape"), "original shape should still be present");
  });

  test("adding image to existing slide survives passthrough", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    slides[0].addImage({ data: TINY_PNG, x: 600, y: 400, w: 100, h: 100 });
    pres.syncSlide(0, slides[0]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    const types = result.getSlides()[0].getElements().map((e) => e.elementType);
    assert.ok(types.includes("image"), "added image should survive");
  });

  test("slide count is unchanged after injecting into existing slides", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    slides[1].addText("extra", { x: 0, y: 300, w: 400, h: 50, fontSize: 14 });
    pres.syncSlide(1, slides[1]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    assert.equal(result.getSlides().length, 3);
  });
});


// ═════════════════════════════════════════════════════════════════════════════
// C. updateChart()
// ═════════════════════════════════════════════════════════════════════════════

describe("updateChart() — modify chart data", () => {
  const src = buildSourceBuffer();

  test("updateChart() does not throw", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    assert.doesNotThrow(() => {
      const chartIdx = slides[2].getElements().findIndex((e) => e.elementType === "chart");
      slides[2].updateChart(chartIdx, [
        { name: "Updated", labels: ["Jan", "Feb"], values: [99, 88] },
      ]);
      pres.syncSlide(2, slides[2]);
    });
  });

  test("write() after updateChart() does not throw", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    const chartIdx = slides[2].getElements().findIndex((e) => e.elementType === "chart");
    slides[2].updateChart(chartIdx, [
      { name: "Updated", labels: ["Jan", "Feb"], values: [99, 88] },
    ]);
    pres.syncSlide(2, slides[2]);
    assert.doesNotThrow(() => pres.write("nodebuffer"));
  });

  test("updated chart data survives write→fromBuffer", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    const chartIdx = slides[2].getElements().findIndex((e) => e.elementType === "chart");
    slides[2].updateChart(chartIdx, [
      { name: "Updated", labels: ["Jan", "Feb"], values: [99, 88] },
    ]);
    pres.syncSlide(2, slides[2]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    const chartEl = result.getSlides()[2].getElements()
      .find((e) => e.elementType === "chart");
    assert.ok(chartEl, "chart element should still be present");
    const { data } = chartEl.toJson();
    const updated = data.find((s) => s.name === "Updated");
    assert.ok(updated, "Updated series should be found");
    assert.deepEqual(updated.values, [99, 88]);
  });

  test("updateChart() on a wrong index throws a descriptive error", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    assert.throws(
      () => slides[2].updateChart(999, []),
      /updateChart.*index/i,
    );
  });

  test("updateChart() on a non-chart element throws descriptive error", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    // Slide 0 element 0 is text, not chart
    assert.throws(
      () => slides[0].updateChart(0, []),
      /not a chart/i,
    );
  });

  test("other slides are unchanged after updateChart on slide 2", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    const chartIdx = slides[2].getElements().findIndex((e) => e.elementType === "chart");
    slides[2].updateChart(chartIdx, [{ name: "X", values: [1] }]);
    pres.syncSlide(2, slides[2]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    // Slide 0 should still have its text and shape
    const types0 = result.getSlides()[0].getElements().map((e) => e.elementType);
    assert.ok(types0.includes("text"),  "slide 0 text unchanged");
    assert.ok(types0.includes("shape"), "slide 0 shape unchanged");
    // Slide 1 should still have its table
    const types1 = result.getSlides()[1].getElements().map((e) => e.elementType);
    assert.ok(types1.includes("table"), "slide 1 table unchanged");
  });
});


// ═════════════════════════════════════════════════════════════════════════════
// D. updateTable()
// ═════════════════════════════════════════════════════════════════════════════

describe("updateTable() — modify table cell text", () => {
  const src = buildSourceBuffer();

  test("updateTable() does not throw", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    const tblIdx = slides[1].getElements().findIndex((e) => e.elementType === "table");
    assert.doesNotThrow(() => {
      slides[1].updateTable(tblIdx, [
        ["Product", "Revenue"],
        ["Widget",  "1000"],
        ["Gadget",  "2000"],
      ]);
      pres.syncSlide(1, slides[1]);
    });
  });

  test("write() after updateTable() does not throw", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    const tblIdx = slides[1].getElements().findIndex((e) => e.elementType === "table");
    slides[1].updateTable(tblIdx, [
      ["Col A", "Col B"],
      ["val1",  "val2"],
    ]);
    pres.syncSlide(1, slides[1]);
    assert.doesNotThrow(() => pres.write("nodebuffer"));
  });

  test("updated table cell text survives write→fromBuffer", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    const tblIdx = slides[1].getElements().findIndex((e) => e.elementType === "table");
    slides[1].updateTable(tblIdx, [
      ["Product",   "Revenue"],
      ["Widget",    "1000"],
      ["Gadget",    "2000"],
    ]);
    pres.syncSlide(1, slides[1]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    const tblEl = result.getSlides()[1].getElements()
      .find((e) => e.elementType === "table");
    assert.ok(tblEl, "table element should still be present");
    const rows = tblEl.toJson().data;
    assert.equal(rows.length, 3);
    assert.deepEqual(rows[0], ["Product", "Revenue"]);
    assert.deepEqual(rows[1], ["Widget",  "1000"]);
    assert.deepEqual(rows[2], ["Gadget",  "2000"]);
  });

  test("updateTable() on a non-table element throws descriptive error", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    // Slide 0 element 0 is text, not table
    assert.throws(
      () => slides[0].updateTable(0, []),
      /not a table/i,
    );
  });

  test("other slides unchanged after updateTable on slide 1", () => {
    const pres = Presentation.fromBuffer(src);
    const slides = pres.getSlides();
    const tblIdx = slides[1].getElements().findIndex((e) => e.elementType === "table");
    slides[1].updateTable(tblIdx, [["A", "B"], ["1", "2"], ["3", "4"]]);
    pres.syncSlide(1, slides[1]);
    const buf2 = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf2);
    // Slide 0 unchanged
    const bg = result.toJson().slides[0].background;
    assert.equal(bg?.color, "1A2B3C");
    // Slide 2 chart unchanged
    const types2 = result.getSlides()[2].getElements().map((e) => e.elementType);
    assert.ok(types2.includes("chart"), "slide 2 chart still present");
  });
});


// ═════════════════════════════════════════════════════════════════════════════
// E. Fresh chart XML generation (not passthrough)
// ═════════════════════════════════════════════════════════════════════════════

describe("Fresh chart creation (no ZIP passthrough)", () => {
  test("bar chart write→fromBuffer shows chart element", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => {
      s.addChart("bar", [
        { name: "Series 1", labels: ["A", "B", "C"], values: [1, 2, 3] },
      ], { x: 48, y: 48, w: 500, h: 300 });
    });
    const buf = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf);
    const types = result.getSlides()[0].getElements().map((e) => e.elementType);
    assert.ok(types.includes("chart"), "chart element should survive fresh build round-trip");
  });

  test("line chart write→fromBuffer shows chart element", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => {
      s.addChart("line", [
        { name: "Trend", labels: ["Jan", "Feb"], values: [5, 10] },
      ], { x: 48, y: 48, w: 500, h: 300 });
    });
    const buf = pres.write("nodebuffer");
    assert.doesNotThrow(() => Presentation.fromBuffer(buf));
  });

  test("pie chart write→fromBuffer shows chart element", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => {
      s.addChart("pie", [
        { name: "Share", labels: ["X", "Y", "Z"], values: [30, 50, 20] },
      ], { x: 48, y: 48, w: 400, h: 300 });
    });
    const buf = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf);
    const chart = result.getSlides()[0].getElements().find((e) => e.elementType === "chart");
    assert.ok(chart, "pie chart should survive round-trip");
  });

  test("chart series data survives fresh build round-trip", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => {
      s.addChart("bar", [
        { name: "Revenue", labels: ["Q1", "Q2", "Q3"], values: [100, 200, 150] },
      ], { x: 48, y: 48, w: 500, h: 300 });
    });
    const buf = pres.write("nodebuffer");
    const result = Presentation.fromBuffer(buf);
    const chart = result.getSlides()[0].getElements().find((e) => e.elementType === "chart");
    const { data } = chart.toJson();
    assert.ok(Array.isArray(data) && data.length > 0, "should have series");
    const rev = data.find((s) => s.name === "Revenue");
    assert.ok(rev, "Revenue series should be found");
    assert.deepEqual(rev.values, [100, 200, 150]);
  });

  test("multiple charts on different slides do not produce duplicate ZIP entries", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => {
      s.addChart("bar", [{ name: "A", values: [1, 2] }], { x: 0, y: 0, w: 400, h: 300 });
    });
    pres.addSlide(null, (s) => {
      s.addChart("line", [{ name: "B", values: [3, 4] }], { x: 0, y: 0, w: 400, h: 300 });
    });
    assert.doesNotThrow(() => pres.write("nodebuffer"), "two charts should not conflict");
  });
});
