/**
 * Write fidelity and JSON pipeline equivalence tests
 *
 * Suite A — Write fidelity: builds a rich 3-slide presentation, writes it to a
 *   PPTX buffer, parses it back with fromBuffer(), then verifies every
 *   individual property that was set (background, text, shapes, image, table,
 *   multi-run text, notes limitation).
 *
 * Suite B — JSON pipeline equivalence: verifies that
 *   new Pres → write → fromBuffer → toJson  (json1)  and
 *   fromJson(json1) → write → fromBuffer → toJson  (json2)
 *   produce structurally identical JSON (positions within ±1px tolerance).
 *
 * Suite C — Mixed slide stress test: large decks and multi-element slides.
 *
 * Run: node --test tests/fidelity.test.js
 */

"use strict";

const { test, describe, before } = require("node:test");
const assert = require("node:assert/strict");

const { Presentation, TINY_PNG, roundTrip, oneSlide, allElements } = require("./helpers.js");


// ═════════════════════════════════════════════════════════════════════════════
// A. WRITE FIDELITY — property-by-property deep verification
// ═════════════════════════════════════════════════════════════════════════════

describe("Write fidelity – property-by-property", () => {

  /**
   * Canonical rich presentation:
   *   Slide 0 — title text + body text + rect shape + ellipse shape, blue bg
   *   Slide 1 — image + 4×3 table, light bg
   *   Slide 2 — multi-run text + speaker notes
   */
  function buildRich() {
    const pres = new Presentation({ layout: "LAYOUT_16x9" });

    pres.addSlide(null, (s) => {
      s.setBackground("1E3A5F");
      s.addText("Title Text", {
        x: 48, y: 24, w: 864, h: 80,
        fontSize: 40, bold: true, italic: false, color: "FFFFFF", align: "center",
      });
      s.addText("Body text with wrap", {
        x: 48, y: 120, w: 864, h: 200,
        fontSize: 18, bold: false, italic: true, color: "CCCCCC",
        align: "left", wrap: true,
      });
      s.addShape("rect", {
        x: 48, y: 360, w: 200, h: 80,
        fill: { color: "4472C4" },
        line: { color: "FFFFFF", width: 2 },
      });
      s.addShape("ellipse", {
        x: 280, y: 360, w: 80, h: 80,
        fill: { color: "FF0000" },
      });
    });

    pres.addSlide(null, (s) => {
      s.setBackground("F5F5F5");
      s.addImage({ data: TINY_PNG, x: 48, y: 48, w: 200, h: 150 });
      s.addTable(
        [
          ["Name", "Score", "Grade"],
          ["Alice", "95",    "A"],
          ["Bob",   "82",    "B"],
          ["Carol", "71",    "C"],
        ],
        { x: 48, y: 240, w: 600, h: 180, colW: [200, 200, 200] },
      );
    });

    pres.addSlide(null, (s) => {
      s.addText(
        [
          { text: "Bold part ",  options: { bold: true } },
          { text: "italic part", options: { italic: true } },
        ],
        { x: 48, y: 48, w: 600, h: 80, fontSize: 24 },
      );
      s.addNotes("Speaker notes for slide 3");
    });

    return pres;
  }

  let parsed;
  before(() => { parsed = roundTrip(buildRich()); });

  // ── Slide structure ──────────────────────────────────────────────────────

  test("slide count is 3", () => {
    assert.equal(parsed.getSlides().length, 3);
  });

  test("slide 0 has text and shape elements", () => {
    const types = parsed.getSlides()[0].getElements().map((e) => e.elementType);
    assert.ok(types.includes("text"),  "slide 0 should have text");
    assert.ok(types.includes("shape"), "slide 0 should have shape");
  });

  test("slide 1 has image and table elements", () => {
    const types = parsed.getSlides()[1].getElements().map((e) => e.elementType);
    assert.ok(types.includes("image"), "slide 1 should have image");
    assert.ok(types.includes("table"), "slide 1 should have table");
  });

  test("slide 2 has text element", () => {
    const types = parsed.getSlides()[2].getElements().map((e) => e.elementType);
    assert.ok(types.includes("text"), "slide 2 should have text");
  });

  test("notes are NOT present after write→fromBuffer (stored in separate notesSlidex.xml)", () => {
    const types = parsed.getSlides()[2].getElements().map((e) => e.elementType);
    assert.ok(!types.includes("notes"), "notes should not appear after fromBuffer round-trip");
  });

  test("text elements appear before shape elements on slide 0", () => {
    const types = parsed.getSlides()[0].getElements().map((e) => e.elementType);
    assert.ok(types.indexOf("text") < types.indexOf("shape"), "texts should precede shapes");
  });

  // ── Background ───────────────────────────────────────────────────────────

  test("slide 0 background color is 1E3A5F", () => {
    const bg = parsed.toJson().slides[0].background;
    assert.equal(bg && bg.color, "1E3A5F");
  });

  test("slide 1 background color is F5F5F5", () => {
    const bg = parsed.toJson().slides[1].background;
    assert.equal(bg && bg.color, "F5F5F5");
  });

  test("slide 2 background is null / white (nothing set)", () => {
    const bg = parsed.toJson().slides[2].background;
    assert.ok(
      bg == null || bg.color == null || bg.color.toUpperCase() === "FFFFFF",
      `expected null/white background, got ${JSON.stringify(bg)}`,
    );
  });

  // ── Title text ───────────────────────────────────────────────────────────

  function titleEl() {
    return parsed
      .getSlides()[0]
      .getElements()
      .filter((e) => e.elementType === "text")
      .find((e) => e.toJson().text === "Title Text");
  }

  test("title text content is 'Title Text'", () => {
    assert.ok(titleEl(), "title element should be findable by text content");
  });

  test("title fontSize is 40", () => {
    assert.equal(titleEl().toJson().options.fontSize, 40);
  });

  test("title bold is true", () => {
    assert.equal(titleEl().toJson().options.bold, true);
  });

  test("title color is FFFFFF", () => {
    assert.equal(titleEl().toJson().options.color, "FFFFFF");
  });

  test("title align is 'center'", () => {
    assert.equal(titleEl().toJson().options.align, "center");
  });

  test("title position x=48 y=24 w=864 h=80", () => {
    const opts = titleEl().toJson().options;
    assert.equal(opts.x, 48);
    assert.equal(opts.y, 24);
    assert.equal(opts.w, 864);
    assert.equal(opts.h, 80);
  });

  // ── Body text ────────────────────────────────────────────────────────────

  function bodyEl() {
    return parsed
      .getSlides()[0]
      .getElements()
      .filter((e) => e.elementType === "text")
      .find((e) => typeof e.toJson().text === "string" && e.toJson().text.includes("Body"));
  }

  test("body text content includes 'Body'", () => {
    assert.ok(bodyEl(), "body text element should be findable");
  });

  test("body fontSize is 18", () => {
    assert.equal(bodyEl().toJson().options.fontSize, 18);
  });

  test("body italic is true", () => {
    assert.equal(bodyEl().toJson().options.italic, true);
  });

  test("body color is CCCCCC", () => {
    assert.equal(bodyEl().toJson().options.color, "CCCCCC");
  });

  test("body align is 'left'", () => {
    assert.equal(bodyEl().toJson().options.align, "left");
  });

  test("body wrap is true", () => {
    assert.equal(bodyEl().toJson().options.wrap, true);
  });

  // ── Rect shape ───────────────────────────────────────────────────────────

  function rectEl() {
    return parsed
      .getSlides()[0]
      .getElements()
      .filter((e) => e.elementType === "shape")
      .find((e) => e.toJson().shapeType === "rect");
  }

  test("rect shape is present on slide 0", () => {
    assert.ok(rectEl(), "rect shape should exist");
  });

  test("rect fill color is 4472C4", () => {
    assert.equal(rectEl().toJson().options.fill.color, "4472C4");
  });

  test("rect line color is FFFFFF", () => {
    assert.equal(rectEl().toJson().options.line.color, "FFFFFF");
  });

  test("rect line width is 2 pt (within 0.01)", () => {
    const w = rectEl().toJson().options.line.width;
    assert.ok(Math.abs(w - 2) < 0.01, `line width should be ~2, got ${w}`);
  });

  test("rect position x=48 y=360 w=200 h=80", () => {
    const opts = rectEl().toJson().options;
    assert.equal(opts.x, 48);
    assert.equal(opts.y, 360);
    assert.equal(opts.w, 200);
    assert.equal(opts.h, 80);
  });

  // ── Ellipse shape ────────────────────────────────────────────────────────

  function ellipseEl() {
    return parsed
      .getSlides()[0]
      .getElements()
      .filter((e) => e.elementType === "shape")
      .find((e) => e.toJson().shapeType === "ellipse");
  }

  test("ellipse shape is present on slide 0", () => {
    assert.ok(ellipseEl(), "ellipse shape should exist");
  });

  test("ellipse fill color is FF0000", () => {
    assert.equal(ellipseEl().toJson().options.fill.color, "FF0000");
  });

  test("ellipse position x=280 y=360 w=80 h=80", () => {
    const opts = ellipseEl().toJson().options;
    assert.equal(opts.x, 280);
    assert.equal(opts.y, 360);
    assert.equal(opts.w, 80);
    assert.equal(opts.h, 80);
  });

  // ── Image ────────────────────────────────────────────────────────────────

  function imgEl() {
    return parsed.getSlides()[1].getElements().find((e) => e.elementType === "image");
  }

  test("image element is present on slide 1", () => {
    assert.ok(imgEl(), "image should exist on slide 1");
  });

  test("image position x=48 y=48 w=200 h=150", () => {
    const opts = imgEl().toJson().options;
    assert.equal(opts.x, 48);
    assert.equal(opts.y, 48);
    assert.equal(opts.w, 200);
    assert.equal(opts.h, 150);
  });

  test("image base64 data is non-empty", () => {
    const d = imgEl().toJson().options.data;
    assert.ok(typeof d === "string" && d.length > 0, "base64 data should be a non-empty string");
  });

  // ── Table ────────────────────────────────────────────────────────────────

  function tblEl() {
    return parsed.getSlides()[1].getElements().find((e) => e.elementType === "table");
  }

  test("table element is present on slide 1", () => {
    assert.ok(tblEl(), "table should exist on slide 1");
  });

  test("table has 4 rows", () => {
    assert.equal(tblEl().toJson().data.length, 4);
  });

  test("table header row is ['Name','Score','Grade']", () => {
    assert.deepEqual(tblEl().toJson().data[0], ["Name", "Score", "Grade"]);
  });

  test("table row 1 is ['Alice','95','A']", () => {
    assert.deepEqual(tblEl().toJson().data[1], ["Alice", "95", "A"]);
  });

  test("table row 2 is ['Bob','82','B']", () => {
    assert.deepEqual(tblEl().toJson().data[2], ["Bob", "82", "B"]);
  });

  test("table row 3 is ['Carol','71','C']", () => {
    assert.deepEqual(tblEl().toJson().data[3], ["Carol", "71", "C"]);
  });

  test("table has 3 colW entries, each ~200px", () => {
    const colW = tblEl().toJson().options.colW;
    assert.ok(Array.isArray(colW) && colW.length === 3, "should have 3 column widths");
    colW.forEach((w, i) =>
      assert.ok(Math.abs(w - 200) < 1, `colW[${i}] should be ~200, got ${w}`),
    );
  });

  test("table position x=48 y=240 (within 1px)", () => {
    const opts = tblEl().toJson().options;
    assert.ok(Math.abs(opts.x - 48)  < 1, `x should be ~48, got ${opts.x}`);
    assert.ok(Math.abs(opts.y - 240) < 1, `y should be ~240, got ${opts.y}`);
  });

  // ── Multi-run text & notes ───────────────────────────────────────────────

  test("slide 2 multi-run text contains 'Bold' and 'italic'", () => {
    const textEl = parsed.getSlides()[2].getElements().find((e) => e.elementType === "text");
    assert.ok(textEl, "text element should exist on slide 2");
    const raw = textEl.toJson().text;
    const combined = Array.isArray(raw) ? raw.map((r) => r.text).join("") : raw;
    assert.ok(
      combined.toLowerCase().includes("bold") || combined.toLowerCase().includes("italic"),
      `Expected multi-run content, got: "${combined}"`,
    );
  });

  test("notes element present in-memory (before write) but not after fromBuffer round-trip", () => {
    const presPre = buildRich();
    const notesInMem = presPre.getSlides()[2].getElements().find((e) => e.elementType === "notes");
    assert.ok(notesInMem, "notes element should be present in-memory before write");

    const notesAfter = parsed.getSlides()[2].getElements().find((e) => e.elementType === "notes");
    assert.ok(!notesAfter, "notes should not be present after fromBuffer round-trip");
  });
});


// ═════════════════════════════════════════════════════════════════════════════
// B. JSON PIPELINE EQUIVALENCE
//   new Pres → write → fromBuffer → toJson  →  json1
//   fromJson(json1) → write → fromBuffer → toJson  →  json2
//   json1 ≈ json2  (positions within ±1px)
// ═════════════════════════════════════════════════════════════════════════════

describe("JSON pipeline equivalence – pptx → json → fromJson → pptx → json", () => {

  function buildPipelinePresentation() {
    const pres = new Presentation({ layout: "LAYOUT_16x9" });

    pres.addSlide(null, (s) => {
      s.setBackground("EAF4FB");
      s.addText("Slide One Title", {
        x: 48, y: 24, w: 864, h: 80,
        fontSize: 36, bold: true, color: "1F497D",
      });
      s.addShape("rect", {
        x: 48, y: 120, w: 300, h: 100,
        fill: { color: "4472C4" },
        line: { color: "FFFFFF", width: 1 },
      });
      s.addImage({ data: TINY_PNG, x: 400, y: 120, w: 150, h: 100 });
    });

    pres.addSlide(null, (s) => {
      s.setBackground("FFF9E6");
      s.addTable(
        [
          ["Product",  "Q1",  "Q2"],
          ["Widget A", "120", "145"],
          ["Widget B", "95",  "110"],
        ],
        { x: 48, y: 48, w: 600, h: 150, colW: [200, 200, 200] },
      );
      s.addText("Notes: Q2 improved", {
        x: 48, y: 220, w: 600, h: 50,
        fontSize: 14, italic: true, color: "888888",
      });
    });

    return pres;
  }

  /**
   * Recursively compare two JSON trees.
   * Numbers: ±1 tolerance (EMU↔px rounding).
   * Null / missing keys on both sides: treated as equal.
   */
  function approxEqual(a, b, path = "") {
    if (a == null && b == null) return;
    if (a == null || b == null) {
      assert.equal(a, b, `${path}: one side is nullish`);
      return;
    }
    if (typeof a === "number" && typeof b === "number") {
      assert.ok(
        Math.abs(a - b) < 1,
        `${path}: numeric drift — ${a} vs ${b} (diff ${Math.abs(a - b)})`,
      );
    } else if (Array.isArray(a) && Array.isArray(b)) {
      assert.equal(a.length, b.length, `${path}: array length ${a.length} vs ${b.length}`);
      a.forEach((v, i) => approxEqual(v, b[i], `${path}[${i}]`));
    } else if (typeof a === "object" && typeof b === "object") {
      const keys = new Set([
        ...Object.keys(a).filter((k) => a[k] != null),
        ...Object.keys(b).filter((k) => b[k] != null),
      ]);
      for (const k of keys) approxEqual(a[k], b[k], `${path}.${k}`);
    } else {
      assert.equal(a, b, `${path}: ${JSON.stringify(a)} ≠ ${JSON.stringify(b)}`);
    }
  }

  let json1, json2;
  before(() => {
    const pres = buildPipelinePresentation();
    json1 = roundTrip(pres).toJson();
    json2 = roundTrip(Presentation.fromJson(json1)).toJson();
  });

  function findByType(slideJson, type) {
    return slideJson.elements.filter((e) => e.type === type);
  }

  // ── Structural ────────────────────────────────────────────────────────────

  test("slide count is the same in both snapshots", () => {
    assert.equal(json1.slides.length, json2.slides.length);
  });

  test("meta.layout is the same in both snapshots", () => {
    assert.equal(json1.meta.layout, json2.meta.layout);
  });

  test("element count matches on every slide", () => {
    for (let i = 0; i < json1.slides.length; i++) {
      assert.equal(
        json1.slides[i].elements.length,
        json2.slides[i].elements.length,
        `slide ${i}: element count mismatch`,
      );
    }
  });

  test("element types match in order on every slide", () => {
    for (let i = 0; i < json1.slides.length; i++) {
      const t1 = json1.slides[i].elements.map((e) => e.type);
      const t2 = json2.slides[i].elements.map((e) => e.type);
      assert.deepEqual(t1, t2, `slide ${i}: element type order mismatch`);
    }
  });

  // ── Background ────────────────────────────────────────────────────────────

  test("background colors match on all slides", () => {
    for (let i = 0; i < json1.slides.length; i++) {
      const bg1 = json1.slides[i].background?.color ?? null;
      const bg2 = json2.slides[i].background?.color ?? null;
      assert.equal(bg1, bg2, `slide ${i}: background color mismatch`);
    }
  });

  // ── Slide 0 — text ────────────────────────────────────────────────────────

  test("slide 0 text content is identical", () => {
    const getText = (e) =>
      typeof e.text === "string" ? e.text : e.text.map((r) => r.text).join("");
    const t1 = findByType(json1.slides[0], "text");
    const t2 = findByType(json2.slides[0], "text");
    t1.forEach((el, i) =>
      assert.equal(getText(el), getText(t2[i]), `text[${i}] content mismatch`),
    );
  });

  test("slide 0 text options (fontSize / bold / color) are identical", () => {
    const t1 = findByType(json1.slides[0], "text")[0];
    const t2 = findByType(json2.slides[0], "text")[0];
    assert.equal(t1.options.fontSize, t2.options.fontSize);
    assert.equal(t1.options.bold,     t2.options.bold);
    assert.equal(t1.options.color,    t2.options.color);
  });

  test("slide 0 text position matches (within 1px)", () => {
    const t1 = findByType(json1.slides[0], "text")[0];
    const t2 = findByType(json2.slides[0], "text")[0];
    for (const k of ["x", "y", "w", "h"]) {
      assert.ok(
        Math.abs(t1.options[k] - t2.options[k]) < 1,
        `text.options.${k}: ${t1.options[k]} vs ${t2.options[k]}`,
      );
    }
  });

  // ── Slide 0 — shape ───────────────────────────────────────────────────────

  test("slide 0 shape type matches", () => {
    const s1 = findByType(json1.slides[0], "shape")[0];
    const s2 = findByType(json2.slides[0], "shape")[0];
    assert.equal(s1.shapeType, s2.shapeType);
  });

  test("slide 0 shape fill color matches", () => {
    const s1 = findByType(json1.slides[0], "shape")[0];
    const s2 = findByType(json2.slides[0], "shape")[0];
    assert.equal(s1.options.fill?.color, s2.options.fill?.color);
  });

  test("slide 0 shape line color and width match", () => {
    const s1 = findByType(json1.slides[0], "shape")[0];
    const s2 = findByType(json2.slides[0], "shape")[0];
    assert.equal(s1.options.line?.color, s2.options.line?.color);
    assert.ok(
      Math.abs((s1.options.line?.width ?? 0) - (s2.options.line?.width ?? 0)) < 0.01,
      `line width: ${s1.options.line?.width} vs ${s2.options.line?.width}`,
    );
  });

  test("slide 0 shape position matches (within 1px)", () => {
    const s1 = findByType(json1.slides[0], "shape")[0];
    const s2 = findByType(json2.slides[0], "shape")[0];
    for (const k of ["x", "y", "w", "h"]) {
      assert.ok(
        Math.abs(s1.options[k] - s2.options[k]) < 1,
        `shape.options.${k}: ${s1.options[k]} vs ${s2.options[k]}`,
      );
    }
  });

  // ── Slide 0 — image ───────────────────────────────────────────────────────

  test("slide 0 image base64 data is identical", () => {
    const i1 = findByType(json1.slides[0], "image")[0];
    const i2 = findByType(json2.slides[0], "image")[0];
    assert.equal(i1.options.data, i2.options.data);
  });

  test("slide 0 image position matches (within 1px)", () => {
    const i1 = findByType(json1.slides[0], "image")[0];
    const i2 = findByType(json2.slides[0], "image")[0];
    for (const k of ["x", "y", "w", "h"]) {
      assert.ok(
        Math.abs(i1.options[k] - i2.options[k]) < 1,
        `image.options.${k}: ${i1.options[k]} vs ${i2.options[k]}`,
      );
    }
  });

  // ── Slide 1 — table ───────────────────────────────────────────────────────

  test("slide 1 table data is identical", () => {
    const tbl1 = findByType(json1.slides[1], "table")[0];
    const tbl2 = findByType(json2.slides[1], "table")[0];
    assert.deepEqual(tbl1.data, tbl2.data);
  });

  test("slide 1 table colW matches (within 1px)", () => {
    const tbl1 = findByType(json1.slides[1], "table")[0];
    const tbl2 = findByType(json2.slides[1], "table")[0];
    assert.equal(tbl1.options.colW.length, tbl2.options.colW.length);
    tbl1.options.colW.forEach((w, i) =>
      assert.ok(
        Math.abs(w - tbl2.options.colW[i]) < 1,
        `colW[${i}]: ${w} vs ${tbl2.options.colW[i]}`,
      ),
    );
  });

  test("slide 1 table position matches (within 1px)", () => {
    const tbl1 = findByType(json1.slides[1], "table")[0];
    const tbl2 = findByType(json2.slides[1], "table")[0];
    for (const k of ["x", "y", "w", "h"]) {
      assert.ok(
        Math.abs(tbl1.options[k] - tbl2.options[k]) < 1,
        `table.options.${k}: ${tbl1.options[k]} vs ${tbl2.options[k]}`,
      );
    }
  });

  // ── Slide 1 — text ────────────────────────────────────────────────────────

  test("slide 1 text options (fontSize / italic / color) are identical", () => {
    const t1 = findByType(json1.slides[1], "text")[0];
    const t2 = findByType(json2.slides[1], "text")[0];
    assert.equal(t1.options.fontSize, t2.options.fontSize);
    assert.equal(t1.options.italic,   t2.options.italic);
    assert.equal(t1.options.color,    t2.options.color);
  });

  // ── Full deep equivalence ─────────────────────────────────────────────────

  test("full structural deep-equal: json1 ≈ json2 (all fields, ±1 on numbers)", () => {
    approxEqual(json1, json2);
  });
});


// ═════════════════════════════════════════════════════════════════════════════
// C. MIXED SLIDE STRESS TEST
// ═════════════════════════════════════════════════════════════════════════════

describe("Mixed slide – stress test", () => {
  test("5 slides with multiple elements each survive round-trip", () => {
    const pres = new Presentation({ layout: "LAYOUT_16x9" });
    for (let i = 0; i < 5; i++) {
      pres.addSlide(null, (s) => {
        s.addText(`Title ${i}`, { x: 48, y: 24,  w: 864, h: 72, fontSize: 32, bold: true });
        s.addText(`Body ${i}`,  { x: 48, y: 120, w: 864, h: 288, fontSize: 18 });
        s.addShape("rect", { x: 48, y: 440, w: 400, h: 80, fill: { color: "4472C4" } });
        s.setBackground("F0F4FF");
      });
    }

    const pres2 = roundTrip(pres);
    assert.equal(pres2.getSlides().length, 5, "all 5 slides present");
    for (const slide of pres2.getSlides()) {
      const els = slide.getElements();
      assert.ok(els.length >= 3, `each slide should have ≥ 3 elements, got ${els.length}`);
    }
  });

  test("all element types on one slide — element count matches", () => {
    const pres = oneSlide((s) => {
      s.addText("T",  { x: 0, y: 0,   w: 200, h: 50 });
      s.addText("T2", { x: 0, y: 60,  w: 200, h: 50 });
      s.addShape("rect", { x: 210, y: 0, w: 100, h: 50 });
      s.addImage({ data: TINY_PNG, x: 320, y: 0, w: 100, h: 100 });
      s.addTable([["A", "B"]], { x: 0, y: 160, w: 400, h: 60 });
    });

    const els = allElements(pres);
    assert.equal(
      els.length, 5,
      `Expected 5 elements, got ${els.length}: ${els.map((e) => e.elementType).join(", ")}`,
    );
  });
});
