/**
 * Text element tests
 *
 * Covers: addText() — plain strings, TextRun arrays, styling (fontSize, bold,
 * italic, color, align, wrap), height estimation, position, percentage coords.
 *
 * Run: node --test tests/text.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { Presentation, roundTrip, oneSlide, firstElement, allElements } = require("./helpers.js");


// ── addText() — content & styling ────────────────────────────────────────────

describe("Text – addText()", () => {
  test("plain string survives round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Hello World", { x: 96, y: 96, w: 480, h: 72 }),
    );
    const el = firstElement(pres);
    assert.equal(el.elementType, "text");
    assert.equal(el.toJson().text, "Hello World");
  });

  test("fontSize is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Text", { x: 0, y: 0, w: 480, h: 72, fontSize: 36 }),
    );
    assert.equal(firstElement(pres).toJson().options.fontSize, 36);
  });

  test("bold: true is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Bold", { x: 0, y: 0, w: 480, h: 72, bold: true }),
    );
    assert.equal(firstElement(pres).toJson().options.bold, true);
  });

  test("bold: false is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Not bold", { x: 0, y: 0, w: 480, h: 72, bold: false }),
    );
    assert.equal(firstElement(pres).toJson().options.bold, false);
  });

  test("italic is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Italic", { x: 0, y: 0, w: 480, h: 72, italic: true }),
    );
    assert.equal(firstElement(pres).toJson().options.italic, true);
  });

  test("color is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Red", { x: 0, y: 0, w: 480, h: 72, color: "FF0000" }),
    );
    assert.equal(firstElement(pres).toJson().options.color, "FF0000");
  });

  test("align: center is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Centered", { x: 0, y: 0, w: 480, h: 72, align: "center" }),
    );
    assert.equal(firstElement(pres).toJson().options.align, "center");
  });

  test("align: right is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Right", { x: 0, y: 0, w: 480, h: 72, align: "right" }),
    );
    assert.equal(firstElement(pres).toJson().options.align, "right");
  });

  test("wrap: false is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("No wrap", { x: 0, y: 0, w: 480, h: 72, wrap: false }),
    );
    assert.equal(firstElement(pres).toJson().options.wrap, false);
  });

  test("wrap: true is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Wrap", { x: 0, y: 0, w: 480, h: 72, wrap: true }),
    );
    assert.equal(firstElement(pres).toJson().options.wrap, true);
  });

  test("position (x/y/w/h) is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addText("Pos", { x: 100, y: 200, w: 300, h: 80 }),
    );
    const opts = firstElement(pres).toJson().options;
    assert.equal(opts.x, 100);
    assert.equal(opts.y, 200);
    assert.equal(opts.w, 300);
    assert.equal(opts.h, 80);
  });

  test("TextRun array — runs are preserved", () => {
    const pres = oneSlide((s) =>
      s.addText(
        [
          { text: "Normal, ", options: {} },
          { text: "bold",     options: { bold: true } },
        ],
        { x: 0, y: 0, w: 480, h: 72 },
      ),
    );
    const data = firstElement(pres).toJson();
    assert.ok(Array.isArray(data.text), "text should be an array of runs");
    const combined = data.text.map((r) => r.text).join("");
    assert.ok(
      combined.includes("Normal") || combined.includes("bold"),
      "run text should contain original content",
    );
  });

  test("multiple text elements on one slide", () => {
    const pres = oneSlide((s) => {
      s.addText("First",  { x: 0, y: 0,   w: 200, h: 50 });
      s.addText("Second", { x: 0, y: 60,  w: 200, h: 50 });
      s.addText("Third",  { x: 0, y: 120, w: 200, h: 50 });
    });
    const els = allElements(pres);
    assert.equal(els.length, 3);
    assert.ok(els.every((e) => e.elementType === "text"));
  });
});


// ── Height estimation (h omitted) ─────────────────────────────────────────────

describe("Text – height estimation when h is omitted", () => {
  test("h is optional — no error when omitted", () => {
    assert.doesNotThrow(() =>
      oneSlide((s) => s.addText("No H", { x: 0, y: 0, w: 480, fontSize: 18 })),
    );
  });

  test("h omitted — getHeight() estimates from fontSize (18pt → 36px)", () => {
    const pres = oneSlide((s) =>
      s.addText("Auto H", { x: 0, y: 0, w: 480, fontSize: 18 }),
    );
    assert.equal(firstElement(pres).getHeight(), 36, "18pt → 18/72*96*1.5 = 36px");
  });

  test("h omitted — getHeight() estimates from fontSize (36pt → 72px)", () => {
    const pres = oneSlide((s) =>
      s.addText("Auto H", { x: 0, y: 0, w: 480, fontSize: 36 }),
    );
    assert.equal(firstElement(pres).getHeight(), 72, "36pt → 36/72*96*1.5 = 72px");
  });

  test("h omitted, no fontSize — getHeight() uses 18pt default (→ 36px)", () => {
    const pres = oneSlide((s) =>
      s.addText("Default font", { x: 0, y: 0, w: 480 }),
    );
    assert.equal(firstElement(pres).getHeight(), 36);
  });
});


// ── Percentage coordinates ────────────────────────────────────────────────────

describe("Percentage coordinates", () => {
  test("percentage x/y/w/h do not throw", () => {
    assert.doesNotThrow(() =>
      oneSlide((s) => s.addText("Pct", { x: "10%", y: "10%", w: "80%", h: "10%" })),
    );
  });

  test("getWidth() with '100%' returns slide width in pixels (LAYOUT_16x9 = 960px)", () => {
    // 16:9: 9144000 EMU wide → 9144000 / 9525 = 960 px
    const pres = new Presentation({ layout: "LAYOUT_16x9" });
    pres.addSlide(null, (s) =>
      s.addText("Full width", { x: "0%", y: "0%", w: "100%", h: "10%" }),
    );
    const el = roundTrip(pres).getSlides()[0].getElements()[0];
    assert.ok(
      Math.abs(el.getWidth() - 960) < 2,
      `Expected ~960px for 100% width, got ${el.getWidth()}`,
    );
  });
});
