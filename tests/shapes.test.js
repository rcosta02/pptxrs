/**
 * Shape and background tests
 *
 * Covers: addShape() — shape types, fill, line options, position;
 *         setBackground() — color persistence across round-trips.
 *
 * Run: node --test tests/shapes.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { roundTrip, oneSlide, firstElement } = require("./helpers.js");


// ── addShape() ────────────────────────────────────────────────────────────────

describe("Shapes – addShape()", () => {
  test("rect shape survives round-trip", () => {
    const pres = oneSlide((s) =>
      s.addShape("rect", { x: 96, y: 96, w: 192, h: 96 }),
    );
    const el = firstElement(pres);
    assert.equal(el.elementType, "shape");
  });

  test("shapeType is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addShape("rect", { x: 0, y: 0, w: 192, h: 96 }),
    );
    assert.equal(firstElement(pres).toJson().shapeType, "rect");
  });

  test("ellipse shape type is preserved", () => {
    const pres = oneSlide((s) =>
      s.addShape("ellipse", { x: 0, y: 0, w: 96, h: 96 }),
    );
    assert.equal(firstElement(pres).toJson().shapeType, "ellipse");
  });

  test("fill color is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addShape("rect", { x: 0, y: 0, w: 192, h: 96, fill: { color: "4472C4" } }),
    );
    assert.equal(firstElement(pres).toJson().options.fill.color, "4472C4");
  });

  test("line width and color are preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addShape("rect", {
        x: 0, y: 0, w: 192, h: 96,
        line: { color: "FF0000", width: 2 },
      }),
    );
    const line = firstElement(pres).toJson().options.line;
    assert.ok(line, "line options should be present");
    assert.equal(line.color, "FF0000");
    // Width is pt; 2pt → written as 25400 EMU → read back as 25400/12700 = 2pt
    assert.ok(Math.abs(line.width - 2) < 0.01, `width should be ~2, got ${line.width}`);
  });

  test("shape position (x/y/w/h) is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addShape("ellipse", { x: 50, y: 60, w: 150, h: 100 }),
    );
    const opts = firstElement(pres).toJson().options;
    assert.equal(opts.x, 50);
    assert.equal(opts.y, 60);
    assert.equal(opts.w, 150);
    assert.equal(opts.h, 100);
  });
});


// ── setBackground() ───────────────────────────────────────────────────────────

describe("Background – setBackground()", () => {
  test("background color survives round-trip", () => {
    const pres = oneSlide((s) => s.setBackground("F5F5F5"));
    const bg = JSON.parse(roundTrip(pres).toJsonString()).slides[0].background;
    assert.equal(bg.color, "F5F5F5");
  });

  test("background color FF0000 survives round-trip", () => {
    const pres = oneSlide((s) => s.setBackground("FF0000"));
    const bg = JSON.parse(roundTrip(pres).toJsonString()).slides[0].background;
    assert.equal(bg.color, "FF0000");
  });

  test("default background is white (FFFFFF)", () => {
    const pres = oneSlide(() => {});
    const bg = JSON.parse(roundTrip(pres).toJsonString()).slides[0].background;
    assert.ok(
      bg == null || bg.color == null || bg.color.toUpperCase() === "FFFFFF",
      `Expected white/null background, got: ${JSON.stringify(bg)}`,
    );
  });
});
