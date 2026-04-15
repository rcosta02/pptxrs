/**
 * fromBuffer() / PPTX parsing tests
 *
 * Covers: round-trip element type detection, non-empty options after parsing,
 * layout auto-detection for 16:9 / 4:3 / WIDE, and the toJsonString regression.
 *
 * Run: node --test tests/parse.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { Presentation, TINY_PNG, roundTrip, oneSlide, allElements } = require("./helpers.js");


describe("fromBuffer() – parsing", () => {
  test("slide count is correct after fromBuffer()", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    pres.addSlide(null, () => {});
    assert.equal(roundTrip(pres).getSlides().length, 2);
  });

  test("all 4 element types survive a round-trip via fromBuffer()", () => {
    const pres = oneSlide((s) => {
      s.addText("text el",  { x: 0, y: 0,   w: 400, h: 60 });
      s.addShape("rect",    { x: 0, y: 100,  w: 100, h: 60 });
      s.addImage({ data: TINY_PNG, x: 0, y: 200, w: 100, h: 100 });
      s.addTable([["A"]], { x: 0, y: 320, w: 200, h: 60 });
    });
    const types = allElements(pres).map((e) => e.elementType).sort();
    assert.ok(types.includes("text"),  "should have text");
    assert.ok(types.includes("shape"), "should have shape");
    assert.ok(types.includes("image"), "should have image");
    assert.ok(types.includes("table"), "should have table");
  });

  test("text parsed from pptx has non-empty options", () => {
    const pres = oneSlide((s) =>
      s.addText("Parsed", { x: 96, y: 96, w: 480, h: 72, fontSize: 20, bold: true, color: "1F497D" }),
    );
    const opts = allElements(pres)[0].toJson().options;
    assert.ok(opts && typeof opts === "object", "options should be a non-null object");
    assert.equal(opts.fontSize, 20);
    assert.equal(opts.bold, true);
    assert.equal(opts.color, "1F497D");
  });

  test("layout is detected correctly from a 16:9 PPTX", () => {
    const pres = new Presentation({ layout: "LAYOUT_16x9" });
    pres.addSlide(null, () => {});
    assert.equal(roundTrip(pres).layout, "LAYOUT_16x9");
  });

  test("layout is detected correctly from a 4:3 PPTX", () => {
    const pres = new Presentation({ layout: "LAYOUT_4x3" });
    pres.addSlide(null, () => {});
    assert.equal(roundTrip(pres).layout, "LAYOUT_4x3");
  });

  test("layout is detected correctly from a WIDE PPTX", () => {
    const pres = new Presentation({ layout: "LAYOUT_WIDE" });
    pres.addSlide(null, () => {});
    assert.equal(roundTrip(pres).layout, "LAYOUT_WIDE");
  });

  test("toJsonString() from a parsed PPTX includes all options (regression)", () => {
    const pres = oneSlide((s) =>
      s.addText("Str", { x: 10, y: 20, w: 200, h: 50, fontSize: 14, align: "right" }),
    );
    const el = JSON.parse(roundTrip(pres).toJsonString()).slides[0].elements[0];
    assert.equal(el.options.x, 10);
    assert.equal(el.options.y, 20);
    assert.equal(el.options.fontSize, 14);
    assert.equal(el.options.align, "right");
  });
});
