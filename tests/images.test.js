/**
 * Image element tests
 *
 * Covers: addImage() — round-trip survival, position, base64 data preservation.
 *
 * Run: node --test tests/images.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { Presentation, TINY_PNG, oneSlide, firstElement, roundTrip } = require("./helpers.js");


describe("Images – addImage()", () => {
  test("image survives round-trip", () => {
    const pres = oneSlide((s) =>
      s.addImage({ data: TINY_PNG, x: 48, y: 48, w: 96, h: 96 }),
    );
    const el = firstElement(pres);
    assert.equal(el.elementType, "image");
  });

  test("image position (x/y/w/h) is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addImage({ data: TINY_PNG, x: 50, y: 60, w: 100, h: 120 }),
    );
    const opts = firstElement(pres).toJson().options;
    assert.equal(opts.x, 50);
    assert.equal(opts.y, 60);
    assert.equal(opts.w, 100);
    assert.equal(opts.h, 120);
  });

  test("image data (base64) is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addImage({ data: TINY_PNG, x: 0, y: 0, w: 96, h: 96 }),
    );
    const data = firstElement(pres).toJson().options.data;
    assert.ok(
      typeof data === "string" && data.length > 0,
      "base64 image data should be a non-empty string",
    );
  });

  test("multiple slides with images do not produce duplicate ZIP entries", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => s.addImage({ data: TINY_PNG, x: 0, y: 0, w: 96, h: 96 }));
    pres.addSlide(null, (s) => s.addImage({ data: TINY_PNG, x: 0, y: 0, w: 96, h: 96 }));
    assert.doesNotThrow(() => pres.write("nodebuffer"), "write() should not throw with images on multiple slides");
  });

  test("fromJson() with images on multiple slides writes without error", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => s.addImage({ data: TINY_PNG, x: 0, y: 0, w: 96, h: 96 }));
    pres.addSlide(null, (s) => s.addImage({ data: TINY_PNG, x: 96, y: 96, w: 96, h: 96 }));
    const json = pres.toJson();
    const pres2 = Presentation.fromJson(json);
    assert.doesNotThrow(() => pres2.write("nodebuffer"), "fromJson() → write() should not throw with images on multiple slides");
  });

  test("all images survive round-trip when on different slides", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => s.addImage({ data: TINY_PNG, x: 0,  y: 0, w: 96, h: 96 }));
    pres.addSlide(null, (s) => s.addImage({ data: TINY_PNG, x: 48, y: 48, w: 64, h: 64 }));
    const parsed = roundTrip(pres);
    const slides = parsed.getSlides();
    assert.equal(slides[0].getElements()[0].elementType, "image", "slide 0 should have an image");
    assert.equal(slides[1].getElements()[0].elementType, "image", "slide 1 should have an image");
  });
});
