/**
 * JSON serialisation / interchange tests
 *
 * Covers: toJson(), toJsonString(), fromJson() — structure, option fidelity,
 * hand-crafted JSON, layout round-trips, and the key regression (flatten bug).
 *
 * Run: node --test tests/json.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { Presentation, roundTrip, oneSlide } = require("./helpers.js");


// ── toJson() / toJsonString() ─────────────────────────────────────────────────

describe("toJson() and toJsonString()", () => {
  test("toJson() returns an object with slides array", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    const json = pres.toJson();
    assert.ok(Array.isArray(json.slides));
    assert.equal(json.slides.length, 1);
  });

  test("toJson() includes meta with layout", () => {
    const pres = new Presentation({ layout: "LAYOUT_4x3" });
    pres.addSlide(null, () => {});
    assert.equal(pres.toJson().meta.layout, "LAYOUT_4x3");
  });

  test("toJson() element options are NOT empty — x/y/w/h present (flatten regression)", () => {
    // Regression: serde-wasm-bindgen v0.6 silently drops #[serde(flatten)] fields.
    // Fixed by routing through serde_json::to_string + js_sys::JSON::parse.
    const pres = oneSlide((s) =>
      s.addText("Test", { x: 100, y: 200, w: 300, h: 80, fontSize: 24 }),
    );
    const el = pres.toJson().slides[0].elements[0];
    assert.equal(el.options.x, 100);
    assert.equal(el.options.y, 200);
    assert.equal(el.options.w, 300);
    assert.equal(el.options.h, 80);
    assert.equal(el.options.fontSize, 24);
  });

  test("toJson() on fromBuffer result — options are present (key regression)", () => {
    const pres = oneSlide((s) =>
      s.addText("Regression", { x: 96, y: 96, w: 480, h: 72, fontSize: 32, bold: true, color: "FF0000" }),
    );
    const el = roundTrip(pres).toJson().slides[0].elements[0];
    assert.equal(el.options.fontSize, 32, "fontSize must survive fromBuffer → toJson");
    assert.equal(el.options.bold, true,   "bold must survive fromBuffer → toJson");
    assert.equal(el.options.color, "FF0000", "color must survive fromBuffer → toJson");
    assert.equal(el.options.x, 96, "x must survive fromBuffer → toJson");
  });

  test("toJsonString() is valid JSON", () => {
    const pres = oneSlide((s) => s.addText("JSON", { x: 0, y: 0, w: 480, h: 72 }));
    assert.doesNotThrow(() => JSON.parse(pres.toJsonString()));
  });

  test("toJson() and JSON.parse(toJsonString()) are equivalent", () => {
    const pres = oneSlide((s) =>
      s.addText("Equiv", { x: 10, y: 20, w: 100, h: 50, fontSize: 12 }),
    );
    const fromObj = pres.toJson();
    const fromStr = JSON.parse(pres.toJsonString());
    assert.equal(
      fromObj.slides[0].elements[0].options.x,
      fromStr.slides[0].elements[0].options.x,
    );
    assert.equal(
      fromObj.slides[0].elements[0].options.fontSize,
      fromStr.slides[0].elements[0].options.fontSize,
    );
  });
});


// ── fromJson() — JSON interchange ─────────────────────────────────────────────

describe("fromJson() – JSON interchange", () => {
  test("fromJson() reconstructs a presentation from toJson() output", () => {
    const pres = oneSlide((s) =>
      s.addText("From JSON", { x: 96, y: 96, w: 480, h: 72, fontSize: 20 }),
    );
    const pres2 = Presentation.fromJson(pres.toJson());
    const slides = pres2.getSlides();
    assert.equal(slides.length, 1);
    const el = slides[0].getElements()[0];
    assert.equal(el.elementType, "text");
    assert.equal(el.toJson().text, "From JSON");
  });

  test("fromJson() can write() after reconstruction", () => {
    const pres = oneSlide((s) => s.addText("Writable", { x: 0, y: 0, w: 200, h: 50 }));
    const pres2 = Presentation.fromJson(pres.toJson());
    const buf = pres2.write("nodebuffer");
    assert.ok(buf.length > 0);
    assert.equal(Buffer.from(buf)[0], 0x50); // PK
  });

  test("fromJson() from a hand-crafted JSON object", () => {
    const json = {
      meta: { layout: "LAYOUT_16x9" },
      slides: [{
        elements: [{
          type: "text",
          text: "Hand-crafted",
          options: { x: 96, y: 96, w: 480, h: 72, fontSize: 24 },
        }],
      }],
    };
    const pres = Presentation.fromJson(json);
    const el = pres.getSlides()[0].getElements()[0];
    assert.equal(el.elementType, "text");
    assert.equal(el.toJson().text, "Hand-crafted");
    assert.equal(el.toJson().options.fontSize, 24);
  });

  test("full round-trip: new → write → fromBuffer → toJson → fromJson → write", () => {
    const pres = oneSlide((s) => {
      s.addText("Full trip", { x: 50, y: 50, w: 400, h: 80, fontSize: 28, bold: true });
      s.addShape("ellipse", { x: 500, y: 50, w: 100, h: 100, fill: { color: "70AD47" } });
    });

    const pres2 = roundTrip(pres);
    const pres3 = Presentation.fromJson(pres2.toJson());
    const buf = pres3.write("nodebuffer");
    assert.ok(buf.length > 0);

    const pres4 = Presentation.fromBuffer(buf);
    const elements = pres4.getSlides()[0].getElements();
    assert.equal(elements.length, 2);
    assert.equal(elements[0].elementType, "text");
    assert.equal(elements[1].elementType, "shape");
  });

  test("layout is preserved through JSON round-trip", () => {
    const pres = new Presentation({ layout: "LAYOUT_4x3" });
    pres.addSlide(null, () => {});
    const pres2 = Presentation.fromJson(pres.toJson());
    assert.equal(pres2.layout, "LAYOUT_4x3");
  });
});
