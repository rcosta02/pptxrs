/**
 * Presentation lifecycle tests
 *
 * Covers: construction, metadata, slide management (add / sync / remove / get),
 * and export (write / writeFile).
 *
 * Run: node --test tests/presentation.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");
const { writeFileSync, readFileSync, mkdtempSync, rmSync, existsSync } = require("fs");
const { tmpdir } = require("os");
const { join } = require("path");

const { Presentation, roundTrip } = require("./helpers.js");


// ── Construction & metadata ───────────────────────────────────────────────────

describe("Presentation – construction", () => {
  test("new Presentation() with no options", () => {
    const pres = new Presentation();
    assert.equal(pres.layout, "LAYOUT_16x9");
    assert.ok(pres.title == null,   "title should be null/undefined");
    assert.ok(pres.author == null,  "author should be null/undefined");
    assert.ok(pres.company == null, "company should be null/undefined");
  });

  test("new Presentation() with all options", () => {
    const pres = new Presentation({
      layout: "LAYOUT_4x3",
      title: "My Title",
      author: "Alice",
      company: "ACME",
    });
    assert.equal(pres.layout, "LAYOUT_4x3");
    assert.equal(pres.title, "My Title");
    assert.equal(pres.author, "Alice");
    assert.equal(pres.company, "ACME");
  });

  test("metadata setters", () => {
    const pres = new Presentation();
    pres.title = "T";
    pres.author = "A";
    pres.company = "C";
    pres.layout = "LAYOUT_WIDE";
    assert.equal(pres.title, "T");
    assert.equal(pres.author, "A");
    assert.equal(pres.company, "C");
    assert.equal(pres.layout, "LAYOUT_WIDE");
  });

  test("LAYOUT_16x9 is default", () => {
    const pres = new Presentation({ layout: "LAYOUT_16x9" });
    assert.equal(pres.layout, "LAYOUT_16x9");
  });

  test("LAYOUT_WIDE is accepted", () => {
    const pres = new Presentation({ layout: "LAYOUT_WIDE" });
    assert.equal(pres.layout, "LAYOUT_WIDE");
  });
});


// ── Slide management ──────────────────────────────────────────────────────────

describe("Slide management", () => {
  test("addSlide() manual form — slide count is 0 until syncSlide is called", () => {
    const pres = new Presentation();
    const slide = pres.addSlide();
    assert.equal(pres.getSlides().length, 0, "slide must not be tracked until syncSlide");
    pres.syncSlide(0, slide);
    assert.equal(pres.getSlides().length, 1);
  });

  test("addSlide() callback form — auto-syncs the slide", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    assert.equal(pres.getSlides().length, 1);
  });

  test("addSlide() callback form — slide modifications are preserved", () => {
    const pres = new Presentation();
    pres.addSlide(null, (slide) => {
      slide.addText("hello", { x: 0, y: 0, w: 96, h: 48 });
    });
    const slides = pres.getSlides();
    assert.equal(slides.length, 1);
    const elements = slides[0].getElements();
    assert.equal(elements.length, 1);
    assert.equal(elements[0].elementType, "text");
  });

  test("multiple slides added via callback form", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => s.addText("Slide 1", { x: 0, y: 0, w: 96, h: 48 }));
    pres.addSlide(null, (s) => s.addText("Slide 2", { x: 0, y: 0, w: 96, h: 48 }));
    pres.addSlide(null, (s) => s.addText("Slide 3", { x: 0, y: 0, w: 96, h: 48 }));
    assert.equal(pres.getSlides().length, 3);
  });

  test("removeSlide() removes by index", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => s.addText("A", { x: 0, y: 0, w: 96, h: 48 }));
    pres.addSlide(null, (s) => s.addText("B", { x: 0, y: 0, w: 96, h: 48 }));
    pres.removeSlide(0);
    assert.equal(pres.getSlides().length, 1);
  });

  test("getSlides() returns clones — modifications require syncSlide", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => s.addText("original", { x: 0, y: 0, w: 96, h: 48 }));

    const [slide] = pres.getSlides();
    slide.addText("added", { x: 0, y: 48, w: 96, h: 48 });
    assert.equal(pres.getSlides()[0].getElements().length, 1,
      "without syncSlide, pres is unaware of the added element");

    pres.syncSlide(0, slide);
    assert.equal(pres.getSlides()[0].getElements().length, 2,
      "after syncSlide, pres reflects the change");
  });

  test("slides survive write → fromBuffer round-trip", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    pres.addSlide(null, () => {});
    const pres2 = roundTrip(pres);
    assert.equal(pres2.getSlides().length, 2);
  });
});


// ── Export / write ────────────────────────────────────────────────────────────

describe("Export – write()", () => {
  test("write('nodebuffer') returns a non-empty Buffer", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    const buf = pres.write("nodebuffer");
    assert.ok(buf instanceof Uint8Array || Buffer.isBuffer(buf));
    assert.ok(buf.length > 0);
  });

  test("write('uint8array') returns a Uint8Array", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    const buf = pres.write("uint8array");
    assert.ok(buf instanceof Uint8Array);
    assert.ok(buf.length > 0);
  });

  test("write('base64') returns a non-empty string", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    const b64 = pres.write("base64");
    assert.equal(typeof b64, "string");
    assert.ok(b64.length > 0);
    assert.match(b64, /^[A-Za-z0-9+/]+=*$/);
  });

  test("write() output starts with PK (ZIP magic bytes)", () => {
    const pres = new Presentation();
    pres.addSlide(null, () => {});
    const buf = Buffer.from(pres.write("nodebuffer"));
    assert.equal(buf[0], 0x50); // 'P'
    assert.equal(buf[1], 0x4b); // 'K'
  });

  test("writeFile() writes a valid file to disk", async () => {
    const tmpDir = mkdtempSync(join(tmpdir(), "pptxrs-"));
    const outPath = join(tmpDir, "out.pptx");
    try {
      const pres = new Presentation();
      pres.addSlide(null, () => {});
      await pres.writeFile(outPath);
      assert.ok(existsSync(outPath), "file should exist");
      const buf = readFileSync(outPath);
      assert.ok(buf.length > 0);
      assert.equal(buf[0], 0x50); // 'P'
      assert.equal(buf[1], 0x4b); // 'K'
    } finally {
      rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});
