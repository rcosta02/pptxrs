/**
 * Speaker notes tests
 *
 * Covers: addNotes() — in-memory presence, and the documented limitation that
 * notes are stored in separate notesSlideX.xml files and are NOT re-parsed by
 * fromBuffer().
 *
 * Run: node --test tests/notes.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { Presentation, roundTrip } = require("./helpers.js");


describe("Notes – addNotes()", () => {
  test("notes element is present in-memory after addNotes()", () => {
    const pres = new Presentation();
    pres.addSlide(null, (s) => {
      s.addText("Slide content", { x: 0, y: 0, w: 480, h: 72 });
      s.addNotes("Speaker note here");
    });
    const notesEl = pres.getSlides()[0].getElements().find((e) => e.elementType === "notes");
    assert.ok(notesEl, "should have a notes element in-memory");
  });

  test("notes are NOT present after write→fromBuffer (stored in separate notesSlideX.xml)", () => {
    // Speaker notes are written to a separate notesSlideX.xml ZIP entry but are
    // not re-parsed by fromBuffer(). This test documents that known behaviour.
    const pres = new Presentation();
    pres.addSlide(null, (s) => {
      s.addText("Content", { x: 0, y: 0, w: 480, h: 72 });
      s.addNotes("These notes won't survive round-trip");
    });
    const parsed = roundTrip(pres);
    const notesEl = parsed.getSlides()[0].getElements().find((e) => e.elementType === "notes");
    assert.ok(!notesEl, "notes should not be present after fromBuffer round-trip");
  });
});
