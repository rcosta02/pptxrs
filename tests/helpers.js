"use strict";

const { Presentation } = require("../pkg/index.js");

/** 1×1 transparent PNG (base64) */
const TINY_PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";

/** Write a Presentation to a Buffer and read it back via fromBuffer(). */
function roundTrip(pres) {
  const buf = pres.write("nodebuffer");
  return Presentation.fromBuffer(buf);
}

/** Create a one-slide presentation using the callback form (auto-sync). */
function oneSlide(setupFn) {
  const pres = new Presentation();
  pres.addSlide(null, setupFn);
  return pres;
}

/** Return the first element of the first slide after a write→fromBuffer round-trip. */
function firstElement(pres) {
  return roundTrip(pres).getSlides()[0].getElements()[0];
}

/** Return all elements of the first slide after a write→fromBuffer round-trip. */
function allElements(pres) {
  return roundTrip(pres).getSlides()[0].getElements();
}

module.exports = { Presentation, TINY_PNG, roundTrip, oneSlide, firstElement, allElements };
