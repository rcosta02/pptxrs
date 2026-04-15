/**
 * SlideElementObject accessor tests
 *
 * Covers: getWidth/getHeight/getX/getY in pixels, and the *Inches() variants,
 * across all element types (text, shape, table, image).
 *
 * Run: node --test tests/elements.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { TINY_PNG, oneSlide, firstElement } = require("./helpers.js");


describe("SlideElementObject – dimension methods", () => {
  test("getWidth() returns correct pixels", () => {
    const pres = oneSlide((s) => s.addText("W", { x: 0, y: 0, w: 480, h: 72 }));
    assert.equal(firstElement(pres).getWidth(), 480);
  });

  test("getHeight() returns correct pixels (explicit h)", () => {
    const pres = oneSlide((s) => s.addText("H", { x: 0, y: 0, w: 480, h: 96 }));
    assert.equal(firstElement(pres).getHeight(), 96);
  });

  test("getX() returns correct pixels", () => {
    const pres = oneSlide((s) => s.addText("X", { x: 128, y: 0, w: 480, h: 72 }));
    assert.equal(firstElement(pres).getX(), 128);
  });

  test("getY() returns correct pixels", () => {
    const pres = oneSlide((s) => s.addText("Y", { x: 0, y: 200, w: 480, h: 72 }));
    assert.equal(firstElement(pres).getY(), 200);
  });

  test("getWidthInches() is getWidth() / 96", () => {
    const pres = oneSlide((s) => s.addText("Inches", { x: 0, y: 0, w: 480, h: 72 }));
    const el = firstElement(pres);
    assert.ok(
      Math.abs(el.getWidthInches() - el.getWidth() / 96) < 0.001,
      `getWidthInches=${el.getWidthInches()} but getWidth/96=${el.getWidth() / 96}`,
    );
  });

  test("getHeightInches() is getHeight() / 96", () => {
    const pres = oneSlide((s) => s.addText("Inches H", { x: 0, y: 0, w: 480, h: 96 }));
    const el = firstElement(pres);
    assert.ok(Math.abs(el.getHeightInches() - el.getHeight() / 96) < 0.001);
  });

  test("getXInches() is getX() / 96", () => {
    const pres = oneSlide((s) => s.addText("XI", { x: 192, y: 0, w: 480, h: 72 }));
    const el = firstElement(pres);
    assert.ok(Math.abs(el.getXInches() - el.getX() / 96) < 0.001);
  });

  test("getYInches() is getY() / 96", () => {
    const pres = oneSlide((s) => s.addText("YI", { x: 0, y: 288, w: 480, h: 72 }));
    const el = firstElement(pres);
    assert.ok(Math.abs(el.getYInches() - el.getY() / 96) < 0.001);
  });

  test("shape dimensions via getElements() — correct pixels", () => {
    const pres = oneSlide((s) => s.addShape("rect", { x: 0, y: 0, w: 192, h: 96 }));
    const el = firstElement(pres);
    assert.equal(el.getWidth(), 192);
    assert.equal(el.getHeight(), 96);
  });

  test("table dimensions via getElements() — correct pixels", () => {
    const pres = oneSlide((s) => s.addTable([["A"]], { x: 0, y: 0, w: 480, h: 144 }));
    const el = firstElement(pres);
    assert.equal(el.getWidth(), 480);
    assert.equal(el.getHeight(), 144);
  });

  test("image dimensions via getElements() — correct pixels", () => {
    const pres = oneSlide((s) =>
      s.addImage({ data: TINY_PNG, x: 0, y: 0, w: 200, h: 150 }),
    );
    const el = firstElement(pres);
    assert.equal(el.getWidth(), 200);
    assert.equal(el.getHeight(), 150);
  });
});
