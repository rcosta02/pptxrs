/**
 * Table element tests
 *
 * Covers: addTable() — round-trip survival, cell data, column widths, position.
 *
 * Run: node --test tests/tables.test.js
 */

"use strict";

const { test, describe } = require("node:test");
const assert = require("node:assert/strict");

const { oneSlide, firstElement } = require("./helpers.js");


describe("Tables – addTable()", () => {
  test("table survives round-trip", () => {
    const pres = oneSlide((s) =>
      s.addTable([["A", "B"], ["1", "2"]], { x: 96, y: 96, w: 480, h: 144 }),
    );
    const el = firstElement(pres);
    assert.equal(el.elementType, "table");
  });

  test("table cell data is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addTable(
        [["Name", "Score"], ["Alice", "100"], ["Bob", "95"]],
        { x: 96, y: 96, w: 480, h: 144 },
      ),
    );
    const data = firstElement(pres).toJson().data;
    assert.equal(data.length, 3, "3 rows");
    assert.deepEqual(data[0], ["Name", "Score"]);
    assert.deepEqual(data[1], ["Alice", "100"]);
    assert.deepEqual(data[2], ["Bob", "95"]);
  });

  test("colW is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addTable([["A", "B"]], {
        x: 96, y: 96, w: 480, h: 96,
        colW: [240, 240],
      }),
    );
    const opts = firstElement(pres).toJson().options;
    assert.ok(Array.isArray(opts.colW), "colW should be an array");
    assert.equal(opts.colW.length, 2);
    assert.ok(Math.abs(opts.colW[0] - 240) < 1, `colW[0] should be ~240, got ${opts.colW[0]}`);
    assert.ok(Math.abs(opts.colW[1] - 240) < 1, `colW[1] should be ~240, got ${opts.colW[1]}`);
  });

  test("table position (x/y/w/h) is preserved in round-trip", () => {
    const pres = oneSlide((s) =>
      s.addTable([["X"]], { x: 50, y: 80, w: 300, h: 100 }),
    );
    const opts = firstElement(pres).toJson().options;
    assert.ok(Math.abs(opts.x - 50)  < 1, `x should be ~50, got ${opts.x}`);
    assert.ok(Math.abs(opts.y - 80)  < 1, `y should be ~80, got ${opts.y}`);
    assert.ok(Math.abs(opts.w - 300) < 1, `w should be ~300, got ${opts.w}`);
    assert.ok(Math.abs(opts.h - 100) < 1, `h should be ~100, got ${opts.h}`);
  });
});
