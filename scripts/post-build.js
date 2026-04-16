#!/usr/bin/env node
/**
 * post-build.js — runs automatically after `npm run build` (wasm-pack).
 *
 * 1. Copies js/index.js  → pkg/index.js
 * 2. Copies js/index.d.ts → pkg/index.d.ts
 * 3. Rewrites pkg/package.json to match the publishing template.
 */

"use strict";

const fs   = require("fs");
const path = require("path");

const root   = path.resolve(__dirname, "..");
const jsDir  = path.join(root, "js");
const pkgDir = path.join(root, "pkg");

// ── 1 & 2. Copy index files ──────────────────────────────────────────────────
for (const name of ["index.js", "index.d.ts"]) {
  const src  = path.join(jsDir, name);
  const dest = path.join(pkgDir, name);
  fs.copyFileSync(src, dest);
  console.log(`  copied  js/${name}  →  pkg/${name}`);
}

// ── 3. Rewrite pkg/package.json ──────────────────────────────────────────────
const manifestPath = path.join(pkgDir, "package.json");
const generated = JSON.parse(fs.readFileSync(manifestPath, "utf8"));

const manifest = {
  name:        "pptxrs",
  description: "Create, read, modify, and export .pptx files — Rust/WASM npm library for Node.js",
  version:     generated.version ?? "0.1.8",
  license:     "MIT",
  author: {
    name:  "Rafael Costa",
    email: "dev.rcosta@gmail.com",
    url:   "https://github.com/rcosta02",
  },
  repository: {
    type: "git",
    url:  "https://github.com/rcosta02/pptxrs",
  },
  homepage: "https://github.com/rcosta02/pptxrs#readme",
  bugs: {
    url: "https://github.com/rcosta02/pptxrs/issues",
  },
  keywords: [
    "pptx", "powerpoint", "presentation", "office",
    "rust", "wasm", "webassembly", "nodejs",
    "ppt", "slides", "openxml", "ooxml",
  ],
  files: [
    "pptxrs_bg.wasm",
    "pptxrs.js",
    "pptxrs.d.ts",
    "index.js",
    "index.d.ts",
  ],
  main:  "index.js",
  types: "index.d.ts",
};

fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + "\n");
console.log("  wrote   pkg/package.json");

console.log("post-build done.");
