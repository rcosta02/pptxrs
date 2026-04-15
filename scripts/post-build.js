#!/usr/bin/env node
/**
 * post-build.js — runs automatically after `npm run build` (wasm-pack).
 *
 * 1. Copies js/index.js  → pkg/index.js
 * 2. Copies js/index.d.ts → pkg/index.d.ts
 * 3. Patches pkg/package.json so that:
 *      "main"  → "index.js"
 *      "types" → "index.d.ts"
 *      "files" includes both new entries
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

// ── 3. Patch pkg/package.json ────────────────────────────────────────────────
const manifestPath = path.join(pkgDir, "package.json");
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));

manifest.main  = "index.js";
manifest.types = "index.d.ts";

manifest.files = Array.from(
  new Set([...(manifest.files ?? []), "index.js", "index.d.ts"]),
);

fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + "\n");
console.log("  patched pkg/package.json  (main, types, files)");

console.log("post-build done.");
