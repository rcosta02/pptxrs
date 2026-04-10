'use strict';
/**
 * 02-from-json.js
 *
 * Demonstrates the JSON interchange workflow:
 *
 *   A. Build a presentation in code → toJson() → store JSON
 *   B. Load JSON → fromJson() → export .pptx
 *   C. Build directly from a hand-crafted JSON object (no Presentation API needed)
 *   D. Use JSON as a reusable template, injecting dynamic data
 *
 * Run:  node examples/02-from-json.js
 * Out:  examples/out/02-*.pptx  +  examples/out/02-deck.json
 */

const { Presentation } = require('../pkg/index.js');
const fs   = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

// ── A. Build → toJson() → save JSON ──────────────────────────────────────────
function buildAndSerialize() {
  const pres = new Presentation({ title: 'JSON Demo', layout: 'LAYOUT_16x9' });

  pres.addSlide(null, slide => {
    slide.addText('Serialized to JSON', {
      x: 1, y: 1, w: 8, h: 1.2,
      fontSize: 36, bold: true, color: '002060',
    });
    slide.addText('This presentation was built in code, serialized to JSON, then rebuilt.', {
      x: 1, y: 2.5, w: 8, h: 1.5,
      fontSize: 18, color: '444444', wrap: true,
    });
    slide.addShape('rect', {
      x: 1, y: 4.2, w: 8, h: 0.8,
      fill: { color: '4472C4' },
    });
  });

  pres.addSlide(null, slide => {
    slide.addTable(
      [
        [
          { text: 'Key',   options: { bold: true, fill: '002060', color: 'FFFFFF' } },
          { text: 'Value', options: { bold: true, fill: '002060', color: 'FFFFFF' } },
        ],
        ['format',  'OOXML / .pptx'],
        ['runtime', 'Rust + WASM'],
        ['target',  'Node.js ≥ 16'],
      ],
      { x: 1, y: 1, w: 8, h: 3.5, colW: [4, 4], fontSize: 16 }
    );
  });

  // Save JSON alongside the output
  const json    = pres.toJson();
  const jsonStr = pres.toJsonString(); // pretty-printed not built-in, use JSON.stringify
  fs.writeFileSync(
    path.join(OUT, '02-deck.json'),
    JSON.stringify(json, null, 2)
  );
  console.log('Saved: 02-deck.json');

  return json;
}

// ── B. fromJson() → export .pptx ─────────────────────────────────────────────
async function rebuildFromJson(json) {
  const pres = Presentation.fromJson(json);
  const out  = path.join(OUT, '02-rebuilt-from-json.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}`);
}

// ── C. Hand-crafted JSON object → Presentation ───────────────────────────────
async function fromHandcraftedJson() {
  /**
   * You can construct a PresentationJson object directly without ever
   * using the Presentation API, then hydrate it into a real .pptx.
   * Useful when your data is already in JSON (e.g. from a database or CMS).
   */
  const json = {
    meta: {
      title:  'Hand-crafted JSON',
      layout: 'LAYOUT_16x9',
    },
    slides: [
      {
        background: { color: 'F5F5F5' },
        elements: [
          {
            type: 'text',
            text: 'Built directly from a JSON object',
            options: { x: 1, y: 1, w: 8, h: 1.2, fontSize: 32, bold: true, color: '002060' },
          },
          {
            type: 'text',
            text: 'No Presentation API calls were used to create this slide.',
            options: { x: 1, y: 2.5, w: 8, h: 1, fontSize: 18, color: '555555' },
          },
          {
            type: 'shape',
            shapeType: 'ellipse',
            options: { x: 4, y: 3.8, w: 2, h: 1.5, fill: { color: 'ED7D31' } },
          },
        ],
      },
      {
        elements: [
          {
            type: 'text',
            text: [
              { text: 'Mixed ', options: {} },
              { text: 'runs ', options: { bold: true, color: '4472C4' } },
              { text: 'also work', options: { italic: true, color: 'ED7D31' } },
            ],
            options: { x: 1, y: 1, w: 8, h: 1, fontSize: 28 },
          },
          {
            type: 'table',
            data: [
              ['Name', 'Score'],
              ['Alice', '95'],
              ['Bob',   '87'],
            ],
            options: { x: 1, y: 2.5, w: 8, h: 2.5, colW: [6, 2], fontSize: 16 },
          },
        ],
      },
    ],
  };

  const pres = Presentation.fromJson(json);
  const out  = path.join(OUT, '02-from-handcrafted-json.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}`);
}

// ── D. JSON as a reusable template ────────────────────────────────────────────
async function jsonTemplate() {
  /**
   * Define a slide template as a plain JS function that returns a PresentationJson.
   * Inject dynamic data — works great for reports, invoices, status updates, etc.
   */
  function buildReportJson({ title, author, rows, quarter }) {
    return {
      meta: { title, layout: 'LAYOUT_16x9' },
      slides: [
        // Cover slide
        {
          background: { color: '002060' },
          elements: [
            {
              type: 'text',
              text: title,
              options: { x: 1, y: 1.5, w: 8, h: 1.5, fontSize: 44, bold: true, color: 'FFFFFF', align: 'center' },
            },
            {
              type: 'text',
              text: `Prepared by ${author} · ${quarter}`,
              options: { x: 1, y: 3.2, w: 8, h: 0.6, fontSize: 16, color: '8090B0', align: 'center', italic: true },
            },
          ],
        },
        // Data table slide
        {
          elements: [
            {
              type: 'text',
              text: 'Results',
              options: { x: 0.5, y: 0.2, w: 9, h: 0.7, fontSize: 28, bold: true, color: '002060' },
            },
            {
              type: 'table',
              data: [
                [
                  { text: 'Region',  options: { bold: true, fill: '002060', color: 'FFFFFF', align: 'center' } },
                  { text: 'Target',  options: { bold: true, fill: '002060', color: 'FFFFFF', align: 'center' } },
                  { text: 'Actual',  options: { bold: true, fill: '002060', color: 'FFFFFF', align: 'center' } },
                  { text: 'Delta',   options: { bold: true, fill: '002060', color: 'FFFFFF', align: 'center' } },
                ],
                ...rows.map(r => [
                  r.region,
                  `$${r.target}k`,
                  `$${r.actual}k`,
                  {
                    text: `${r.actual >= r.target ? '+' : ''}${r.actual - r.target}k`,
                    options: { color: r.actual >= r.target ? '00B050' : 'FF0000', bold: true },
                  },
                ]),
              ],
              options: { x: 0.5, y: 1, w: 9, h: 4.4, colW: [3, 2, 2, 2], fontSize: 14 },
            },
          ],
        },
        // Chart slide
        {
          elements: [
            {
              type: 'text',
              text: `${quarter} Performance`,
              options: { x: 0.5, y: 0.1, w: 9, h: 0.6, fontSize: 24, bold: true, color: '002060' },
            },
            {
              type: 'chart',
              chartType: 'bar',
              data: [
                { name: 'Target', labels: rows.map(r => r.region), values: rows.map(r => r.target) },
                { name: 'Actual', labels: rows.map(r => r.region), values: rows.map(r => r.actual) },
              ],
              options: {
                x: 0.5, y: 0.8, w: 9, h: 4.7,
                barDir: 'col', barGrouping: 'clustered',
                showLegend: true, legendPos: 'b',
                showValue:  true,
                chartColors: ['4472C4', 'ED7D31'],
              },
            },
          ],
        },
      ],
    };
  }

  const reportData = {
    title:   'Q3 Sales Report',
    author:  'Sales Team',
    quarter: 'Q3 2024',
    rows: [
      { region: 'North',  target: 120, actual: 138 },
      { region: 'South',  target: 100, actual:  92 },
      { region: 'East',   target:  90, actual:  97 },
      { region: 'West',   target: 110, actual: 125 },
    ],
  };

  const pres = Presentation.fromJson(buildReportJson(reportData));
  const out  = path.join(OUT, '02-template-report.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}`);
}

// ── E. Load a saved JSON file and export ──────────────────────────────────────
async function loadFromDisk() {
  const jsonPath = path.join(OUT, '02-deck.json');
  if (!fs.existsSync(jsonPath)) {
    console.log('Skipping load-from-disk (run again after first pass).');
    return;
  }
  const json = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
  const pres = Presentation.fromJson(json);
  const out  = path.join(OUT, '02-loaded-from-disk.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}`);
}

async function main() {
  const json = buildAndSerialize();
  await rebuildFromJson(json);
  await fromHandcraftedJson();
  await jsonTemplate();
  await loadFromDisk();
}

main().catch(err => { console.error(err); process.exit(1); });
