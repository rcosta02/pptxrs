'use strict';
/**
 * 03-import-existing.js
 *
 * Demonstrates importing an existing .pptx file:
 *
 *   A. Open a .pptx and inspect its structure (slide count, element types)
 *   B. Read all text elements across every slide
 *   C. Read images (base64) and save them to disk
 *   D. Convert an existing .pptx to a JSON snapshot
 *
 * The example first generates a "source" .pptx so it is self-contained —
 * swap `SOURCE` for any real .pptx path you have.
 *
 * Run:  node examples/03-import-existing.js
 * Out:  examples/out/03-*.pptx  +  examples/out/03-extracted/
 */

const { Presentation } = require('../pkg/index.js');
const fs   = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

// ── Generate a source deck to import ─────────────────────────────────────────
async function createSourceDeck() {
  const pres = new Presentation({ title: 'Source Deck', author: 'Alice' });

  pres.addSlide(null, slide => {
    slide.addText('Introduction', {
      x: 1, y: 1, w: 8, h: 1.2,
      fontSize: 40, bold: true, color: '002060',
    });
    slide.addText('This deck will be imported and inspected by example 03.', {
      x: 1, y: 2.5, w: 8, h: 1,
      fontSize: 18, color: '555555',
    });
    slide.addShape('rect', {
      x: 1, y: 4, w: 8, h: 0.7,
      fill: { color: '4472C4' },
    });
    slide.addNotes('Speaker note for slide 1.');
  });

  pres.addSlide(null, slide => {
    slide.addText('Data Table', {
      x: 0.5, y: 0.2, w: 9, h: 0.7,
      fontSize: 28, bold: true, color: '002060',
    });
    slide.addTable(
      [
        [{ text: 'Product', options: { bold: true, fill: '002060', color: 'FFFFFF' } },
         { text: 'Sales',   options: { bold: true, fill: '002060', color: 'FFFFFF' } }],
        ['Widget A', '1,200'],
        ['Widget B', '980'],
        ['Widget C', '2,100'],
      ],
      { x: 0.5, y: 1, w: 9, h: 4, colW: [6, 3], fontSize: 16 }
    );
  });

  pres.addSlide(null, slide => {
    slide.addText('Chart Slide', {
      x: 0.5, y: 0.1, w: 9, h: 0.6,
      fontSize: 28, bold: true, color: '002060',
    });
    slide.addChart('bar', [
      { name: 'Sales', labels: ['Jan','Feb','Mar'], values: [30, 50, 40] },
    ], { x: 0.5, y: 0.8, w: 9, h: 4.7, showLegend: true, showTitle: true, title: 'Monthly Sales' });
  });

  const sourcePath = path.join(OUT, '03-source.pptx');
  await pres.writeFile(sourcePath);
  console.log(`Source deck created: ${sourcePath}`);
  return sourcePath;
}

// ── A. Inspect structure ──────────────────────────────────────────────────────
function inspect(sourcePath) {
  const buf  = fs.readFileSync(sourcePath);
  const pres = Presentation.fromBuffer(buf);

  console.log('\n── Inspection ─────────────────────────────────────────────');
  console.log('  layout :', pres.layout);
  console.log('  title  :', pres.title  ?? '(none)');
  console.log('  author :', pres.author ?? '(none)');

  const slides = pres.getSlides();
  console.log(`  slides : ${slides.length}`);

  slides.forEach((slide, i) => {
    const elements = slide.getElements();
    const counts   = {};
    elements.forEach(el => { counts[el.type] = (counts[el.type] ?? 0) + 1; });
    console.log(`  slide ${i + 1}: ${JSON.stringify(counts)}`);
  });
}

// ── B. Extract all text ───────────────────────────────────────────────────────
function extractText(sourcePath) {
  const buf    = fs.readFileSync(sourcePath);
  const pres   = Presentation.fromBuffer(buf);
  const slides = pres.getSlides();

  console.log('\n── Extracted text ─────────────────────────────────────────');
  slides.forEach((slide, i) => {
    console.log(`  Slide ${i + 1}:`);
    slide.getElements().forEach(el => {
      if (el.type === 'text') {
        const str = typeof el.text === 'string'
          ? el.text
          : el.text.map(r => r.text).join('');
        console.log(`    [text] "${str.slice(0, 80)}${str.length > 80 ? '…' : ''}"`);
        console.log(`           at (${el.options.x}, ${el.options.y})  ${el.options.w}×${el.options.h} in`);
      }
      if (el.type === 'notes') {
        console.log(`    [notes] "${el.text}"`);
      }
    });
  });
}

// ── C. Extract images ─────────────────────────────────────────────────────────
function extractImages(sourcePath) {
  const buf    = fs.readFileSync(sourcePath);
  const pres   = Presentation.fromBuffer(buf);
  const slides = pres.getSlides();
  const imgDir = path.join(OUT, '03-extracted-images');
  fs.mkdirSync(imgDir, { recursive: true });

  let count = 0;
  slides.forEach((slide, si) => {
    slide.getElements().forEach((el, ei) => {
      if (el.type === 'image' && el.options.data) {
        const imgBytes = Buffer.from(el.options.data, 'base64');
        // Detect extension from magic bytes
        let ext = 'png';
        if (imgBytes[0] === 0xFF && imgBytes[1] === 0xD8) ext = 'jpg';
        if (imgBytes.slice(0, 3).toString() === 'GIF') ext = 'gif';

        const fname = `slide${si + 1}-image${ei + 1}.${ext}`;
        fs.writeFileSync(path.join(imgDir, fname), imgBytes);
        console.log(`  Extracted: ${fname}  (${imgBytes.length} bytes)`);
        count++;
      }
    });
  });

  if (count === 0) {
    console.log('\n── Image extraction ───────────────────────────────────────');
    console.log('  (no embedded images in this deck)');
  } else {
    console.log(`\n── Image extraction ───────────────────────────────────────`);
    console.log(`  Saved ${count} image(s) to ${imgDir}`);
  }
}

// ── D. Export to JSON snapshot ────────────────────────────────────────────────
function exportToJson(sourcePath) {
  const buf    = fs.readFileSync(sourcePath);
  const pres   = Presentation.fromBuffer(buf);
  const json   = pres.toJson();

  const outPath = path.join(OUT, '03-snapshot.json');
  fs.writeFileSync(outPath, JSON.stringify(json, null, 2));

  console.log('\n── JSON snapshot ──────────────────────────────────────────');
  console.log(`  Saved: ${outPath}`);
  console.log(`  Slides: ${json.slides.length}`);
  json.slides.forEach((s, i) => {
    console.log(`  Slide ${i + 1}: ${s.elements.length} elements`);
  });
}

async function main() {
  const sourcePath = await createSourceDeck();
  inspect(sourcePath);
  extractText(sourcePath);
  extractImages(sourcePath);
  exportToJson(sourcePath);
}

main().catch(err => { console.error(err); process.exit(1); });
