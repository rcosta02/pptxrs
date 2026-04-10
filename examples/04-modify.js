'use strict';
/**
 * 04-modify.js
 *
 * Demonstrates modifying an imported .pptx file:
 *
 *   A. Add a watermark to every slide
 *   B. Append a new slide to an existing deck
 *   C. Prepend a title/cover slide
 *   D. Add a footer bar to every slide
 *   E. Replace text content (find & replace across all slides)
 *   F. Remove a slide by index
 *   G. Reorder slides
 *
 * Run:  node examples/04-modify.js
 * Out:  examples/out/04-*.pptx
 */

const { Presentation } = require('../pkg/index.js');
const fs   = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

// ── Generate a source deck ────────────────────────────────────────────────────
async function createSource() {
  const pres = new Presentation({ title: 'Original Deck' });

  ['Alpha', 'Beta', 'Gamma', 'Delta'].forEach((name, i) => {
    pres.addSlide(null, slide => {
      slide.setBackground(i % 2 === 0 ? 'FFFFFF' : 'F0F4FF');
      slide.addText(`Section: ${name}`, {
        x: 1, y: 1.5, w: 8, h: 1.2,
        fontSize: 36, bold: true, color: '002060',
      });
      slide.addText(`Slide ${i + 1} of 4 — original content.`, {
        x: 1, y: 3, w: 8, h: 0.7,
        fontSize: 16, color: '555555',
      });
    });
  });

  const src = path.join(OUT, '04-source.pptx');
  await pres.writeFile(src);
  console.log(`Source created: ${src}`);
  return src;
}

// ── A. Watermark every slide ──────────────────────────────────────────────────
async function watermark(srcPath) {
  const pres   = Presentation.fromBuffer(fs.readFileSync(srcPath));
  const slides = pres.getSlides();

  slides.forEach((slide, i) => {
    slide.addText('CONFIDENTIAL', {
      x: 1.5, y: 1.5, w: 7, h: 3,
      fontSize:     72,
      bold:         true,
      color:        'FF0000',
      transparency: 75,
      rotate:       45,
      align:        'center',
      valign:       'middle',
      isTextBox:    true,
    });
    pres.syncSlide(i, slide);
  });

  const out = path.join(OUT, '04-watermarked.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}`);
}

// ── B. Append a new slide ─────────────────────────────────────────────────────
async function appendSlide(srcPath) {
  const pres   = Presentation.fromBuffer(fs.readFileSync(srcPath));
  const count  = pres.getSlides().length;

  // addSlide with callback auto-appends and syncs
  pres.addSlide(null, slide => {
    slide.setBackground('002060');
    slide.addText('Thank You', {
      x: 1, y: 1.8, w: 8, h: 1.5,
      fontSize: 52, bold: true, color: 'FFFFFF', align: 'center',
    });
    slide.addText('Questions?', {
      x: 1, y: 3.5, w: 8, h: 0.8,
      fontSize: 24, color: 'C0D0F0', align: 'center',
    });
  });
  pres.syncSlide(count, pres.getSlides()[pres.getSlides().length - 1]
    ?? (() => { throw new Error('sync failed'); })());

  // Simpler pattern: rebuild the slide list
  const newSlide = pres.addSlide();
  newSlide.setBackground('002060');
  newSlide.addText('Thank You', {
    x: 1, y: 1.8, w: 8, h: 1.5,
    fontSize: 52, bold: true, color: 'FFFFFF', align: 'center',
  });
  pres.syncSlide(count, newSlide);

  const out = path.join(OUT, '04-appended.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}  (${pres.getSlides().length} slides)`);
}

// ── C. Prepend a cover slide ──────────────────────────────────────────────────
async function prependCover(srcPath) {
  const pres   = Presentation.fromBuffer(fs.readFileSync(srcPath));
  const slides = pres.getSlides();

  // Build the cover
  const cover = pres.addSlide();
  cover.setBackground('002060');
  cover.addText('Q3 Business Review', {
    x: 1, y: 1.5, w: 8, h: 1.4,
    fontSize: 44, bold: true, color: 'FFFFFF', align: 'center',
  });
  cover.addText('Prepared for the Board · October 2024', {
    x: 1, y: 3.2, w: 8, h: 0.6,
    fontSize: 16, color: '8090B0', align: 'center', italic: true,
  });

  // Rebuild: cover first, then existing slides
  pres.syncSlide(0, cover);
  slides.forEach((s, i) => pres.syncSlide(i + 1, s));

  const out = path.join(OUT, '04-prepended-cover.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}  (${pres.getSlides().length} slides)`);
}

// ── D. Add footer bar to every slide ─────────────────────────────────────────
async function addFooter(srcPath) {
  const pres      = Presentation.fromBuffer(fs.readFileSync(srcPath));
  const slides    = pres.getSlides();
  const slideH    = 5.625; // inches for 16:9
  const footerH   = 0.32;
  const footerY   = slideH - footerH;

  slides.forEach((slide, i) => {
    // Dark footer bar
    slide.addShape('rect', {
      x: 0, y: footerY, w: 10, h: footerH,
      fill: { color: '002060' },
    });
    // Company name left
    slide.addText('Acme Corp · Confidential', {
      x: 0.15, y: footerY, w: 5, h: footerH,
      fontSize: 9, color: 'AABBCC', valign: 'middle',
    });
    // Slide number right
    slide.addText(`${i + 1} / ${slides.length}`, {
      x: 5, y: footerY, w: 4.8, h: footerH,
      fontSize: 9, color: 'AABBCC', align: 'right', valign: 'middle',
    });
    pres.syncSlide(i, slide);
  });

  const out = path.join(OUT, '04-with-footer.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}`);
}

// ── E. Find & replace text across all slides ──────────────────────────────────
async function findAndReplace(srcPath, find, replace) {
  const pres   = Presentation.fromBuffer(fs.readFileSync(srcPath));
  const slides = pres.getSlides();
  let   hits   = 0;

  slides.forEach((slide, si) => {
    let changed = false;

    slide.getElements().forEach(el => {
      if (el.type !== 'text') return;

      if (typeof el.text === 'string' && el.text.includes(find)) {
        el.text = el.text.replaceAll(find, replace);
        changed = true;
        hits++;
      }

      if (Array.isArray(el.text)) {
        el.text.forEach(run => {
          if (run.text.includes(find)) {
            run.text = run.text.replaceAll(find, replace);
            changed = true;
            hits++;
          }
        });
      }
    });

    if (changed) pres.syncSlide(si, slide);
  });

  const out = path.join(OUT, '04-find-replaced.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}  (${hits} replacement(s): "${find}" → "${replace}")`);
}

// ── F. Remove a slide ────────────────────────────────────────────────────────
async function removeSlide(srcPath, index) {
  const pres = Presentation.fromBuffer(fs.readFileSync(srcPath));
  pres.removeSlide(index);

  const out = path.join(OUT, '04-removed-slide.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}  (removed slide ${index + 1}, now ${pres.getSlides().length} slides)`);
}

// ── G. Reorder slides ────────────────────────────────────────────────────────
async function reorderSlides(srcPath) {
  const pres   = Presentation.fromBuffer(fs.readFileSync(srcPath));
  const slides = pres.getSlides();

  // Reverse the order
  const reversed = [...slides].reverse();
  reversed.forEach((s, i) => pres.syncSlide(i, s));

  const out = path.join(OUT, '04-reordered.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}  (slides reversed)`);
}

async function main() {
  const src = await createSource();

  await watermark(src);
  await appendSlide(src);
  await prependCover(src);
  await addFooter(src);
  await findAndReplace(src, 'original content', 'updated content ✓');
  await removeSlide(src, 1);   // remove second slide (0-indexed)
  await reorderSlides(src);
}

main().catch(err => { console.error(err); process.exit(1); });
