'use strict';
/**
 * 06-charts.js
 *
 * One slide per chart type with realistic data and fully configured options:
 *   bar / column, stacked bar, line, area, pie, doughnut,
 *   radar, scatter, bubble, and a bar+line combo chart.
 *
 * Run:  node examples/06-charts.js
 * Out:  examples/out/06-charts.pptx
 */

const { Presentation } = require('../pkg/index.js');
const fs   = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

const QUARTERS = ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'];
const MONTHS   = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
const REGIONS  = ['North', 'South', 'East', 'West', 'Central'];

async function main() {
  const pres = new Presentation({ title: 'Chart Showcase', layout: 'LAYOUT_16x9' });

  // ── Clustered column ────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Clustered Column Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('bar', [
      { name: 'Revenue',  labels: QUARTERS, values: [120, 190, 160, 230] },
      { name: 'Expenses', labels: QUARTERS, values: [80,  110,  90, 130] },
      { name: 'Profit',   labels: QUARTERS, values: [40,   80,  70, 100] },
    ], {
      x: 0.3, y: 0.7, w: 9.4, h: 4.9,
      barDir:          'col',
      barGrouping:     'clustered',
      showTitle:       false,
      showLegend:      true,
      legendPos:       'b',
      showValue:       true,
      dataLabelPosition: 'outEnd',
      dataLabelFontSize: 9,
      chartColors:     ['4472C4', 'ED7D31', 'A9D18E'],
      catAxisTitle:    'Quarter',
      valAxisTitle:    'USD (thousands)',
      valAxisMaxVal:   260,
      valAxisMajorUnit: 50,
    });
  });

  // ── Stacked bar ─────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Stacked Bar Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('bar', [
      { name: 'Product A', labels: REGIONS, values: [30, 45, 20, 55, 40] },
      { name: 'Product B', labels: REGIONS, values: [25, 30, 35, 20, 45] },
      { name: 'Product C', labels: REGIONS, values: [15, 20, 25, 30, 10] },
    ], {
      x: 0.3, y: 0.7, w: 9.4, h: 4.9,
      barDir:       'bar',           // horizontal bars
      barGrouping:  'stacked',
      showLegend:   true,
      legendPos:    'b',
      showPercent:  false,
      chartColors:  ['4472C4', 'ED7D31', 'A9D18E'],
      catAxisTitle: 'Region',
      valAxisTitle: 'Units sold',
    });
  });

  // ── 100% stacked ────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('100% Stacked Column', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('bar', [
      { name: 'Direct',     labels: QUARTERS, values: [40, 50, 45, 55] },
      { name: 'Partner',    labels: QUARTERS, values: [35, 30, 40, 25] },
      { name: 'Digital',    labels: QUARTERS, values: [25, 20, 15, 20] },
    ], {
      x: 0.3, y: 0.7, w: 9.4, h: 4.9,
      barDir:       'col',
      barGrouping:  'percentStacked',
      showLegend:   true,
      legendPos:    'b',
      showPercent:  true,
      chartColors:  ['4472C4', 'ED7D31', 'A9D18E'],
    });
  });

  // ── Line chart ──────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Line Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('line', [
      { name: 'NPS Score',   labels: MONTHS, values: [62, 65, 70, 68, 75, 80] },
      { name: 'Satisfaction', labels: MONTHS, values: [55, 58, 62, 60, 70, 74] },
      { name: 'Retention',   labels: MONTHS, values: [88, 90, 87, 91, 93, 95] },
    ], {
      x: 0.3, y: 0.7, w: 9.4, h: 4.9,
      lineSmooth:          true,
      lineSize:            2.5,
      lineDataSymbol:      'circle',
      lineDataSymbolSize:  7,
      showTitle:           false,
      showLegend:          true,
      legendPos:           'b',
      showValue:           false,
      chartColors:         ['4472C4', 'ED7D31', '70AD47'],
      catAxisTitle:        'Month',
      valAxisTitle:        'Score',
      valAxisMinVal:       50,
      valAxisMaxVal:       100,
    });
  });

  // ── Area chart ──────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Area Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('area', [
      { name: 'Cloud',    labels: MONTHS, values: [10, 18, 25, 30, 38, 48] },
      { name: 'On-Prem',  labels: MONTHS, values: [45, 43, 40, 38, 35, 30] },
    ], {
      x: 0.3, y: 0.7, w: 9.4, h: 4.9,
      showLegend:   true,
      legendPos:    'b',
      chartColors:  ['4472C4', 'ED7D31'],
      chartColorsOpacity: 60,
      catAxisTitle: 'Month',
      valAxisTitle: 'Deployments',
    });
  });

  // ── Pie chart ───────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Pie Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('pie', [
      {
        labels: ['Engineering', 'Sales', 'Marketing', 'Support', 'G&A'],
        values: [35, 25, 18, 12, 10],
      },
    ], {
      x: 1, y: 0.7, w: 8, h: 4.9,
      showTitle:     true,
      title:         'Headcount by Department',
      showLegend:    true,
      legendPos:     'r',
      showPercent:   true,
      showLabel:     true,
      dataLabelPosition: 'bestFit',
      chartColors:   ['4472C4', 'ED7D31', 'A9D18E', 'FFC000', '5B9BD5'],
    });
  });

  // ── Doughnut ────────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Doughnut Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('doughnut', [
      {
        labels: ['North America', 'EMEA', 'APAC', 'LATAM'],
        values: [42, 28, 22, 8],
      },
    ], {
      x: 1, y: 0.7, w: 8, h: 4.9,
      showTitle:   true,
      title:       'Revenue by Region',
      showLegend:  true,
      legendPos:   'b',
      showPercent: true,
      chartColors: ['002060', '4472C4', '8FAADC', 'D6E4F7'],
    });
  });

  // ── Radar chart ─────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Radar Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('radar', [
      {
        name:   'Product X',
        labels: ['Performance', 'Reliability', 'Usability', 'Features', 'Support', 'Price'],
        values: [80, 90, 70, 85, 75, 65],
      },
      {
        name:   'Competitor Y',
        labels: ['Performance', 'Reliability', 'Usability', 'Features', 'Support', 'Price'],
        values: [70, 75, 85, 70, 80, 90],
      },
    ], {
      x: 1, y: 0.7, w: 8, h: 4.9,
      showTitle:   true,
      title:       'Competitive Analysis',
      showLegend:  true,
      legendPos:   'b',
      chartColors: ['4472C4', 'ED7D31'],
    });
  });

  // ── Scatter chart ────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Scatter Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('scatter', [
      {
        name:   'Group A',
        labels: ['1','2','3','4','5','6','7'],
        values: [10, 22, 18, 35, 28, 42, 38],
      },
      {
        name:   'Group B',
        labels: ['1','2','3','4','5','6','7'],
        values: [5,  15, 25, 20, 45, 30, 50],
      },
    ], {
      x: 0.3, y: 0.7, w: 9.4, h: 4.9,
      showLegend:   true,
      legendPos:    'b',
      chartColors:  ['4472C4', 'ED7D31'],
      catAxisTitle: 'X Axis',
      valAxisTitle: 'Y Axis',
    });
  });

  // ── Bubble chart ─────────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Bubble Chart', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addChart('bubble', [
      {
        name:   'Segment A',
        labels: ['P1','P2','P3','P4'],
        values: [10, 30, 50, 20],
        sizes:  [15, 30, 25, 10],
      },
      {
        name:   'Segment B',
        labels: ['P1','P2','P3'],
        values: [20, 40, 15],
        sizes:  [20, 10, 35],
      },
    ], {
      x: 0.3, y: 0.7, w: 9.4, h: 4.9,
      showLegend:   true,
      legendPos:    'b',
      chartColors:  ['4472C4', 'ED7D31'],
      catAxisTitle: 'Market Size',
      valAxisTitle: 'Growth Rate',
    });
  });

  // ── Combo: bar + line ────────────────────────────────────────────────────────
  pres.addSlide(null, slide => {
    slide.addText('Combo Chart  (Bar + Line)', {
      x: 0.3, y: 0.1, w: 9.4, h: 0.55, fontSize: 22, bold: true, color: '002060',
    });
    slide.addComboChart(
      ['bar', 'line'],
      [
        // Data for 'bar'
        [{ name: 'Revenue', labels: QUARTERS, values: [120, 190, 160, 230] }],
        // Data for 'line' — plotted on secondary axis
        [{ name: 'Growth %', labels: QUARTERS, values: [5, 8, 6, 10] }],
      ],
      {
        x: 0.3, y: 0.7, w: 9.4, h: 4.9,
        barDir:           'col',
        barGrouping:      'clustered',
        showLegend:       true,
        legendPos:        'b',
        showValue:        false,
        secondaryValAxis: true,
        chartColors:      ['4472C4', 'ED7D31'],
        lineSize:         2.5,
        lineDataSymbol:   'diamond',
        catAxisTitle:     'Quarter',
        valAxisTitle:     'Revenue (USD k)',
      }
    );
  });

  const out = path.join(OUT, '06-charts.pptx');
  await pres.writeFile(out);
  console.log(`Written: ${out}  (${pres.getSlides().length} slides, one per chart type)`);
}

main().catch(err => { console.error(err); process.exit(1); });
