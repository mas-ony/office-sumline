import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { SpreadsheetDocument } from '../../src/core/document';
import { chartFromRange, guessLayout, layoutOfChart, readChartData, seriesForLayout } from '../../src/core/chartData';
import { CHART_TYPES, CHARTEX_TYPES, type ChartType } from '../../src/core/model/drawing';
import { makeRange, parseRange } from '../../src/core/refs';
import { readXlsx } from '../../src/core/xlsx/import';
import { parseXml } from '../../src/core/xml/xml';
import { chartSvg, fitTrend, histogramBins, niceScale, quartiles, squarify } from '../../src/grid/chartRender';
import { SHAPE_GALLERY, shapeSvg } from '../../src/grid/shapes';
import { readFixture } from './helpers';

function salesDoc(): SpreadsheetDocument {
  const d = SpreadsheetDocument.blank();
  d.setValues(0, 0, [
    ['Region', 'Q1', 'Q2', 'Q3'],
    ['North', 120, 150, 90],
    ['South', 90, 110, 130],
    ['East', 60, 80, 100],
  ]);
  return d;
}

const rect = { x: 500, y: 0, width: 400, height: 300 };

function addChart(d: SpreadsheetDocument, type: ChartType, ref = 'A1:D4') {
  return d.addChart({ ...chartFromRange(d, d.sheet, parseRange(ref)!, type), name: 'Chart 1' }, rect)!;
}

function roundTrip(d: SpreadsheetDocument) {
  const bytes = d.save();
  return { bytes, files: unzipSync(bytes), wb: readXlsx(bytes) };
}

describe('chart data', () => {
  it('builds one series per row when there are as many rows as columns (headers on both sides)', () => {
    const d = salesDoc();
    const def = chartFromRange(d, d.sheet, makeRange(0, 0, 3, 3), 'column');
    expect(def.seriesIn).toBe('rows');
    expect(def.headers).toBe(true);
    expect(def.categories).toBe('Sheet1!$B$1:$D$1');
    expect(def.series.map((s) => [s.nameRef, s.values])).toEqual([
      ['Sheet1!$A$2', 'Sheet1!$B$2:$D$2'],
      ['Sheet1!$A$3', 'Sheet1!$B$3:$D$3'],
      ['Sheet1!$A$4', 'Sheet1!$B$4:$D$4'],
    ]);
  });

  it('uses one series per column for tall ranges and reads current values', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['Month', 'Sales', 'Cost'], ['Jan', 10, 4], ['Feb', 12, 5], ['Mar', 9, 6], ['Apr', 14, 7]]);
    const chart = d.addChart({ ...chartFromRange(d, d.sheet, makeRange(0, 0, 4, 2), 'line'), name: 'c' }, rect)!;
    expect(chart.seriesIn).toBe('columns');
    const data = readChartData(d, chart);
    expect(data.categories).toEqual(['Jan', 'Feb', 'Mar', 'Apr']);
    expect(data.series.map((s) => s.name)).toEqual(['Sales', 'Cost']);
    expect(data.series[0].values).toEqual([10, 12, 9, 14]);
    // Values follow edits.
    d.setInput(1, 1, '=5*5');
    expect(readChartData(d, chart).series[0].values[0]).toBe(25);
  });

  it('takes the first column as x values for scatter charts', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [['x', 'y'], [1, 3], [2, 5], [3, 4]]);
    const def = chartFromRange(d, d.sheet, makeRange(0, 0, 3, 1), 'scatter');
    expect(def.series).toHaveLength(1);
    expect(def.series[0]).toMatchObject({ values: 'Sheet1!$B$2:$B$4', xValues: 'Sheet1!$A$2:$A$4', nameRef: 'Sheet1!$B$1' });
  });

  it('recovers the layout of a chart for its settings', () => {
    const d = salesDoc();
    const chart = addChart(d, 'column');
    expect(layoutOfChart(d, chart)).toEqual({ seriesIn: 'rows', headers: true, labels: 0, series: [1, 2, 3] });
    const vertical = seriesForLayout(d.sheet, makeRange(0, 0, 3, 3), { seriesIn: 'columns', headers: true, labels: 0, series: [1, 3] }, 'line');
    expect(vertical.categories).toBe('Sheet1!$A$2:$A$4');
    expect(vertical.series.map((s) => s.values)).toEqual(['Sheet1!$B$2:$B$4', 'Sheet1!$D$2:$D$4']);
  });

  it('guesses headers only when the first row holds text', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[1, 2], [3, 4], [5, 6]]);
    const layout = guessLayout(d, d.sheet, makeRange(0, 0, 2, 1), 'column');
    expect(layout).toMatchObject({ seriesIn: 'columns', headers: false, labels: null, series: [0, 1] });
  });

  it('keeps chart references in step with inserted rows and renamed sheets', () => {
    const d = salesDoc();
    const chart = addChart(d, 'column');
    const row = chart.from.r;
    d.insertRows(0, 2);
    d.renameSheet(0, 'Data');
    const moved = d.sheet.charts.find((c) => c.id === chart.id)!;
    expect(moved.series[0].values).toBe('Data!$B$4:$D$4');
    expect(moved.categories).toBe('Data!$B$3:$D$3');
    expect(moved.from.r).toBe(row + 2);
    expect(readChartData(d, moved).series[0].values).toEqual([120, 150, 90]);
  });

  it('is undone with the rest of the edit', () => {
    const d = salesDoc();
    addChart(d, 'pie');
    expect(d.sheet.charts).toHaveLength(1);
    d.undo();
    expect(d.sheet.charts).toHaveLength(0);
    d.redo();
    expect(d.sheet.charts[0].type).toBe('pie');
  });
});

describe('chart rendering', () => {
  it('draws every chart type', () => {
    const d = salesDoc();
    for (const type of CHART_TYPES) {
      const chart = addChart(d, type);
      const svg = chartSvg({ ...chart, title: 'T', valueLabels: true }, readChartData(d, chart), 400, 300);
      expect(svg.startsWith('<svg'), type).toBe(true);
      expect(svg, type).not.toContain('NaN');
      expect(svg, type).not.toContain('No data');
      expect(() => parseXml(svg), type).not.toThrow();
    }
  });

  it('computes nice scales, quartiles, bins, trends and treemap tiles', () => {
    expect(niceScale(0, 170, 5)).toMatchObject({ min: 0, max: 200, step: 50 });
    expect(quartiles([1, 2, 3, 4, 5, 6, 7, 8, 9, 100])).toMatchObject({ median: 5.5, min: 1, max: 100, outliers: [100] });
    const bins = histogramBins([1, 2, 2, 3, 3, 3, 4, 4, 5, 9]);
    expect(bins.reduce((a, b) => a + b.count, 0)).toBe(10);
    expect(fitTrend('linear', [1, 2, 3], [2, 4, 6])!(4)).toBeCloseTo(8);
    expect(fitTrend('polynomial', [1, 2, 3, 4], [1, 4, 9, 16])!(5)).toBeCloseTo(25);
    const tiles = squarify([6, 6, 4, 3, 2, 2, 1], { x: 0, y: 0, w: 600, h: 400 });
    expect(tiles.reduce((a, t) => a + t.w * t.h, 0)).toBeCloseTo(240000);
  });

  it('draws every gallery shape', () => {
    for (const group of SHAPE_GALLERY) {
      for (const item of group.items) {
        const svg = shapeSvg({ preset: item.preset, fill: '#4F81BD', line: { color: '#385D8A', width: 3 }, fontColor: '#000', text: 'A & B', arrowEnd: item.arrowEnd }, 120, 80);
        expect(svg, item.preset).not.toContain('NaN');
        expect(() => parseXml(svg), item.preset).not.toThrow();
      }
    }
  });
});

describe('charts in XLSX files', () => {
  it('writes classic chart parts that read back as the same charts', () => {
    for (const type of CHART_TYPES.filter((t) => !CHARTEX_TYPES.has(t))) {
      const d = salesDoc();
      const chart = addChart(d, type);
      d.updateChart(chart.id, { title: 'Sales', subtitle: 'By quarter', xAxisTitle: 'Quarter', yAxisTitle: 'Units', legendPosition: 'right', valueLabels: true });
      const { files, wb } = roundTrip(d);
      const xml = strFromU8(files['xl/charts/chart1.xml']);
      expect(xml, type).toMatch(/<c:f>Sheet1!\$B\$2:\$[BD]\$[24]<\/c:f>/);
      expect(() => parseXml(xml), type).not.toThrow();
      expect(strFromU8(files['[Content_Types].xml'])).toContain('/xl/charts/chart1.xml');
      expect(strFromU8(files['xl/drawings/_rels/drawing1.xml.rels'])).toContain('relationships/chart');
      const back = wb.sheets[0].charts[0];
      expect(back.type, type).toBe(type);
      expect(back).toMatchObject({ title: 'Sales', subtitle: 'By quarter', legendPosition: 'right', valueLabels: true, showLegend: true });
      expect(back.series.length, type).toBeGreaterThan(0);
      expect(back.source?.chartPath).toBe('xl/charts/chart1.xml');
      expect(back.from).toEqual(d.sheet.charts[0].from);
    }
  });

  it('writes cached values for viewers that do not recalculate', () => {
    const d = salesDoc();
    addChart(d, 'column');
    const xml = strFromU8(roundTrip(d).files['xl/charts/chart1.xml']);
    expect(xml).toContain('<c:pt idx="0"><c:v>120</c:v></c:pt>');
    expect(xml).toContain('<c:v>North</c:v>');
    expect(xml).toContain('<c:v>Q1</c:v>');
  });

  it('writes the Office 2016 chart types as chartex parts with a fallback', () => {
    for (const type of [...CHARTEX_TYPES]) {
      const d = salesDoc();
      addChart(d, type, type === 'histogram' || type === 'box-whisker' ? 'B1:B4' : 'A1:B4');
      const { files, wb } = roundTrip(d);
      const part = Object.keys(files).find((f) => /^xl\/charts\/chartEx\d+\.xml$/.test(f));
      expect(part, type).toBeDefined();
      const xml = strFromU8(files[part!]);
      expect(() => parseXml(xml), type).not.toThrow();
      expect(xml).toContain('<cx:f>Sheet1!$B$2:$B$4</cx:f>');
      expect(strFromU8(files['[Content_Types].xml'])).toContain('application/vnd.ms-office.chartex+xml');
      const drawing = strFromU8(files['xl/drawings/drawing1.xml']);
      expect(drawing).toContain('mc:AlternateContent');
      expect(drawing).toContain('mc:Fallback');
      expect(wb.sheets[0].charts[0]?.type, type).toBe(type);
    }
  });

  it('keeps an unchanged chart part of a file and rewrites an edited one', () => {
    const bytes = readFixture('kitchen-sink.xlsx');
    const original = strFromU8(unzipSync(bytes)['xl/charts/chart1.xml']);
    const d = SpreadsheetDocument.fromXlsx(bytes, 'kitchen-sink.xlsx');
    const chart = d.wb.sheets[0].charts[0];
    // Moving the chart keeps its part.
    d.setObjectRect(chart.id, { x: 10, y: 400, width: 300, height: 200 });
    let files = unzipSync(d.save());
    expect(strFromU8(files['xl/charts/chart1.xml'])).toBe(original);
    expect(strFromU8(files['xl/drawings/drawing1.xml'])).toContain('<xdr:col>0</xdr:col>');
    // Editing it writes a new part and drops the old one.
    d.updateChart(chart.id, { type: 'line', title: 'Revenue' });
    files = unzipSync(d.save());
    expect(files['xl/charts/chart1.xml']).toBeUndefined();
    const again = readXlsx(d.save());
    expect(again.sheets[0].charts[0]).toMatchObject({ type: 'line', title: 'Revenue' });
    // Deleting it removes the part.
    d.deleteObject(chart.id);
    files = unzipSync(d.save());
    expect(Object.keys(files).filter((f) => f.startsWith('xl/charts/'))).toEqual([]);
  });
});

describe('shapes in XLSX files', () => {
  it('round-trips preset shapes and connectors', () => {
    const d = salesDoc();
    d.addShape({ name: 'Box', preset: 'roundRect', fill: '#4F81BD', line: { color: '#385D8A', width: 3 }, fontColor: '#FFFFFF', text: 'Hello\nworld' }, { x: 10, y: 150, width: 200, height: 100 });
    d.addShape({ name: 'Arrow', preset: 'straightConnector1', fill: null, line: { color: '#385D8A', width: 2, dash: 'dashed' }, fontColor: '#000000', arrowEnd: true }, { x: 250, y: 150, width: 100, height: 50 });
    const { files, wb } = roundTrip(d);
    const drawing = strFromU8(files['xl/drawings/drawing1.xml']);
    expect(drawing).toContain('<a:prstGeom prst="roundRect">');
    expect(drawing).toContain('<xdr:cxnSp');
    expect(drawing).toContain('<a:tailEnd type="triangle"/>');
    const shapes = wb.sheets[0].shapes;
    expect(shapes).toHaveLength(2);
    expect(shapes[0]).toMatchObject({ name: 'Box', preset: 'roundRect', fill: '#4F81BD', fontColor: '#FFFFFF', text: 'Hello\nworld' });
    expect(shapes[0].line).toMatchObject({ color: '#385D8A', width: 3 });
    expect(shapes[1]).toMatchObject({ preset: 'straightConnector1', fill: null, arrowEnd: true });
    expect(shapes[1].line).toMatchObject({ width: 2, dash: 'dashed' });
  });

  it('keeps the stacking order of pictures, charts and shapes', () => {
    const d = salesDoc();
    const chart = addChart(d, 'column');
    const shape = d.addShape({ name: 'S', preset: 'ellipse', fill: '#FF0000', line: null, fontColor: '#000000' }, { x: 520, y: 20, width: 80, height: 80 })!;
    d.reorderObject(shape.id, 'back');
    const { wb } = roundTrip(d);
    const back = wb.sheets[0];
    expect(back.shapes[0].z).toBeLessThan(back.charts[0].z);
    void chart;
  });
});
