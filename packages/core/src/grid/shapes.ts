/**
 * Shapes: the gallery of the Insert Shape dialog and SVG geometry for the DrawingML preset shapes
 * it offers (the same preset names are written to the file, so Excel draws the same shapes).
 */
import type { SheetShape } from '../core/model/drawing';

export interface GalleryItem {
  /** Name shown as the button label in the gallery. */
  label: string;
  preset: string;
  arrowStart?: boolean;
  arrowEnd?: boolean;
}

const item = (label: string, preset: string, arrows?: 'end' | 'both'): GalleryItem => ({ label, preset, arrowEnd: !!arrows, arrowStart: arrows === 'both' });

export const SHAPE_GALLERY: { title: string; items: GalleryItem[] }[] = [
  {
    title: 'Lines',
    items: [
      item('line', 'line'),
      item('line with arrow', 'straightConnector1', 'end'),
      item('double arrow line', 'straightConnector1', 'both'),
      item('bent connector3', 'bentConnector3'),
      item('elbow arrow connector', 'bentConnector3', 'end'),
      item('double arrow elbow connector', 'bentConnector3', 'both'),
      item('curved connector3', 'curvedConnector3'),
      item('curved arrow connector', 'curvedConnector3', 'end'),
      item('curved double arrow connector', 'curvedConnector3', 'both'),
    ],
  },
  {
    title: 'Rectangles',
    items: [
      item('rect', 'rect'),
      item('round rect', 'roundRect'),
      item('snip1 rect', 'snip1Rect'),
      item('snip2 same rect', 'snip2SameRect'),
      item('snip2 diag rect', 'snip2DiagRect'),
      item('round1 rect', 'round1Rect'),
      item('rounded top left rectangle', 'snipRoundRect'),
      item('round2 diag rect', 'round2DiagRect'),
    ],
  },
  {
    title: 'Basic Forms',
    items: [
      ['ellipse', 'ellipse'], ['triangle', 'triangle'], ['rt triangle', 'rtTriangle'], ['parallelogram', 'parallelogram'], ['trapezoid', 'trapezoid'],
      ['diamond', 'diamond'], ['pentagon', 'pentagon'], ['hexagon', 'hexagon'], ['heptagon', 'heptagon'], ['octagon', 'octagon'], ['decagon', 'decagon'],
      ['pie', 'pie'], ['chord', 'chord'], ['teardrop', 'teardrop'], ['frame', 'frame'], ['half frame', 'halfFrame'], ['corner', 'corner'],
      ['diag stripe', 'diagStripe'], ['plus', 'plus'], ['plaque', 'plaque'], ['can', 'can'], ['cube', 'cube'], ['bevel', 'bevel'], ['donut', 'donut'],
      ['no smoking', 'noSmoking'], ['block arc', 'blockArc'], ['smiley face', 'smileyFace'], ['heart', 'heart'], ['lightning bolt', 'lightningBolt'],
      ['sun', 'sun'], ['moon', 'moon'], ['cloud', 'cloud'], ['arc', 'arc'], ['left bracket', 'leftBracket'], ['right bracket', 'rightBracket'],
      ['bracket pair', 'bracketPair'], ['left brace', 'leftBrace'], ['right brace', 'rightBrace'], ['brace pair', 'bracePair'],
    ].map(([l, p]) => item(l, p)),
  },
  {
    title: 'Block Arrows',
    items: [
      ['right arrow', 'rightArrow'], ['left arrow', 'leftArrow'], ['up arrow', 'upArrow'], ['down arrow', 'downArrow'], ['up down arrow', 'upDownArrow'],
      ['left right arrow', 'leftRightArrow'], ['quad arrow', 'quadArrow'], ['left right up arrow', 'leftRightUpArrow'], ['bent arrow', 'bentArrow'],
      ['uturn arrow', 'uturnArrow'], ['left up arrow', 'leftUpArrow'], ['bent up arrow', 'bentUpArrow'], ['curved right arrow', 'curvedRightArrow'],
      ['curved left arrow', 'curvedLeftArrow'], ['curved up arrow', 'curvedUpArrow'], ['curved down arrow', 'curvedDownArrow'],
      ['striped right arrow', 'stripedRightArrow'], ['notched right arrow', 'notchedRightArrow'], ['home plate', 'homePlate'], ['chevron', 'chevron'],
      ['right arrow callout', 'rightArrowCallout'], ['down arrow callout', 'downArrowCallout'], ['left arrow callout', 'leftArrowCallout'],
      ['up arrow callout', 'upArrowCallout'], ['left right arrow callout', 'leftRightArrowCallout'], ['quad arrow callout', 'quadArrowCallout'],
      ['circular arrow', 'circularArrow'],
    ].map(([l, p]) => item(l, p)),
  },
  { title: 'Stars', items: [['star4', 'star4'], ['star5', 'star5'], ['star6', 'star6'], ['star8', 'star8'], ['irregular seal1', 'irregularSeal1']].map(([l, p]) => item(l, p)) },
  { title: 'Callouts', items: [['wedge rect callout', 'wedgeRectCallout'], ['wedge round rect callout', 'wedgeRoundRectCallout'], ['wedge ellipse callout', 'wedgeEllipseCallout']].map(([l, p]) => item(l, p)) },
  {
    title: 'Equation Shapes',
    items: [['math plus', 'mathPlus'], ['math minus', 'mathMinus'], ['math multiply', 'mathMultiply'], ['math divide', 'mathDivide'], ['math equal', 'mathEqual'], ['math not equal', 'mathNotEqual']].map(([l, p]) => item(l, p)),
  },
  {
    title: 'Flowchart',
    items: [
      ['rect', 'flowChartProcess'], ['round rect', 'flowChartAlternateProcess'], ['diamond', 'flowChartDecision'], ['parallelogram', 'flowChartInputOutput'],
      ['predefined process', 'flowChartPredefinedProcess'], ['internal storage', 'flowChartInternalStorage'], ['document', 'flowChartDocument'],
      ['multidocument', 'flowChartMultidocument'], ['terminator', 'flowChartTerminator'], ['preparation', 'flowChartPreparation'],
      ['manual input', 'flowChartManualInput'], ['manual operation', 'flowChartManualOperation'], ['ellipse', 'flowChartConnector'],
      ['offpage connector', 'flowChartOffpageConnector'], ['punched card', 'flowChartPunchedCard'], ['punched tape', 'flowChartPunchedTape'],
      ['summing junction', 'flowChartSummingJunction'], ['or', 'flowChartOr'], ['collate', 'flowChartCollate'], ['sort', 'flowChartSort'],
      ['extract', 'flowChartExtract'], ['merge', 'flowChartMerge'], ['online storage', 'flowChartOnlineStorage'], ['delay', 'flowChartDelay'],
      ['magnetic tape', 'flowChartMagneticTape'], ['can', 'flowChartMagneticDisk'], ['magnetic drum', 'flowChartMagneticDrum'], ['display', 'flowChartDisplay'],
    ].map(([l, p]) => item(l, p)),
  },
];

/** Presets drawn as lines (no fill, open path). */
export const LINE_PRESETS = new Set(['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3']);

type Shade = 'norm' | 'darken' | 'darkenLess' | 'lighten' | 'lightenLess' | 'none';

interface GeoPath {
  d: string;
  fill: Shade;
  stroke: boolean;
  evenOdd?: boolean;
}

interface Geo {
  paths: GeoPath[];
  /** Text rectangle (defaults to the whole shape). */
  text?: [number, number, number, number];
  /** Arrowhead tips and directions of line presets. */
  ends?: { start: [number, number, number, number]; end: [number, number, number, number] };
}

const r = (n: number) => String(Math.round(n * 100) / 100);
const P = (x: number, y: number) => `${r(x)} ${r(y)}`;

function poly(pts: [number, number][]): string {
  return `M${pts.map(([x, y]) => P(x, y)).join(' L')} Z`;
}

function ellipsePath(cx: number, cy: number, rx: number, ry: number): string {
  return `M${P(cx - rx, cy)} A${r(rx)} ${r(ry)} 0 1 0 ${P(cx + rx, cy)} A${r(rx)} ${r(ry)} 0 1 0 ${P(cx - rx, cy)} Z`;
}

function roundRectPath(x: number, y: number, w: number, h: number, rad: number): string {
  const q = Math.max(0, Math.min(rad, w / 2, h / 2));
  return `M${P(x + q, y)} L${P(x + w - q, y)} A${r(q)} ${r(q)} 0 0 1 ${P(x + w, y + q)} L${P(x + w, y + h - q)} A${r(q)} ${r(q)} 0 0 1 ${P(x + w - q, y + h)} L${P(x + q, y + h)} A${r(q)} ${r(q)} 0 0 1 ${P(x, y + h - q)} L${P(x, y + q)} A${r(q)} ${r(q)} 0 0 1 ${P(x + q, y)} Z`;
}

/** Point on the ellipse inscribed in the box at an angle in degrees (clockwise from 3 o'clock). */
function onEllipse(cx: number, cy: number, rx: number, ry: number, deg: number): [number, number] {
  const a = (deg * Math.PI) / 180;
  return [cx + rx * Math.cos(a), cy + ry * Math.sin(a)];
}

function regular(n: number, w: number, h: number): string {
  const pts: [number, number][] = [];
  for (let i = 0; i < n; i++) pts.push(onEllipse(w / 2, h / 2, w / 2, h / 2, -90 + (360 * i) / n));
  return poly(pts);
}

function star(n: number, w: number, h: number, inner: number): string {
  const pts: [number, number][] = [];
  for (let i = 0; i < n * 2; i++) {
    const k = i % 2 ? inner : 1;
    pts.push(onEllipse(w / 2, h / 2, (w / 2) * k, (h / 2) * k, -90 + (180 * i) / n));
  }
  return poly(pts);
}

/** Path from commands in unit coordinates, mapped into the box (optionally mirrored or transposed). */
function unitPath(cmds: (string | number)[], w: number, h: number, map: (x: number, y: number) => [number, number] = (x, y) => [x, y]): string {
  const out: string[] = [];
  for (let i = 0; i < cmds.length; ) {
    const c = cmds[i++];
    if (c === 'Z') {
      out.push('Z');
      continue;
    }
    const count = c === 'Q' ? 2 : c === 'C' ? 3 : 1;
    const pts: string[] = [];
    for (let k = 0; k < count; k++) {
      const [x, y] = map(cmds[i++] as number, cmds[i++] as number);
      pts.push(P(x * w, y * h));
    }
    out.push(`${c}${pts.join(' ')}`);
  }
  return out.join(' ');
}

const fillPath = (d: string, evenOdd = false): GeoPath => ({ d, fill: 'norm', stroke: true, evenOdd });

function geometry(preset: string, w: number, h: number): Geo {
  const ss = Math.min(w, h);
  const cx = w / 2;
  const cy = h / 2;
  const one = (d: string, text?: Geo['text']): Geo => ({ paths: [fillPath(d)], text });
  switch (preset) {
    // Lines
    case 'line':
    case 'straightConnector1':
      return { paths: [{ d: `M0 0 L${P(w, h)}`, fill: 'none', stroke: true }], ends: { start: [0, 0, -w, -h], end: [w, h, w, h] } };
    case 'bentConnector3':
      return { paths: [{ d: `M0 0 L${P(w / 2, 0)} L${P(w / 2, h)} L${P(w, h)}`, fill: 'none', stroke: true }], ends: { start: [0, 0, -1, 0], end: [w, h, 1, 0] } };
    case 'curvedConnector3':
      return { paths: [{ d: `M0 0 C${P(w / 2, 0)} ${P(w / 2, h)} ${P(w, h)}`, fill: 'none', stroke: true }], ends: { start: [0, 0, -1, 0], end: [w, h, 1, 0] } };

    // Rectangles
    case 'rect':
    case 'flowChartProcess':
      return one(poly([[0, 0], [w, 0], [w, h], [0, h]]));
    case 'roundRect':
    case 'flowChartAlternateProcess':
      return one(roundRectPath(0, 0, w, h, ss * 0.1667));
    case 'snip1Rect': {
      const a = ss * 0.1667;
      return one(poly([[0, 0], [w - a, 0], [w, a], [w, h], [0, h]]));
    }
    case 'snip2SameRect': {
      const a = ss * 0.1667;
      return one(poly([[a, 0], [w - a, 0], [w, a], [w, h], [0, h], [0, a]]));
    }
    case 'snip2DiagRect': {
      const a = ss * 0.1667;
      return one(poly([[0, 0], [w - a, 0], [w, a], [w, h], [a, h], [0, h - a]]));
    }
    case 'round1Rect': {
      const a = ss * 0.1667;
      return one(`M0 0 L${P(w - a, 0)} A${r(a)} ${r(a)} 0 0 1 ${P(w, a)} L${P(w, h)} L${P(0, h)} Z`);
    }
    case 'snipRoundRect': {
      const a = ss * 0.1667;
      return one(`M${P(a, 0)} L${P(w - a, 0)} L${P(w, a)} L${P(w, h)} L${P(0, h)} L${P(0, a)} A${r(a)} ${r(a)} 0 0 1 ${P(a, 0)} Z`);
    }
    case 'round2DiagRect': {
      const a = ss * 0.1667;
      return one(`M${P(a, 0)} L${P(w, 0)} L${P(w, h - a)} A${r(a)} ${r(a)} 0 0 1 ${P(w - a, h)} L${P(0, h)} L${P(0, a)} A${r(a)} ${r(a)} 0 0 1 ${P(a, 0)} Z`);
    }

    // Basic forms
    case 'ellipse':
    case 'flowChartConnector':
      return one(ellipsePath(cx, cy, cx, cy), [w * 0.15, h * 0.15, w * 0.7, h * 0.7]);
    case 'triangle':
    case 'flowChartExtract':
      return one(poly([[cx, 0], [w, h], [0, h]]), [w * 0.25, h * 0.5, w * 0.5, h * 0.45]);
    case 'flowChartMerge':
      return one(poly([[0, 0], [w, 0], [cx, h]]), [w * 0.25, h * 0.05, w * 0.5, h * 0.45]);
    case 'rtTriangle':
      return one(poly([[0, 0], [w, h], [0, h]]), [w * 0.08, h * 0.5, w * 0.5, h * 0.45]);
    case 'parallelogram': {
      const a = ss * 0.25;
      return one(poly([[a, 0], [w, 0], [w - a, h], [0, h]]), [a, 0, w - 2 * a, h]);
    }
    case 'flowChartInputOutput':
      return one(poly([[w * 0.2, 0], [w, 0], [w * 0.8, h], [0, h]]), [w * 0.2, 0, w * 0.6, h]);
    case 'trapezoid': {
      const a = ss * 0.25;
      return one(poly([[0, h], [a, 0], [w - a, 0], [w, h]]), [a * 0.7, h * 0.2, w - a * 1.4, h * 0.8]);
    }
    case 'diamond':
    case 'flowChartDecision':
      return one(poly([[cx, 0], [w, cy], [cx, h], [0, cy]]), [w * 0.25, h * 0.25, w * 0.5, h * 0.5]);
    case 'pentagon':
      return one(poly([[cx, 0], [w, h * 0.3804], [w * 0.809, h], [w * 0.191, h], [0, h * 0.3804]]), [w * 0.2, h * 0.3, w * 0.6, h * 0.6]);
    case 'hexagon': {
      const a = ss * 0.25;
      return one(poly([[0, cy], [a, 0], [w - a, 0], [w, cy], [w - a, h], [a, h]]), [a * 0.6, h * 0.15, w - a * 1.2, h * 0.7]);
    }
    case 'heptagon':
      return one(regular(7, w, h), [w * 0.2, h * 0.2, w * 0.6, h * 0.6]);
    case 'decagon':
      return one(regular(10, w, h), [w * 0.15, h * 0.15, w * 0.7, h * 0.7]);
    case 'octagon': {
      const a = ss * 0.2929;
      return one(poly([[a, 0], [w - a, 0], [w, a], [w, h - a], [w - a, h], [a, h], [0, h - a], [0, a]]), [a * 0.5, a * 0.5, w - a, h - a]);
    }
    case 'pie': {
      const [sx, sy] = onEllipse(cx, cy, cx, cy, 0);
      const [ex, ey] = onEllipse(cx, cy, cx, cy, 270);
      return one(`M${P(cx, cy)} L${P(sx, sy)} A${r(cx)} ${r(cy)} 0 1 1 ${P(ex, ey)} Z`, [w * 0.15, h * 0.15, w * 0.7, h * 0.7]);
    }
    case 'chord': {
      const [sx, sy] = onEllipse(cx, cy, cx, cy, 45);
      const [ex, ey] = onEllipse(cx, cy, cx, cy, 270);
      return one(`M${P(sx, sy)} A${r(cx)} ${r(cy)} 0 1 1 ${P(ex, ey)} Z`, [w * 0.15, h * 0.15, w * 0.7, h * 0.7]);
    }
    case 'teardrop':
      return one(`M0 ${r(cy)} A${r(cx)} ${r(cy)} 0 0 1 ${P(cx, 0)} L${P(w, 0)} L${P(w, cy)} A${r(cx)} ${r(cy)} 0 0 1 ${P(cx, h)} A${r(cx)} ${r(cy)} 0 0 1 ${P(0, cy)} Z`, [w * 0.15, h * 0.15, w * 0.7, h * 0.7]);
    case 'frame': {
      const t = ss * 0.125;
      return { paths: [fillPath(`${poly([[0, 0], [w, 0], [w, h], [0, h]])} ${poly([[t, t], [t, h - t], [w - t, h - t], [w - t, t]])}`, true)], text: [t, t, w - 2 * t, h - 2 * t] };
    }
    case 'halfFrame': {
      const t = ss * 0.2;
      return one(poly([[0, 0], [w, 0], [w - t, t], [t, t], [t, h - t], [0, h]]));
    }
    case 'corner': {
      const t = ss * 0.5;
      return one(poly([[0, 0], [t, 0], [t, h - t], [w, h - t], [w, h], [0, h]]));
    }
    case 'diagStripe':
      return one(poly([[0, cy], [cx, 0], [w, 0], [0, h]]));
    case 'plus': {
      const a = ss * 0.25;
      return one(poly([[0, a], [a, a], [a, 0], [w - a, 0], [w - a, a], [w, a], [w, h - a], [w - a, h - a], [w - a, h], [a, h], [a, h - a], [0, h - a]]), [a, a, w - 2 * a, h - 2 * a]);
    }
    case 'plaque': {
      const q = ss * 0.1667;
      return one(`M0 ${r(q)} A${r(q)} ${r(q)} 0 0 0 ${P(q, 0)} L${P(w - q, 0)} A${r(q)} ${r(q)} 0 0 0 ${P(w, q)} L${P(w, h - q)} A${r(q)} ${r(q)} 0 0 0 ${P(w - q, h)} L${P(q, h)} A${r(q)} ${r(q)} 0 0 0 ${P(0, h - q)} Z`, [q, q, w - 2 * q, h - 2 * q]);
    }
    case 'can':
    case 'flowChartMagneticDisk': {
      const ry = preset === 'can' ? ss * 0.125 : h / 6;
      return {
        paths: [
          fillPath(`M0 ${r(ry)} L0 ${r(h - ry)} A${r(cx)} ${r(ry)} 0 0 0 ${P(w, h - ry)} L${P(w, ry)} A${r(cx)} ${r(ry)} 0 0 0 ${P(0, ry)} Z`),
          { d: ellipsePath(cx, ry, cx, ry), fill: 'lighten', stroke: true },
        ],
        text: [0, ry * 2, w, h - ry * 3],
      };
    }
    case 'flowChartMagneticDrum': {
      const rx = w / 6;
      return {
        paths: [
          fillPath(`M${P(rx, 0)} L${P(w - rx, 0)} A${r(rx)} ${r(cy)} 0 0 1 ${P(w - rx, h)} L${P(rx, h)} A${r(rx)} ${r(cy)} 0 0 1 ${P(rx, 0)} Z`),
          { d: ellipsePath(w - rx, cy, rx, cy), fill: 'lighten', stroke: true },
        ],
        text: [rx, 0, w - rx * 3, h],
      };
    }
    case 'cube': {
      const d = ss * 0.25;
      return {
        paths: [
          fillPath(poly([[0, d], [w - d, d], [w - d, h], [0, h]])),
          { d: poly([[0, d], [d, 0], [w, 0], [w - d, d]]), fill: 'lighten', stroke: true },
          { d: poly([[w - d, d], [w, 0], [w, h - d], [w - d, h]]), fill: 'darken', stroke: true },
        ],
        text: [0, d, w - d, h - d],
      };
    }
    case 'bevel': {
      const b = ss * 0.125;
      return {
        paths: [
          fillPath(poly([[b, b], [w - b, b], [w - b, h - b], [b, h - b]])),
          { d: poly([[0, 0], [w, 0], [w - b, b], [b, b]]), fill: 'lightenLess', stroke: true },
          { d: poly([[0, 0], [b, b], [b, h - b], [0, h]]), fill: 'lighten', stroke: true },
          { d: poly([[w, 0], [w, h], [w - b, h - b], [w - b, b]]), fill: 'darkenLess', stroke: true },
          { d: poly([[0, h], [b, h - b], [w - b, h - b], [w, h]]), fill: 'darken', stroke: true },
        ],
        text: [b, b, w - 2 * b, h - 2 * b],
      };
    }
    case 'donut': {
      const t = ss * 0.25;
      return { paths: [fillPath(`${ellipsePath(cx, cy, cx, cy)} ${ellipsePath(cx, cy, Math.max(1, cx - t), Math.max(1, cy - t))}`, true)] };
    }
    case 'noSmoking': {
      const t = ss * 0.1875;
      const irx = Math.max(1, cx - t);
      const iry = Math.max(1, cy - t);
      const delta = (Math.asin(Math.min(1, t / 2 / Math.min(irx, iry))) * 180) / Math.PI;
      const band = poly([onEllipse(cx, cy, irx, iry, 225 - delta), onEllipse(cx, cy, irx, iry, 225 + delta), onEllipse(cx, cy, irx, iry, 45 - delta), onEllipse(cx, cy, irx, iry, 45 + delta)]);
      return { paths: [fillPath(`${ellipsePath(cx, cy, cx, cy)} ${ellipsePath(cx, cy, irx, iry)}`, true), fillPath(band)] };
    }
    case 'blockArc': {
      const t = ss * 0.25;
      return one(`M0 ${r(cy)} A${r(cx)} ${r(cy)} 0 0 1 ${P(w, cy)} L${P(w - t, cy)} A${r(cx - t)} ${r(cy - t)} 0 0 0 ${P(t, cy)} Z`);
    }
    case 'smileyFace':
      return {
        paths: [
          fillPath(ellipsePath(cx, cy, cx, cy)),
          { d: `${ellipsePath(w * 0.35, h * 0.37, w * 0.05, h * 0.05)} ${ellipsePath(w * 0.65, h * 0.37, w * 0.05, h * 0.05)}`, fill: 'darkenLess', stroke: true },
          { d: `M${P(w * 0.27, h * 0.68)} Q${P(cx, h * 0.86)} ${P(w * 0.73, h * 0.68)}`, fill: 'none', stroke: true },
        ],
      };
    case 'heart':
      return one(
        `M${P(cx, h * 0.3)} C${P(cx, h * 0.1)} ${P(w * 0.4, 0)} ${P(w * 0.25, 0)} C${P(w * 0.1, 0)} ${P(0, h * 0.12)} ${P(0, h * 0.3)} C${P(0, h * 0.55)} ${P(w * 0.3, h * 0.75)} ${P(cx, h)} C${P(w * 0.7, h * 0.75)} ${P(w, h * 0.55)} ${P(w, h * 0.3)} C${P(w, h * 0.12)} ${P(w * 0.9, 0)} ${P(w * 0.75, 0)} C${P(w * 0.6, 0)} ${P(cx, h * 0.1)} ${P(cx, h * 0.3)} Z`,
        [w * 0.2, h * 0.2, w * 0.6, h * 0.5],
      );
    case 'lightningBolt':
      return one(poly([[0.39, 0], [0.56, 0.31], [0.5, 0.34], [0.74, 0.56], [0.67, 0.6], [1, 1], [0.46, 0.67], [0.54, 0.63], [0.22, 0.43], [0.3, 0.39], [0, 0.18]].map(([x, y]) => [x * w, y * h] as [number, number])));
    case 'sun': {
      const rays: string[] = [];
      for (let i = 0; i < 8; i++) {
        const a = i * 45;
        rays.push(poly([onEllipse(cx, cy, cx, cy, a), onEllipse(cx, cy, cx * 0.7, cy * 0.7, a - 9), onEllipse(cx, cy, cx * 0.7, cy * 0.7, a + 9)]));
      }
      return { paths: [fillPath(rays.join(' ')), fillPath(ellipsePath(cx, cy, cx * 0.5, cy * 0.5))], text: [w * 0.3, h * 0.3, w * 0.4, h * 0.4] };
    }
    case 'moon':
      return one(`M${P(w, 0)} A${r(w)} ${r(cy)} 0 0 0 ${P(w, h)} A${r(w * 0.5)} ${r(cy)} 0 0 1 ${P(w, 0)} Z`, [w * 0.12, h * 0.25, w * 0.35, h * 0.5]);
    case 'cloud': {
      const n = 9;
      const pts = Array.from({ length: n }, (_, i) => onEllipse(cx, cy, cx * 0.8, cy * 0.72, -90 + (360 * i) / n));
      let d = `M${P(pts[0][0], pts[0][1])}`;
      for (let i = 0; i < n; i++) {
        const a = pts[i];
        const b = pts[(i + 1) % n];
        const rad = Math.hypot(b[0] - a[0], b[1] - a[1]) * 0.62;
        d += ` A${r(rad)} ${r(rad)} 0 0 1 ${P(b[0], b[1])}`;
      }
      return one(`${d} Z`, [w * 0.2, h * 0.25, w * 0.6, h * 0.5]);
    }
    case 'arc': {
      const [sx, sy] = onEllipse(cx, cy, cx, cy, 270);
      return { paths: [{ d: `M${P(sx, sy)} A${r(cx)} ${r(cy)} 0 0 1 ${P(w, cy)}`, fill: 'none', stroke: true }] };
    }
    case 'leftBracket': {
      const a = Math.min(h / 2, ss * 0.0833 * 2);
      return { paths: [{ d: `M${P(w, 0)} A${r(w)} ${r(a)} 0 0 0 ${P(0, a)} L${P(0, h - a)} A${r(w)} ${r(a)} 0 0 0 ${P(w, h)}`, fill: 'none', stroke: true }] };
    }
    case 'rightBracket': {
      const a = Math.min(h / 2, ss * 0.0833 * 2);
      return { paths: [{ d: `M0 0 A${r(w)} ${r(a)} 0 0 1 ${P(w, a)} L${P(w, h - a)} A${r(w)} ${r(a)} 0 0 1 ${P(0, h)}`, fill: 'none', stroke: true }] };
    }
    case 'bracketPair': {
      const a = ss * 0.1667;
      return {
        paths: [
          { d: roundRectPath(0, 0, w, h, a), fill: 'norm', stroke: false },
          { d: `M${P(a, h)} A${r(a)} ${r(a)} 0 0 1 ${P(0, h - a)} L${P(0, a)} A${r(a)} ${r(a)} 0 0 1 ${P(a, 0)} M${P(w - a, 0)} A${r(a)} ${r(a)} 0 0 1 ${P(w, a)} L${P(w, h - a)} A${r(a)} ${r(a)} 0 0 1 ${P(w - a, h)}`, fill: 'none', stroke: true },
        ],
      };
    }
    case 'leftBrace':
    case 'rightBrace': {
      const a = Math.min(h / 4, ss * 0.0833 * 2);
      const half = w / 2;
      const d = `M${P(w, 0)} A${r(half)} ${r(a)} 0 0 0 ${P(half, a)} L${P(half, cy - a)} A${r(half)} ${r(a)} 0 0 1 ${P(0, cy)} A${r(half)} ${r(a)} 0 0 1 ${P(half, cy + a)} L${P(half, h - a)} A${r(half)} ${r(a)} 0 0 0 ${P(w, h)}`;
      if (preset === 'leftBrace') return { paths: [{ d, fill: 'none', stroke: true }] };
      const m = `M0 0 A${r(half)} ${r(a)} 0 0 1 ${P(half, a)} L${P(half, cy - a)} A${r(half)} ${r(a)} 0 0 0 ${P(w, cy)} A${r(half)} ${r(a)} 0 0 0 ${P(half, cy + a)} L${P(half, h - a)} A${r(half)} ${r(a)} 0 0 1 ${P(0, h)}`;
      return { paths: [{ d: m, fill: 'none', stroke: true }] };
    }
    case 'bracePair': {
      const bw = Math.min(w * 0.12, ss * 0.1667);
      const a = Math.min(h / 4, bw);
      const hb = bw / 2;
      const left = `M${P(bw, h)} A${r(hb)} ${r(a)} 0 0 1 ${P(hb, h - a)} L${P(hb, cy + a)} A${r(hb)} ${r(a)} 0 0 0 ${P(0, cy)} A${r(hb)} ${r(a)} 0 0 0 ${P(hb, cy - a)} L${P(hb, a)} A${r(hb)} ${r(a)} 0 0 1 ${P(bw, 0)}`;
      const right = `M${P(w - bw, 0)} A${r(hb)} ${r(a)} 0 0 1 ${P(w - hb, a)} L${P(w - hb, cy - a)} A${r(hb)} ${r(a)} 0 0 0 ${P(w, cy)} A${r(hb)} ${r(a)} 0 0 0 ${P(w - hb, cy + a)} L${P(w - hb, h - a)} A${r(hb)} ${r(a)} 0 0 1 ${P(w - bw, h)}`;
      return { paths: [{ d: poly([[bw, 0], [w - bw, 0], [w - bw, h], [bw, h]]), fill: 'norm', stroke: false }, { d: `${left} ${right}`, fill: 'none', stroke: true }] };
    }

    // Block arrows
    case 'rightArrow':
    case 'leftArrow': {
      const hl = Math.min(w, ss * 0.5);
      const pts: [number, number][] = [[0, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, cy], [w - hl, h], [w - hl, h * 0.75], [0, h * 0.75]];
      return one(poly(preset === 'leftArrow' ? pts.map(([x, y]) => [w - x, y]) : pts), [0, h * 0.25, w - hl * 0.5, h * 0.5]);
    }
    case 'upArrow':
    case 'downArrow': {
      const hl = Math.min(h, ss * 0.5);
      const pts: [number, number][] = [[w * 0.25, h], [w * 0.25, hl], [0, hl], [cx, 0], [w, hl], [w * 0.75, hl], [w * 0.75, h]];
      return one(poly(preset === 'downArrow' ? pts.map(([x, y]) => [x, h - y]) : pts), [w * 0.25, preset === 'upArrow' ? hl * 0.5 : 0, w * 0.5, h - hl * 0.5]);
    }
    case 'leftRightArrow': {
      const hl = Math.min(w / 2, ss * 0.5);
      return one(poly([[0, cy], [hl, 0], [hl, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, cy], [w - hl, h], [w - hl, h * 0.75], [hl, h * 0.75], [hl, h]]), [hl * 0.5, h * 0.25, w - hl, h * 0.5]);
    }
    case 'upDownArrow': {
      const hl = Math.min(h / 2, ss * 0.5);
      return one(poly([[cx, 0], [w, hl], [w * 0.75, hl], [w * 0.75, h - hl], [w, h - hl], [cx, h], [0, h - hl], [w * 0.25, h - hl], [w * 0.25, hl], [0, hl]]), [w * 0.25, hl * 0.5, w * 0.5, h - hl]);
    }
    case 'quadArrow': {
      const a1 = ss * 0.1125;
      const a2 = ss * 0.225;
      const a3 = ss * 0.225;
      return one(
        poly([
          [cx, 0], [cx + a2, a3], [cx + a1, a3], [cx + a1, cy - a1], [w - a3, cy - a1], [w - a3, cy - a2], [w, cy], [w - a3, cy + a2], [w - a3, cy + a1], [cx + a1, cy + a1],
          [cx + a1, h - a3], [cx + a2, h - a3], [cx, h], [cx - a2, h - a3], [cx - a1, h - a3], [cx - a1, cy + a1], [a3, cy + a1], [a3, cy + a2], [0, cy], [a3, cy - a2], [a3, cy - a1],
          [cx - a1, cy - a1], [cx - a1, a3], [cx - a2, a3],
        ]),
      );
    }
    case 'leftRightUpArrow': {
      const a1 = ss * 0.1125;
      const a2 = ss * 0.225;
      const a3 = ss * 0.225;
      const yb = h - a2;
      return one(poly([[cx, 0], [cx + a2, a3], [cx + a1, a3], [cx + a1, yb - a1], [w - a3, yb - a1], [w - a3, yb - a2], [w, yb], [w - a3, h], [w - a3, yb + a1], [a3, yb + a1], [a3, h], [0, yb], [a3, yb - a2], [a3, yb - a1], [cx - a1, yb - a1], [cx - a1, a3], [cx - a2, a3]]));
    }
    case 'bentArrow': {
      const th = ss * 0.25;
      const hw = ss * 0.25;
      const hl = ss * 0.25;
      const rad = Math.min(ss * 0.4375, h - hw);
      const ri = Math.max(0, rad - th);
      const top = hw - th / 2;
      return one(`M0 ${r(h)} L${P(0, top + rad)} A${r(rad)} ${r(rad)} 0 0 1 ${P(rad, top)} L${P(w - hl, top)} L${P(w - hl, 0)} L${P(w, hw)} L${P(w - hl, hw * 2)} L${P(w - hl, top + th)} L${P(th + ri, top + th)} A${r(ri)} ${r(ri)} 0 0 0 ${P(th, top + th + ri)} L${P(th, h)} Z`);
    }
    case 'uturnArrow': {
      const th = ss * 0.25;
      const hw = ss * 0.25;
      const hl = ss * 0.25;
      const rad = Math.min(ss * 0.4375, (w - hw + th / 2) / 2);
      const ri = Math.max(0, rad - th);
      const xr2 = w - hw + th / 2;
      const xr1 = w - hw - th / 2;
      const tipY = h * 0.75;
      const ya = Math.max(rad, tipY - hl);
      return one(`M0 ${r(h)} L${P(0, rad)} A${r(rad)} ${r(rad)} 0 0 1 ${P(rad, 0)} L${P(xr2 - rad, 0)} A${r(rad)} ${r(rad)} 0 0 1 ${P(xr2, rad)} L${P(xr2, ya)} L${P(w, ya)} L${P(w - hw, tipY)} L${P(w - 2 * hw, ya)} L${P(xr1, ya)} L${P(xr1, rad)} A${r(ri)} ${r(ri)} 0 0 0 ${P(xr2 - rad, th)} L${P(rad, th)} A${r(ri)} ${r(ri)} 0 0 0 ${P(th, rad)} L${P(th, h)} Z`);
    }
    case 'leftUpArrow': {
      const th = ss * 0.25;
      const hw = ss * 0.25;
      const hl = ss * 0.25;
      return one(poly([[0, h - hw], [hl, h - 2 * hw], [hl, h - hw - th / 2], [w - hw - th / 2, h - hw - th / 2], [w - hw - th / 2, hl], [w - 2 * hw, hl], [w - hw, 0], [w, hl], [w - hw + th / 2, hl], [w - hw + th / 2, h - hw + th / 2], [hl, h - hw + th / 2], [hl, h]]));
    }
    case 'bentUpArrow': {
      const th = ss * 0.25;
      const hw = ss * 0.25;
      const hl = ss * 0.25;
      return one(poly([[0, h - th], [w - hw - th / 2, h - th], [w - hw - th / 2, hl], [w - 2 * hw, hl], [w - hw, 0], [w, hl], [w - hw + th / 2, hl], [w - hw + th / 2, h], [0, h]]));
    }
    case 'curvedRightArrow':
    case 'curvedLeftArrow':
    case 'curvedUpArrow':
    case 'curvedDownArrow': {
      // Band along the left edge bending into an arrowhead at the bottom right (unit square).
      const cmds = ['M', 1, 0, 'Q', 0, 0, 0, 0.5, 'Q', 0, 0.78, 0.7, 0.78, 'L', 0.7, 0.6, 'L', 1, 0.82, 'L', 0.7, 1, 'L', 0.7, 0.9, 'Q', 0.25, 0.9, 0.25, 0.5, 'Q', 0.25, 0.25, 1, 0.25, 'Z'];
      const maps: Record<string, (x: number, y: number) => [number, number]> = {
        curvedRightArrow: (x, y) => [x, y],
        curvedLeftArrow: (x, y) => [1 - x, y],
        curvedDownArrow: (x, y) => [y, x],
        curvedUpArrow: (x, y) => [y, 1 - x],
      };
      return one(unitPath(cmds, w, h, maps[preset]));
    }
    case 'stripedRightArrow': {
      const hl = Math.min(w, ss * 0.5);
      const s1 = ss * 0.03125;
      const s2 = ss * 0.0625;
      const x0 = ss * 0.15625;
      return {
        paths: [
          fillPath(`${poly([[0, h * 0.25], [s1, h * 0.25], [s1, h * 0.75], [0, h * 0.75]])} ${poly([[s2, h * 0.25], [s2 * 2, h * 0.25], [s2 * 2, h * 0.75], [s2, h * 0.75]])}`),
          fillPath(poly([[x0, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, cy], [w - hl, h], [w - hl, h * 0.75], [x0, h * 0.75]])),
        ],
        text: [x0, h * 0.25, w - x0 - hl * 0.5, h * 0.5],
      };
    }
    case 'notchedRightArrow': {
      const hl = Math.min(w, ss * 0.5);
      const nx = ss * 0.2;
      return one(poly([[0, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, cy], [w - hl, h], [w - hl, h * 0.75], [0, h * 0.75], [nx, cy]]), [nx, h * 0.25, w - nx - hl * 0.5, h * 0.5]);
    }
    case 'homePlate': {
      const a = Math.min(w, ss * 0.5);
      return one(poly([[0, 0], [w - a, 0], [w, cy], [w - a, h], [0, h]]), [0, 0, w - a * 0.5, h]);
    }
    case 'chevron': {
      const a = Math.min(w / 2, ss * 0.5);
      return one(poly([[0, 0], [w - a, 0], [w, cy], [w - a, h], [0, h], [a, cy]]), [a, 0, w - 2 * a, h]);
    }
    case 'rightArrowCallout':
    case 'leftArrowCallout': {
      const bw = w * 0.65;
      const th = ss * 0.125;
      const hw = ss * 0.25;
      const hl = ss * 0.25;
      const pts: [number, number][] = [[0, 0], [bw, 0], [bw, cy - th], [w - hl, cy - th], [w - hl, cy - hw], [w, cy], [w - hl, cy + hw], [w - hl, cy + th], [bw, cy + th], [bw, h], [0, h]];
      return one(poly(preset === 'leftArrowCallout' ? pts.map(([x, y]) => [w - x, y]) : pts), preset === 'leftArrowCallout' ? [w - bw, 0, bw, h] : [0, 0, bw, h]);
    }
    case 'downArrowCallout':
    case 'upArrowCallout': {
      const bh = h * 0.65;
      const th = ss * 0.125;
      const hw = ss * 0.25;
      const hl = ss * 0.25;
      const pts: [number, number][] = [[0, 0], [w, 0], [w, bh], [cx + th, bh], [cx + th, h - hl], [cx + hw, h - hl], [cx, h], [cx - hw, h - hl], [cx - th, h - hl], [cx - th, bh], [0, bh]];
      return one(poly(preset === 'upArrowCallout' ? pts.map(([x, y]) => [x, h - y]) : pts), preset === 'upArrowCallout' ? [0, h - bh, w, bh] : [0, 0, w, bh]);
    }
    case 'leftRightArrowCallout': {
      const bw = w * 0.48;
      const x0 = (w - bw) / 2;
      const th = ss * 0.125;
      const hw = ss * 0.25;
      const hl = ss * 0.25;
      return one(poly([[0, cy], [hl, cy - hw], [hl, cy - th], [x0, cy - th], [x0, 0], [x0 + bw, 0], [x0 + bw, cy - th], [w - hl, cy - th], [w - hl, cy - hw], [w, cy], [w - hl, cy + hw], [w - hl, cy + th], [x0 + bw, cy + th], [x0 + bw, h], [x0, h], [x0, cy + th], [hl, cy + th], [hl, cy + hw]]), [x0, 0, bw, h]);
    }
    case 'quadArrowCallout': {
      const bx = w * 0.26;
      const by = h * 0.26;
      const a1 = ss * 0.09;
      const a2 = ss * 0.18;
      const a3 = ss * 0.18;
      return one(
        poly([
          [cx, 0], [cx + a2, a3], [cx + a1, a3], [cx + a1, by], [w - bx, by], [w - bx, cy - a1], [w - a3, cy - a1], [w - a3, cy - a2], [w, cy], [w - a3, cy + a2], [w - a3, cy + a1], [w - bx, cy + a1],
          [w - bx, h - by], [cx + a1, h - by], [cx + a1, h - a3], [cx + a2, h - a3], [cx, h], [cx - a2, h - a3], [cx - a1, h - a3], [cx - a1, h - by], [bx, h - by], [bx, cy + a1], [a3, cy + a1],
          [a3, cy + a2], [0, cy], [a3, cy - a2], [a3, cy - a1], [bx, cy - a1], [bx, by], [cx - a1, by], [cx - a1, a3], [cx - a2, a3],
        ]),
        [bx, by, w - 2 * bx, h - 2 * by],
      );
    }
    case 'circularArrow': {
      const t = ss * 0.125;
      const rx = cx - t * 0.6;
      const ry = cy - t * 0.6;
      const irx = Math.max(1, rx - t);
      const iry = Math.max(1, ry - t);
      const end = 330;
      const [ox0, oy0] = onEllipse(cx, cy, rx, ry, 180);
      const [ox1, oy1] = onEllipse(cx, cy, rx, ry, end);
      const [ix1, iy1] = onEllipse(cx, cy, irx, iry, end);
      const [ix0, iy0] = onEllipse(cx, cy, irx, iry, 180);
      const ext = t * 0.6;
      const [ex1, ey1] = onEllipse(cx, cy, rx + ext, ry + ext, end);
      const [ex2, ey2] = onEllipse(cx, cy, irx - ext, iry - ext, end);
      const [tx, ty] = onEllipse(cx, cy, (rx + irx) / 2, (ry + iry) / 2, end + 22);
      return one(`M${P(ox0, oy0)} A${r(rx)} ${r(ry)} 0 0 1 ${P(ox1, oy1)} L${P(ex1, ey1)} L${P(tx, ty)} L${P(ex2, ey2)} L${P(ix1, iy1)} A${r(irx)} ${r(iry)} 0 0 0 ${P(ix0, iy0)} Z`);
    }

    // Stars
    case 'star4':
      return one(star(4, w, h, 0.25), [w * 0.3, h * 0.3, w * 0.4, h * 0.4]);
    case 'star5':
      return one(star(5, w, h, 0.382), [w * 0.3, h * 0.35, w * 0.4, h * 0.35]);
    case 'star6':
      return one(star(6, w, h, 0.577), [w * 0.25, h * 0.25, w * 0.5, h * 0.5]);
    case 'star8':
      return one(star(8, w, h, 0.765), [w * 0.2, h * 0.2, w * 0.6, h * 0.6]);
    case 'irregularSeal1': {
      const pts = [[0.42, 0.24], [0.56, 0], [0.6, 0.2], [0.79, 0.12], [0.73, 0.3], [1, 0.35], [0.8, 0.48], [0.95, 0.64], [0.72, 0.62], [0.74, 0.86], [0.58, 0.7], [0.5, 1], [0.42, 0.74], [0.28, 0.9], [0.27, 0.68], [0.04, 0.73], [0.18, 0.55], [0, 0.42], [0.21, 0.36], [0.08, 0.14], [0.3, 0.22]];
      return one(poly(pts.map(([x, y]) => [x * w, y * h] as [number, number])), [w * 0.25, h * 0.3, w * 0.5, h * 0.4]);
    }

    // Callouts (the pointer reaches below the box, like Office's defaults)
    case 'wedgeRectCallout':
      return one(poly([[0, 0], [w, 0], [w, h], [w * 0.4167, h], [w * 0.2917, h * 1.125], [w * 0.1667, h], [0, h]]));
    case 'wedgeRoundRectCallout': {
      const q = ss * 0.1667;
      return one(`M${P(q, 0)} L${P(w - q, 0)} A${r(q)} ${r(q)} 0 0 1 ${P(w, q)} L${P(w, h - q)} A${r(q)} ${r(q)} 0 0 1 ${P(w - q, h)} L${P(w * 0.4167, h)} L${P(w * 0.2917, h * 1.125)} L${P(w * 0.1667, h)} L${P(q, h)} A${r(q)} ${r(q)} 0 0 1 ${P(0, h - q)} L${P(0, q)} A${r(q)} ${r(q)} 0 0 1 ${P(q, 0)} Z`);
    }
    case 'wedgeEllipseCallout': {
      const tx = w * 0.2917;
      const ty = h * 1.125;
      const a = (Math.atan2((ty - cy) / cy, (tx - cx) / cx) * 180) / Math.PI;
      const [sx, sy] = onEllipse(cx, cy, cx, cy, a + 12);
      const [ex, ey] = onEllipse(cx, cy, cx, cy, a - 12);
      return one(`M${P(sx, sy)} A${r(cx)} ${r(cy)} 0 1 1 ${P(ex, ey)} L${P(tx, ty)} Z`, [w * 0.15, h * 0.15, w * 0.7, h * 0.7]);
    }

    // Equation shapes
    case 'mathPlus': {
      const t = ss * 0.1176;
      const dx = w * 0.3675;
      const dy = h * 0.3675;
      return one(poly([[cx - t, cy - dy], [cx + t, cy - dy], [cx + t, cy - t], [cx + dx, cy - t], [cx + dx, cy + t], [cx + t, cy + t], [cx + t, cy + dy], [cx - t, cy + dy], [cx - t, cy + t], [cx - dx, cy + t], [cx - dx, cy - t], [cx - t, cy - t]]));
    }
    case 'mathMinus': {
      const t = ss * 0.1176;
      const dx = w * 0.3675;
      return one(poly([[cx - dx, cy - t], [cx + dx, cy - t], [cx + dx, cy + t], [cx - dx, cy + t]]));
    }
    case 'mathMultiply': {
      const t = ss * 0.1176;
      const len = ss * 0.42;
      const local: [number, number][] = [[-t, -len], [t, -len], [t, -t], [len, -t], [len, t], [t, t], [t, len], [-t, len], [-t, t], [-len, t], [-len, -t], [-t, -t]];
      const c = Math.SQRT1_2;
      return one(poly(local.map(([x, y]) => [cx + ((x * c - y * c) * w) / ss, cy + ((x * c + y * c) * h) / ss] as [number, number])));
    }
    case 'mathDivide': {
      const t = ss * 0.1;
      const dx = w * 0.3675;
      const dot = t * 1.1;
      return one(`${poly([[cx - dx, cy - t], [cx + dx, cy - t], [cx + dx, cy + t], [cx - dx, cy + t]])} ${ellipsePath(cx, cy - t * 2.6 - dot, dot, dot)} ${ellipsePath(cx, cy + t * 2.6 + dot, dot, dot)}`);
    }
    case 'mathEqual':
    case 'mathNotEqual': {
      const t = ss * 0.1;
      const dx = w * 0.3675;
      const g = t * 0.9;
      const bars = `${poly([[cx - dx, cy - g - 2 * t], [cx + dx, cy - g - 2 * t], [cx + dx, cy - g], [cx - dx, cy - g]])} ${poly([[cx - dx, cy + g], [cx + dx, cy + g], [cx + dx, cy + g + 2 * t], [cx - dx, cy + g + 2 * t]])}`;
      if (preset === 'mathEqual') return one(bars);
      const s = t * 0.9;
      const slash = poly([[cx + w * 0.12 - s, h * 0.1], [cx + w * 0.12 + s, h * 0.1], [cx - w * 0.12 + s, h * 0.9], [cx - w * 0.12 - s, h * 0.9]]);
      return { paths: [fillPath(bars), fillPath(slash)] };
    }

    // Flowchart
    case 'flowChartPredefinedProcess':
      return {
        paths: [fillPath(poly([[0, 0], [w, 0], [w, h], [0, h]])), { d: `M${P(w / 8, 0)} L${P(w / 8, h)} M${P((w * 7) / 8, 0)} L${P((w * 7) / 8, h)}`, fill: 'none', stroke: true }],
        text: [w / 8, 0, (w * 3) / 4, h],
      };
    case 'flowChartInternalStorage':
      return { paths: [fillPath(poly([[0, 0], [w, 0], [w, h], [0, h]])), { d: `M${P(w / 8, 0)} L${P(w / 8, h)} M${P(0, h / 8)} L${P(w, h / 8)}`, fill: 'none', stroke: true }], text: [w / 8, h / 8, (w * 7) / 8, (h * 7) / 8] };
    case 'flowChartDocument':
      return one(`M0 0 L${P(w, 0)} L${P(w, h * 0.8)} C${P(w * 0.7, h * 0.62)} ${P(w * 0.38, h * 1.06)} ${P(0, h * 0.9)} Z`, [0, 0, w, h * 0.8]);
    case 'flowChartMultidocument': {
      const doc = (x: number, y: number, dw: number, dh: number) => `M${P(x, y)} L${P(x + dw, y)} L${P(x + dw, y + dh * 0.8)} C${P(x + dw * 0.7, y + dh * 0.62)} ${P(x + dw * 0.38, y + dh * 1.06)} ${P(x, y + dh * 0.9)} Z`;
      const dw = w * 0.85;
      const dh = h * 0.82;
      return { paths: [fillPath(doc(w * 0.15, 0, dw, dh)), fillPath(doc(w * 0.075, h * 0.09, dw, dh)), fillPath(doc(0, h * 0.18, dw, dh))], text: [0, h * 0.18, dw, dh * 0.8] };
    }
    case 'flowChartTerminator': {
      const rx = Math.min(w * 0.16, w / 2);
      return one(`M${P(rx, 0)} L${P(w - rx, 0)} A${r(rx)} ${r(cy)} 0 0 1 ${P(w - rx, h)} L${P(rx, h)} A${r(rx)} ${r(cy)} 0 0 1 ${P(rx, 0)} Z`, [rx * 0.6, 0, w - rx * 1.2, h]);
    }
    case 'flowChartPreparation':
      return one(poly([[0, cy], [w * 0.2, 0], [w * 0.8, 0], [w, cy], [w * 0.8, h], [w * 0.2, h]]), [w * 0.2, 0, w * 0.6, h]);
    case 'flowChartManualInput':
      return one(poly([[0, h * 0.2], [w, 0], [w, h], [0, h]]), [0, h * 0.2, w, h * 0.8]);
    case 'flowChartManualOperation':
      return one(poly([[0, 0], [w, 0], [w * 0.8, h], [w * 0.2, h]]), [w * 0.2, 0, w * 0.6, h]);
    case 'flowChartOffpageConnector':
      return one(poly([[0, 0], [w, 0], [w, h * 0.8], [cx, h], [0, h * 0.8]]), [0, 0, w, h * 0.8]);
    case 'flowChartPunchedCard':
      return one(poly([[0, h * 0.2], [w * 0.2, 0], [w, 0], [w, h], [0, h]]), [0, h * 0.2, w, h * 0.8]);
    case 'flowChartPunchedTape':
      return one(`M0 ${r(h * 0.1)} Q${P(w * 0.25, h * 0.3)} ${P(cx, h * 0.1)} Q${P(w * 0.75, -h * 0.1)} ${P(w, h * 0.1)} L${P(w, h * 0.9)} Q${P(w * 0.75, h * 0.7)} ${P(cx, h * 0.9)} Q${P(w * 0.25, h * 1.1)} ${P(0, h * 0.9)} Z`, [0, h * 0.2, w, h * 0.6]);
    case 'flowChartSummingJunction':
    case 'flowChartOr': {
      const [ax, ay] = onEllipse(cx, cy, cx, cy, 225);
      const [bx, by] = onEllipse(cx, cy, cx, cy, 45);
      const [ex, ey] = onEllipse(cx, cy, cx, cy, 315);
      const [fx, fy] = onEllipse(cx, cy, cx, cy, 135);
      const cross = preset === 'flowChartOr' ? `M${P(cx, 0)} L${P(cx, h)} M${P(0, cy)} L${P(w, cy)}` : `M${P(ax, ay)} L${P(bx, by)} M${P(ex, ey)} L${P(fx, fy)}`;
      return { paths: [fillPath(ellipsePath(cx, cy, cx, cy)), { d: cross, fill: 'none', stroke: true }] };
    }
    case 'flowChartCollate':
      return one(`${poly([[0, 0], [w, 0], [cx, cy]])} ${poly([[cx, cy], [w, h], [0, h]])}`);
    case 'flowChartSort':
      return { paths: [fillPath(poly([[cx, 0], [w, cy], [cx, h], [0, cy]])), { d: `M0 ${r(cy)} L${P(w, cy)}`, fill: 'none', stroke: true }], text: [w * 0.25, h * 0.25, w * 0.5, h * 0.5] };
    case 'flowChartOnlineStorage':
      return one(`M${P(w / 6, 0)} L${P(w, 0)} A${r(w / 6)} ${r(cy)} 0 0 0 ${P(w, h)} L${P(w / 6, h)} A${r(w / 6)} ${r(cy)} 0 0 1 ${P(w / 6, 0)} Z`, [w / 6, 0, (w * 2) / 3, h]);
    case 'flowChartDelay':
      return one(`M0 0 L${P(cx, 0)} A${r(cx)} ${r(cy)} 0 0 1 ${P(cx, h)} L0 ${r(h)} Z`, [0, h * 0.15, w * 0.85, h * 0.7]);
    case 'flowChartMagneticTape': {
      const [px, py] = onEllipse(cx, cy, cx, cy, 45);
      return one(`M${P(cx, h)} A${r(cx)} ${r(cy)} 0 1 1 ${P(px, py)} L${P(w, py)} L${P(w, h)} Z`, [w * 0.15, h * 0.15, w * 0.7, h * 0.7]);
    }
    case 'flowChartDisplay':
      return one(`M0 ${r(cy)} L${P(w / 6, 0)} L${P((w * 5) / 6, 0)} A${r(w / 6)} ${r(cy)} 0 0 1 ${P((w * 5) / 6, h)} L${P(w / 6, h)} Z`, [w / 6, 0, (w * 2) / 3, h]);
    default:
      return one(poly([[0, 0], [w, 0], [w, h], [0, h]]));
  }
}

function hexToRgb(hex: string): [number, number, number] {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  const n = m ? parseInt(m[1], 16) : 0x4f81bd;
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function mix(hex: string, toward: number, amount: number): string {
  const [r0, g0, b0] = hexToRgb(hex);
  const f = (c: number) => Math.round(c + (toward - c) * amount).toString(16).padStart(2, '0');
  return `#${f(r0)}${f(g0)}${f(b0)}`;
}

function shadeColor(base: string, shade: Shade): string {
  switch (shade) {
    case 'darken':
      return mix(base, 0, 0.4);
    case 'darkenLess':
      return mix(base, 0, 0.2);
    case 'lighten':
      return mix(base, 255, 0.4);
    case 'lightenLess':
      return mix(base, 255, 0.2);
    default:
      return base;
  }
}

const escXml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Wraps text to a width (approximate glyph widths, no layout engine needed). */
export function wrapText(s: string, width: number, size: number): string[] {
  const out: string[] = [];
  const charW = size * 0.52;
  const max = Math.max(1, Math.floor(width / charW));
  for (const para of s.split('\n')) {
    let line = '';
    for (const word of para.split(/(\s+)/)) {
      if (!word) continue;
      if ((line + word).length > max && line.trim()) {
        out.push(line.trimEnd());
        line = word.trimStart();
      } else line += word;
      while (line.length > max) {
        out.push(line.slice(0, max));
        line = line.slice(max);
      }
    }
    out.push(line.trimEnd());
  }
  return out;
}

/** The shape as an SVG document of `w` x `h` pixels (pointers and thick lines may overflow it). */
export function shapeSvg(shape: Pick<SheetShape, 'preset' | 'fill' | 'line' | 'fontColor' | 'text' | 'arrowStart' | 'arrowEnd' | 'rotation' | 'flipH' | 'flipV'>, w: number, h: number): string {
  const W = Math.max(1, w);
  const H = Math.max(1, h);
  const geo = geometry(shape.preset, W, H);
  const isLine = LINE_PRESETS.has(shape.preset);
  const lineColor = shape.line?.color ?? (isLine ? '#385D8A' : 'none');
  const lineWidth = shape.line ? Math.max(0.5, shape.line.width) : isLine ? 1.5 : 0;
  const dash = shape.line?.dash === 'dashed' ? ` stroke-dasharray="${lineWidth * 4} ${lineWidth * 3}"` : shape.line?.dash === 'dotted' ? ` stroke-dasharray="${lineWidth} ${lineWidth * 2}"` : shape.line?.dash === 'dash-dot' ? ` stroke-dasharray="${lineWidth * 4} ${lineWidth * 2} ${lineWidth} ${lineWidth * 2}"` : '';
  const parts: string[] = [];
  for (const p of geo.paths) {
    const fill = p.fill === 'none' || !shape.fill ? 'none' : shadeColor(shape.fill, p.fill);
    const stroke = p.stroke && shape.line !== null && lineWidth > 0 ? ` stroke="${lineColor}" stroke-width="${lineWidth}" stroke-linejoin="round"${dash}` : '';
    parts.push(`<path d="${p.d}" fill="${fill}"${p.evenOdd ? ' fill-rule="evenodd"' : ''}${stroke}/>`);
  }
  if (geo.ends && shape.line !== null) {
    const head = (x: number, y: number, dx: number, dy: number) => {
      const len = Math.hypot(dx, dy) || 1;
      const ux = dx / len;
      const uy = dy / len;
      const L = Math.max(8, lineWidth * 4);
      const hw = Math.max(4, lineWidth * 2.2);
      const bx = x - ux * L;
      const by = y - uy * L;
      return `<path d="M${P(x, y)} L${P(bx - uy * hw, by + ux * hw)} L${P(bx + uy * hw, by - ux * hw)} Z" fill="${lineColor}"/>`;
    };
    if (shape.arrowEnd) parts.push(head(...geo.ends.end));
    if (shape.arrowStart) parts.push(head(...geo.ends.start));
  }
  if (shape.text && !isLine) {
    const [tx, ty, tw, th] = geo.text ?? [0, 0, W, H];
    const size = 14;
    const lines = wrapText(shape.text, Math.max(10, tw - 12), size);
    const lh = size * 1.25;
    const top = ty + th / 2 - (lines.length * lh) / 2 + size * 0.85;
    parts.push(
      `<text x="${r(tx + tw / 2)}" y="${r(top)}" font-size="${size}" text-anchor="middle" fill="${escXml(shape.fontColor || '#000000')}">${lines
        .map((l, i) => `<tspan x="${r(tx + tw / 2)}" dy="${i ? r(lh) : 0}">${escXml(l)}</tspan>`)
        .join('')}</text>`,
    );
  }
  const t: string[] = [];
  if (shape.rotation) t.push(`rotate(${r(shape.rotation)} ${r(W / 2)} ${r(H / 2)})`);
  if (shape.flipH || shape.flipV) t.push(`translate(${shape.flipH ? r(W) : 0} ${shape.flipV ? r(H) : 0}) scale(${shape.flipH ? -1 : 1} ${shape.flipV ? -1 : 1})`);
  const body = t.length ? `<g transform="${t.join(' ')}">${parts.join('')}</g>` : parts.join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${r(W)} ${r(H)}" width="${r(W)}" height="${r(H)}" overflow="visible" font-family="Carlito, Calibri, Arimo, Arial, sans-serif">${body}</svg>`;
}

/** Small preview of a gallery item (gallery buttons). */
export function shapeIconSvg(it: GalleryItem, size = 20): string {
  const line = LINE_PRESETS.has(it.preset);
  return shapeSvg({ preset: it.preset, fill: line ? null : '#ffffff', line: { color: '#333333', width: 1.2 }, fontColor: '#000', arrowStart: it.arrowStart, arrowEnd: it.arrowEnd }, size, size)
    .replace(/<svg /, '<svg class="sl-shape-icon" ');
}
