/**
 * jsdom has no canvas, no observers and no scrolling: stand-ins for what the grid and the interface
 * call, so the mount() scenarios can run the whole editor (setup file of the React and Vue tests).
 */

const noop = () => undefined;
const gradient = { addColorStop: noop };
const pixels = (width: number, height: number) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) });

/** A 2D context that takes every drawing call, keeps what is set on it and measures 7 px per character. */
function context2d(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const props: Record<PropertyKey, unknown> = {
    canvas,
    font: '10px sans-serif',
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    lineCap: 'butt',
    lineJoin: 'miter',
    lineDashOffset: 0,
    miterLimit: 10,
    globalAlpha: 1,
    globalCompositeOperation: 'source-over',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    filter: 'none',
    imageSmoothingEnabled: true,
    shadowBlur: 0,
    shadowColor: 'rgba(0, 0, 0, 0)',
    shadowOffsetX: 0,
    shadowOffsetY: 0,
    measureText: (text: string) => {
      const width = text.length * 7;
      return { width, actualBoundingBoxLeft: 0, actualBoundingBoxRight: width, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 9, fontBoundingBoxDescent: 3 };
    },
    createLinearGradient: () => gradient,
    createRadialGradient: () => gradient,
    createConicGradient: () => gradient,
    createPattern: () => ({ setTransform: noop }),
    createImageData: (width: number, height: number) => pixels(width, height),
    getImageData: (_x: number, _y: number, width: number, height: number) => pixels(width, height),
    getLineDash: () => [],
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    isPointInPath: () => false,
    isPointInStroke: () => false,
  };
  return new Proxy(props, {
    get: (target, key) => (key in target ? target[key] : noop),
    set: (target, key, value) => {
      target[key] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
}

const contexts = new WeakMap<HTMLCanvasElement, CanvasRenderingContext2D>();
Object.defineProperty(HTMLCanvasElement.prototype, 'getContext', {
  configurable: true,
  writable: true,
  value(this: HTMLCanvasElement, type: string) {
    if (type !== '2d') return null;
    let ctx = contexts.get(this);
    if (!ctx) contexts.set(this, (ctx = context2d(this)));
    return ctx;
  },
});
Object.defineProperty(HTMLCanvasElement.prototype, 'toDataURL', { configurable: true, writable: true, value: () => 'data:image/png;base64,' });

/** Observers that never report: jsdom has no layout to observe. */
class SilentObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
  takeRecords(): [] {
    return [];
  }
}
for (const name of ['ResizeObserver', 'IntersectionObserver']) {
  if (!(name in globalThis)) Object.defineProperty(globalThis, name, { configurable: true, writable: true, value: SilentObserver });
}

if (typeof Element.prototype.scrollIntoView !== 'function') {
  Object.defineProperty(Element.prototype, 'scrollIntoView', { configurable: true, writable: true, value: noop });
}

/** Object URLs for pictures (jsdom does not load them, so they point nowhere). */
if (typeof URL.createObjectURL !== 'function') {
  let next = 0;
  Object.defineProperty(URL, 'createObjectURL', { configurable: true, writable: true, value: () => `blob:http://localhost/${++next}` });
  Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, writable: true, value: noop });
}
