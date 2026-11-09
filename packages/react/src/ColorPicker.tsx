/**
 * Colour picker pop-up: a palette grid, a spectrum and RGB sliders, with Reset and Done.
 */
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { hexToRgb, normalizeHex, PALETTE_ROWS, rgbToHex } from '@sumline/core/views';
import { usePersistent, useUI } from './context';

function Spectrum({ value, onPick }: { value: string; onPick: (hex: string) => void }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const [point, setPoint] = useState<{ x: number; y: number } | null>(null);
  useEffect(() => {
    const c = canvas.current;
    if (!c) return;
    const ctx = c.getContext('2d');
    if (!ctx) return;
    const w = c.width;
    const h = c.height;
    const hue = ctx.createLinearGradient(0, 0, w, 0);
    for (let i = 0; i <= 6; i++) hue.addColorStop(i / 6, `hsl(${i * 60}, 100%, 50%)`);
    ctx.fillStyle = hue;
    ctx.fillRect(0, 0, w, h);
    const shade = ctx.createLinearGradient(0, 0, 0, h);
    shade.addColorStop(0, 'rgba(255,255,255,1)');
    shade.addColorStop(0.5, 'rgba(255,255,255,0)');
    shade.addColorStop(0.5, 'rgba(0,0,0,0)');
    shade.addColorStop(1, 'rgba(0,0,0,1)');
    ctx.fillStyle = shade;
    ctx.fillRect(0, 0, w, h);
  }, []);
  void value;
  const pick = (e: React.PointerEvent) => {
    const c = canvas.current;
    if (!c) return;
    const rect = c.getBoundingClientRect();
    const x = Math.max(0, Math.min(rect.width - 1, e.clientX - rect.left));
    const y = Math.max(0, Math.min(rect.height - 1, e.clientY - rect.top));
    setPoint({ x, y });
    const data = c.getContext('2d')?.getImageData(Math.floor((x / rect.width) * c.width), Math.floor((y / rect.height) * c.height), 1, 1).data;
    if (data) onPick(rgbToHex(data[0], data[1], data[2]));
  };
  return (
    <div
      className="sl-color-spectrum"
      onPointerDown={(e) => {
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        pick(e);
      }}
      onPointerMove={(e) => {
        if (e.buttons) pick(e);
      }}
    >
      <canvas ref={canvas} width={232} height={160} />
      {point && <div className="sl-color-point" style={{ left: point.x, top: point.y }} />}
    </div>
  );
}

export interface ColorPickerProps {
  anchor: DOMRect;
  value: string | null;
  onChange: (hex: string | null) => void;
  onClose: () => void;
  /** Applies immediately on every pick (true) or only on Done (false). */
  live?: boolean;
  testId?: string;
}

export function ColorPicker({ anchor, value, onChange, onClose, live = true, testId }: ColorPickerProps) {
  const { layer } = useUI();
  const ref = useRef<HTMLDivElement>(null);
  const [tab, setTab] = useState<'grid' | 'spectrum' | 'sliders'>('grid');
  const [current, setCurrent] = useState<string>(value ?? '#000000');
  const [text, setText] = useState(value ?? '');
  const [recent, setRecent] = usePersistent<string[]>('recent-colors', []);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    let x = anchor.left;
    let y = anchor.bottom + 4;
    if (x + w > window.innerWidth - 4) x = Math.max(4, window.innerWidth - w - 4);
    if (y + h > window.innerHeight - 4) y = Math.max(4, anchor.top - h - 4);
    el.style.left = `${x}px`;
    el.style.top = `${y}px`;
  }, [anchor]);

  useEffect(() => {
    const down = (e: PointerEvent) => {
      if (!(e.target as HTMLElement).closest?.('.sl-color')) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    const id = window.setTimeout(() => {
      window.addEventListener('pointerdown', down, true);
      window.addEventListener('keydown', key);
    }, 0);
    return () => {
      window.clearTimeout(id);
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key);
    };
  }, [onClose]);

  const pick = (hex: string) => {
    setCurrent(hex);
    setText(hex);
    if (live) onChange(hex);
  };
  const done = () => {
    const hex = normalizeHex(text) ?? current;
    onChange(hex);
    setRecent([hex, ...recent.filter((c) => c !== hex)].slice(0, 12));
    onClose();
  };
  const [r, g, b] = hexToRgb(current);

  return createPortal(
    <div ref={ref} className="sl-color" role="dialog" aria-label="Colour" data-testid={testId}>
      <div className="sl-color-head">
        <button type="button" data-testid="color-reset" onClick={() => { onChange(null); onClose(); }}>Reset</button>
        <button type="button" data-testid="color-done" onClick={done}>Done</button>
      </div>
      <div className="sl-segmented" role="tablist">
        {(['grid', 'spectrum', 'sliders'] as const).map((t) => (
          <button key={t} type="button" role="tab" aria-selected={tab === t} className={tab === t ? 'selected' : ''} onClick={() => setTab(t)}>
            {t[0].toUpperCase() + t.slice(1)}
          </button>
        ))}
      </div>
      {tab === 'grid' && (
        <div className="sl-color-grid" style={{ gridTemplateColumns: `repeat(${PALETTE_ROWS[0].length}, 1fr)` }}>
          {PALETTE_ROWS.flatMap((row, i) =>
            row.map((hex, j) => (
              <button key={`${i}-${j}`} type="button" title={hex} aria-label={hex} data-color={hex} className={current.toUpperCase() === hex ? 'selected' : ''} style={{ background: hex }} onClick={() => pick(hex)} />
            )),
          )}
        </div>
      )}
      {tab === 'spectrum' && <Spectrum value={current} onPick={pick} />}
      {tab === 'sliders' && (
        <div className="sl-color-sliders">
          {(['Red', 'Green', 'Blue'] as const).map((label, k) => {
            const vals = [r, g, b];
            return (
              <label key={label}>
                {label}
                <input
                  type="range"
                  min={0}
                  max={255}
                  value={vals[k]}
                  aria-label={label}
                  onChange={(e) => {
                    const next = [...vals];
                    next[k] = Number(e.target.value);
                    pick(rgbToHex(next[0], next[1], next[2]));
                  }}
                />
                <span>{vals[k]}</span>
              </label>
            );
          })}
        </div>
      )}
      {recent.length > 0 && (
        <div className="sl-color-recent" aria-label="Recent colours">
          {recent.map((hex) => (
            <button key={hex} type="button" title={hex} style={{ background: hex }} onClick={() => pick(hex)} />
          ))}
        </div>
      )}
      <div className="sl-color-foot">
        <span className="swatch" style={{ background: normalizeHex(text) ?? current }} />
        <input
          value={text}
          placeholder="#000000"
          aria-label="Hex colour"
          onChange={(e) => {
            setText(e.target.value);
            const hex = normalizeHex(e.target.value);
            if (hex) setCurrent(hex);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') done();
          }}
        />
      </div>
    </div>,
    layer,
  );
}
