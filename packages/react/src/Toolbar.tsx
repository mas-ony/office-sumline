/**
 * The toolbar: undo, redo, save | font, size, alignment, bold, colours, vertical alignment, merge,
 * borders | image, search, chart, validation, shape. Items that do not fit move to a "more" panel.
 */
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import {
  ALIGN_ICONS, BORDER_KINDS, type BorderKind, type BorderLine, borderStyleFor, cmd, fitToolbar, type IconName, LINE_KINDS, type LineKind,
  rgbSpec, TOOLBAR_FONTS, TOOLBAR_SIZES, type ToolKey, VALIGN_ICONS,
} from '@sumline/core/views';
import { Icon } from './Icon';
import { useUI } from './context';
import { ColorPicker } from './ColorPicker';

const keep = (e: React.MouseEvent | React.PointerEvent) => e.preventDefault();

export function ToolButton({ icon, label, onClick, active, disabled, testId, children, buttonRef, expanded }: {
  icon?: IconName;
  label: string;
  onClick?: (e: React.MouseEvent<HTMLButtonElement>) => void;
  active?: boolean;
  disabled?: boolean;
  testId?: string;
  children?: ReactNode;
  buttonRef?: React.Ref<HTMLButtonElement>;
  expanded?: boolean;
}) {
  return (
    <button
      ref={buttonRef}
      type="button"
      className={`sl-tool${active ? ' active' : ''}`}
      title={label}
      aria-label={label}
      aria-pressed={active}
      aria-expanded={expanded}
      disabled={disabled}
      data-testid={testId}
      onMouseDown={keep}
      onClick={onClick}
    >
      {icon && <Icon name={icon} />}
      {children}
    </button>
  );
}

/** A pop-up anchored under a toolbar button. */
export function Popup({ anchor, onClose, children, className, label, testId }: { anchor: DOMRect; onClose: () => void; children: ReactNode; className?: string; label?: string; testId?: string }) {
  const { layer } = useUI();
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    let x = anchor.left;
    let y = anchor.bottom + 2;
    if (x + w > window.innerWidth - 4) x = Math.max(4, window.innerWidth - w - 4);
    if (y + h > window.innerHeight - 4) y = Math.max(4, anchor.top - h - 2);
    el.style.left = `${x}px`;
    el.style.top = `${y}px`;
  }, [anchor]);
  useEffect(() => {
    const down = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest('.sl-picker') || t.closest('.sl-color')) return;
      onClose();
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
  return createPortal(
    <div ref={ref} className={`sl-picker${className ? ` ${className}` : ''}`} role="listbox" aria-label={label} data-testid={testId} onMouseDown={keep}>
      {children}
    </div>,
    layer,
  );
}

/** A toolbar button that opens a pop-up. */
function PickerTool({ label, testId, content, children, className }: { label: string; testId: string; content: (close: () => void) => ReactNode; children: ReactNode; className?: string }) {
  const btn = useRef<HTMLButtonElement>(null);
  const [rect, setRect] = useState<DOMRect | null>(null);
  return (
    <>
      <button
        ref={btn}
        type="button"
        className={`sl-tool${className ? ` ${className}` : ''}`}
        title={label}
        aria-label={label}
        aria-haspopup="listbox"
        aria-expanded={!!rect}
        data-testid={testId}
        onMouseDown={keep}
        onClick={() => setRect(rect ? null : btn.current!.getBoundingClientRect())}
      >
        {children}
      </button>
      {rect && (
        <Popup anchor={rect} onClose={() => setRect(null)} label={label} testId={`${testId}-list`}>
          {content(() => setRect(null))}
        </Popup>
      )}
    </>
  );
}

function ColorTool({ kind }: { kind: 'text' | 'fill' }) {
  const ui = useUI();
  const btn = useRef<HTMLButtonElement>(null);
  const [rect, setRect] = useState<DOMRect | null>(null);
  const value = kind === 'text' ? ui.state.format.color : ui.state.format.fill;
  const apply = (hex: string | null) => {
    if (kind === 'text') cmd.format(ui, { font: { color: hex ? rgbSpec(hex) ?? null : null } }, 'Text color');
    else cmd.format(ui, { fill: hex ? { pattern: 'solid', fg: rgbSpec(hex) } : null }, 'Fill color');
  };
  const label = kind === 'text' ? 'Text Color' : 'Background Color';
  return (
    <>
      <button
        ref={btn}
        type="button"
        className="sl-tool sl-tool-color"
        title={label}
        aria-label={label}
        aria-expanded={!!rect}
        data-testid={kind === 'text' ? 'tb-text-color' : 'tb-fill-color'}
        disabled={ui.state.readOnly}
        onMouseDown={keep}
        onClick={() => setRect(rect ? null : btn.current!.getBoundingClientRect())}
      >
        <Icon name={kind === 'text' ? 'format_color_text' : 'format_color_fill'} />
        <span className="swatch" style={{ background: value ?? (kind === 'text' ? '#000000' : 'transparent') }} />
      </button>
      {rect && <ColorPicker anchor={rect} value={value} onChange={apply} onClose={() => setRect(null)} live={false} testId={`${kind}-color-picker`} />}
    </>
  );
}

function BorderPicker({ close }: { close: () => void }) {
  const ui = useUI();
  const [color, setColor] = useState('#000000');
  const [width, setWidth] = useState(1);
  const [kind, setKind] = useState<LineKind>('solid');
  const [colorRect, setColorRect] = useState<DOMRect | null>(null);
  const [menu, setMenu] = useState<'width' | 'kind' | null>(null);
  const colorBtn = useRef<HTMLButtonElement>(null);
  const multi = ui.state.multiCell;
  const apply = (k: BorderKind) => {
    const line: BorderLine = { style: borderStyleFor(width, kind), color: rgbSpec(color) };
    if (ui.grid?.isEditing()) ui.grid.commitEdit();
    ui.doc.setBorder(k, k === 'none' ? null : line);
    close();
  };
  const sample = (w: number, k: LineKind) => (
    <span className="sl-line-sample" style={k === 'solid' ? { height: w, background: '#333' } : { borderTop: `${k === 'double' ? 3 : 2}px ${k} #333` }} />
  );
  return (
    <div className="sl-border-picker" data-testid="border-picker">
      <div className="sl-picker-icons cols-5">
        {BORDER_KINDS.map((b) => (
          <button
            key={b.kind}
            type="button"
            className="sl-icon-btn"
            title={b.label}
            aria-label={b.label}
            data-testid={`border-${b.kind}`}
            disabled={!multi && (b.kind === 'inside' || b.kind === 'horizontal' || b.kind === 'vertical')}
            onClick={() => apply(b.kind)}
          >
            <Icon name={b.icon} />
          </button>
        ))}
      </div>
      <div className="sl-border-options">
        <button ref={colorBtn} type="button" className="sl-icon-btn" title="Border color" aria-label="Border color" data-testid="border-color" onClick={() => setColorRect(colorRect ? null : colorBtn.current!.getBoundingClientRect())}>
          <Icon name="color_lens" />
          <span style={{ position: 'absolute', width: 14, height: 3, marginTop: 22, background: color }} />
        </button>
        <div style={{ position: 'relative' }}>
          <button type="button" className="sl-select-btn" aria-label="Border width" data-testid="border-width" onClick={() => setMenu(menu === 'width' ? null : 'width')}>
            {sample(width, 'solid')}
            <Icon name="arrow_drop_down" />
          </button>
          {menu === 'width' && (
            <div className="sl-picker" style={{ position: 'absolute', top: '100%', left: 0 }}>
              {[1, 2, 3, 4, 5].map((w) => (
                <button key={w} type="button" className={`sl-picker-item${w === width ? ' selected' : ''}`} aria-label={`${w}px`} onClick={() => { setWidth(w); setMenu(null); }}>
                  {sample(w, 'solid')}
                </button>
              ))}
            </div>
          )}
        </div>
        <div style={{ position: 'relative' }}>
          <button type="button" className="sl-select-btn" aria-label="Border style" data-testid="border-style" onClick={() => setMenu(menu === 'kind' ? null : 'kind')}>
            {sample(2, kind)}
            <Icon name="arrow_drop_down" />
          </button>
          {menu === 'kind' && (
            <div className="sl-picker" style={{ position: 'absolute', top: '100%', left: 0 }}>
              {LINE_KINDS.map((k) => (
                <button key={k} type="button" className={`sl-picker-item${k === kind ? ' selected' : ''}`} aria-label={k} data-testid={`border-style-${k}`} onClick={() => { setKind(k); setMenu(null); }}>
                  {sample(2, k)}
                </button>
              ))}
            </div>
          )}
        </div>
      </div>
      {colorRect && <ColorPicker anchor={colorRect} value={color} onChange={(hex) => setColor(hex ?? '#000000')} onClose={() => setColorRect(null)} />}
    </div>
  );
}

export function Toolbar() {
  const ui = useUI();
  const { state, doc } = ui;
  const ro = state.readOnly;
  const f = state.format;
  const fonts = useMemo(() => {
    const set = new Set<string>(TOOLBAR_FONTS);
    if (!f.defaultFont) set.add(f.fontName);
    return [...set];
  }, [f.defaultFont, f.fontName]);
  const align = f.hAlign === 'general' ? 'left' : f.hAlign;

  const nodes: Record<ToolKey, ReactNode> = {
    undo: <ToolButton icon="undo" label="Undo" testId="tb-undo" disabled={ro || (!state.canUndo && !ui.grid?.isEditing())} onClick={() => cmd.undo(ui)} />,
    redo: <ToolButton icon="redo" label="Redo" testId="tb-redo" disabled={ro || !state.canRedo} onClick={() => cmd.redo(ui)} />,
    save: <ToolButton icon="save" label="Save" testId="tb-save" onClick={() => ui.files.save()} />,
    d1: <span className="sl-divider" />,
    font: (
      <PickerTool
        label="Font Family"
        testId="tb-font"
        className="sl-tool-font"
        content={(close) =>
          fonts.map((name) => {
            const selected = name === 'Default' ? f.defaultFont : !f.defaultFont && name === f.fontName;
            return (
              <button
                key={name}
                type="button"
                role="option"
                aria-selected={selected}
                className={`sl-picker-item${selected ? ' selected' : ''}`}
                style={{ fontFamily: name === 'Default' ? undefined : name }}
                onClick={() => {
                  close();
                  const def = doc.wb.styles.defaultStyle.font.name;
                  cmd.format(ui, { font: { name: name === 'Default' ? def : name } }, 'Font');
                }}
              >
                {name}
              </button>
            );
          })
        }
      >
        <span className="label">{f.defaultFont ? 'Default' : f.fontName}</span>
        <Icon name="arrow_drop_down" className="caret" />
      </PickerTool>
    ),
    size: (
      <PickerTool
        label="Font Size"
        testId="tb-size"
        content={(close) =>
          TOOLBAR_SIZES.map((s) => (
            <button
              key={s.label}
              type="button"
              role="option"
              aria-selected={Math.abs(f.fontSize - s.pt) < 0.01}
              className={`sl-picker-item${Math.abs(f.fontSize - s.pt) < 0.01 ? ' selected' : ''}`}
              style={{ fontSize: s.css }}
              onClick={() => {
                close();
                cmd.format(ui, { font: { size: s.pt } }, 'Font size');
              }}
            >
              {s.label}
            </button>
          ))
        }
      >
        <Icon name="format_size" />
        <Icon name="arrow_drop_down" className="caret" />
      </PickerTool>
    ),
    align: (
      <PickerTool
        label="Text Align"
        testId="tb-align"
        content={(close) => (
          <div className="sl-picker-icons cols-1">
            {(['left', 'center', 'right', 'justify'] as const).map((a) => (
              <button key={a} type="button" role="option" aria-label={`Align ${a}`} data-testid={`align-${a}`} className={`sl-icon-btn${align === a ? ' selected' : ''}`} onClick={() => { close(); cmd.format(ui, { align: { horizontal: a } }, 'Align'); }}>
                <Icon name={ALIGN_ICONS[a]} />
              </button>
            ))}
          </div>
        )}
      >
        <Icon name={ALIGN_ICONS[align] ?? 'format_align_left'} />
        <Icon name="arrow_drop_down" className="caret" />
      </PickerTool>
    ),
    bold: <ToolButton icon="format_bold" label="Bold" testId="tb-bold" active={f.bold} disabled={ro} onClick={() => cmd.toggleFont(ui, 'bold')} />,
    color: <ColorTool kind="text" />,
    fill: <ColorTool kind="fill" />,
    valign: (
      <PickerTool
        label="Vertical Align"
        testId="tb-valign"
        content={(close) => (
          <div className="sl-picker-icons cols-1">
            {(['top', 'center', 'bottom'] as const).map((a) => (
              <button key={a} type="button" role="option" aria-label={`Align ${a}`} data-testid={`valign-${a}`} className={`sl-icon-btn${f.vAlign === a ? ' selected' : ''}`} onClick={() => { close(); cmd.format(ui, { align: { vertical: a } }, 'Vertical align'); }}>
                <Icon name={VALIGN_ICONS[a]} />
              </button>
            ))}
          </div>
        )}
      >
        <Icon name={VALIGN_ICONS[f.vAlign] ?? 'vertical_align_bottom'} />
        <Icon name="arrow_drop_down" className="caret" />
      </PickerTool>
    ),
    merge: <ToolButton icon="web" label="Merge the selected cells" testId="tb-merge" active={f.merged} disabled={ro} onClick={() => void cmd.merge(ui)} />,
    border: (
      <PickerTool label="Border Style" testId="tb-border" content={(close) => <BorderPicker close={close} />}>
        <Icon name="border_all" />
        <Icon name="arrow_drop_down" className="caret" />
      </PickerTool>
    ),
    d2: <span className="sl-divider" />,
    image: <ToolButton icon="add_photo_alternate" label="Upload image" testId="tb-image" disabled={ro} onClick={() => cmd.uploadImage(ui)} />,
    search: <ToolButton icon="search" label="Toggle Search" testId="tb-search" active={ui.dialog?.kind === 'search'} onClick={() => (ui.dialog?.kind === 'search' ? ui.closeDialog() : ui.openDialog({ kind: 'search', replace: true }))} />,
    chart: <ToolButton icon="addchart" label="Insert chart" testId="tb-chart" disabled={ro} onClick={() => ui.openDialog({ kind: 'chart-insert' })} />,
    rule: <ToolButton icon="rule" label="Add new validation" testId="tb-validation" disabled={ro} onClick={() => ui.openDialog({ kind: 'validations' })} />,
    shape: <ToolButton icon="interests" label="Insert shape" testId="tb-shape" disabled={ro} onClick={() => ui.openDialog({ kind: 'shapes' })} />,
  };

  const bar = useRef<HTMLDivElement>(null);
  const [avail, setAvail] = useState(10000);
  const [more, setMore] = useState(false);
  // Tools that open a dialog close the overflow panel.
  useEffect(() => {
    if (ui.dialog) setMore(false);
  }, [ui.dialog]);
  useLayoutEffect(() => {
    const el = bar.current;
    if (!el) return;
    const measure = () => setAvail(el.clientWidth - 16);
    measure();
    const ro2 = new ResizeObserver(measure);
    ro2.observe(el);
    return () => ro2.disconnect();
  }, []);
  const { shown, overflow } = fitToolbar(avail);
  useEffect(() => {
    if (!overflow.length) setMore(false);
  }, [overflow.length]);

  return (
    <div ref={bar} className="sl-toolbar" role="toolbar" aria-label="Spreadsheet Tools" data-testid="toolbar">
      <div className="sl-toolbar-items">
        {shown.map((k) => <span key={k} style={{ display: 'contents' }}>{nodes[k]}</span>)}
      </div>
      {overflow.length > 0 && (
        <>
          <span className="sl-divider" />
          <ToolButton icon="more_vert" label="More tools" testId="tb-more" active={more} expanded={more} onClick={() => setMore(!more)} />
          {more && (
            <div className="sl-overflow-panel" data-testid="tb-overflow" onMouseDown={keep}>
              {overflow.map((k) => <span key={k} style={{ display: 'contents' }}>{nodes[k]}</span>)}
            </div>
          )}
        </>
      )}
    </div>
  );
}
