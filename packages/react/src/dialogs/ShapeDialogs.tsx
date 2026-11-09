/**
 * Shapes: the Insert Shape gallery (lines, rectangles, basic forms, block arrows, stars, callouts,
 * equation shapes, flowchart) and the Shape Settings dialog (fill, border, font colour, text).
 */
import { useState } from 'react';
import { DEFAULT_SHAPE_FILL, DEFAULT_SHAPE_LINE, type GalleryItem, LINE_PRESETS, SHAPE_GALLERY, shapeIconSvg, SheetLayout } from '@sumline/core/views';
import { ColorField } from '../ColorField';
import { useUI } from '../context';
import { Modal, Switch } from '../Modal';

const SHAPE_W = 400;
const SHAPE_H = 300;

export function ShapesDialog({ onClose }: { onClose: () => void }) {
  const { doc, grid, phone, selectObject } = useUI();

  const insert = (it: GalleryItem) => {
    const sheet = doc.sheet;
    const layout = SheetLayout.of(sheet);
    const { r, c } = doc.selection.active;
    const line = LINE_PRESETS.has(it.preset);
    const w = phone ? Math.min(SHAPE_W, Math.round(window.innerWidth * 0.7)) : SHAPE_W;
    const h = phone ? Math.round((w * 3) / 4) : SHAPE_H;
    const shape = doc.addShape(
      {
        name: `${it.label.replace(/\b\w/g, (m) => m.toUpperCase())} ${sheet.shapes.length + 1}`,
        preset: it.preset,
        fill: line ? null : DEFAULT_SHAPE_FILL,
        line: { color: DEFAULT_SHAPE_LINE, width: line ? 2 : 3 },
        fontColor: '#000000',
        arrowStart: it.arrowStart || undefined,
        arrowEnd: it.arrowEnd || undefined,
      },
      { x: layout.colX(c), y: layout.rowY(r), width: w, height: line ? Math.round(h / 2) : h },
    );
    onClose();
    if (shape) {
      grid?.scrollIntoView(shape.from.r, shape.from.c);
      selectObject(shape.id);
    }
  };

  return (
    <Modal title="Insert Shape" icon="interests" onClose={onClose} width={400} testId="shapes-dialog" bodyClassName="sl-shapes-body">
      {SHAPE_GALLERY.map((group) => (
        <div key={group.title} className="sl-shape-group">
          <h4>{group.title}</h4>
          <div className="sl-shape-grid">
            {group.items.map((it, i) => (
              <button
                key={`${it.preset}-${i}`}
                type="button"
                title={it.label}
                aria-label={it.label}
                data-testid={`shape-${it.preset}${it.arrowEnd ? (it.arrowStart ? '-both' : '-end') : ''}`}
                onClick={() => insert(it)}
                // Generated from the shape geometry (no user content).
                dangerouslySetInnerHTML={{ __html: shapeIconSvg(it) }}
              />
            ))}
          </div>
        </div>
      ))}
    </Modal>
  );
}

export function ShapeSettings({ id, onClose }: { id: string; onClose: () => void }) {
  const { doc } = useUI();
  const shape = doc.sheet.shapes.find((s) => s.id === id);
  const line = shape ? LINE_PRESETS.has(shape.preset) : false;
  const [fill, setFill] = useState(shape?.fill ?? DEFAULT_SHAPE_FILL);
  const [transparent, setTransparent] = useState(shape ? shape.fill === null : false);
  const [border, setBorder] = useState(shape?.line?.color ?? DEFAULT_SHAPE_LINE);
  const [width, setWidth] = useState(shape?.line ? String(Math.round(shape.line.width)) : '0');
  const [fontColor, setFontColor] = useState(shape?.fontColor ?? '#000000');
  const [text, setText] = useState(shape?.text ?? '');
  if (!shape) return null;

  const save = () => {
    const w = Number(width);
    doc.updateShape(id, {
      fill: line || transparent ? null : fill || DEFAULT_SHAPE_FILL,
      line: w > 0 ? { color: border || DEFAULT_SHAPE_LINE, width: w, dash: shape.line?.dash } : null,
      fontColor: fontColor || '#000000',
      text: line ? undefined : text || undefined,
    });
    onClose();
  };

  return (
    <Modal title="Shape Settings" icon="interests" onClose={onClose} width={380} testId="shape-settings">
      <div className="sl-form">
        {!line && (
          <div className="sl-row sl-row-center">
            <div className="sl-field">
              <label>Background Color</label>
              <ColorField value={fill} disabled={transparent} testId="shape-fill" onChange={(hex) => setFill(hex)} />
            </div>
            <Switch label="Transparent" checked={transparent} testId="shape-transparent" onChange={setTransparent} />
          </div>
        )}
        <div className="sl-row">
          <div className="sl-field">
            <label>Border Color</label>
            <ColorField value={border} testId="shape-border" onChange={(hex) => setBorder(hex)} />
          </div>
          <div className="sl-field">
            <label htmlFor="shape-width">Border Width</label>
            <select id="shape-width" className="sl-select" value={width} data-testid="shape-border-width" onChange={(e) => setWidth(e.target.value)}>
              {['0', '1', '2', '3', '4', '5', '6', '8', '10'].map((v) => <option key={v} value={v}>{v}</option>)}
            </select>
          </div>
        </div>
        {!line && (
          <>
            <div className="sl-row">
              <div className="sl-field">
                <label>Font Color</label>
                <ColorField value={fontColor} testId="shape-font-color" onChange={(hex) => setFontColor(hex)} />
              </div>
            </div>
            <div className="sl-field">
              <label htmlFor="shape-text">Text</label>
              <textarea id="shape-text" className="sl-textarea" rows={2} value={text} data-testid="shape-text" onChange={(e) => setText(e.target.value)} />
            </div>
          </>
        )}
      </div>
      <div className="sl-dialog-buttons">
        <button type="button" className="sl-button wide" data-testid="shape-save" onClick={save}>Save</button>
        <button type="button" className="sl-button wide" onClick={onClose}>Cancel</button>
      </div>
    </Modal>
  );
}
