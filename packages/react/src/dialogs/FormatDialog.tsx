/**
 * Format cells: the Mask tab (General, Number, Currency, Accounting, Percentage, Text, Custom with
 * a live sample) and the Editor tab (how cells are edited: text, numbers, percentages, a drop-down
 * list, a checkbox or a date, and text wrapping).
 */
import { useMemo, useState } from 'react';
import { applyCellEditor, CUSTOM_FORMATS, EDITOR_TYPES, type EditorType, FORMAT_CATEGORIES, formatValue, maskCode, maskFromCode, type MaskState } from '@sumline/core/views';
import { useUI } from '../context';
import { Modal, Switch } from '../Modal';

export function FormatDialog({ onClose, initialTab = 'mask' }: { onClose: () => void; initialTab?: 'mask' | 'editor' }) {
  const ui = useUI();
  const { doc, state } = ui;
  const [tab, setTab] = useState<'mask' | 'editor'>(initialTab);
  const [mask, setMask] = useState<MaskState>(() => maskFromCode(state.format.numFmt));
  const [editor, setEditor] = useState<EditorType>('text');
  const [wrap, setWrap] = useState(state.format.wrap === 'wrap');
  const [options, setOptions] = useState('');
  const { r, c } = doc.selection.active;
  const value = doc.cell(r, c)?.v;
  const sampleValue = typeof value === 'number' ? value : 1234;
  const code = maskCode(mask);
  const sample = useMemo(() => {
    try {
      const f = formatValue(mask.category === 'text' && typeof value === 'string' ? value : sampleValue, code, doc.wb.date1904);
      return { text: f.text, color: f.color };
    } catch {
      return { text: String(sampleValue), color: undefined };
    }
  }, [code, sampleValue, value, mask.category, doc]);
  const set = (patch: Partial<MaskState>) => setMask((m) => ({ ...m, ...patch }));

  const apply = () => {
    if (ui.grid?.isEditing()) ui.grid.commitEdit();
    doc.transact('Format cells', () => {
      if (tab === 'mask') doc.setNumberFormat(code);
      else applyCellEditor(doc, editor, options, wrap);
    });
    onClose();
  };

  return (
    <Modal
      title="Format"
      icon="settings"
      onClose={onClose}
      width={500}
      testId="format-dialog"
      footer={
        <>
          <button type="button" className="sl-button" data-testid="format-ok" onClick={apply}>Ok</button>
          <button type="button" className="sl-button" onClick={onClose}>Cancel</button>
        </>
      }
    >
      <div className="sl-subtabs" role="tablist">
        <button type="button" role="tab" aria-selected={tab === 'mask'} className={tab === 'mask' ? 'selected' : ''} onClick={() => setTab('mask')}>Mask</button>
        <button type="button" role="tab" aria-selected={tab === 'editor'} className={tab === 'editor' ? 'selected' : ''} data-testid="format-editor-tab" onClick={() => setTab('editor')}>Editor</button>
      </div>
      {tab === 'mask' ? (
        <div className="sl-format">
          <div className="sl-list" role="listbox" aria-label="Category" data-testid="format-categories">
            {FORMAT_CATEGORIES.map((cat) => (
              <div
                key={cat.key}
                role="option"
                aria-selected={mask.category === cat.key}
                tabIndex={0}
                className={`sl-list-item${mask.category === cat.key ? ' selected' : ''}`}
                onClick={() => set({ category: cat.key, custom: cat.key === 'custom' ? code : mask.custom })}
                onKeyDown={(e) => {
                  const i = FORMAT_CATEGORIES.findIndex((x) => x.key === mask.category);
                  if (e.key === 'ArrowDown') set({ category: FORMAT_CATEGORIES[Math.min(FORMAT_CATEGORIES.length - 1, i + 1)].key });
                  else if (e.key === 'ArrowUp') set({ category: FORMAT_CATEGORIES[Math.max(0, i - 1)].key });
                }}
              >
                {cat.label}
              </div>
            ))}
          </div>
          <div className="sl-form" style={{ alignContent: 'start' }}>
            <div className="sl-field">
              <label>Sample</label>
              <input className="sl-input" readOnly value={sample.text} data-testid="format-sample" style={{ color: sample.color }} />
            </div>
            {['number', 'currency', 'accounting', 'percentage'].includes(mask.category) && (
              <div className="sl-field">
                <label htmlFor="sl-dec">Decimal Place:</label>
                <input id="sl-dec" className="sl-input" type="number" min={0} max={30} inputMode="numeric" value={mask.decimals} data-testid="format-decimals" onChange={(e) => set({ decimals: e.target.value })} />
              </div>
            )}
            {mask.category === 'number' && <Switch label="Use thousands separator" checked={mask.thousands} onChange={(v) => set({ thousands: v })} testId="format-thousands" />}
            {(mask.category === 'currency' || mask.category === 'accounting') && (
              <div className="sl-field">
                <label htmlFor="sl-sym">Symbol:</label>
                <input id="sl-sym" className="sl-input" value={mask.symbol} data-testid="format-symbol" onChange={(e) => set({ symbol: e.target.value })} />
              </div>
            )}
            {(mask.category === 'number' || mask.category === 'currency') && (
              <div className="sl-field">
                <label>Negative Numbers:</label>
                <div className="sl-list" role="listbox" aria-label="Negative numbers" style={{ maxHeight: 110 }}>
                  {['-1234', '1234', '(1234)', '(1234)'].map((t, i) => (
                    <div key={i} role="option" aria-selected={mask.negative === i} className={`sl-list-item${mask.negative === i ? ' selected' : ''}`} style={{ color: mask.negative === i ? undefined : i % 2 ? '#d32f2f' : undefined }} onClick={() => set({ negative: i })}>
                      {t}
                    </div>
                  ))}
                </div>
              </div>
            )}
            {mask.category === 'custom' && (
              <div className="sl-field">
                <label htmlFor="sl-type">Type:</label>
                <input id="sl-type" className="sl-input" value={mask.custom} spellCheck={false} data-testid="format-custom" onChange={(e) => set({ custom: e.target.value })} />
                <div className="sl-list codes" role="listbox" aria-label="Format codes">
                  {CUSTOM_FORMATS.map((f) => (
                    <div key={f} role="option" aria-selected={mask.custom === f} className={`sl-list-item${mask.custom === f ? ' selected' : ''}`} onClick={() => set({ custom: f })}>
                      {f}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      ) : (
        <div className="sl-form">
          <div className="sl-field">
            <label htmlFor="sl-coltype">Column type</label>
            <select id="sl-coltype" className="sl-select" value={editor} data-testid="format-editor-type" onChange={(e) => setEditor(e.target.value as EditorType)}>
              {EDITOR_TYPES.map((t) => <option key={t.key} value={t.key}>{t.label}</option>)}
            </select>
          </div>
          {editor === 'dropdown' && (
            <div className="sl-field">
              <label htmlFor="sl-options">Options (one per line)</label>
              <textarea id="sl-options" className="sl-textarea" value={options} data-testid="format-options" onChange={(e) => setOptions(e.target.value)} />
            </div>
          )}
          <Switch label="Wrap content" checked={wrap} onChange={setWrap} testId="format-wrap" />
          <p className="sl-hint">Drop-down, checkbox and calendar cells are stored as data validation rules, so Excel keeps them.</p>
        </div>
      )}
    </Modal>
  );
}
