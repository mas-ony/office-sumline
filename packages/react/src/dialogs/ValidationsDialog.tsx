/**
 * Data validations: the rules of the selected cells, and a form to add or edit one (cell range,
 * action, message, type, criteria, values, blanks, list source, format style).
 */
import { useMemo, useState } from 'react';
import {
  blankValidationForm, describeValidation, quoteSheetName, rangeName, rangesIntersect, ruleFromForm, TEXT_CRITERIA, type TextCriteria, type ValidationAction,
  VALIDATION_CRITERIA, VALIDATION_KINDS, type ValidationCriteria, type ValidationForm, type ValidationKind, validationFormOf,
} from '@sumline/core/views';
import { useUI } from '../context';
import { Icon } from '../Icon';
import { Modal, Switch } from '../Modal';
import { ColorPicker } from '../ColorPicker';

export function ValidationsDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const sheet = doc.sheet;
  const selection = doc.range;
  const [mode, setMode] = useState<{ edit: number | null } | null>(null);
  const [error, setError] = useState('');
  const [colorPick, setColorPick] = useState<{ which: 'color' | 'background'; rect: DOMRect } | null>(null);

  const list = useMemo(
    () => sheet.validations.map((v, i) => ({ v, i })).filter(({ v }) => v.sqref.some((r) => rangesIntersect(r, selection))),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sheet.validations, selection, ui.state.version],
  );

  const blank = () => blankValidationForm(sheet.name, selection);

  const [form, setForm] = useState<ValidationForm>(blank);
  const set = (patch: Partial<ValidationForm>) => setForm((f) => ({ ...f, ...patch }));

  const save = () => {
    const res = ruleFromForm(form, sheet.name, doc.wb.date1904);
    if ('error' in res) {
      setError(res.error);
      return;
    }
    if (mode?.edit !== null && mode?.edit !== undefined) doc.updateValidation(mode.edit, res.rule);
    else doc.addValidation(res.rule);
    setMode(null);
    setError('');
  };

  const valueInputs = () => {
    const two = form.criteria === 'between' || form.criteria === 'notBetween';
    const type = form.kind === 'date' ? 'date' : form.kind === 'time' ? 'time' : 'text';
    return (
      <div className="sl-field">
        <label>Value</label>
        <div className="sl-row" style={{ gridTemplateColumns: two ? '1fr 1fr' : '1fr' }}>
          <input className="sl-input" type={type} step={form.kind === 'time' ? 1 : undefined} value={form.value1} data-testid="dv-value1" onChange={(e) => set({ value1: e.target.value })} />
          {two && <input className="sl-input" type={type} step={form.kind === 'time' ? 1 : undefined} value={form.value2} data-testid="dv-value2" onChange={(e) => set({ value2: e.target.value })} />}
        </div>
      </div>
    );
  };

  if (!mode) {
    return (
      <Modal title="Validations" icon="rule" onClose={onClose} width={480} testId="validations-dialog">
        {list.length === 0 ? (
          <p data-testid="dv-empty" style={{ margin: '0 0 22px' }}>No validation found for this selection</p>
        ) : (
          <div style={{ marginBottom: 16 }}>
            {list.map(({ v, i }) => (
              <div key={i} className="sl-validation-card" data-testid="dv-card">
                <div>
                  <div>{describeValidation(v)}</div>
                  <div className="meta">
                    {v.sqref.map((r) => rangeName(r)).join(' ')} · {v.formatStyle ? 'Format style' : v.errorStyle === 'stop' ? 'Reject entry' : v.errorStyle === 'information' ? 'Information' : 'Warning message'}
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 2 }}>
                  <button type="button" className="sl-icon-btn" aria-label="Edit validation" data-testid="dv-edit" onClick={() => { setForm(validationFormOf(v, sheet.name, doc.wb.date1904)); setMode({ edit: i }); }}>
                    <Icon name="edit" />
                  </button>
                  <button type="button" className="sl-icon-btn" aria-label="Delete validation" data-testid="dv-delete" onClick={() => doc.removeValidation(i)}>
                    <Icon name="delete" />
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
        <button type="button" className="sl-button wide" data-testid="dv-add" onClick={() => { setForm(blank()); setMode({ edit: null }); }}>
          Add new validation
        </button>
      </Modal>
    );
  }

  return (
    <Modal title="Validations" icon="rule" onClose={onClose} width={480} testId="validations-dialog">
      <div className="sl-form">
        <div className="sl-row">
          <div className="sl-field">
            <label htmlFor="dv-range">Cell range</label>
            <div className="sl-range-input">
              <input id="dv-range" className="sl-input" value={form.range} data-testid="dv-range" onChange={(e) => set({ range: e.target.value })} />
              <button type="button" title="Use the selection" aria-label="Use the selection" onClick={() => set({ range: `${quoteSheetName(sheet.name)}!${rangeName(doc.range)}` })}>
                <Icon name="highlight_alt" size={18} />
              </button>
            </div>
          </div>
          <div className="sl-field">
            <label htmlFor="dv-action">Action</label>
            <select id="dv-action" className="sl-select" value={form.action} data-testid="dv-action" onChange={(e) => set({ action: e.target.value as ValidationAction })}>
              <option value="warning">Warning message</option>
              <option value="reject">Reject entry</option>
              <option value="format">Format style</option>
            </select>
          </div>
        </div>
        {form.action !== 'format' && (
          <div className="sl-field">
            <label htmlFor="dv-message">{form.action === 'reject' ? 'Error message' : 'Warning message'}</label>
            <input id="dv-message" className="sl-input" value={form.message} data-testid="dv-message" onChange={(e) => set({ message: e.target.value })} />
          </div>
        )}
        <div className="sl-row">
          <div className="sl-field">
            <label htmlFor="dv-type">Type</label>
            <select id="dv-type" className="sl-select" value={form.kind} data-testid="dv-type" onChange={(e) => set({ kind: e.target.value as ValidationKind, value1: '', value2: '' })}>
              {VALIDATION_KINDS.map((k) => <option key={k.key} value={k.key}>{k.label}</option>)}
            </select>
          </div>
          {(form.kind === 'number' || form.kind === 'date' || form.kind === 'time' || form.kind === 'textLength') && (
            <div className="sl-field">
              <label htmlFor="dv-criteria">Criteria</label>
              <select id="dv-criteria" className="sl-select" value={form.criteria} data-testid="dv-criteria" onChange={(e) => set({ criteria: e.target.value as ValidationCriteria })}>
                {VALIDATION_CRITERIA.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
              </select>
            </div>
          )}
          {form.kind === 'text' && (
            <div className="sl-field">
              <label htmlFor="dv-tcriteria">Criteria</label>
              <select id="dv-tcriteria" className="sl-select" value={form.textCriteria} data-testid="dv-text-criteria" onChange={(e) => set({ textCriteria: e.target.value as TextCriteria })}>
                {TEXT_CRITERIA.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
              </select>
            </div>
          )}
        </div>
        {(form.kind === 'number' || form.kind === 'date' || form.kind === 'time' || form.kind === 'textLength') && valueInputs()}
        {form.kind === 'text' && form.textCriteria !== 'email' && form.textCriteria !== 'url' && (
          <div className="sl-field">
            <label htmlFor="dv-text">Value</label>
            <input id="dv-text" className="sl-input" value={form.value1} data-testid="dv-value1" onChange={(e) => set({ value1: e.target.value })} />
          </div>
        )}
        {form.kind === 'list' && (
          <>
            <div className="sl-field">
              <label htmlFor="dv-list">Items (comma separated) or a range like =$D$1:$D$5</label>
              <input id="dv-list" className="sl-input" value={form.value1} data-testid="dv-value1" onChange={(e) => set({ value1: e.target.value })} />
            </div>
            <Switch label="Create list as a dropdown" checked={form.dropdown} onChange={(v) => set({ dropdown: v })} testId="dv-dropdown" />
          </>
        )}
        {form.kind === 'formula' && (
          <div className="sl-field">
            <label htmlFor="dv-formula">Formula (TRUE when the value is valid)</label>
            <input id="dv-formula" className="sl-input" value={form.value1} placeholder="=A1>0" spellCheck={false} data-testid="dv-value1" onChange={(e) => set({ value1: e.target.value })} />
          </div>
        )}
        {form.kind !== 'empty' && form.kind !== 'notEmpty' && <Switch label="Allow blank" checked={form.allowBlank} onChange={(v) => set({ allowBlank: v })} testId="dv-blank" />}
        {form.action === 'format' && (
          <div className="sl-field">
            <div className="label">Formatting Options</div>
            <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
              {(['color', 'background'] as const).map((which) => (
                <button
                  key={which}
                  type="button"
                  className="sl-select-btn"
                  data-testid={`dv-style-${which}`}
                  onClick={(e) => setColorPick({ which, rect: e.currentTarget.getBoundingClientRect() })}
                >
                  <span style={{ width: 16, height: 16, borderRadius: 2, border: '1px solid #ccc', background: form.style[which] ?? 'transparent' }} />
                  {which === 'color' ? 'Color' : 'Background color'}
                </button>
              ))}
              <Switch label="Bold" checked={!!form.style.bold} onChange={(v) => set({ style: { ...form.style, bold: v } })} testId="dv-style-bold" />
            </div>
          </div>
        )}
        {error && <div className="sl-error" role="alert">{error}</div>}
        <button type="button" className="sl-button wide" data-testid="dv-save" onClick={save}>Save</button>
        <button type="button" className="sl-button light wide" data-testid="dv-back" onClick={() => { setMode(null); setError(''); }}>Back</button>
      </div>
      {colorPick && (
        <ColorPicker
          anchor={colorPick.rect}
          value={form.style[colorPick.which] ?? null}
          live
          onChange={(hex) => set({ style: { ...form.style, [colorPick.which]: hex ?? undefined } })}
          onClose={() => setColorPick(null)}
        />
      )}
    </Modal>
  );
}
