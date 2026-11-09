/** Small dialogs: defined names, links, confirm/prompt, about, shortcuts and the function list. */
import { useMemo, useState } from 'react';
import { exampleFor, functionCatalog, functionCategories, MOD, quoteSheetName, rangeName } from '@sumline/core/views';
import { useUI } from '../context';
import { Modal } from '../Modal';

export function NameDialog({ onClose, name }: { onClose: () => void; name?: string }) {
  const ui = useUI();
  const { doc } = ui;
  const existing = name ? doc.wb.names.find((n) => n.localSheet === undefined && n.name === name) : undefined;
  const [value, setValue] = useState(existing?.name ?? '');
  const [ref, setRef] = useState(existing ? `=${existing.formula}` : `=${quoteSheetName(doc.sheet.name)}!${rangeName(doc.range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`);
  const [error, setError] = useState('');
  const save = () => {
    const err = doc.setDefinedName(value, ref, existing?.name);
    if (err) setError(err);
    else onClose();
  };
  return (
    <Modal title="Defined names" onClose={onClose} width={380} testId="name-dialog">
      <div className="sl-form">
        <div className="sl-field">
          <label htmlFor="dn-name">Interval Name</label>
          <input id="dn-name" className="sl-input" value={value} data-testid="dn-name" onChange={(e) => setValue(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && save()} />
        </div>
        <div className="sl-field">
          <label htmlFor="dn-range">Cell range</label>
          <input id="dn-range" className="sl-input" value={ref} spellCheck={false} data-testid="dn-range" onChange={(e) => setRef(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && save()} />
        </div>
        {error && <div className="sl-error" role="alert">{error}</div>}
        <div style={{ display: 'flex', gap: 6 }}>
          <button type="button" className="sl-button" data-testid="dn-save" onClick={save}>Save</button>
          <button type="button" className="sl-button" onClick={onClose}>Cancel</button>
          {existing && (
            <button type="button" className="sl-button danger" style={{ marginLeft: 'auto' }} data-testid="dn-delete" onClick={() => { doc.deleteDefinedName(existing.name); onClose(); }}>
              Delete
            </button>
          )}
        </div>
      </div>
    </Modal>
  );
}

export function LinkDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const { r, c } = doc.selection.active;
  const current = doc.sheet.hyperlinkAt(r, c);
  const [target, setTarget] = useState(current?.target ?? (current?.location ? `#${current.location}` : ''));
  const [text, setText] = useState(doc.displayText(r, c));
  const save = () => {
    const t = target.trim();
    if (!t) return;
    const url = /^(#|[a-z][a-z0-9+.-]*:)/i.test(t) || /!/.test(t) ? t : /^www\./i.test(t) || /\.[a-z]{2,}(\/|$)/i.test(t) ? `https://${t}` : t;
    doc.setLink(url, text && text !== doc.displayText(r, c) ? text : text ? undefined : url);
    onClose();
  };
  return (
    <Modal title="Link" icon="link" onClose={onClose} width={420} testId="link-dialog">
      <div className="sl-form">
        <div className="sl-field">
          <label htmlFor="ln-text">Text</label>
          <input id="ln-text" className="sl-input" value={text} data-testid="link-text" onChange={(e) => setText(e.target.value)} />
        </div>
        <div className="sl-field">
          <label htmlFor="ln-url">Web address or place in this workbook (Sheet2!A1)</label>
          <input id="ln-url" className="sl-input" value={target} placeholder="https://" inputMode="url" data-testid="link-url" onChange={(e) => setTarget(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && save()} />
        </div>
        <div style={{ display: 'flex', gap: 6 }}>
          <button type="button" className="sl-button" data-testid="link-save" onClick={save}>Save</button>
          <button type="button" className="sl-button" onClick={onClose}>Cancel</button>
          {current && (
            <button type="button" className="sl-button light" style={{ marginLeft: 'auto' }} data-testid="link-remove" onClick={() => { doc.removeLink(); onClose(); }}>
              Remove link
            </button>
          )}
        </div>
      </div>
    </Modal>
  );
}

export function ConfirmDialog({ message, title, ok = 'Ok', cancel = 'Cancel', danger, onAnswer }: { message: string; title?: string; ok?: string; cancel?: string; danger?: boolean; onAnswer: (v: boolean) => void }) {
  return (
    <Modal title={title ?? 'Confirm'} onClose={() => onAnswer(false)} width={400} blocking testId="confirm-dialog">
      <p style={{ margin: '0 0 20px', whiteSpace: 'pre-wrap' }}>{message}</p>
      <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
        <button type="button" className={`sl-button${danger ? ' danger' : ''}`} data-testid="confirm-ok" data-autofocus onClick={() => onAnswer(true)}>{ok}</button>
        {cancel && <button type="button" className="sl-button light" data-testid="confirm-cancel" onClick={() => onAnswer(false)}>{cancel}</button>}
      </div>
    </Modal>
  );
}

export function PromptDialog({ message, value, title, ok = 'Ok', onAnswer }: { message: string; value: string; title?: string; ok?: string; onAnswer: (v: string | null) => void }) {
  const [text, setText] = useState(value);
  return (
    <Modal title={title ?? message} onClose={() => onAnswer(null)} width={400} blocking testId="prompt-dialog">
      <div className="sl-form">
        <div className="sl-field">
          <label htmlFor="sl-prompt">{message}</label>
          <input id="sl-prompt" className="sl-input" value={text} data-testid="prompt-input" onFocus={(e) => e.target.select()} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && onAnswer(text)} />
        </div>
        <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
          <button type="button" className="sl-button" data-testid="prompt-ok" onClick={() => onAnswer(text)}>{ok}</button>
          <button type="button" className="sl-button light" onClick={() => onAnswer(null)}>Cancel</button>
        </div>
      </div>
    </Modal>
  );
}

export function AboutDialog({ onClose }: { onClose: () => void }) {
  return (
    <Modal title="About" icon="info" onClose={onClose} width={420} testId="about-dialog">
      <p style={{ marginTop: 0 }}><b>Sumline</b></p>
      <p>Open, edit and save Excel workbooks (XLSX, XLSM, XLSB, XLS) and CSV files in the browser, on desktop and on phones. Files never leave your device.</p>
      <p className="sl-hint" style={{ marginBottom: 0 }}>Formulas, charts, shapes, comments, notes, filters, data validation and conditional formatting are kept in the saved XLSX file.</p>
    </Modal>
  );
}

const SHORTCUTS: [string, string][] = [
  ['Undo / Redo', `${MOD}+Z / ${MOD}+Y`],
  ['Cut / Copy / Paste', `${MOD}+X / ${MOD}+C / ${MOD}+V`],
  ['Find and replace', `${MOD}+F`],
  ['Save (download XLSX)', `${MOD}+S`],
  ['Print / PDF', `${MOD}+P`],
  ['Bold / Italic / Underline', `${MOD}+B / ${MOD}+I / ${MOD}+U`],
  ['Remove all format', `${MOD}+\\`],
  ['Select all', `${MOD}+A`],
  ['Edit the cell', 'F2 or double click (double tap on phones)'],
  ['New line in a cell', 'Alt+Enter'],
  ['Fill down / right', `${MOD}+D / ${MOD}+R`],
  ['Jump to the data edge', `${MOD}+Arrow keys`],
  ['Extend the selection', 'Shift+Arrow keys'],
  ['Absolute / relative reference', 'F4 while editing a formula'],
  ['AutoSum', 'Alt+='],
  ['Next / previous worksheet', `${MOD}+Page Down / ${MOD}+Page Up`],
];

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  return (
    <Modal title="Keyboard shortcuts" icon="article" onClose={onClose} width={520} testId="shortcuts-dialog">
      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13.5 }}>
        <tbody>
          {SHORTCUTS.map(([what, keys]) => (
            <tr key={what} style={{ borderBottom: '1px solid #eee' }}>
              <td style={{ padding: '7px 8px 7px 0' }}>{what}</td>
              <td style={{ padding: '7px 0', color: '#555', textAlign: 'right' }}>{keys}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </Modal>
  );
}

export function FunctionsDialog({ onClose }: { onClose: () => void }) {
  const ui = useUI();
  const [query, setQuery] = useState('');
  const [category, setCategory] = useState('All');
  const categories = useMemo(() => ['All', ...functionCategories()], []);
  const items = useMemo(() => {
    const q = query.trim().toUpperCase();
    return [...functionCatalog().values()].filter((f) => (category === 'All' || f.category === category) && (!q || f.name.includes(q) || f.description.toUpperCase().includes(q)));
  }, [query, category]);
  const insert = (name: string) => {
    onClose();
    if (ui.state.readOnly || !ui.grid) return;
    if (ui.grid.isEditing() && ui.grid.editState.text.startsWith('=')) {
      const { text, caret } = ui.grid.editState;
      ui.grid.replaceEditText(`${text.slice(0, caret)}${name}(${text.slice(caret)}`, caret + name.length + 1);
    } else {
      ui.grid.startEdit(`=${name}(`, 'enter');
    }
  };
  return (
    <Modal title={`Function list (${items.length})`} icon="functions" onClose={onClose} width={560} height="min(640px, calc(100vh - 16px))" testId="functions-dialog">
      <div className="sl-row" style={{ marginBottom: 12 }}>
        <input className="sl-input" placeholder="Search functions" value={query} data-testid="functions-search" onChange={(e) => setQuery(e.target.value)} />
        <select className="sl-select" value={category} aria-label="Category" onChange={(e) => setCategory(e.target.value)}>
          {categories.map((c) => <option key={c}>{c}</option>)}
        </select>
      </div>
      <div>
        {items.slice(0, 400).map((f) => (
          <button key={f.name} type="button" className="sl-suggest-item" style={{ display: 'block', width: '100%', border: 0, background: 'none', textAlign: 'left' }} onClick={() => insert(f.name)}>
            <div className="sl-suggest-sig">{f.name}({f.args.join(', ')})</div>
            {f.description && <div className="sl-suggest-desc">{f.description}</div>}
            <div className="sl-suggest-example">{exampleFor(f)} · {f.category}</div>
          </button>
        ))}
      </div>
    </Modal>
  );
}
