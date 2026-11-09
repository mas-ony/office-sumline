/** Pop-ups anchored to cells: the comment thread box, the note editor and the column filter. */
import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { cellName, type ColumnFilter, type FilterOperator } from '@sumline/core/views';
import { useUI } from './context';
import { Icon } from './Icon';
import { Menu } from './Menu';

/** A fixed pop-up next to a rectangle, closed by clicks outside and Escape. */
function Anchored({ rect, className, onClose, children, testId, label, width }: { rect: DOMRect; className: string; onClose: () => void; children: ReactNode; testId?: string; label?: string; width?: number }) {
  const { layer } = useUI();
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.visualViewport?.height ?? window.innerHeight;
    let x = rect.right + 4;
    let y = rect.top;
    if (x + w > vw - 4) x = rect.left - w - 4;
    if (x < 4) {
      x = Math.max(4, Math.min(rect.left, vw - w - 4));
      y = rect.bottom + 4;
    }
    if (y + h > vh - 4) y = Math.max(4, vh - h - 4);
    el.style.left = `${x}px`;
    el.style.top = `${y}px`;
  });
  useEffect(() => {
    const down = (e: PointerEvent) => {
      const t = e.target as HTMLElement;
      if (t.closest(`.${className.split(' ')[0]}`) || t.closest('.sl-popup') || t.closest('.sl-bottom-sheet') || t.closest('.sl-sheet-backdrop')) return;
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
  }, [onClose, className]);
  return createPortal(
    <div ref={ref} className={className} role="dialog" aria-label={label} data-testid={testId} style={width ? { width } : undefined}>
      {children}
    </div>,
    layer,
  );
}

function when(iso?: string): string {
  if (!iso) return '';
  const t = Date.parse(iso.endsWith('Z') || /[+-]\d\d:\d\d$/.test(iso) ? iso : `${iso}Z`);
  if (!Number.isFinite(t)) return '';
  const s = Math.round((Date.now() - t) / 1000);
  if (s < 60) return 'Just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  return new Date(t).toLocaleDateString();
}

export function CommentsBox({ r, c, rect, onClose }: { r: number; c: number; rect: DOMRect; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const [text, setText] = useState('');
  const [editing, setEditing] = useState<{ id: string; text: string } | null>(null);
  const [menu, setMenu] = useState<{ id: string; at: DOMRect } | null>(null);
  const thread = doc.sheet.commentAt(r, c);
  const ro = ui.state.readOnly;
  const save = () => {
    if (!text.trim()) return;
    doc.addComment(r, c, text.trim());
    setText('');
  };
  return (
    <Anchored rect={rect} className="sl-comments" onClose={onClose} testId="comments-box" label={`Comments ${cellName(r, c)}`}>
      {thread?.comments.map((cm, i) => (
        <div key={cm.id} className="sl-comment" data-testid="comment">
          <div className="who">
            <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ width: 28, height: 28, borderRadius: '50%', border: '1px solid #ccc', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, color: '#555' }}>
                {(cm.author || '?').slice(0, 1).toUpperCase()}
              </span>
              <span>
                {cm.author || 'Guest'}
                <div className="when">{when(cm.created) || (thread.kind === 'note' ? 'Note' : '')}{i === 0 && thread.resolved ? ' · Resolved' : ''}</div>
              </span>
            </span>
            {!ro && (
              <button type="button" className="sl-icon-btn" style={{ width: 26, height: 26 }} aria-label="Comment options" data-testid="comment-options" onClick={(e) => setMenu({ id: cm.id, at: e.currentTarget.getBoundingClientRect() })}>
                <Icon name="arrow_drop_down" />
              </button>
            )}
          </div>
          {editing?.id === cm.id ? (
            <div style={{ display: 'grid', gap: 6, marginTop: 6 }}>
              <textarea className="sl-textarea" value={editing.text} onChange={(e) => setEditing({ ...editing, text: e.target.value })} />
              <div style={{ display: 'flex', gap: 6 }}>
                <button type="button" className="sl-button" onClick={() => { doc.editComment(cm.id, editing.text); setEditing(null); }}>Save</button>
                <button type="button" className="sl-button light" onClick={() => setEditing(null)}>Cancel</button>
              </div>
            </div>
          ) : (
            <div className="body">{cm.text}</div>
          )}
        </div>
      ))}
      {!ro && (
        <div className="sl-comments-input">
          <textarea className="sl-textarea" value={text} autoFocus placeholder={thread ? 'Reply' : 'Comment'} aria-label="Comment" data-testid="comment-input" onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) save(); }} />
          <div style={{ display: 'flex', gap: 6 }}>
            <button type="button" className="sl-button" style={{ flex: 1 }} disabled={!text.trim()} data-testid="comment-save" onClick={save}>Save</button>
            <button type="button" className="sl-button" style={{ flex: 1 }} onClick={onClose}>Cancel</button>
          </div>
        </div>
      )}
      {menu && (
        <Menu
          at={menu.at}
          onClose={() => setMenu(null)}
          items={[
            { title: 'Edit', icon: 'edit', onClick: () => setEditing({ id: menu.id, text: thread?.comments.find((x) => x.id === menu.id)?.text ?? '' }) },
            { title: thread?.resolved ? 'Reopen' : 'Resolve', icon: 'check', onClick: () => doc.setResolved(menu.id, !thread?.resolved) },
            { title: 'Delete', icon: 'delete', onClick: () => doc.deleteComment(menu.id) },
          ]}
        />
      )}
    </Anchored>
  );
}

export function NoteEditor({ r, c, rect, onClose }: { r: number; c: number; rect: DOMRect; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const existing = doc.sheet.commentAt(r, c);
  const [text, setText] = useState(existing?.comments.map((x) => x.text).join('\n') ?? '');
  const save = () => {
    doc.setNote(r, c, text.trim());
    onClose();
  };
  return (
    <Anchored rect={rect} className="sl-note" onClose={onClose} testId="note-editor" label={`Notes ${cellName(r, c)}`}>
      <strong style={{ fontSize: 13 }}>Notes</strong>
      <textarea value={text} autoFocus aria-label="Note" data-testid="note-input" onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) save(); }} />
      <div className="sl-note-actions">
        {existing && <button type="button" className="sl-button light" onClick={() => { doc.setNote(r, c, ''); onClose(); }}>Clear</button>}
        <button type="button" className="sl-button" data-testid="note-save" onClick={save}>Save</button>
        <button type="button" className="sl-button light" onClick={onClose}>Cancel</button>
      </div>
    </Anchored>
  );
}

const OPERATORS: { key: FilterOperator; label: string }[] = [
  { key: 'contains', label: 'Contains' },
  { key: 'notContains', label: 'Does not contain' },
  { key: 'beginsWith', label: 'Begins with' },
  { key: 'endsWith', label: 'Ends with' },
  { key: 'equal', label: 'Equal' },
  { key: 'notEqual', label: 'Not equal' },
  { key: 'greaterThan', label: 'Greater than' },
  { key: 'lessThan', label: 'Lower than' },
];

export function FilterPopup({ col, at, onClose }: { col: number; at: { x: number; y: number }; onClose: () => void }) {
  const ui = useUI();
  const { doc } = ui;
  const af = doc.sheet.autoFilter;
  const current: ColumnFilter | undefined = af?.columns.get(col);
  const values = useMemo(() => doc.filterValues(col), [doc, col]);
  const [search, setSearch] = useState(current?.conditions?.[0]?.value ?? '');
  const [op, setOp] = useState<FilterOperator>(current?.conditions?.[0]?.op ?? 'contains');
  const [checked, setChecked] = useState<Set<string>>(() => {
    if (current?.values) return new Set([...current.values.map((v) => v.toLowerCase()), ...(current.blanks ? [''] : [])]);
    return new Set(values.map((v) => v.text.toLowerCase()));
  });
  const shown = values.filter((v) => !search || op !== 'contains' || v.text.toLowerCase().includes(search.toLowerCase()));
  const all = values.every((v) => checked.has(v.text.toLowerCase()));
  const rect = new DOMRect(at.x, at.y, 0, 0);
  const apply = () => {
    if (search.trim()) {
      doc.setColumnFilter(col, { conditions: [{ op, value: search.trim() }] });
    } else if (all) {
      doc.setColumnFilter(col, null);
    } else {
      const keep = values.filter((v) => checked.has(v.text.toLowerCase()) && v.text !== '').map((v) => v.text);
      doc.setColumnFilter(col, { values: keep, blanks: checked.has('') });
    }
    onClose();
  };
  return (
    <Anchored rect={rect} className="sl-filter" onClose={onClose} testId="filter-popup" label="Filter Options">
      <div className="sl-filter-sort">
        <button type="button" data-testid="filter-sort-asc" onClick={() => { doc.sortFilterColumn(col, true); onClose(); }}>Sort smallest to largest</button>
        <button type="button" data-testid="filter-sort-desc" onClick={() => { doc.sortFilterColumn(col, false); onClose(); }}>Sort largest to smallest</button>
      </div>
      <div className="sl-filter-search">
        <input className="sl-input" placeholder="Search" value={search} aria-label="Search" data-testid="filter-search" onChange={(e) => setSearch(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && apply()} />
        <select className="sl-select" value={op} aria-label="Condition" data-testid="filter-operator" onChange={(e) => setOp(e.target.value as FilterOperator)}>
          {OPERATORS.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
        </select>
      </div>
      <div className="sl-filter-values" data-testid="filter-values">
        <label>
          <input type="checkbox" checked={all} onChange={(e) => setChecked(e.target.checked ? new Set(values.map((v) => v.text.toLowerCase())) : new Set())} />
          (Select all)
        </label>
        {shown.map((v) => (
          <label key={v.text || '(blank)'}>
            <input
              type="checkbox"
              checked={checked.has(v.text.toLowerCase())}
              onChange={(e) => {
                const next = new Set(checked);
                if (e.target.checked) next.add(v.text.toLowerCase());
                else next.delete(v.text.toLowerCase());
                setChecked(next);
              }}
            />
            {v.text === '' ? '(Blanks)' : v.text}
          </label>
        ))}
      </div>
      <div className="sl-filter-actions">
        <button type="button" className="sl-button" data-testid="filter-ok" onClick={apply}>Ok</button>
        <button type="button" className="sl-button" onClick={onClose}>Cancel</button>
      </div>
    </Anchored>
  );
}
