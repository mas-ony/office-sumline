/**
 * Formula bar: the name box (with the list of defined names and "Add new entry"), fx, the formula
 * input with coloured references, and the expand button. Typing a function name shows matching
 * functions with their signature, description and an example.
 */
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { exampleFor, formulaSegments, highlightedArg, type UIContextValue } from '@sumline/core/views';
import { useUI } from './context';
import { Icon } from './Icon';
import { Popup } from './Toolbar';

/** Colours cell references like the reference highlights drawn on the grid. */
function colorize(text: string): ReactNode[] {
  return formulaSegments(text).map((s) => (s.color ? <span key={s.start} style={{ color: s.color }}>{s.text}</span> : s.text));
}

function NameBox() {
  const ui = useUI();
  const { doc, grid, state } = ui;
  const [name, setName] = useState(state.rangeRef);
  const [focused, setFocused] = useState(false);
  const [list, setList] = useState<DOMRect | null>(null);
  const wrap = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!focused) setName(state.rangeRef);
  }, [state.rangeRef, focused]);
  const names = doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.'));
  const goTo = () => {
    const text = name.trim();
    if (text && !doc.goTo(text)) ui.notify(`“${text}” is not a valid cell reference or name.`);
    setFocused(false);
    grid?.focusIdle();
  };
  return (
    <div className="sl-namebox" ref={wrap}>
      <input
        aria-label="Name box"
        data-testid="name-box"
        value={name}
        spellCheck={false}
        autoCapitalize="characters"
        autoComplete="off"
        enterKeyHint="go"
        onFocus={(e) => {
          setFocused(true);
          e.target.select();
        }}
        onBlur={() => setFocused(false)}
        onChange={(e) => setName(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            goTo();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setName(state.rangeRef);
            grid?.focusIdle();
          }
        }}
      />
      <button
        type="button"
        className="sl-namebox-caret"
        aria-label="Defined names"
        data-testid="names-button"
        aria-expanded={!!list}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setList(list ? null : wrap.current!.getBoundingClientRect())}
      >
        <Icon name="arrow_drop_down" />
      </button>
      {list && (
        <Popup anchor={list} onClose={() => setList(null)} label="Defined names" testId="names-list">
          <div style={{ padding: '6px 0', minWidth: 260 }}>
            {names.length === 0 && <div className="sl-hint" style={{ padding: '6px 14px' }}>No defined names</div>}
            {names.map((n) => (
              <button
                key={`${n.name}-${n.localSheet ?? ''}`}
                type="button"
                className="sl-picker-item"
                onClick={() => {
                  setList(null);
                  doc.goTo(n.name);
                  grid?.focusIdle();
                }}
              >
                <span style={{ flex: 1 }}>{n.name}</span>
                <span className="sl-hint" style={{ fontSize: 12 }}>{n.formula}</span>
                {!state.readOnly && n.localSheet === undefined && (
                  <span
                    role="button"
                    tabIndex={0}
                    aria-label={`Edit ${n.name}`}
                    style={{ display: 'flex', color: '#666' }}
                    onClick={(e) => {
                      e.stopPropagation();
                      setList(null);
                      ui.openDialog({ kind: 'name-edit', name: n.name });
                    }}
                  >
                    <Icon name="edit" size={16} />
                  </span>
                )}
              </button>
            ))}
            {!state.readOnly && (
              <div style={{ padding: '8px 12px 4px' }}>
                <button type="button" className="sl-button wide" data-testid="names-add" onClick={() => { setList(null); ui.openDialog({ kind: 'name-edit' }); }}>
                  Add new entry
                </button>
              </div>
            )}
          </div>
        </Popup>
      )}
    </div>
  );
}

export function FormulaBar() {
  const ui = useUI();
  const { grid, state, edit } = ui;
  const input = useRef<HTMLTextAreaElement>(null);
  const mirror = useRef<HTMLDivElement>(null);
  const reported = useRef<{ text: string; caret: number } | null>(null);
  const [expanded, setExpanded] = useState(false);
  const value = edit.active ? edit.text : state.editText;

  // Keep focus and caret in the bar when the edit changes from elsewhere (autocomplete, clicking
  // cells to insert references).
  useLayoutEffect(() => {
    const el = input.current;
    if (!el || !edit.active || edit.origin !== 'bar') return;
    const r = reported.current;
    if (r && r.text === edit.text && r.caret === edit.caret) return;
    reported.current = { text: edit.text, caret: edit.caret };
    if (document.activeElement !== el) el.focus({ preventScroll: true });
    el.setSelectionRange(edit.caret, edit.caret);
  }, [edit]);

  useLayoutEffect(() => {
    if (mirror.current && input.current) {
      mirror.current.scrollTop = input.current.scrollTop;
      mirror.current.scrollLeft = input.current.scrollLeft;
    }
  });

  const report = () => {
    const el = input.current;
    if (!el || state.readOnly || !grid) return;
    const caret = el.selectionStart ?? el.value.length;
    reported.current = { text: el.value, caret };
    grid.setEditFromBar(el.value, caret);
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (!grid) return;
    if (edit.active && ui.ui.assist.handleKey(e.nativeEvent)) {
      e.preventDefault();
      return;
    }
    if (e.nativeEvent.isComposing) return;
    if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      const el = e.currentTarget;
      const s = el.selectionStart ?? el.value.length;
      el.value = `${el.value.slice(0, s)}\n${el.value.slice(el.selectionEnd ?? s)}`;
      el.setSelectionRange(s + 1, s + 1);
      report();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      grid.commitEdit({ dr: e.shiftKey ? -1 : 1, dc: 0 });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      grid.commitEdit({ dr: 0, dc: e.shiftKey ? -1 : 1 });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      grid.cancelEdit();
    }
  };

  return (
    <div className={`sl-formulabar${expanded ? ' expanded' : ''}`} role="group" aria-label="Formula editor" data-testid="formula-bar">
      <NameBox />
      <div className="sl-fx" aria-hidden="true">fx</div>
      <div style={{ position: 'relative', flex: '1 1 auto', minWidth: 0, display: 'flex' }}>
        <div ref={mirror} className="sl-formula-input" aria-hidden="true" style={{ position: 'absolute', inset: 0, pointerEvents: 'none', color: '#111' }}>
          {colorize(value)}
          {value.endsWith('\n') ? '​' : null}
        </div>
        <textarea
          ref={input}
          className="sl-formula-input"
          aria-label="Formula bar"
          data-testid="formula-input"
          rows={1}
          spellCheck={false}
          autoCapitalize="off"
          autoComplete="off"
          autoCorrect="off"
          enterKeyHint={ui.phone ? 'done' : undefined}
          readOnly={state.readOnly}
          value={value}
          style={{ color: 'transparent', caretColor: '#111', position: 'relative' }}
          onFocus={() => {
            if (state.readOnly || !grid) return;
            if (!grid.isEditing()) {
              grid.startEdit(undefined, 'edit', undefined, 'bar');
              reported.current = { text: grid.editState.text, caret: grid.editState.caret };
            }
          }}
          onChange={report}
          onSelect={() => edit.active && report()}
          onKeyDown={onKeyDown}
          onScroll={(e) => {
            if (mirror.current) {
              mirror.current.scrollTop = e.currentTarget.scrollTop;
              mirror.current.scrollLeft = e.currentTarget.scrollLeft;
            }
          }}
        />
      </div>
      <button type="button" className="sl-formula-expand" aria-label={expanded ? 'Collapse formula bar' : 'Expand formula bar'} data-testid="fx-expand" onMouseDown={(e) => e.preventDefault()} onClick={() => setExpanded((x) => !x)}>
        <Icon name={expanded ? 'expand_less' : 'expand_more'} />
      </button>
    </div>
  );
}

// -----------------------------------------------------------------------------------------------
// Suggestions and argument hints

function anchorRect(origin: 'cell' | 'bar', ctx: UIContextValue): DOMRect | null {
  if (origin === 'bar') return ctx.ui.rootElement?.querySelector<HTMLElement>('[data-testid="formula-input"]')?.getBoundingClientRect() ?? null;
  return ctx.grid?.editorClientRect() ?? null;
}

export function FormulaAssist() {
  const ctx = useUI();
  const { edit, layer } = ctx;
  const assist = ctx.ui.assist;
  const { showList, suggestions, index, call, visible } = assist.view;
  const box = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);

  useLayoutEffect(() => {
    if (!visible) {
      if (pos) setPos(null);
      return;
    }
    const a = anchorRect(edit.origin, ctx);
    const el = box.current;
    if (!a || !el) return;
    const vh = window.visualViewport?.height ?? window.innerHeight;
    const vw = window.innerWidth;
    const h = el.offsetHeight;
    const w = el.offsetWidth;
    let top = a.bottom + 2;
    if (top + h > vh - 4 && a.top - h - 2 > 4) top = a.top - h - 2;
    const left = Math.max(4, Math.min(a.left, vw - w - 4));
    if (!pos || pos.left !== left || pos.top !== top) setPos({ left, top });
  });

  useEffect(() => {
    box.current?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' });
  }, [index, showList]);

  if (!visible) return null;
  const style = pos ? { left: pos.left, top: pos.top } : { left: -9999, top: 0 };
  return createPortal(
    showList ? (
      <div ref={box} className="sl-suggest" style={{ position: 'fixed', ...style }} role="listbox" aria-label="Functions" data-testid="formula-suggestions">
        {suggestions.map((s, i) => (
          <div
            key={s.name}
            role="option"
            aria-selected={i === index}
            className={`sl-suggest-item${i === index ? ' active' : ''}`}
            onMouseDown={(e) => e.preventDefault()}
            onPointerDown={(e) => e.preventDefault()}
            onClick={() => assist.apply(s)}
          >
            <div className="sl-suggest-sig">{s.category === 'Name' ? s.name : `${s.name}(${s.args.join(', ')})`}</div>
            {s.description && <div className="sl-suggest-desc">{s.description}</div>}
            {s.category !== 'Name' && <div className="sl-suggest-example">{exampleFor(s)}</div>}
          </div>
        ))}
      </div>
    ) : call ? (
      <div ref={box} className="sl-arghint" style={{ position: 'fixed', ...style }} data-testid="formula-hint">
        <div className="sl-suggest-sig">
          <b>{call.info.name}</b>(
          {call.info.args.map((a, i) => (
            <span key={i}>
              {i > 0 && ', '}
              <span className={i === highlightedArg(call.info, call.argIndex) ? 'cur' : undefined}>{a}</span>
            </span>
          ))}
          )
        </div>
        {call.info.description && <div className="sl-suggest-desc">{call.info.description}</div>}
      </div>
    ) : (
      <span />
    ),
    layer,
  );
}
