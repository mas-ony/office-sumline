/**
 * Phones: a row of keys above the on-screen keyboard while a cell is edited (formula characters
 * that are hard to reach on mobile keyboards, plus done/cancel and the function list).
 */
import { useEffect, useState } from 'react';
import { useUI } from './context';
import { Icon } from './Icon';

const KEYS = ['=', '(', ')', ',', ':', '+', '-', '*', '/', '&', '$', '%', '"', '<', '>', ';'];

/** Height of the on-screen keyboard overlapping the layout viewport (iOS does not resize it). */
export function useKeyboardOffset(): number {
  const [offset, setOffset] = useState(0);
  useEffect(() => {
    const vv = window.visualViewport;
    if (!vv) return;
    const update = () => setOffset(Math.max(0, Math.round(window.innerHeight - vv.height - vv.offsetTop)));
    vv.addEventListener('resize', update);
    vv.addEventListener('scroll', update);
    update();
    return () => {
      vv.removeEventListener('resize', update);
      vv.removeEventListener('scroll', update);
    };
  }, []);
  return offset;
}

export function KeyBar() {
  const ui = useUI();
  const { grid } = ui;
  const offset = useKeyboardOffset();
  const insert = (s: string) => {
    if (!grid) return;
    const { text, caret } = grid.editState;
    grid.replaceEditText(text.slice(0, caret) + s + text.slice(caret), caret + s.length);
  };
  const keep = (e: React.PointerEvent | React.MouseEvent) => e.preventDefault();
  return (
    <div className="sl-keybar" style={{ transform: offset ? `translateY(${-offset}px)` : undefined }} role="toolbar" aria-label="Cell editing keys" data-testid="keybar">
      <button type="button" className="sl-key" aria-label="Cancel" data-testid="key-cancel" onPointerDown={keep} onMouseDown={keep} onClick={() => grid?.cancelEdit()}>
        <Icon name="close" />
      </button>
      <button type="button" className="sl-key primary" aria-label="Done" data-testid="key-done" onPointerDown={keep} onMouseDown={keep} onClick={() => grid?.commitEdit({ dr: 1, dc: 0 })}>
        <Icon name="check" />
      </button>
      {KEYS.map((k) => (
        <button key={k} type="button" className="sl-key" aria-label={`Insert ${k}`} onPointerDown={keep} onMouseDown={keep} onClick={() => insert(k)}>
          {k}
        </button>
      ))}
      <button type="button" className="sl-key" aria-label="Functions" data-testid="key-functions" onPointerDown={keep} onMouseDown={keep} onClick={() => ui.openDialog({ kind: 'functions' })}>
        <Icon name="functions" />
      </button>
      <button type="button" className="sl-key" aria-label="Hide keyboard" onPointerDown={keep} onMouseDown={keep} onClick={() => (document.activeElement as HTMLElement | null)?.blur()}>
        <Icon name="keyboard" />
      </button>
    </div>
  );
}
