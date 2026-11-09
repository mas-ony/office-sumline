/**
 * Worksheet tabs above the toolbar: select, add (+), rename (double click or the tab menu),
 * reorder by dragging, and the tab context menu (right click or long press).
 */
import { useEffect, useRef, useState } from 'react';
import { useUI } from './context';
import { Icon } from './Icon';
import { Menu } from './Menu';
import { cmd, tabMenu } from '@sumline/core/views';

export function SheetTabs() {
  const ui = useUI();
  const { doc, state } = ui;
  const [editing, setEditing] = useState<number | null>(null);
  const [draft, setDraft] = useState('');
  const [menu, setMenu] = useState<{ index: number; at: { x: number; y: number } } | null>(null);
  const [drag, setDrag] = useState<{ from: number; over: number; side: 'before' | 'after' } | null>(null);
  const press = useRef<{ timer: number; index: number; x: number; y: number; moved: boolean; pointerId: number } | null>(null);
  const list = useRef<HTMLDivElement>(null);

  // Keep the active tab in view.
  useEffect(() => {
    list.current?.querySelector<HTMLElement>('.sl-tab.selected')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [state.activeIndex]);

  const commitRename = () => {
    if (editing === null) return;
    const index = editing;
    const name = draft.trim();
    setEditing(null);
    if (!name || name === doc.wb.sheets[index]?.name) return;
    const error = doc.renameSheet(index, name);
    if (error) ui.notify(error);
  };

  const openMenu = (index: number, x: number, y: number) => {
    doc.setActiveSheet(index);
    setMenu({ index, at: { x, y } });
  };

  const onPointerDown = (e: React.PointerEvent, index: number) => {
    if (e.button !== 0 || editing !== null) return;
    const timer = window.setTimeout(() => {
      if (press.current && !press.current.moved) {
        openMenu(index, press.current.x, press.current.y);
        press.current = null;
      }
    }, 500);
    press.current = { timer, index, x: e.clientX, y: e.clientY, moved: false, pointerId: e.pointerId };
  };

  const onPointerMove = (e: React.PointerEvent) => {
    const p = press.current;
    if (!p || p.pointerId !== e.pointerId) return;
    if (!p.moved && Math.hypot(e.clientX - p.x, e.clientY - p.y) > 8) {
      p.moved = true;
      window.clearTimeout(p.timer);
      // Mouse drags reorder; touch moves scroll the tab strip.
      if (e.pointerType === 'mouse' && !state.readOnly) {
        (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
        setDrag({ from: p.index, over: p.index, side: 'after' });
      }
    }
    if (drag) {
      const el = document.elementFromPoint(e.clientX, e.clientY)?.closest<HTMLElement>('[data-tab-index]');
      if (el) {
        const over = Number(el.dataset.tabIndex);
        const rect = el.getBoundingClientRect();
        setDrag({ ...drag, over, side: e.clientX < rect.left + rect.width / 2 ? 'before' : 'after' });
      }
    }
  };

  const onPointerUp = (e: React.PointerEvent, index: number) => {
    const p = press.current;
    if (p) window.clearTimeout(p.timer);
    press.current = null;
    if (drag) {
      const { from, over, side } = drag;
      setDrag(null);
      let to = side === 'before' ? over : over + 1;
      if (from < to) to--;
      if (to !== from) doc.moveSheet(from, Math.max(0, Math.min(doc.wb.sheets.length - 1, to)));
      return;
    }
    if (p && !p.moved && e.button === 0) {
      if (ui.grid?.isEditing() && !ui.grid.editState.text.startsWith('=')) ui.grid.commitEdit();
      doc.setActiveSheet(index);
    }
  };

  return (
    <div className="sl-tabs" data-testid="sheet-tabs">
      <div ref={list} className="sl-tabs-list" role="tablist" aria-label="Worksheets">
        {state.sheets.map((s) => {
          if (s.state !== 'visible') return null;
          const selected = s.index === state.activeIndex;
          if (editing === s.index) {
            return (
              <input
                key={s.index}
                className="sl-tab-input"
                aria-label="Worksheet name"
                data-testid="tab-rename-input"
                autoFocus
                value={draft}
                onChange={(e) => setDraft(e.target.value)}
                onBlur={commitRename}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') commitRename();
                  else if (e.key === 'Escape') setEditing(null);
                }}
              />
            );
          }
          const dropClass = drag && drag.over === s.index && drag.from !== s.index ? ` drop-${drag.side}` : '';
          return (
            <button
              key={s.index}
              type="button"
              role="tab"
              aria-selected={selected}
              data-tab-index={s.index}
              data-testid={`tab-${s.name}`}
              className={`sl-tab${selected ? ' selected' : ''}${dropClass}`}
              onPointerDown={(e) => onPointerDown(e, s.index)}
              onPointerMove={onPointerMove}
              onPointerUp={(e) => onPointerUp(e, s.index)}
              onPointerCancel={() => {
                if (press.current) window.clearTimeout(press.current.timer);
                press.current = null;
                setDrag(null);
              }}
              onDoubleClick={() => {
                if (state.readOnly) return;
                setDraft(s.name);
                setEditing(s.index);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                openMenu(s.index, e.clientX, e.clientY);
              }}
            >
              {s.color && <span className="sl-tab-color" style={{ background: s.color }} />}
              {s.name}
            </button>
          );
        })}
      </div>
      {!state.readOnly && (
        <button type="button" className="sl-tab-add" aria-label="Add worksheet" title="Add worksheet" data-testid="tab-add" onClick={() => cmd.newWorksheet(ui)}>
          <Icon name="add" />
        </button>
      )}
      {menu && <Menu items={tabMenu(ui, menu.index)} at={menu.at} onClose={() => setMenu(null)} sheet={ui.phone} title={doc.wb.sheets[menu.index]?.name} label="Worksheet" testId="tab-menu" />}
    </div>
  );
}
