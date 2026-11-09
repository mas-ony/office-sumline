/**
 * The editor of one workbook: menu bar, worksheet tabs, toolbar, formula bar and grid, with the
 * context menus, dialogs and cell pop-ups.
 */
import { useEffect, useRef } from 'react';
import { topMenus, type UIState } from '@sumline/core/views';
import { useUI } from './context';
import { Menu } from './Menu';
import { Toolbar } from './Toolbar';
import { FormulaAssist, FormulaBar } from './FormulaBar';
import { SheetTabs } from './SheetTabs';
import { FooterRow } from './FooterRow';
import { KeyBar } from './KeyBar';
import { CommentsBox, FilterPopup, NoteEditor } from './Popovers';
import { FormatDialog } from './dialogs/FormatDialog';
import { SearchDialog } from './dialogs/SearchDialog';
import { ValidationsDialog } from './dialogs/ValidationsDialog';
import { AboutDialog, ConfirmDialog, FunctionsDialog, LinkDialog, NameDialog, PromptDialog, ShortcutsDialog } from './dialogs/SmallDialogs';
import { ChartInsertDialog, ChartSettings } from './dialogs/ChartDialogs';
import { ShapesDialog, ShapeSettings } from './dialogs/ShapeDialogs';
import { PrintDialog } from './dialogs/PrintDialog';

/** The grid's element; the grid of the workbook is created in it. */
function GridHost() {
  const { ui } = useUI();
  const host = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = host.current!;
    ui.attachGrid(el);
    return () => ui.detachGrid(el);
  }, [ui]);
  return <div className="sl-grid-host" ref={host} data-testid="grid" />;
}

export function Editor({ state }: { state: UIState }) {
  const ctx = useUI();
  const { ui, doc, phone, state: docState } = ctx;
  const { edit, menu, topMenu, dialog, comments, note, filter, pending, toast } = state;
  const close = () => ui.closeDialog();
  const menus = topMenus(ctx);

  return (
    <>
      <div className="sl-editor" data-testid="editor">
        <div className="sl-menubar" role="menubar" aria-label="Menu">
          <div className="sl-menubar-items">
            {menus.map((m, i) => (
              <button
                key={m.title}
                type="button"
                role="menuitem"
                aria-haspopup="menu"
                aria-expanded={topMenu === i}
                className="sl-menu-title"
                data-testid={`menu-${m.title.toLowerCase()}`}
                onMouseDown={(e) => e.preventDefault()}
                onClick={(e) => ui.clickTopMenu(i, e.currentTarget.getBoundingClientRect())}
                onMouseEnter={(e) => ui.hoverTopMenu(i, e.currentTarget.getBoundingClientRect())}
              >
                {m.title}
              </button>
            ))}
          </div>
          <div className="sl-filename" title={docState.fileName}>
            {docState.modified && <span className="sl-dirty" title="Unsaved changes" />}
            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} data-testid="file-name">{docState.fileName}</span>
            {docState.readOnly && <span className="sl-hint">(view only)</span>}
          </div>
        </div>
        <SheetTabs />
        <div className="sl-sheet-frame">
          <Toolbar />
          <FormulaBar />
          <GridHost />
          {doc.sheet.footer && <FooterRow />}
        </div>
      </div>
      <FormulaAssist />
      {phone && edit.active && <KeyBar />}
      {menu && <Menu items={menu.items} at={menu.at} sheet={phone} title={menu.title} testId={menu.testId} onClose={ui.closeMenu} />}
      {dialog?.kind === 'format' && <FormatDialog onClose={close} initialTab={dialog.tab} />}
      {dialog?.kind === 'search' && <SearchDialog onClose={close} />}
      {dialog?.kind === 'validations' && <ValidationsDialog onClose={close} />}
      {dialog?.kind === 'name-edit' && <NameDialog onClose={close} name={dialog.name} />}
      {dialog?.kind === 'link' && <LinkDialog onClose={close} />}
      {dialog?.kind === 'about' && <AboutDialog onClose={close} />}
      {dialog?.kind === 'shortcuts' && <ShortcutsDialog onClose={close} />}
      {dialog?.kind === 'functions' && <FunctionsDialog onClose={close} />}
      {dialog?.kind === 'chart-insert' && <ChartInsertDialog onClose={close} />}
      {dialog?.kind === 'chart-settings' && <ChartSettings id={dialog.id} onClose={close} />}
      {dialog?.kind === 'shapes' && <ShapesDialog onClose={close} />}
      {dialog?.kind === 'shape-settings' && <ShapeSettings id={dialog.id} onClose={close} />}
      {dialog?.kind === 'print' && <PrintDialog onClose={close} />}
      {dialog?.kind === 'column-title' && (
        <PromptDialog
          title="Rename this column"
          message="Column title (shown in the column header)"
          value={doc.sheet.colTitles?.get(dialog.col) ?? ''}
          onAnswer={(v) => {
            if (v !== null) doc.setColumnTitle(dialog.col, v.trim() || null);
            close();
          }}
        />
      )}
      {comments && <CommentsBox r={comments.r} c={comments.c} rect={comments.rect} onClose={ui.closeComments} />}
      {note && <NoteEditor r={note.r} c={note.c} rect={note.rect} onClose={ui.closeNote} />}
      {filter && <FilterPopup col={filter.col} at={filter.at} onClose={ui.closeFilter} />}
      {pending?.kind === 'confirm' && (
        <ConfirmDialog message={pending.message} title={pending.opts?.title} ok={pending.opts?.ok} cancel={pending.opts?.cancel} danger={pending.opts?.danger} onAnswer={(v) => ui.answer(v)} />
      )}
      {pending?.kind === 'prompt' && <PromptDialog message={pending.message} value={pending.value} title={pending.opts?.title} ok={pending.opts?.ok} onAnswer={(v) => ui.answer(v)} />}
      {toast && (
        <div className="sl-toast" role="status" data-testid="toast">
          <span>{toast}</span>
          <button type="button" onClick={ui.dismissToast}>OK</button>
        </div>
      )}
    </>
  );
}
