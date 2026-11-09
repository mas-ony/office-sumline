/**
 * The complete editor: the start screen until a workbook is open (with the startScreen option),
 * then the editor of the workbook, and the editor's own messages (errors, dragged files, files
 * being read).
 */
import { useEffect, useRef, useState, useSyncExternalStore, type CSSProperties } from 'react';
import { type SumlineOptions, SumlineUI, type UIState } from '@sumline/core/views';
import { UIContext } from './context';
import { Editor } from './Editor';
import { StartScreen } from './StartScreen';

export interface SumlineEditorProps extends SumlineOptions {
  className?: string;
  style?: CSSProperties;
}

function Messages({ ui, state }: { ui: SumlineUI; state: UIState }) {
  const { notice, dragging, busy } = state;
  return (
    <>
      {notice && (
        <div className="sl-toast" role="alert" data-testid="notice">
          <span>{notice}</span>
          <button type="button" onClick={() => ui.notice(null)}>OK</button>
        </div>
      )}
      {dragging && <div className="sl-drop-hint">Drop a spreadsheet to open it, or an image to insert it</div>}
      {busy && (
        <div className="sl-busy" role="status" aria-live="polite" data-testid="busy">
          <span className="sl-spinner" aria-hidden="true" />
          {busy}
        </div>
      )}
    </>
  );
}

export function SumlineEditor({ className, style, ...options }: SumlineEditorProps) {
  const [ui] = useState(() => new SumlineUI(options));
  const state = useSyncExternalStore(ui.subscribe, ui.getState, ui.getState);
  const ctx = ui.getCtx();
  const root = useRef<HTMLDivElement>(null);

  // Later props are passed on (the first ones created the interface).
  const mounted = useRef(false);
  useEffect(() => {
    if (mounted.current) ui.setOptions(options);
  });

  useEffect(() => {
    mounted.current = true;
    ui.attach(root.current!);
    return () => ui.detach();
  }, [ui]);

  return (
    <div ref={root} className={`sl-app${className ? ` ${className}` : ''}`} style={style}>
      {ctx ? (
        <UIContext.Provider value={ctx}>
          <Editor key={state.docKey} state={state} />
        </UIContext.Provider>
      ) : (
        <StartScreen
          onNew={ui.startNew}
          onOpen={() => void ui.files.open()}
          onFile={(f) => void ui.openFile(f)}
          restore={state.restore}
          onRestore={ui.recover}
          onDiscard={ui.discardRecovery}
        />
      )}
      <Messages ui={ui} state={state} />
    </div>
  );
}
