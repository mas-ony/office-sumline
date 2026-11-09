/**
 * What makes the editor a full-page app: the workbook and mode in the address, stored settings and
 * autosave, the window title, warning before leaving with unsaved changes, files passed by the
 * installed app, and the page API on `window`.
 */
import { installPageApi, nameOfUrl, type SpreadsheetApi, type SumlineOptions } from '@sumline/core';

const APP = 'Sumline';

/**
 * Options from the address: `?src=` opens a workbook (the start screen shows while it loads),
 * `?new=1` starts with a blank workbook instead of the start screen, `?readonly=1` or `?mode=view`
 * opens view-only.
 */
export function shellOptions(): SumlineOptions {
  const params = new URLSearchParams(location.search);
  const src = params.get('src');
  return {
    storageKey: 'sumline',
    autosave: 'sumline',
    startScreen: !!src || params.get('new') !== '1',
    readOnly: params.get('readonly') === '1' || params.get('mode') === 'view',
    document: src ? { url: src, name: nameOfUrl(src) } : undefined,
  };
}

interface LaunchQueue {
  setConsumer(fn: (p: { files: FileSystemFileHandle[] }) => void): void;
}

/** Files opened through the installed app (PWA file handling), also from the start screen. */
export function watchLaunches(open: (handle: FileSystemFileHandle) => void): void {
  (window as unknown as { launchQueue?: LaunchQueue }).launchQueue?.setConsumer((p) => {
    const handle = p.files?.[0];
    if (handle) open(handle);
  });
}

/** Connect the editor to the page once a workbook is open; returns the function that disconnects it. */
export function connectShell(api: SpreadsheetApi): () => void {
  const title = (name: string, modified: boolean) => {
    document.title = `${modified ? '• ' : ''}${name} – ${APP}`;
  };
  title(api.doc.fileName, false);
  const offs = [
    api.on('open', ({ fileName }) => title(fileName, false)),
    api.on('change', () => {
      if (api.doc.modified) title(api.doc.fileName, true);
    }),
    api.on('file', ({ name }) => title(name, api.doc.modified)),
    installPageApi(api),
  ];
  const before = (e: BeforeUnloadEvent) => {
    if (api.doc.modified) {
      e.preventDefault();
      e.returnValue = '';
    }
  };
  window.addEventListener('beforeunload', before);
  return () => {
    for (const off of offs) off();
    window.removeEventListener('beforeunload', before);
  };
}

/** The installed app works offline (production builds in secure contexts; plain-HTTP LAN addresses skip it). */
export function registerServiceWorker(): void {
  if (window.isSecureContext && 'serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js').catch(() => undefined);
    });
  }
}
