/** Crash/tab-kill recovery: the latest unsaved workbook is kept in an IndexedDB database (the `autosave` option). */

export interface AutosaveRecord {
  fileName: string;
  bytes: Uint8Array;
  savedAt: number;
}

const STORE = 'autosave';
const KEY = 'current';

function open(db: string): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(db, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function run<T>(name: string, mode: IDBTransactionMode, fn: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open(name);
  try {
    return await new Promise<T>((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const req = fn(tx.objectStore(STORE));
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

export async function saveAutosave(db: string, record: AutosaveRecord): Promise<void> {
  try {
    await run(db, 'readwrite', (s) => s.put(record, KEY));
  } catch {
    // Storage can be unavailable (private mode, quota); recovery is best effort.
  }
}

export async function loadAutosave(db: string): Promise<AutosaveRecord | null> {
  try {
    return ((await run(db, 'readonly', (s) => s.get(KEY))) as AutosaveRecord | undefined) ?? null;
  } catch {
    return null;
  }
}

export async function clearAutosave(db: string): Promise<void> {
  try {
    await run(db, 'readwrite', (s) => s.delete(KEY));
  } catch {
    // Ignore.
  }
}
