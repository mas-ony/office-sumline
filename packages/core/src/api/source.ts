/** The ways a workbook can be given (URL, file handle, File, Blob, bytes), and saving files. */

/**
 * A workbook to open: a URL, a file handle (File System Access API: Save writes back to the file),
 * a File or Blob, or the bytes.
 */
export type DocumentSource = string | FileSystemFileHandle | Blob | ArrayBuffer | Uint8Array;

/** A document option: a source, or a source with the file name to show. */
export type DocumentInput = DocumentSource | { data?: DocumentSource; url?: string; name?: string };

export const isFileHandle = (f: unknown): f is FileSystemFileHandle => typeof FileSystemFileHandle !== 'undefined' && f instanceof FileSystemFileHandle;

/** The source and name of a document option (no source for an empty object). */
export function documentOf(input: DocumentInput): { source?: DocumentSource; name?: string } {
  if (typeof input === 'string' || isFileHandle(input) || input instanceof Blob || input instanceof ArrayBuffer || input instanceof Uint8Array) return { source: input };
  return { source: input.data ?? input.url, name: typeof input.name === 'string' ? input.name : undefined };
}

/** File name in a URL: its last path segment. */
export function nameOfUrl(url: string): string {
  return decodeURIComponent(new URL(url, location.href).pathname.split('/').pop() || 'Workbook.xlsx');
}

export function downloadBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}
