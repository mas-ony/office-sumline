/** Small browser helpers of the interface: media queries and touch input. */

/** Phones: narrow screens (bottom sheets, full-screen dialogs and the key bar). */
export const PHONE_QUERY = '(max-width: 720px)';

export function matchesMedia(query: string): boolean {
  return typeof window !== 'undefined' && !!window.matchMedia && window.matchMedia(query).matches;
}

/** Calls `fn` when the query starts or stops matching; returns the function that stops watching. */
export function watchMedia(query: string, fn: (matches: boolean) => void): () => void {
  const mq = typeof window !== 'undefined' ? window.matchMedia?.(query) : undefined;
  if (!mq) return () => undefined;
  const on = () => fn(mq.matches);
  mq.addEventListener?.('change', on);
  return () => mq.removeEventListener?.('change', on);
}

/** Whether the primary input is touch (phones and tablets). */
export function isTouchDevice(): boolean {
  if (typeof window === 'undefined') return false;
  return window.matchMedia?.('(pointer: coarse)').matches || 'ontouchstart' in window;
}
