/** The interface's context for the React components. */
import { createContext, useContext, useEffect, useState } from 'react';
import type { UIContextValue } from '@sumline/core/views';

export const UIContext = createContext<UIContextValue | null>(null);

export function useUI(): UIContextValue {
  const v = useContext(UIContext);
  if (!v) throw new Error('UIContext missing');
  return v;
}

/** A value remembered by the interface (in localStorage with the storageKey option). */
export function usePersistent<T>(name: string, initial: T): [T, (v: T) => void] {
  const { prefs } = useUI();
  const [value, setValue] = useState<T>(() => prefs.get(name, initial));
  useEffect(() => prefs.set(name, value), [prefs, name, value]);
  return [value, setValue];
}
