/** The interface's context for the Vue components. */
import type { UIContextValue } from '@sumline/core/views';
import { inject, type InjectionKey, type Ref, type ShallowRef, shallowRef, watch } from 'vue';

/** The interface: a snapshot of its state with the controller's actions, replaced on every change. */
export type UIRef = Readonly<ShallowRef<UIContextValue>>;

export const UI_KEY: InjectionKey<UIRef> = Symbol('sumline-ui');

export function useUI(): UIRef {
  const v = inject(UI_KEY);
  if (!v) throw new Error('UIContext missing');
  return v;
}

/** A value remembered by the interface (in localStorage with the storageKey option). */
export function usePersistent<T>(name: string, initial: T): Ref<T> {
  const { prefs } = useUI().value;
  const value = shallowRef(prefs.get(name, initial)) as Ref<T>;
  watch(value, (v) => prefs.set(name, v), { immediate: true });
  return value;
}
