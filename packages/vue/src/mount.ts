import type { SpreadsheetApi, SumlineOptions } from '@sumline/core';
import { createApp, h, shallowReactive } from 'vue';
import SumlineEditor from './SumlineEditor.vue';

export interface MountedEditor {
  /** Resolves with the page API once a workbook is open. */
  ready: Promise<SpreadsheetApi>;
  /** Change options, as new props would. */
  update(options: SumlineOptions): void;
  /** Remove the editor from the page. */
  destroy(): void;
}

/** Render the editor into an element, without writing Vue code. */
export function mount(element: HTMLElement, options: SumlineOptions = {}): MountedEditor {
  let resolve!: (api: SpreadsheetApi) => void;
  const ready = new Promise<SpreadsheetApi>((r) => (resolve = r));
  const current = shallowReactive<SumlineOptions>({ ...options });
  const app = createApp({
    render() {
      const { onReady, ...props } = current;
      return h(SumlineEditor, {
        ...props,
        onReady: (api: SpreadsheetApi) => {
          resolve(api);
          onReady?.(api);
        },
      });
    },
  });
  app.mount(element);
  return {
    ready,
    update(next) {
      Object.assign(current, next);
    },
    destroy() {
      app.unmount();
    },
  };
}
