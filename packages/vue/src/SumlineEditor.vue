<script setup lang="ts">
/**
 * The complete editor: the start screen until a workbook is open (with the startScreen option),
 * then the editor of the workbook, and the editor's own messages (errors, dragged files, files
 * being read).
 */
import { type DocumentInput, type FileActions, type SpreadsheetApi, type SumlineOptions, SumlineUI } from '@sumline/core/views';
import { onMounted, onUnmounted, type PropType, provide, ref, shallowRef, watch } from 'vue';
import { UI_KEY, type UIRef } from './context';
import Editor from './Editor.vue';
import StartScreen from './StartScreen.vue';

/** The complete editor. */
defineOptions({ name: 'SumlineEditor' });

// The options of SumlineOptions; its callback is the event below.
const props = defineProps({
  /** Opened at start, and again whenever it changes. */
  document: { type: null as unknown as PropType<DocumentInput>, default: undefined },
  /** Start on the start screen instead of a blank workbook; it also shows while `document` loads. Read at start. */
  startScreen: { type: Boolean, default: undefined },
  /** Workbooks open view-only. */
  readOnly: { type: Boolean, default: undefined },
  /** Shown on comments and notes. */
  author: { type: String, default: undefined },
  /** Keeps the recent colours and print options in localStorage under `<storageKey>-…`. */
  storageKey: { type: String, default: undefined },
  /** Keeps unsaved changes in an IndexedDB database of this name, to recover them after a crash. */
  autosave: { type: String, default: undefined },
  /** With `autosave` and `startScreen`: the start screen offers to recover unsaved changes (default true). */
  offerRecovery: { type: Boolean, default: undefined },
  /** Replace some of the built-in file actions (New, Open, Save, Download, Print…). */
  files: { type: Object as PropType<Partial<FileActions>>, default: undefined },
});

const emit = defineEmits<{
  /** A workbook is open (the initial document, or the first one opened from the start screen). */
  ready: [api: SpreadsheetApi];
}>();

function options(): SumlineOptions {
  return { ...props, onReady: (api) => emit('ready', api) };
}

// The first props create the interface; later ones are passed on.
const ui = new SumlineUI(options());
watch(() => ({ ...props }), () => ui.setOptions(options()));

const state = shallowRef(ui.getState());
const ctx = shallowRef(ui.getCtx());
const off = ui.subscribe(() => {
  state.value = ui.getState();
  ctx.value = ui.getCtx();
});
// Only the editor of a workbook (rendered while there is a context) uses it.
provide(UI_KEY, ctx as UIRef);

const root = ref<HTMLDivElement | null>(null);
onMounted(() => ui.attach(root.value!));
onUnmounted(() => {
  off();
  ui.detach();
});
</script>

<template>
  <div ref="root" class="sl-app">
    <Editor v-if="ctx" :key="state.docKey" :state="state" />
    <StartScreen
      v-else
      :restore="state.restore"
      @new="ui.startNew"
      @open="ui.files.open()"
      @file="ui.openFile($event)"
      @restore="ui.recover"
      @discard="ui.discardRecovery"
    />
    <div v-if="state.notice" class="sl-toast" role="alert" data-testid="notice">
      <span>{{ state.notice }}</span>
      <button type="button" @click="ui.notice(null)">OK</button>
    </div>
    <div v-if="state.dragging" class="sl-drop-hint">Drop a spreadsheet to open it, or an image to insert it</div>
    <div v-if="state.busy" class="sl-busy" role="status" aria-live="polite" data-testid="busy">
      <span class="sl-spinner" aria-hidden="true" />{{ state.busy }}
    </div>
  </div>
</template>
