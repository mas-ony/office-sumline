<script setup lang="ts">
/**
 * Formula bar: the name box (with the list of defined names and "Add new entry"), fx, the formula
 * input with coloured references, and the expand button.
 */
import { formulaSegments } from '@sumline/core/views';
import { computed, onMounted, onUnmounted, onUpdated, ref, watch } from 'vue';
import { useUI } from './context';
import Icon from './Icon.vue';
import NameBox from './NameBox.vue';

const ui = useUI();
const input = ref<HTMLTextAreaElement | null>(null);
const mirror = ref<HTMLDivElement | null>(null);
let reported: { text: string; caret: number } | null = null;
const expanded = ref(false);
const value = computed(() => (ui.value.edit.active ? ui.value.edit.text : ui.value.state.editText));
/** The text with cell references coloured like the reference highlights drawn on the grid. */
const segments = computed(() => formulaSegments(value.value));
/** A last empty line needs a character to take up its height. */
const tail = computed(() => (value.value.endsWith('\n') ? '​' : ''));

// Keep focus and caret in the bar when the edit changes from elsewhere (autocomplete, clicking
// cells to insert references).
const followEdit = () => {
  const el = input.value;
  const { edit } = ui.value;
  if (!el || !edit.active || edit.origin !== 'bar') return;
  const r = reported;
  if (r && r.text === edit.text && r.caret === edit.caret) return;
  reported = { text: edit.text, caret: edit.caret };
  if (document.activeElement !== el) el.focus({ preventScroll: true });
  el.setSelectionRange(edit.caret, edit.caret);
};
onMounted(followEdit);
watch(() => ui.value.edit, followEdit, { flush: 'post' });

const syncScroll = () => {
  if (mirror.value && input.value) {
    mirror.value.scrollTop = input.value.scrollTop;
    mirror.value.scrollLeft = input.value.scrollLeft;
  }
};
onMounted(syncScroll);
onUpdated(syncScroll);

const report = () => {
  const el = input.value;
  const { state, grid } = ui.value;
  if (!el || state.readOnly || !grid) return;
  const caret = el.selectionStart ?? el.value.length;
  reported = { text: el.value, caret };
  grid.setEditFromBar(el.value, caret);
};

// The caret is reported as it moves, as React's onSelect does: on selectionchange, keys and mouse
// up (not while a mouse button is down), once per change.
let lastSelection: { start: number; end: number } | null = null;
let mouseDown = false;
const checkSelection = () => {
  const el = input.value;
  if (mouseDown || !el || document.activeElement !== el) return;
  const start = el.selectionStart ?? 0;
  const end = el.selectionEnd ?? 0;
  if (lastSelection && lastSelection.start === start && lastSelection.end === end) return;
  lastSelection = { start, end };
  if (ui.value.edit.active) report();
};
const pressed = () => (mouseDown = true);
const released = () => {
  mouseDown = false;
  checkSelection();
};
const RELEASE = ['mouseup', 'contextmenu', 'dragend'] as const;
onMounted(() => {
  document.addEventListener('selectionchange', checkSelection);
  document.addEventListener('keydown', checkSelection);
  document.addEventListener('keyup', checkSelection);
  document.addEventListener('mousedown', pressed);
  for (const t of RELEASE) document.addEventListener(t, released);
});
onUnmounted(() => {
  document.removeEventListener('selectionchange', checkSelection);
  document.removeEventListener('keydown', checkSelection);
  document.removeEventListener('keyup', checkSelection);
  document.removeEventListener('mousedown', pressed);
  for (const t of RELEASE) document.removeEventListener(t, released);
});

const focus = () => {
  lastSelection = null;
  const { state, grid } = ui.value;
  if (state.readOnly || !grid) return;
  if (!grid.isEditing()) {
    grid.startEdit(undefined, 'edit', undefined, 'bar');
    reported = { text: grid.editState.text, caret: grid.editState.caret };
  }
};
const blurred = () => (lastSelection = null);
const keydown = (e: KeyboardEvent) => {
  const { grid, edit } = ui.value;
  if (!grid) return;
  if (edit.active && ui.value.ui.assist.handleKey(e)) {
    e.preventDefault();
    return;
  }
  if (e.isComposing) return;
  if (e.key === 'Enter' && e.altKey) {
    e.preventDefault();
    const el = e.currentTarget as HTMLTextAreaElement;
    const s = el.selectionStart ?? el.value.length;
    el.value = `${el.value.slice(0, s)}\n${el.value.slice(el.selectionEnd ?? s)}`;
    el.setSelectionRange(s + 1, s + 1);
    report();
  } else if (e.key === 'Enter') {
    e.preventDefault();
    grid.commitEdit({ dr: e.shiftKey ? -1 : 1, dc: 0 });
  } else if (e.key === 'Tab') {
    e.preventDefault();
    grid.commitEdit({ dr: 0, dc: e.shiftKey ? -1 : 1 });
  } else if (e.key === 'Escape') {
    e.preventDefault();
    grid.cancelEdit();
  }
};
const scrolled = (e: Event) => {
  const el = e.currentTarget as HTMLTextAreaElement;
  if (mirror.value) {
    mirror.value.scrollTop = el.scrollTop;
    mirror.value.scrollLeft = el.scrollLeft;
  }
};
</script>

<template>
  <div :class="`sl-formulabar${expanded ? ' expanded' : ''}`" role="group" aria-label="Formula editor" data-testid="formula-bar">
    <NameBox />
    <div class="sl-fx" aria-hidden="true">fx</div>
    <div :style="{ position: 'relative', flex: '1 1 auto', minWidth: 0, display: 'flex' }">
      <div ref="mirror" class="sl-formula-input" aria-hidden="true" :style="{ position: 'absolute', inset: 0, pointerEvents: 'none', color: '#111' }"><template v-for="s in segments" :key="s.start"><span v-if="s.color" :style="{ color: s.color }">{{ s.text }}</span><template v-else>{{ s.text }}</template></template>{{ tail }}</div>
      <textarea
        ref="input"
        class="sl-formula-input"
        aria-label="Formula bar"
        data-testid="formula-input"
        rows="1"
        spellcheck="false"
        autocapitalize="off"
        autocomplete="off"
        autocorrect="off"
        :enterkeyhint="ui.phone ? 'done' : undefined"
        :readonly="ui.state.readOnly"
        :value="value"
        :style="{ color: 'transparent', caretColor: '#111', position: 'relative' }"
        @focus="focus"
        @blur="blurred"
        @input="report"
        @keydown="keydown"
        @scroll="scrolled"
      />
    </div>
    <button
      type="button"
      class="sl-formula-expand"
      :aria-label="expanded ? 'Collapse formula bar' : 'Expand formula bar'"
      data-testid="fx-expand"
      @mousedown.prevent
      @click="expanded = !expanded"
    >
      <Icon :name="expanded ? 'expand_less' : 'expand_more'" />
    </button>
  </div>
</template>
