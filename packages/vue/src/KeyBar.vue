<script setup lang="ts">
/**
 * Phones: a row of keys above the on-screen keyboard while a cell is edited (formula characters
 * that are hard to reach on mobile keyboards, plus done/cancel and the function list).
 */
import { onMounted, onUnmounted, ref } from 'vue';
import { useUI } from './context';
import Icon from './Icon.vue';

const KEYS = ['=', '(', ')', ',', ':', '+', '-', '*', '/', '&', '$', '%', '"', '<', '>', ';'];

const ui = useUI();

// Height of the on-screen keyboard overlapping the layout viewport (iOS does not resize it).
const offset = ref(0);
const vv = window.visualViewport;
const update = () => {
  if (vv) offset.value = Math.max(0, Math.round(window.innerHeight - vv.height - vv.offsetTop));
};
onMounted(() => {
  vv?.addEventListener('resize', update);
  vv?.addEventListener('scroll', update);
  update();
});
onUnmounted(() => {
  vv?.removeEventListener('resize', update);
  vv?.removeEventListener('scroll', update);
});

const insert = (s: string) => {
  const { grid } = ui.value;
  if (!grid) return;
  const { text, caret } = grid.editState;
  grid.replaceEditText(text.slice(0, caret) + s + text.slice(caret), caret + s.length);
};
const keep = (e: Event) => e.preventDefault();
const blur = () => (document.activeElement as HTMLElement | null)?.blur();
</script>

<template>
  <div class="sl-keybar" :style="{ transform: offset ? `translateY(${-offset}px)` : undefined }" role="toolbar" aria-label="Cell editing keys" data-testid="keybar">
    <button type="button" class="sl-key" aria-label="Cancel" data-testid="key-cancel" @pointerdown="keep" @mousedown="keep" @click="ui.grid?.cancelEdit()">
      <Icon name="close" />
    </button>
    <button type="button" class="sl-key primary" aria-label="Done" data-testid="key-done" @pointerdown="keep" @mousedown="keep" @click="ui.grid?.commitEdit({ dr: 1, dc: 0 })">
      <Icon name="check" />
    </button>
    <button v-for="k in KEYS" :key="k" type="button" class="sl-key" :aria-label="`Insert ${k}`" @pointerdown="keep" @mousedown="keep" @click="insert(k)">{{ k }}</button>
    <button type="button" class="sl-key" aria-label="Functions" data-testid="key-functions" @pointerdown="keep" @mousedown="keep" @click="ui.openDialog({ kind: 'functions' })">
      <Icon name="functions" />
    </button>
    <button type="button" class="sl-key" aria-label="Hide keyboard" @pointerdown="keep" @mousedown="keep" @click="blur">
      <Icon name="keyboard" />
    </button>
  </div>
</template>
