<script setup lang="ts">
/** The note editor at a cell (legacy notes). */
import { cellName } from '@sumline/core/views';
import { computed, ref } from 'vue';
import Anchored from './Anchored.vue';
import { useUI } from './context';
import { vFocus } from './directives';

const props = defineProps<{ r: number; c: number; rect: DOMRect }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const existing = computed(() => ui.value.doc.sheet.commentAt(props.r, props.c));
const text = ref(existing.value?.comments.map((x) => x.text).join('\n') ?? '');
const save = () => {
  doc.setNote(props.r, props.c, text.value.trim());
  emit('close');
};
const clear = () => {
  doc.setNote(props.r, props.c, '');
  emit('close');
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) save();
};
</script>

<template>
  <Anchored :rect="rect" box-class="sl-note" test-id="note-editor" :label="`Notes ${cellName(r, c)}`" @close="emit('close')">
    <strong :style="{ fontSize: '13px' }">Notes</strong>
    <textarea v-focus :value="text" aria-label="Note" data-testid="note-input" @input="text = ($event.target as HTMLTextAreaElement).value" @keydown="key" />
    <div class="sl-note-actions">
      <button v-if="existing" type="button" class="sl-button light" @click="clear">Clear</button>
      <button type="button" class="sl-button" data-testid="note-save" @click="save">Save</button>
      <button type="button" class="sl-button light" @click="emit('close')">Cancel</button>
    </div>
  </Anchored>
</template>
