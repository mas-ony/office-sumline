<script setup lang="ts">
/** The link of the active cell: its text and a web address or a place in the workbook. */
import { ref } from 'vue';
import { useUI } from '../context';
import Modal from '../Modal.vue';

const emit = defineEmits<{ close: [] }>();

const { doc } = useUI().value;
const { r, c } = doc.selection.active;
const current = doc.sheet.hyperlinkAt(r, c);
const target = ref(current?.target ?? (current?.location ? `#${current.location}` : ''));
const text = ref(doc.displayText(r, c));
const save = () => {
  const t = target.value.trim();
  if (!t) return;
  const url = /^(#|[a-z][a-z0-9+.-]*:)/i.test(t) || /!/.test(t) ? t : /^www\./i.test(t) || /\.[a-z]{2,}(\/|$)/i.test(t) ? `https://${t}` : t;
  doc.setLink(url, text.value && text.value !== doc.displayText(r, c) ? text.value : text.value ? undefined : url);
  emit('close');
};
const remove = () => {
  doc.removeLink();
  emit('close');
};
</script>

<template>
  <Modal title="Link" icon="link" :width="420" test-id="link-dialog" @close="emit('close')">
    <div class="sl-form">
      <div class="sl-field">
        <label for="ln-text">Text</label>
        <input id="ln-text" class="sl-input" :value="text" data-testid="link-text" @input="text = ($event.target as HTMLInputElement).value" />
      </div>
      <div class="sl-field">
        <label for="ln-url">Web address or place in this workbook (Sheet2!A1)</label>
        <input
          id="ln-url"
          class="sl-input"
          :value="target"
          placeholder="https://"
          inputmode="url"
          data-testid="link-url"
          @input="target = ($event.target as HTMLInputElement).value"
          @keydown="$event.key === 'Enter' && save()"
        />
      </div>
      <div :style="{ display: 'flex', gap: '6px' }">
        <button type="button" class="sl-button" data-testid="link-save" @click="save">Save</button>
        <button type="button" class="sl-button" @click="emit('close')">Cancel</button>
        <button v-if="current" type="button" class="sl-button light" :style="{ marginLeft: 'auto' }" data-testid="link-remove" @click="remove">Remove link</button>
      </div>
    </div>
  </Modal>
</template>
