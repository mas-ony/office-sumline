<script setup lang="ts">
/** A defined name: its name and the cell range it refers to. */
import { quoteSheetName, rangeName } from '@sumline/core/views';
import { ref } from 'vue';
import { useUI } from '../context';
import Modal from '../Modal.vue';

const props = defineProps<{ name?: string }>();
const emit = defineEmits<{ close: [] }>();

const { doc } = useUI().value;
const existing = props.name ? doc.wb.names.find((n) => n.localSheet === undefined && n.name === props.name) : undefined;
const value = ref(existing?.name ?? '');
const reference = ref(existing ? `=${existing.formula}` : `=${quoteSheetName(doc.sheet.name)}!${rangeName(doc.range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`);
const error = ref('');
const save = () => {
  const err = doc.setDefinedName(value.value, reference.value, existing?.name);
  if (err) error.value = err;
  else emit('close');
};
const remove = (name: string) => {
  doc.deleteDefinedName(name);
  emit('close');
};
</script>

<template>
  <Modal title="Defined names" :width="380" test-id="name-dialog" @close="emit('close')">
    <div class="sl-form">
      <div class="sl-field">
        <label for="dn-name">Interval Name</label>
        <input id="dn-name" class="sl-input" :value="value" data-testid="dn-name" @input="value = ($event.target as HTMLInputElement).value" @keydown="$event.key === 'Enter' && save()" />
      </div>
      <div class="sl-field">
        <label for="dn-range">Cell range</label>
        <input
          id="dn-range"
          class="sl-input"
          :value="reference"
          spellcheck="false"
          data-testid="dn-range"
          @input="reference = ($event.target as HTMLInputElement).value"
          @keydown="$event.key === 'Enter' && save()"
        />
      </div>
      <div v-if="error" class="sl-error" role="alert">{{ error }}</div>
      <div :style="{ display: 'flex', gap: '6px' }">
        <button type="button" class="sl-button" data-testid="dn-save" @click="save">Save</button>
        <button type="button" class="sl-button" @click="emit('close')">Cancel</button>
        <button v-if="existing" type="button" class="sl-button danger" :style="{ marginLeft: 'auto' }" data-testid="dn-delete" @click="remove(existing.name)">Delete</button>
      </div>
    </div>
  </Modal>
</template>
