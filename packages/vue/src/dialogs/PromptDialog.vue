<script setup lang="ts">
/** A one-line prompt: Enter or OK answers with the text, Cancel with null. */
import { ref } from 'vue';
import Modal from '../Modal.vue';

const props = withDefaults(defineProps<{ message: string; value: string; title?: string; ok?: string }>(), { title: undefined, ok: 'Ok' });
const emit = defineEmits<{ answer: [value: string | null] }>();
const text = ref(props.value);
</script>

<template>
  <Modal :title="title ?? message" :width="400" blocking test-id="prompt-dialog" @close="emit('answer', null)">
    <div class="sl-form">
      <div class="sl-field">
        <label for="sl-prompt">{{ message }}</label>
        <input
          id="sl-prompt"
          class="sl-input"
          :value="text"
          data-testid="prompt-input"
          @focus="($event.target as HTMLInputElement).select()"
          @input="text = ($event.target as HTMLInputElement).value"
          @keydown="$event.key === 'Enter' && emit('answer', text)"
        />
      </div>
      <div :style="{ display: 'flex', gap: '6px', justifyContent: 'flex-end' }">
        <button type="button" class="sl-button" data-testid="prompt-ok" @click="emit('answer', text)">{{ ok }}</button>
        <button type="button" class="sl-button light" @click="emit('answer', null)">Cancel</button>
      </div>
    </div>
  </Modal>
</template>
