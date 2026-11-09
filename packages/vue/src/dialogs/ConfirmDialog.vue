<script setup lang="ts">
/** A question with OK and Cancel (no Cancel button when `cancel` is empty). */
import Modal from '../Modal.vue';

withDefaults(defineProps<{ message: string; title?: string; ok?: string; cancel?: string; danger?: boolean }>(), { title: undefined, ok: 'Ok', cancel: 'Cancel' });
const emit = defineEmits<{ answer: [value: boolean] }>();
</script>

<template>
  <Modal :title="title ?? 'Confirm'" :width="400" blocking test-id="confirm-dialog" @close="emit('answer', false)">
    <p :style="{ margin: '0 0 20px', whiteSpace: 'pre-wrap' }">{{ message }}</p>
    <div :style="{ display: 'flex', gap: '6px', justifyContent: 'flex-end' }">
      <button type="button" :class="`sl-button${danger ? ' danger' : ''}`" data-testid="confirm-ok" data-autofocus="true" @click="emit('answer', true)">{{ ok }}</button>
      <button v-if="cancel" type="button" class="sl-button light" data-testid="confirm-cancel" @click="emit('answer', false)">{{ cancel }}</button>
    </div>
  </Modal>
</template>
