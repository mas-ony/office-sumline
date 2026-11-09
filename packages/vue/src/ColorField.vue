<script setup lang="ts">
/** A colour text field with a palette button opening the colour picker. */
import { normalizeHex } from '@sumline/core/views';
import { ref, shallowRef } from 'vue';
import ColorPicker from './ColorPicker.vue';
import Icon from './Icon.vue';

defineProps<{ value: string; placeholder?: string; disabled?: boolean; testId?: string; id?: string }>();
const emit = defineEmits<{ change: [hex: string] }>();

const anchor = shallowRef<DOMRect | null>(null);
const text = ref<string | null>(null);
const typed = (e: Event) => {
  const v = (e.target as HTMLInputElement).value;
  text.value = v;
  const hex = normalizeHex(v);
  if (hex || v === '') emit('change', hex ?? '');
};
const picked = (hex: string | null) => {
  emit('change', hex ?? '');
  text.value = null;
};
</script>

<template>
  <div class="sl-range-input sl-color-field">
    <input
      :id="id"
      class="sl-input"
      :value="text ?? value"
      :placeholder="placeholder"
      :disabled="disabled"
      spellcheck="false"
      :data-testid="testId"
      :style="value ? { color: value } : undefined"
      @input="typed"
      @blur="text = null"
    />
    <button
      type="button"
      aria-label="Pick a colour"
      :disabled="disabled"
      :data-testid="testId ? `${testId}-pick` : undefined"
      @click="anchor = ($event.currentTarget as HTMLElement).getBoundingClientRect()"
    >
      <Icon name="palette" :size="18" />
    </button>
    <ColorPicker v-if="anchor" :anchor="anchor" :value="value || null" :live="false" @change="picked" @close="anchor = null" />
  </div>
</template>
