<script setup lang="ts">
/** A toolbar button that opens a pop-up (the `content` slot, given `close`). */
import { ref, shallowRef } from 'vue';
import Popup from './Popup.vue';

defineProps<{ label: string; testId: string; extraClass?: string }>();

const btn = ref<HTMLButtonElement | null>(null);
const rect = shallowRef<DOMRect | null>(null);
const toggle = () => (rect.value = rect.value ? null : btn.value!.getBoundingClientRect());
const close = () => (rect.value = null);
</script>

<template>
  <button
    ref="btn"
    type="button"
    :class="`sl-tool${extraClass ? ` ${extraClass}` : ''}`"
    :title="label"
    :aria-label="label"
    aria-haspopup="listbox"
    :aria-expanded="!!rect"
    :data-testid="testId"
    @mousedown.prevent
    @click="toggle"
  >
    <slot />
  </button>
  <Popup v-if="rect" :anchor="rect" :label="label" :test-id="`${testId}-list`" @close="close">
    <slot name="content" :close="close" />
  </Popup>
</template>
