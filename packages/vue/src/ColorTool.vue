<script setup lang="ts">
/** The text and background colour buttons of the toolbar, showing the colour of the active cell. */
import { cmd, rgbSpec } from '@sumline/core/views';
import { computed, ref, shallowRef } from 'vue';
import ColorPicker from './ColorPicker.vue';
import { useUI } from './context';
import Icon from './Icon.vue';

const props = defineProps<{ kind: 'text' | 'fill' }>();

const ui = useUI();
const btn = ref<HTMLButtonElement | null>(null);
const rect = shallowRef<DOMRect | null>(null);
const value = computed(() => (props.kind === 'text' ? ui.value.state.format.color : ui.value.state.format.fill));
const label = computed(() => (props.kind === 'text' ? 'Text Color' : 'Background Color'));
const apply = (hex: string | null) => {
  if (props.kind === 'text') cmd.format(ui.value, { font: { color: hex ? rgbSpec(hex) ?? null : null } }, 'Text color');
  else cmd.format(ui.value, { fill: hex ? { pattern: 'solid', fg: rgbSpec(hex) } : null }, 'Fill color');
};
const toggle = () => (rect.value = rect.value ? null : btn.value!.getBoundingClientRect());
</script>

<template>
  <button
    ref="btn"
    type="button"
    class="sl-tool sl-tool-color"
    :title="label"
    :aria-label="label"
    :aria-expanded="!!rect"
    :data-testid="kind === 'text' ? 'tb-text-color' : 'tb-fill-color'"
    :disabled="ui.state.readOnly"
    @mousedown.prevent
    @click="toggle"
  >
    <Icon :name="kind === 'text' ? 'format_color_text' : 'format_color_fill'" />
    <span class="swatch" :style="{ background: value ?? (kind === 'text' ? '#000000' : 'transparent') }" />
  </button>
  <ColorPicker v-if="rect" :anchor="rect" :value="value" :live="false" :test-id="`${kind}-color-picker`" @change="apply" @close="rect = null" />
</template>
