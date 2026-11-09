<script setup lang="ts">
/** Colour picker pop-up: a palette grid, a spectrum and RGB sliders, with Reset and Done. */
import { hexToRgb, normalizeHex, PALETTE_ROWS, rgbToHex } from '@sumline/core/views';
import { computed, onMounted, onUnmounted, ref, watch } from 'vue';
import ColorSpectrum from './ColorSpectrum.vue';
import { usePersistent, useUI } from './context';

const props = withDefaults(
  defineProps<{
    anchor: DOMRect;
    value: string | null;
    /** Applies immediately on every pick (true) or only on Done (false). */
    live?: boolean;
    testId?: string;
  }>(),
  { live: true, testId: undefined },
);
const emit = defineEmits<{ change: [hex: string | null]; close: [] }>();

const { layer } = useUI().value;
const el = ref<HTMLDivElement | null>(null);
const TABS = ['grid', 'spectrum', 'sliders'] as const;
const SLIDERS = ['Red', 'Green', 'Blue'] as const;
const tab = ref<(typeof TABS)[number]>('grid');
const current = ref(props.value ?? '#000000');
const text = ref(props.value ?? '');
const recent = usePersistent<string[]>('recent-colors', []);

const place = () => {
  const box = el.value;
  if (!box) return;
  const w = box.offsetWidth;
  const h = box.offsetHeight;
  const { anchor } = props;
  let x = anchor.left;
  let y = anchor.bottom + 4;
  if (x + w > window.innerWidth - 4) x = Math.max(4, window.innerWidth - w - 4);
  if (y + h > window.innerHeight - 4) y = Math.max(4, anchor.top - h - 4);
  box.style.left = `${x}px`;
  box.style.top = `${y}px`;
};
onMounted(place);
watch(() => props.anchor, place, { flush: 'post' });

const down = (e: PointerEvent) => {
  if (!(e.target as HTMLElement).closest?.('.sl-color')) emit('close');
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Escape') emit('close');
};
const timer = window.setTimeout(() => {
  window.addEventListener('pointerdown', down, true);
  window.addEventListener('keydown', key);
}, 0);
onUnmounted(() => {
  window.clearTimeout(timer);
  window.removeEventListener('pointerdown', down, true);
  window.removeEventListener('keydown', key);
});

const pick = (hex: string) => {
  current.value = hex;
  text.value = hex;
  if (props.live) emit('change', hex);
};
const done = () => {
  const hex = normalizeHex(text.value) ?? current.value;
  emit('change', hex);
  recent.value = [hex, ...recent.value.filter((c) => c !== hex)].slice(0, 12);
  emit('close');
};
const reset = () => {
  emit('change', null);
  emit('close');
};
const rgb = computed(() => hexToRgb(current.value));
const slide = (k: number, e: Event) => {
  const next = [...rgb.value];
  next[k] = Number((e.target as HTMLInputElement).value);
  pick(rgbToHex(next[0], next[1], next[2]));
};
const typed = (e: Event) => {
  text.value = (e.target as HTMLInputElement).value;
  const hex = normalizeHex(text.value);
  if (hex) current.value = hex;
};
</script>

<template>
  <Teleport :to="layer">
    <div ref="el" class="sl-color" role="dialog" aria-label="Colour" :data-testid="testId">
      <div class="sl-color-head">
        <button type="button" data-testid="color-reset" @click="reset">Reset</button>
        <button type="button" data-testid="color-done" @click="done">Done</button>
      </div>
      <div class="sl-segmented" role="tablist">
        <button v-for="t in TABS" :key="t" type="button" role="tab" :aria-selected="tab === t" :class="tab === t ? 'selected' : ''" @click="tab = t">
          {{ t[0].toUpperCase() + t.slice(1) }}
        </button>
      </div>
      <div v-if="tab === 'grid'" class="sl-color-grid" :style="{ gridTemplateColumns: `repeat(${PALETTE_ROWS[0].length}, 1fr)` }">
        <template v-for="(row, i) in PALETTE_ROWS" :key="i">
          <button
            v-for="(hex, j) in row"
            :key="`${i}-${j}`"
            type="button"
            :title="hex"
            :aria-label="hex"
            :data-color="hex"
            :class="current.toUpperCase() === hex ? 'selected' : ''"
            :style="{ background: hex }"
            @click="pick(hex)"
          />
        </template>
      </div>
      <ColorSpectrum v-if="tab === 'spectrum'" @pick="pick" />
      <div v-if="tab === 'sliders'" class="sl-color-sliders">
        <label v-for="(label, k) in SLIDERS" :key="label">
          {{ label }}<input type="range" :min="0" :max="255" :value="rgb[k]" :aria-label="label" @input="slide(k, $event)" />
          <span>{{ rgb[k] }}</span>
        </label>
      </div>
      <div v-if="recent.length > 0" class="sl-color-recent" aria-label="Recent colours">
        <button v-for="hex in recent" :key="hex" type="button" :title="hex" :style="{ background: hex }" @click="pick(hex)" />
      </div>
      <div class="sl-color-foot">
        <span class="swatch" :style="{ background: normalizeHex(text) ?? current }" />
        <input :value="text" placeholder="#000000" aria-label="Hex colour" @input="typed" @keydown.enter="done" />
      </div>
    </div>
  </Teleport>
</template>
