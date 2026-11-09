<script setup lang="ts">
/** The colour spectrum of the colour picker: drag on it to pick. */
import { rgbToHex } from '@sumline/core/views';
import { onMounted, ref, shallowRef } from 'vue';

const emit = defineEmits<{ pick: [hex: string] }>();
const canvas = ref<HTMLCanvasElement | null>(null);
const point = shallowRef<{ x: number; y: number } | null>(null);
onMounted(() => {
  const c = canvas.value;
  const ctx = c?.getContext('2d');
  if (!c || !ctx) return;
  const w = c.width;
  const h = c.height;
  const hue = ctx.createLinearGradient(0, 0, w, 0);
  for (let i = 0; i <= 6; i++) hue.addColorStop(i / 6, `hsl(${i * 60}, 100%, 50%)`);
  ctx.fillStyle = hue;
  ctx.fillRect(0, 0, w, h);
  const shade = ctx.createLinearGradient(0, 0, 0, h);
  shade.addColorStop(0, 'rgba(255,255,255,1)');
  shade.addColorStop(0.5, 'rgba(255,255,255,0)');
  shade.addColorStop(0.5, 'rgba(0,0,0,0)');
  shade.addColorStop(1, 'rgba(0,0,0,1)');
  ctx.fillStyle = shade;
  ctx.fillRect(0, 0, w, h);
});
const pick = (e: PointerEvent) => {
  const c = canvas.value;
  if (!c) return;
  const rect = c.getBoundingClientRect();
  const x = Math.max(0, Math.min(rect.width - 1, e.clientX - rect.left));
  const y = Math.max(0, Math.min(rect.height - 1, e.clientY - rect.top));
  point.value = { x, y };
  const data = c.getContext('2d')?.getImageData(Math.floor((x / rect.width) * c.width), Math.floor((y / rect.height) * c.height), 1, 1).data;
  if (data) emit('pick', rgbToHex(data[0], data[1], data[2]));
};
const down = (e: PointerEvent) => {
  (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  pick(e);
};
const move = (e: PointerEvent) => {
  if (e.buttons) pick(e);
};
</script>

<template>
  <div class="sl-color-spectrum" @pointerdown="down" @pointermove="move">
    <canvas ref="canvas" width="232" height="160" />
    <div v-if="point" class="sl-color-point" :style="{ left: `${point.x}px`, top: `${point.y}px` }" />
  </div>
</template>
