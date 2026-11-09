<script setup lang="ts">
/** A fixed pop-up next to a rectangle, closed by clicks outside and Escape. */
import { onMounted, onUnmounted, onUpdated, ref, watch } from 'vue';
import { useUI } from './context';

const props = defineProps<{ rect: DOMRect; boxClass: string; testId?: string; label?: string; width?: number }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { layer } = ui.value;
const el = ref<HTMLDivElement | null>(null);
// Placed after every render, as the whole interface renders again on each change.
const place = () => {
  const box = el.value;
  if (!box) return;
  const { rect } = props;
  const w = box.offsetWidth;
  const h = box.offsetHeight;
  const vw = window.innerWidth;
  const vh = window.visualViewport?.height ?? window.innerHeight;
  let x = rect.right + 4;
  let y = rect.top;
  if (x + w > vw - 4) x = rect.left - w - 4;
  if (x < 4) {
    x = Math.max(4, Math.min(rect.left, vw - w - 4));
    y = rect.bottom + 4;
  }
  if (y + h > vh - 4) y = Math.max(4, vh - h - 4);
  box.style.left = `${x}px`;
  box.style.top = `${y}px`;
};
onMounted(place);
onUpdated(place);
watch([ui, () => props.rect], place, { flush: 'post' });

const down = (e: PointerEvent) => {
  const t = e.target as HTMLElement;
  if (t.closest(`.${props.boxClass.split(' ')[0]}`) || t.closest('.sl-popup') || t.closest('.sl-bottom-sheet') || t.closest('.sl-sheet-backdrop')) return;
  emit('close');
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
</script>

<template>
  <Teleport :to="layer">
    <div ref="el" :class="boxClass" role="dialog" :aria-label="label" :data-testid="testId" :style="width ? { width: `${width}px` } : undefined">
      <slot />
    </div>
  </Teleport>
</template>
