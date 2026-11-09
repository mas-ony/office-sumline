<script setup lang="ts">
/** A pop-up anchored under a toolbar button. */
import { onMounted, onUnmounted, ref, watch } from 'vue';
import { useUI } from './context';

const props = defineProps<{ anchor: DOMRect; extraClass?: string; label?: string; testId?: string }>();
const emit = defineEmits<{ close: [] }>();

const { layer } = useUI().value;
const el = ref<HTMLDivElement | null>(null);

const place = () => {
  const box = el.value;
  if (!box) return;
  const { anchor } = props;
  const w = box.offsetWidth;
  const h = box.offsetHeight;
  let x = anchor.left;
  let y = anchor.bottom + 2;
  if (x + w > window.innerWidth - 4) x = Math.max(4, window.innerWidth - w - 4);
  if (y + h > window.innerHeight - 4) y = Math.max(4, anchor.top - h - 2);
  box.style.left = `${x}px`;
  box.style.top = `${y}px`;
};
onMounted(place);
watch(() => props.anchor, place, { flush: 'post' });

const down = (e: PointerEvent) => {
  const t = e.target as HTMLElement;
  if (t.closest('.sl-picker') || t.closest('.sl-color')) return;
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
    <div ref="el" :class="['sl-picker', extraClass]" role="listbox" :aria-label="label" :data-testid="testId" @mousedown.prevent>
      <slot />
    </div>
  </Teleport>
</template>
