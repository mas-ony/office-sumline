<script setup lang="ts">
/** A draggable modal window with a title bar (icon, title, close), or a panel on the right. */
import type { IconName } from '@sumline/core/views';
import { onMounted, onUnmounted, ref, type StyleValue, watch } from 'vue';
import { useUI } from './context';
import Icon from './Icon.vue';

const props = withDefaults(
  defineProps<{
    title: string;
    icon?: IconName;
    width?: number;
    height?: number | string;
    /** Docked to the right edge, full height (chart settings). */
    panel?: boolean;
    /** Dims and blocks the page behind. */
    blocking?: boolean;
    testId?: string;
    bodyClass?: string;
    bodyStyle?: StyleValue;
  }>(),
  { width: 430 },
);
const emit = defineEmits<{ close: [] }>();

const { layer } = useUI().value;
const el = ref<HTMLDivElement | null>(null);
let drag: { dx: number; dy: number; id: number } | null = null;

const position = () => {
  const box = el.value;
  if (!box || props.panel) return;
  const w = box.offsetWidth;
  const h = box.offsetHeight;
  box.style.left = `${Math.max(8, (window.innerWidth - w) / 2)}px`;
  box.style.top = `${Math.max(8, Math.min((window.innerHeight - h) / 2, 120))}px`;
};
watch(() => props.panel, position, { flush: 'post' });

const key = (e: KeyboardEvent) => {
  if (e.key === 'Escape' && !e.defaultPrevented) {
    // Escape inside an open drop-down of the dialog closes the drop-down only.
    if ((e.target as HTMLElement | null)?.closest?.('.sl-picker, .sl-color, .sl-popup')) return;
    e.preventDefault();
    emit('close');
  }
};
onMounted(() => {
  position();
  window.addEventListener('keydown', key);
  const first = el.value?.querySelector<HTMLElement>('[data-autofocus], input:not([type=hidden]):not([disabled]), select, textarea');
  if (first && window.matchMedia?.('(pointer: fine)').matches) first.focus({ preventScroll: true });
});
onUnmounted(() => window.removeEventListener('keydown', key));

const startDrag = (e: PointerEvent) => {
  const box = el.value;
  if (!box || props.panel || (e.target as HTMLElement).closest('button')) return;
  const rect = box.getBoundingClientRect();
  drag = { dx: e.clientX - rect.left, dy: e.clientY - rect.top, id: e.pointerId };
  (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
};
const moveDrag = (e: PointerEvent) => {
  const d = drag;
  const box = el.value;
  if (!d || !box || d.id !== e.pointerId) return;
  const x = Math.max(-box.offsetWidth + 80, Math.min(window.innerWidth - 80, e.clientX - d.dx));
  const y = Math.max(0, Math.min(window.innerHeight - 40, e.clientY - d.dy));
  box.style.left = `${x}px`;
  box.style.top = `${y}px`;
};
const endDrag = () => {
  drag = null;
};
const px = (v: number | string | undefined) => (typeof v === 'number' ? `${v}px` : v);
const boxStyle = () => ({ width: px(props.panel ? Math.min(props.width, window.innerWidth) : props.width), height: px(props.height) });
</script>

<template>
  <Teleport :to="layer">
    <div :class="['sl-modal-layer', { blocking }]" @pointerdown="blocking && $event.target === $event.currentTarget && emit('close')">
      <div
        ref="el"
        :class="['sl-modal', { 'panel-right': panel }]"
        role="dialog"
        :aria-modal="blocking || undefined"
        :aria-label="title"
        :data-testid="testId"
        :style="boxStyle()"
      >
        <div class="sl-modal-title" @pointerdown="startDrag" @pointermove="moveDrag" @pointerup="endDrag">
          <Icon v-if="icon" :name="icon" />
          <span class="text">{{ title }}</span>
          <button type="button" class="sl-modal-close" aria-label="Close" data-testid="modal-close" @click="emit('close')">
            <Icon name="close" />
          </button>
        </div>
        <div :class="['sl-modal-body', bodyClass]" :style="bodyStyle"><slot /></div>
        <div v-if="$slots.footer" class="sl-modal-foot"><slot name="footer" /></div>
      </div>
    </div>
  </Teleport>
</template>
