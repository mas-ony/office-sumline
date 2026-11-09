<script setup lang="ts">
/**
 * A menu shown until an item is chosen or the user clicks elsewhere (top menu, context menus). On
 * phones the same items are shown in a bottom sheet.
 */
import type { MenuAnchor, MenuItem } from '@sumline/core/views';
import { onUnmounted, watch } from 'vue';
import { useUI } from './context';
import MenuList from './MenuList.vue';
import SheetMenu from './SheetMenu.vue';

const props = defineProps<{
  items: MenuItem[];
  /** Position (client coordinates) of the top-left corner, or the rectangle to open next to. */
  at: MenuAnchor;
  /** Bottom sheet presentation (phones). */
  sheet?: boolean;
  title?: string;
  label?: string;
  testId?: string;
}>();
const emit = defineEmits<{ close: [] }>();

const { layer } = useUI().value;

let stop = () => {};
const down = (e: PointerEvent) => {
  const t = e.target as HTMLElement | null;
  if (t?.closest('.sl-popup')) return;
  emit('close');
};
const blur = () => emit('close');
watch(
  () => props.sheet,
  (sheet) => {
    stop();
    stop = () => {};
    if (sheet) return;
    // Opening clicks must not close the menu immediately.
    const id = window.setTimeout(() => {
      window.addEventListener('pointerdown', down, true);
      window.addEventListener('resize', blur);
      window.addEventListener('blur', blur);
    }, 0);
    stop = () => {
      window.clearTimeout(id);
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('resize', blur);
      window.removeEventListener('blur', blur);
    };
  },
  { immediate: true },
);
onUnmounted(() => stop());
</script>

<template>
  <Teleport :to="layer">
    <SheetMenu v-if="sheet" :items="items" :title="title ?? label" :test-id="testId" @close="emit('close')" />
    <MenuList v-else :items="items" :at="at" side="below" :depth="0" :label="label ?? title" :test-id="testId" @close="emit('close')" />
  </Teleport>
</template>
