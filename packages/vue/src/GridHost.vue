<script setup lang="ts">
/** The grid's element; the grid of the workbook is created in it. */
import { onMounted, onUnmounted, ref } from 'vue';
import { useUI } from './context';

const { ui } = useUI().value;
const host = ref<HTMLDivElement | null>(null);
let attached: HTMLDivElement | null = null;
onMounted(() => {
  attached = host.value!;
  ui.attachGrid(attached);
});
onUnmounted(() => {
  if (attached) ui.detachGrid(attached);
});
</script>

<template>
  <div ref="host" class="sl-grid-host" data-testid="grid" />
</template>
