<script setup lang="ts">
/**
 * The toolbar: undo, redo, save | font, size, alignment, bold, colours, vertical alignment, merge,
 * borders | image, search, chart, validation, shape. Items that do not fit move to a "more" panel.
 */
import { fitToolbar } from '@sumline/core/views';
import { computed, onMounted, onUnmounted, ref, watch } from 'vue';
import { useUI } from './context';
import ToolbarItem from './ToolbarItem.vue';
import ToolButton from './ToolButton.vue';

const ui = useUI();
const bar = ref<HTMLDivElement | null>(null);
const avail = ref(10000);
const more = ref(false);
// Tools that open a dialog close the overflow panel.
watch(
  () => ui.value.dialog,
  (d) => {
    if (d) more.value = false;
  },
);
let observer: ResizeObserver | null = null;
onMounted(() => {
  const el = bar.value!;
  const measure = () => (avail.value = el.clientWidth - 16);
  measure();
  observer = new ResizeObserver(measure);
  observer.observe(el);
});
onUnmounted(() => observer?.disconnect());
const fit = computed(() => fitToolbar(avail.value));
watch(
  () => fit.value.overflow.length,
  (n) => {
    if (!n) more.value = false;
  },
);
</script>

<template>
  <div ref="bar" class="sl-toolbar" role="toolbar" aria-label="Spreadsheet Tools" data-testid="toolbar">
    <div class="sl-toolbar-items">
      <span v-for="k in fit.shown" :key="k" :style="{ display: 'contents' }"><ToolbarItem :tool="k" /></span>
    </div>
    <template v-if="fit.overflow.length > 0">
      <span class="sl-divider" />
      <ToolButton icon="more_vert" label="More tools" test-id="tb-more" :active="more" :expanded="more" @click="more = !more" />
      <div v-if="more" class="sl-overflow-panel" data-testid="tb-overflow" @mousedown.prevent>
        <span v-for="k in fit.overflow" :key="k" :style="{ display: 'contents' }"><ToolbarItem :tool="k" /></span>
      </div>
    </template>
  </div>
</template>
