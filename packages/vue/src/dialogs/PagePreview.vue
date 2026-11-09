<script setup lang="ts">
/** A page of the print preview, drawn once it comes near the visible part of the preview. */
import { type Page, pageSvg } from '@sumline/core/views';
import { computed, onMounted, onUnmounted, ref, watch } from 'vue';

const props = defineProps<{ page: Page; index: number; width: number; root: HTMLElement | null }>();

const el = ref<HTMLDivElement | null>(null);
const visible = ref(props.index < 2);
let io: IntersectionObserver | null = null;
const observe = () => {
  io?.disconnect();
  io = null;
  const box = el.value;
  if (!box || visible.value || typeof IntersectionObserver === 'undefined') {
    if (!visible.value) visible.value = true;
    return;
  }
  io = new IntersectionObserver(
    (entries) => {
      if (entries.some((e) => e.isIntersecting)) {
        visible.value = true;
        io?.disconnect();
      }
    },
    { root: props.root, rootMargin: '600px 0px' },
  );
  io.observe(box);
};
onMounted(observe);
watch([visible, () => props.root], observe, { flush: 'post' });
onUnmounted(() => io?.disconnect());

const height = computed(() => (props.width * props.page.height) / props.page.width);
// Generated from the document by the print layout (escaped text).
const markup = computed(() => (visible.value ? pageSvg(props.page, { width: '100%', height: '100%' }) : ''));
</script>

<template>
  <div ref="el" class="sl-print-page" :data-page="index + 1" data-testid="print-page" :style="{ width: `${width}px`, height: `${height}px` }" v-html="markup" />
</template>
