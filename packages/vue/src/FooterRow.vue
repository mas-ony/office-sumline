<script setup lang="ts">
/**
 * The footer row (View > Toggle footer row): a band under the grid aligned with the columns,
 * showing a label or the result of a formula per column (e.g. =SUM(B:B)).
 */
import { ArrayValue, CellError, formatValue, RefValue, type Scalar, type ViewTransform } from '@sumline/core/views';
import { computed, onUnmounted, shallowRef, watch } from 'vue';
import { useUI } from './context';

const ui = useUI();
const vt = shallowRef<ViewTransform | null>(ui.value.grid?.transform() ?? null);
let off = () => {};
watch(
  () => ui.value.grid,
  (grid) => {
    off();
    off = ui.value.ui.onGridPaint((v) => (vt.value = v));
    grid?.requestRender();
  },
  { immediate: true },
);
onUnmounted(() => off());

const value = (text: string, c: number): string => {
  if (!text.startsWith('=')) return text;
  const { doc } = ui.value;
  try {
    let v = doc.engine.evaluateText(text.slice(1), doc.sheet, 1048575, c) as unknown;
    if (v instanceof RefValue) v = doc.engine.cellValue(v.single.sheet, v.single.range.r0, v.single.range.c0);
    if (v instanceof ArrayValue) v = v.get(0, 0);
    if (v instanceof CellError) return v.code;
    const numFmt = doc.styleAt(Math.max(0, (doc.sheet.usedRange()?.r1 ?? 0)), c).numFmt;
    return formatValue(v as Scalar, typeof v === 'number' ? numFmt : 'General', doc.wb.date1904).text;
  } catch {
    return '#VALUE!';
  }
};

// The cells of the footer, worked out again whenever the interface or the grid's view changes.
const cells = computed(() => {
  const footer = ui.value.doc.sheet.footer;
  const t = vt.value;
  if (!footer || !t) return null;
  const { cols } = t.visibleMain();
  const frozen = Array.from({ length: t.fc }, (_, i) => i);
  const columns = [...frozen, ...Array.from({ length: Math.max(0, cols[1] - cols[0] + 1) }, (_, i) => cols[0] + i)];
  const out: { c: number; text: string; label: boolean; x: number; w: number }[] = [];
  for (const c of columns) {
    const rect = t.cellRect(0, c);
    if (rect.w <= 0 || c >= footer.length) continue;
    const text = footer[c] ?? '';
    if (!text) continue;
    out.push({ c, text, label: !text.startsWith('='), x: rect.x, w: rect.w });
  }
  return { headerW: t.v.headerW, cells: out };
});
</script>

<template>
  <div v-if="cells" class="sl-footer-row" data-testid="footer-row" aria-label="Footer row">
    <div class="sl-footer-cell label" :style="{ left: '0px', width: `${cells.headerW}px` }">Σ</div>
    <div v-for="f in cells.cells" :key="f.c" :class="`sl-footer-cell${f.label ? ' label' : ''}`" :style="{ left: `${f.x}px`, width: `${f.w}px` }" :title="f.text">{{ value(f.text, f.c) }}</div>
  </div>
</template>
