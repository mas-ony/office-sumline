<script setup lang="ts">
/**
 * The Insert Shape gallery: lines, rectangles, basic forms, block arrows, stars, callouts,
 * equation shapes and flowchart.
 */
import { DEFAULT_SHAPE_FILL, DEFAULT_SHAPE_LINE, type GalleryItem, LINE_PRESETS, SHAPE_GALLERY, shapeIconSvg, SheetLayout } from '@sumline/core/views';
import { useUI } from '../context';
import Modal from '../Modal.vue';

const emit = defineEmits<{ close: [] }>();

const SHAPE_W = 400;
const SHAPE_H = 300;

const ui = useUI();
const insert = (it: GalleryItem) => {
  const { doc, grid, phone, selectObject } = ui.value;
  const sheet = doc.sheet;
  const layout = SheetLayout.of(sheet);
  const { r, c } = doc.selection.active;
  const line = LINE_PRESETS.has(it.preset);
  const w = phone ? Math.min(SHAPE_W, Math.round(window.innerWidth * 0.7)) : SHAPE_W;
  const h = phone ? Math.round((w * 3) / 4) : SHAPE_H;
  const shape = doc.addShape(
    {
      name: `${it.label.replace(/\b\w/g, (m) => m.toUpperCase())} ${sheet.shapes.length + 1}`,
      preset: it.preset,
      fill: line ? null : DEFAULT_SHAPE_FILL,
      line: { color: DEFAULT_SHAPE_LINE, width: line ? 2 : 3 },
      fontColor: '#000000',
      arrowStart: it.arrowStart || undefined,
      arrowEnd: it.arrowEnd || undefined,
    },
    { x: layout.colX(c), y: layout.rowY(r), width: w, height: line ? Math.round(h / 2) : h },
  );
  emit('close');
  if (shape) {
    grid?.scrollIntoView(shape.from.r, shape.from.c);
    selectObject(shape.id);
  }
};
// The buttons' icons (v-html) are generated from the shape geometry, not user content.
const testId = (it: GalleryItem) => `shape-${it.preset}${it.arrowEnd ? (it.arrowStart ? '-both' : '-end') : ''}`;
</script>

<template>
  <Modal title="Insert Shape" icon="interests" :width="400" test-id="shapes-dialog" body-class="sl-shapes-body" @close="emit('close')">
    <div v-for="group in SHAPE_GALLERY" :key="group.title" class="sl-shape-group">
      <h4>{{ group.title }}</h4>
      <div class="sl-shape-grid">
        <button v-for="(it, i) in group.items" :key="`${it.preset}-${i}`" type="button" :title="it.label" :aria-label="it.label" :data-testid="testId(it)" @click="insert(it)" v-html="shapeIconSvg(it)" />
      </div>
    </div>
  </Modal>
</template>
