<script setup lang="ts">
/** One item of the toolbar (in the bar or in the "more" panel). */
import { ALIGN_ICONS, cmd, TOOLBAR_FONTS, TOOLBAR_SIZES, type ToolKey, VALIGN_ICONS } from '@sumline/core/views';
import { computed } from 'vue';
import BorderPicker from './BorderPicker.vue';
import ColorTool from './ColorTool.vue';
import { useUI } from './context';
import Icon from './Icon.vue';
import PickerTool from './PickerTool.vue';
import ToolButton from './ToolButton.vue';

defineProps<{ tool: ToolKey }>();

const ALIGNS = ['left', 'center', 'right', 'justify'] as const;
const VALIGNS = ['top', 'center', 'bottom'] as const;

const ui = useUI();
const f = computed(() => ui.value.state.format);
const fonts = computed(() => {
  const set = new Set<string>(TOOLBAR_FONTS);
  if (!f.value.defaultFont) set.add(f.value.fontName);
  return [...set];
});
const align = computed(() => (f.value.hAlign === 'general' ? 'left' : f.value.hAlign));
const fontSelected = (name: string) => (name === 'Default' ? f.value.defaultFont : !f.value.defaultFont && name === f.value.fontName);
const sizeSelected = (pt: number) => Math.abs(f.value.fontSize - pt) < 0.01;

const setFont = (name: string, close: () => void) => {
  close();
  const def = ui.value.doc.wb.styles.defaultStyle.font.name;
  cmd.format(ui.value, { font: { name: name === 'Default' ? def : name } }, 'Font');
};
const setSize = (pt: number, close: () => void) => {
  close();
  cmd.format(ui.value, { font: { size: pt } }, 'Font size');
};
const setAlign = (a: (typeof ALIGNS)[number], close: () => void) => {
  close();
  cmd.format(ui.value, { align: { horizontal: a } }, 'Align');
};
const setVAlign = (a: (typeof VALIGNS)[number], close: () => void) => {
  close();
  cmd.format(ui.value, { align: { vertical: a } }, 'Vertical align');
};
const toggleSearch = () => (ui.value.dialog?.kind === 'search' ? ui.value.closeDialog() : ui.value.openDialog({ kind: 'search', replace: true }));
</script>

<template>
  <ToolButton
    v-if="tool === 'undo'"
    icon="undo"
    label="Undo"
    test-id="tb-undo"
    :disabled="ui.state.readOnly || (!ui.state.canUndo && !ui.grid?.isEditing())"
    @click="cmd.undo(ui)"
  />
  <ToolButton v-else-if="tool === 'redo'" icon="redo" label="Redo" test-id="tb-redo" :disabled="ui.state.readOnly || !ui.state.canRedo" @click="cmd.redo(ui)" />
  <ToolButton v-else-if="tool === 'save'" icon="save" label="Save" test-id="tb-save" @click="ui.files.save()" />
  <span v-else-if="tool === 'd1' || tool === 'd2'" class="sl-divider" />
  <PickerTool v-else-if="tool === 'font'" label="Font Family" test-id="tb-font" extra-class="sl-tool-font">
    <span class="label">{{ f.defaultFont ? 'Default' : f.fontName }}</span>
    <Icon name="arrow_drop_down" class="caret" />
    <template #content="{ close }">
      <button
        v-for="name in fonts"
        :key="name"
        type="button"
        role="option"
        :aria-selected="fontSelected(name)"
        :class="`sl-picker-item${fontSelected(name) ? ' selected' : ''}`"
        :style="{ fontFamily: name === 'Default' ? undefined : name }"
        @click="setFont(name, close)"
      >
        {{ name }}
      </button>
    </template>
  </PickerTool>
  <PickerTool v-else-if="tool === 'size'" label="Font Size" test-id="tb-size">
    <Icon name="format_size" />
    <Icon name="arrow_drop_down" class="caret" />
    <template #content="{ close }">
      <button
        v-for="s in TOOLBAR_SIZES"
        :key="s.label"
        type="button"
        role="option"
        :aria-selected="sizeSelected(s.pt)"
        :class="`sl-picker-item${sizeSelected(s.pt) ? ' selected' : ''}`"
        :style="{ fontSize: s.css }"
        @click="setSize(s.pt, close)"
      >
        {{ s.label }}
      </button>
    </template>
  </PickerTool>
  <PickerTool v-else-if="tool === 'align'" label="Text Align" test-id="tb-align">
    <Icon :name="ALIGN_ICONS[align] ?? 'format_align_left'" />
    <Icon name="arrow_drop_down" class="caret" />
    <template #content="{ close }">
      <div class="sl-picker-icons cols-1">
        <button
          v-for="a in ALIGNS"
          :key="a"
          type="button"
          role="option"
          :aria-label="`Align ${a}`"
          :data-testid="`align-${a}`"
          :class="`sl-icon-btn${align === a ? ' selected' : ''}`"
          @click="setAlign(a, close)"
        >
          <Icon :name="ALIGN_ICONS[a]" />
        </button>
      </div>
    </template>
  </PickerTool>
  <ToolButton v-else-if="tool === 'bold'" icon="format_bold" label="Bold" test-id="tb-bold" :active="f.bold" :disabled="ui.state.readOnly" @click="cmd.toggleFont(ui, 'bold')" />
  <ColorTool v-else-if="tool === 'color'" kind="text" />
  <ColorTool v-else-if="tool === 'fill'" kind="fill" />
  <PickerTool v-else-if="tool === 'valign'" label="Vertical Align" test-id="tb-valign">
    <Icon :name="VALIGN_ICONS[f.vAlign] ?? 'vertical_align_bottom'" />
    <Icon name="arrow_drop_down" class="caret" />
    <template #content="{ close }">
      <div class="sl-picker-icons cols-1">
        <button
          v-for="a in VALIGNS"
          :key="a"
          type="button"
          role="option"
          :aria-label="`Align ${a}`"
          :data-testid="`valign-${a}`"
          :class="`sl-icon-btn${f.vAlign === a ? ' selected' : ''}`"
          @click="setVAlign(a, close)"
        >
          <Icon :name="VALIGN_ICONS[a]" />
        </button>
      </div>
    </template>
  </PickerTool>
  <ToolButton
    v-else-if="tool === 'merge'"
    icon="web"
    label="Merge the selected cells"
    test-id="tb-merge"
    :active="f.merged"
    :disabled="ui.state.readOnly"
    @click="void cmd.merge(ui)"
  />
  <PickerTool v-else-if="tool === 'border'" label="Border Style" test-id="tb-border">
    <Icon name="border_all" />
    <Icon name="arrow_drop_down" class="caret" />
    <template #content="{ close }">
      <BorderPicker @close="close" />
    </template>
  </PickerTool>
  <ToolButton v-else-if="tool === 'image'" icon="add_photo_alternate" label="Upload image" test-id="tb-image" :disabled="ui.state.readOnly" @click="cmd.uploadImage(ui)" />
  <ToolButton v-else-if="tool === 'search'" icon="search" label="Toggle Search" test-id="tb-search" :active="ui.dialog?.kind === 'search'" @click="toggleSearch" />
  <ToolButton v-else-if="tool === 'chart'" icon="addchart" label="Insert chart" test-id="tb-chart" :disabled="ui.state.readOnly" @click="ui.openDialog({ kind: 'chart-insert' })" />
  <ToolButton v-else-if="tool === 'rule'" icon="rule" label="Add new validation" test-id="tb-validation" :disabled="ui.state.readOnly" @click="ui.openDialog({ kind: 'validations' })" />
  <ToolButton v-else-if="tool === 'shape'" icon="interests" label="Insert shape" test-id="tb-shape" :disabled="ui.state.readOnly" @click="ui.openDialog({ kind: 'shapes' })" />
</template>
