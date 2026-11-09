<script setup lang="ts">
/**
 * Format cells: the Mask tab (General, Number, Currency, Accounting, Percentage, Text, Custom with
 * a live sample) and the Editor tab (how cells are edited: text, numbers, percentages, a drop-down
 * list, a checkbox or a date, and text wrapping).
 */
import {
  applyCellEditor, CUSTOM_FORMATS, EDITOR_TYPES, type EditorType, FORMAT_CATEGORIES, formatValue, maskCode, maskFromCode,
  type MaskState,
} from '@sumline/core/views';
import { computed, ref, shallowRef } from 'vue';
import { useUI } from '../context';
import Modal from '../Modal.vue';
import Switch from '../Switch.vue';

const props = withDefaults(defineProps<{ initialTab?: 'mask' | 'editor' }>(), { initialTab: 'mask' });
const emit = defineEmits<{ close: [] }>();

const NEGATIVES = ['-1234', '1234', '(1234)', '(1234)'];
const ui = useUI();
const { doc } = ui.value;
const tab = ref<'mask' | 'editor'>(props.initialTab);
const mask = shallowRef<MaskState>(maskFromCode(ui.value.state.format.numFmt));
const editor = ref<EditorType>('text');
const wrap = ref(ui.value.state.format.wrap === 'wrap');
const options = ref('');
const code = computed(() => maskCode(mask.value));
// The active cell's value shown in the format (worked out again on every change).
const sample = computed(() => {
  const d = ui.value.doc;
  const { r, c } = d.selection.active;
  const value = d.cell(r, c)?.v;
  const sampleValue = typeof value === 'number' ? value : 1234;
  try {
    const f = formatValue(mask.value.category === 'text' && typeof value === 'string' ? value : sampleValue, code.value, d.wb.date1904);
    return { text: f.text, color: f.color };
  } catch {
    return { text: String(sampleValue), color: undefined };
  }
});
const set = (patch: Partial<MaskState>) => (mask.value = { ...mask.value, ...patch });

const apply = () => {
  const { grid } = ui.value;
  if (grid?.isEditing()) grid.commitEdit();
  doc.transact('Format cells', () => {
    if (tab.value === 'mask') doc.setNumberFormat(code.value);
    else applyCellEditor(doc, editor.value, options.value, wrap.value);
  });
  emit('close');
};
const categoryKey = (e: KeyboardEvent) => {
  const i = FORMAT_CATEGORIES.findIndex((x) => x.key === mask.value.category);
  if (e.key === 'ArrowDown') set({ category: FORMAT_CATEGORIES[Math.min(FORMAT_CATEGORIES.length - 1, i + 1)].key });
  else if (e.key === 'ArrowUp') set({ category: FORMAT_CATEGORIES[Math.max(0, i - 1)].key });
};
const typed = (e: Event) => (e.target as HTMLInputElement).value;
</script>

<template>
  <Modal title="Format" icon="settings" :width="500" test-id="format-dialog" @close="emit('close')">
    <div class="sl-subtabs" role="tablist">
      <button type="button" role="tab" :aria-selected="tab === 'mask'" :class="tab === 'mask' ? 'selected' : ''" @click="tab = 'mask'">Mask</button>
      <button type="button" role="tab" :aria-selected="tab === 'editor'" :class="tab === 'editor' ? 'selected' : ''" data-testid="format-editor-tab" @click="tab = 'editor'">Editor</button>
    </div>
    <div v-if="tab === 'mask'" class="sl-format">
      <div class="sl-list" role="listbox" aria-label="Category" data-testid="format-categories">
        <div
          v-for="cat in FORMAT_CATEGORIES"
          :key="cat.key"
          role="option"
          :aria-selected="mask.category === cat.key"
          tabindex="0"
          :class="`sl-list-item${mask.category === cat.key ? ' selected' : ''}`"
          @click="set({ category: cat.key, custom: cat.key === 'custom' ? code : mask.custom })"
          @keydown="categoryKey"
        >
          {{ cat.label }}
        </div>
      </div>
      <div class="sl-form" :style="{ alignContent: 'start' }">
        <div class="sl-field">
          <label>Sample</label>
          <input class="sl-input" readonly :value="sample.text" data-testid="format-sample" :style="{ color: sample.color }" />
        </div>
        <div v-if="['number', 'currency', 'accounting', 'percentage'].includes(mask.category)" class="sl-field">
          <label for="sl-dec">Decimal Place:</label>
          <input
            id="sl-dec"
            class="sl-input"
            type="number"
            min="0"
            max="30"
            inputmode="numeric"
            :value="mask.decimals"
            data-testid="format-decimals"
            @input="set({ decimals: typed($event) })"
          />
        </div>
        <Switch v-if="mask.category === 'number'" label="Use thousands separator" :checked="mask.thousands" test-id="format-thousands" @change="set({ thousands: $event })" />
        <div v-if="mask.category === 'currency' || mask.category === 'accounting'" class="sl-field">
          <label for="sl-sym">Symbol:</label>
          <input id="sl-sym" class="sl-input" :value="mask.symbol" data-testid="format-symbol" @input="set({ symbol: typed($event) })" />
        </div>
        <div v-if="mask.category === 'number' || mask.category === 'currency'" class="sl-field">
          <label>Negative Numbers:</label>
          <div class="sl-list" role="listbox" aria-label="Negative numbers" :style="{ maxHeight: '110px' }">
            <div
              v-for="(t, i) in NEGATIVES"
              :key="i"
              role="option"
              :aria-selected="mask.negative === i"
              :class="`sl-list-item${mask.negative === i ? ' selected' : ''}`"
              :style="{ color: mask.negative === i ? undefined : i % 2 ? '#d32f2f' : undefined }"
              @click="set({ negative: i })"
            >
              {{ t }}
            </div>
          </div>
        </div>
        <div v-if="mask.category === 'custom'" class="sl-field">
          <label for="sl-type">Type:</label>
          <input id="sl-type" class="sl-input" :value="mask.custom" spellcheck="false" data-testid="format-custom" @input="set({ custom: typed($event) })" />
          <div class="sl-list codes" role="listbox" aria-label="Format codes">
            <div
              v-for="f in CUSTOM_FORMATS"
              :key="f"
              role="option"
              :aria-selected="mask.custom === f"
              :class="`sl-list-item${mask.custom === f ? ' selected' : ''}`"
              @click="set({ custom: f })"
            >
              {{ f }}
            </div>
          </div>
        </div>
      </div>
    </div>
    <div v-else class="sl-form">
      <div class="sl-field">
        <label for="sl-coltype">Column type</label>
        <select id="sl-coltype" class="sl-select" :value="editor" data-testid="format-editor-type" @change="editor = ($event.target as HTMLSelectElement).value as EditorType">
          <option v-for="t in EDITOR_TYPES" :key="t.key" :value="t.key">{{ t.label }}</option>
        </select>
      </div>
      <div v-if="editor === 'dropdown'" class="sl-field">
        <label for="sl-options">Options (one per line)</label>
        <textarea id="sl-options" class="sl-textarea" :value="options" data-testid="format-options" @input="options = ($event.target as HTMLTextAreaElement).value" />
      </div>
      <Switch label="Wrap content" :checked="wrap" test-id="format-wrap" @change="wrap = $event" />
      <p class="sl-hint">Drop-down, checkbox and calendar cells are stored as data validation rules, so Excel keeps them.</p>
    </div>
    <template #footer>
      <button type="button" class="sl-button" data-testid="format-ok" @click="apply">Ok</button>
      <button type="button" class="sl-button" @click="emit('close')">Cancel</button>
    </template>
  </Modal>
</template>
