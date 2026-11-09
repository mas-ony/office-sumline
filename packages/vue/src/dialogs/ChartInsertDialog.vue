<script setup lang="ts">
/** The Insert chart dialog: the data range and a gallery of chart types. */
import { CHART_TYPE_INFO, type ChartType, insertChart, rangeName, trimRange } from '@sumline/core/views';
import { ref } from 'vue';
import { useUI } from '../context';
import Icon from '../Icon.vue';
import Modal from '../Modal.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const selectionRange = () => rangeName(trimRange(ui.value.doc.sheet, doc.range));
const range = ref(selectionRange());
const type = ref<ChartType>('line');
const error = ref('');

const insert = () => {
  const { grid, phone, selectObject } = ui.value;
  const res = insertChart(doc, range.value, type.value, phone);
  if ('error' in res) {
    error.value = res.error;
    return;
  }
  emit('close');
  if (res.chart) {
    grid?.scrollIntoView(res.chart.from.r, res.chart.from.c);
    selectObject(res.chart.id);
  }
};
const typed = (e: Event) => {
  range.value = (e.target as HTMLInputElement).value;
  error.value = '';
};
</script>

<template>
  <Modal title="Insert chart" icon="addchart" :width="430" test-id="chart-insert-dialog" @close="emit('close')">
    <div class="sl-form">
      <div class="sl-field">
        <label for="chart-range">Data to analyze</label>
        <div class="sl-range-input">
          <input id="chart-range" class="sl-input" :value="range" spellcheck="false" data-testid="chart-range" @input="typed" />
          <button type="button" title="Use the selection" aria-label="Use the selection" @click="range = selectionRange()">
            <Icon name="highlight_alt" :size="18" />
          </button>
        </div>
      </div>
    </div>
    <div class="sl-chart-types" role="listbox" aria-label="Chart type">
      <button
        v-for="t in CHART_TYPE_INFO"
        :key="t.type"
        type="button"
        role="option"
        :aria-selected="type === t.type"
        :class="`sl-chart-type${type === t.type ? ' selected' : ''}`"
        :data-testid="`chart-type-${t.type}`"
        @click="type = t.type"
        @dblclick="insert"
      >
        <span class="tile"><Icon :name="t.icon" /></span>
        <span>{{ t.label }}</span>
      </button>
    </div>
    <div v-if="error" class="sl-error" role="alert">{{ error }}</div>
    <div class="sl-stack-buttons">
      <button type="button" class="sl-button wide" data-testid="chart-insert" @click="insert">Insert</button>
      <button type="button" class="sl-button wide" @click="emit('close')">Cancel</button>
    </div>
  </Modal>
</template>
