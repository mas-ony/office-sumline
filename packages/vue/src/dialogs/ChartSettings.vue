<script setup lang="ts">
/** The chart settings panel: Configuration, Style and Series tabs. */
import {
  CHART_SIDES, CHART_TYPE_INFO, type ChartConfig, chartLines, chartSettingsOf, chartSettingsPatch, type ChartStyle, type ChartType, LINE_DASHES,
  type LineDash, localRef, numText, parseNum, rangeName, readChartData, relayoutChart, resolveRef, type SeriesStyle, type Side, type Trendline,
  TRENDLINES, trimRange, withSeriesStyle,
} from '@sumline/core/views';
import { computed, ref, shallowRef } from 'vue';
import ColorField from '../ColorField.vue';
import { useUI } from '../context';
import { vValue } from '../directives';
import Icon from '../Icon.vue';
import Modal from '../Modal.vue';
import Switch from '../Switch.vue';

const props = defineProps<{ id: string }>();
const emit = defineEmits<{ close: [] }>();

const TABS = ['config', 'style', 'series'] as const;
const ui = useUI();
const { doc } = ui.value;
const sheet = computed(() => ui.value.doc.sheet);
const chart = computed(() => sheet.value.charts.find((c) => c.id === props.id));
const tab = ref<(typeof TABS)[number]>('config');
const initial = chartSettingsOf(doc, chart.value, sheet.value);
const config = shallowRef<ChartConfig>(initial.config);
const custom = ref(initial.custom);
const style = shallowRef<ChartStyle>(initial.style);
const seriesStyles = shallowRef<SeriesStyle[]>(initial.seriesStyles);

const target = computed(() => resolveRef(doc, config.value.range, sheet.value));
const lines = computed(() => (target.value ? chartLines(doc, target.value.sheet, target.value.range, config.value.seriesIn, config.value.headers) : []));
const data = computed(() => (chart.value ? readChartData(doc, chart.value) : null));

/** Re-derives labels and series when the range, orientation or header option changes. */
const relayout = (patch: Partial<ChartConfig>) => {
  const next = relayoutChart(doc, sheet.value, config.value, patch);
  if (next.laidOut) custom.value = false;
  config.value = next.config;
};

const seriesNames = computed<string[]>(() =>
  custom.value ? (data.value?.series.map((s) => s.name) ?? []) : config.value.series.map((i) => lines.value[i]?.name ?? `Series ${i + 1}`),
);

const save = () => {
  const c = chart.value;
  if (!c) return;
  doc.updateChart(props.id, chartSettingsPatch(c, { config: config.value, custom: custom.value, style: style.value, seriesStyles: seriesStyles.value }, target.value));
  emit('close');
};

const setConfig = (patch: Partial<ChartConfig>) => (config.value = { ...config.value, ...patch });
const setStyle = (patch: Partial<ChartStyle>) => (style.value = { ...style.value, ...patch });
const setSeriesStyle = (i: number, p: Partial<SeriesStyle>) => (seriesStyles.value = withSeriesStyle(seriesStyles.value, i, p));
const styleOf = (i: number): SeriesStyle => seriesStyles.value[i] ?? {};
const value = (e: Event) => (e.target as HTMLInputElement | HTMLSelectElement).value;

const setSeries = (k: number, e: Event) => setConfig({ series: config.value.series.map((x, j) => (j === k ? Number(value(e)) : x)) });
const removeSeries = (k: number) => {
  setConfig({ series: config.value.series.filter((_, j) => j !== k) });
  seriesStyles.value = seriesStyles.value.filter((_, j) => j !== k);
};
const addSeries = () => {
  const c = config.value;
  const unused = lines.value.find((l) => l.index !== c.labels && !c.series.includes(l.index)) ?? lines.value[0];
  if (unused) setConfig({ series: [...c.series, unused.index] });
};
const panelWidth = () => (ui.value.phone ? window.innerWidth : 440);
</script>

<template>
  <Modal v-if="chart" title="Chart settings" icon="bar_chart" panel :width="panelWidth()" test-id="chart-settings" @close="emit('close')">
    <div class="sl-subtabs" role="tablist">
      <button v-for="t in TABS" :key="t" type="button" role="tab" :aria-selected="tab === t" :class="tab === t ? 'selected' : ''" :data-testid="`chart-tab-${t}`" @click="tab = t">
        {{ t === 'config' ? 'Configuration' : t === 'style' ? 'Style' : 'Series' }}
      </button>
    </div>
    <div v-if="tab === 'config'" class="sl-form">
      <div class="sl-row">
        <div class="sl-field">
          <label for="chart-type">Chart type</label>
          <select id="chart-type" class="sl-select" :value="config.type" data-testid="chart-type" @change="setConfig({ type: value($event) as ChartType })">
            <option v-for="t in CHART_TYPE_INFO" :key="t.type" :value="t.type">{{ t.name }}</option>
          </select>
        </div>
        <div class="sl-field">
          <label for="chart-orientation">Data orientation</label>
          <select
            id="chart-orientation"
            class="sl-select"
            :value="config.seriesIn"
            data-testid="chart-orientation"
            :disabled="custom && !target"
            @change="relayout({ seriesIn: value($event) as 'rows' | 'columns' })"
          >
            <option value="rows">Horizontal</option>
            <option value="columns">Vertical</option>
          </select>
        </div>
      </div>
      <div class="sl-field">
        <label for="chart-data-range">Data range</label>
        <div class="sl-range-input">
          <input
            id="chart-data-range"
            :class="`sl-input${config.range && !target ? ' invalid' : ''}`"
            :value="config.range"
            spellcheck="false"
            data-testid="chart-data-range"
            @input="setConfig({ range: value($event) })"
            @blur="target && relayout({})"
          />
          <button type="button" title="Use the selection" aria-label="Use the selection" @click="relayout({ range: rangeName(trimRange(sheet, doc.range)) })">
            <Icon name="highlight_alt" :size="18" />
          </button>
        </div>
      </div>
      <template v-if="!custom && target">
        <div class="sl-field">
          <label for="chart-labels">Labels</label>
          <select
            id="chart-labels"
            class="sl-select"
            :value="config.labels ?? ''"
            data-testid="chart-labels"
            @change="setConfig({ labels: value($event) === '' ? null : Number(value($event)) })"
          >
            <option value="">None</option>
            <option v-for="l in lines" :key="l.index" :value="l.index">{{ l.name }}</option>
          </select>
        </div>
        <div class="sl-field">
          <div class="label">Series</div>
          <div v-for="(s, k) in config.series" :key="k" class="sl-series-row">
            <select class="sl-select" :value="s" :aria-label="`Series ${k + 1}`" data-testid="chart-series" @change="setSeries(k, $event)">
              <option v-for="l in lines" :key="l.index" :value="l.index">{{ l.name }}</option>
            </select>
            <button type="button" class="sl-icon-btn" aria-label="Remove series" data-testid="chart-series-remove" @click="removeSeries(k)">
              <Icon name="close" />
            </button>
          </div>
          <button type="button" class="sl-link" data-testid="chart-series-add" @click="addSeries">Add new serie</button>
        </div>
        <Switch
          :label="config.seriesIn === 'columns' ? 'Use the first row as headers' : 'Use the first column as headers'"
          :checked="config.headers"
          test-id="chart-headers"
          @change="setConfig({ headers: $event })"
        />
      </template>
      <div v-else class="sl-field">
        <div class="label">Series</div>
        <div v-for="s in data?.series ?? []" :key="s.index" class="sl-hint">{{ s.name }}: {{ localRef(s.source.values, sheet) }}</div>
      </div>
    </div>
    <div v-if="tab === 'style'" class="sl-form">
      <div class="sl-row">
        <div class="sl-field">
          <label for="chart-title">Chart title</label>
          <input id="chart-title" class="sl-input" :value="style.title ?? ''" data-testid="chart-title" @input="setStyle({ title: value($event) })" />
        </div>
        <div class="sl-field">
          <label for="chart-title-pos">Title position</label>
          <select id="chart-title-pos" class="sl-select" :value="style.titlePosition ?? ''" @change="setStyle({ titlePosition: (value($event) || undefined) as Side | undefined })">
            <option value="" />
            <option v-for="s in CHART_SIDES" :key="s.key" :value="s.key">{{ s.label }}</option>
          </select>
        </div>
      </div>
      <div class="sl-field">
        <label for="chart-subtitle">Subtitle</label>
        <input id="chart-subtitle" class="sl-input" :value="style.subtitle ?? ''" @input="setStyle({ subtitle: value($event) })" />
      </div>
      <div class="sl-row">
        <div class="sl-field">
          <label for="chart-legend-title">Legend title</label>
          <input id="chart-legend-title" class="sl-input" :value="style.legendTitle ?? ''" @input="setStyle({ legendTitle: value($event) })" />
        </div>
        <div class="sl-field">
          <label for="chart-legend-pos">Legend position</label>
          <select
            id="chart-legend-pos"
            class="sl-select"
            :value="style.legendPosition ?? ''"
            data-testid="chart-legend-position"
            @change="setStyle({ legendPosition: (value($event) || undefined) as Side | undefined })"
          >
            <option value="" />
            <option v-for="s in CHART_SIDES" :key="s.key" :value="s.key">{{ s.label }}</option>
          </select>
        </div>
      </div>
      <div class="sl-row">
        <Switch label="Show legend" :checked="!!style.showLegend" test-id="chart-show-legend" @change="setStyle({ showLegend: $event })" />
        <Switch label="Gridlines" :checked="!!style.gridlines" test-id="chart-gridlines" @change="setStyle({ gridlines: $event })" />
      </div>
      <Switch label="Value labels" :checked="!!style.valueLabels" test-id="chart-value-labels" @change="setStyle({ valueLabels: $event })" />
      <div class="sl-row">
        <div class="sl-field">
          <label for="chart-x-title">X axis title</label>
          <input id="chart-x-title" class="sl-input" :value="style.xAxisTitle ?? ''" @input="setStyle({ xAxisTitle: value($event) })" />
        </div>
        <div class="sl-field">
          <label for="chart-y-title">Y axis title</label>
          <input id="chart-y-title" class="sl-input" :value="style.yAxisTitle ?? ''" @input="setStyle({ yAxisTitle: value($event) })" />
        </div>
      </div>
      <div class="sl-field">
        <label for="chart-sec-title">Secondary axis title</label>
        <input id="chart-sec-title" class="sl-input" :value="style.secondaryAxisTitle ?? ''" @input="setStyle({ secondaryAxisTitle: value($event) })" />
      </div>
      <div class="sl-row">
        <div class="sl-field">
          <label for="chart-ymin">Y min</label>
          <input id="chart-ymin" class="sl-input" inputmode="decimal" :value="style.yMin" @input="setStyle({ yMin: value($event) })" />
        </div>
        <div class="sl-field">
          <label for="chart-ymax">Y max</label>
          <input id="chart-ymax" class="sl-input" inputmode="decimal" :value="style.yMax" @input="setStyle({ yMax: value($event) })" />
        </div>
      </div>
    </div>
    <div v-if="tab === 'series'" class="sl-form">
      <div v-for="(name, i) in seriesNames" :key="i" class="sl-series-style">
        <h4>{{ name }}</h4>
        <div class="sl-row">
          <div class="sl-field">
            <label>Color</label>
            <ColorField :value="styleOf(i).color ?? ''" :test-id="`chart-series-color-${i}`" @change="setSeriesStyle(i, { color: $event || undefined })" />
          </div>
          <div class="sl-field">
            <label :for="`sw-${i}`">Line width</label>
            <input :id="`sw-${i}`" v-value="numText(styleOf(i).lineWidth)" class="sl-input" inputmode="numeric" @input="setSeriesStyle(i, { lineWidth: parseNum(value($event)) })" />
          </div>
        </div>
        <div class="sl-row">
          <div class="sl-field">
            <label :for="`sd-${i}`">Line style</label>
            <select :id="`sd-${i}`" class="sl-select" :value="styleOf(i).lineStyle ?? 'solid'" @change="setSeriesStyle(i, { lineStyle: value($event) as LineDash })">
              <option v-for="d in LINE_DASHES" :key="d.key" :value="d.key">{{ d.label }}</option>
            </select>
          </div>
          <div class="sl-field">
            <label :for="`st-${i}`">Trendline</label>
            <select
              :id="`st-${i}`"
              class="sl-select"
              :value="styleOf(i).trendline ?? 'none'"
              :data-testid="`chart-series-trend-${i}`"
              @change="setSeriesStyle(i, { trendline: value($event) as Trendline })"
            >
              <option v-for="d in TRENDLINES" :key="d.key" :value="d.key">{{ d.label }}</option>
            </select>
          </div>
        </div>
        <div class="sl-row">
          <Switch label="Markers" :checked="!!styleOf(i).markers" @change="setSeriesStyle(i, { markers: $event })" />
          <Switch label="Secondary axis" :checked="!!styleOf(i).secondaryAxis" @change="setSeriesStyle(i, { secondaryAxis: $event })" />
        </div>
      </div>
    </div>
    <div class="sl-dialog-buttons">
      <button type="button" class="sl-button wide" data-testid="chart-save" @click="save">Save</button>
      <button type="button" class="sl-button wide" @click="emit('close')">Cancel</button>
    </div>
  </Modal>
</template>
