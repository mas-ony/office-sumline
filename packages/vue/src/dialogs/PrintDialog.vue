<script setup lang="ts">
/**
 * Print: a preview of the printed pages with the page options (scope, paper, orientation,
 * scaling, margins, page order, alignment, gridlines, notes, headings, header and footer,
 * repeated frozen rows and columns). The pages can be saved as a PDF or sent to the printer.
 */
import {
  buildPages, canvasMeasure, DEFAULT_PRINT_OPTIONS, loadPrintFonts, type Measure, type Page, pagesToPdf, PRINT_CHOICES, PRINT_SWITCHES,
  type PrintOptions, printPages, PT_TO_PX, rasterize, saveFile,
} from '@sumline/core/views';
import { computed, onMounted, onUnmounted, ref, shallowRef, watch } from 'vue';
import { usePersistent, useUI } from '../context';
import Icon from '../Icon.vue';
import Modal from '../Modal.vue';
import Switch from '../Switch.vue';
import PagePreview from './PagePreview.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const stored = usePersistent<Partial<PrintOptions>>('print-options', {});
const opts = computed<PrintOptions>(() => ({ ...DEFAULT_PRINT_OPTIONS, ...stored.value }));
const set = <K extends keyof PrintOptions>(key: K, value: PrintOptions[K]) => (stored.value = { ...stored.value, [key]: value });
const pages = shallowRef<Page[] | null>(null);
const busy = ref(false);
const zoom = ref<number | null>(null);
const current = ref(1);
const paneWidth = ref(600);
const showOptions = ref(!ui.value.phone);
let measure: Measure | null = null;
const pane = ref<HTMLDivElement | null>(null);
const title = doc.fileName.replace(/\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i, '');

watch(
  opts,
  (o, _, onCleanup) => {
    let alive = true;
    const t = setTimeout(async () => {
      await loadPrintFonts();
      if (!alive) return;
      measure ??= canvasMeasure();
      pages.value = buildPages(doc, o, measure, { title: doc.fileName });
    }, 30);
    onCleanup(() => {
      alive = false;
      clearTimeout(t);
    });
  },
  { immediate: true },
);

let observer: ResizeObserver | null = null;
onMounted(() => {
  const el = pane.value;
  if (!el) return;
  observer = new ResizeObserver(() => (paneWidth.value = el.clientWidth));
  observer.observe(el);
  paneWidth.value = el.clientWidth;
});
onUnmounted(() => observer?.disconnect());

const scale = computed(() => {
  const first = pages.value?.[0];
  const fit = first ? Math.max(0.1, (paneWidth.value - 36) / (first.width * PT_TO_PX)) : 1;
  return zoom.value ?? fit;
});
const pageWidth = (p: Page) => p.width * PT_TO_PX * scale.value;

const onScroll = () => {
  const el = pane.value;
  const list = pages.value;
  if (!el || !list) return;
  const mid = el.scrollTop + el.clientHeight / 3;
  let y = 18;
  for (let i = 0; i < list.length; i++) {
    const h = (pageWidth(list[i]) * list[i].height) / list[i].width + 16;
    if (mid < y + h) {
      current.value = i + 1;
      return;
    }
    y += h;
  }
  current.value = list.length;
};

const download = async () => {
  const list = pages.value;
  if (!list || busy.value) return;
  busy.value = true;
  try {
    const bytes = await pagesToPdf(list, { rasterize, measure: measure ?? canvasMeasure() }, { title });
    saveFile(bytes, `${title || 'spreadsheet'}.pdf`, 'application/pdf');
  } catch (e) {
    ui.value.notify(`The PDF could not be created: ${(e as Error).message}`);
  } finally {
    busy.value = false;
  }
};

const print = () => {
  if (pages.value) printPages(pages.value);
};

const zoomBy = (f: number) => (zoom.value = Math.max(0.25, Math.min(4, Math.round(scale.value * f * 100) / 100)));
const modalWidth = () => Math.min(1300, window.innerWidth - 16);
const choose = (key: keyof PrintOptions, e: Event) => set(key, (e.target as HTMLSelectElement).value as never);
</script>

<template>
  <Modal
    title="Print"
    icon="print"
    :width="modalWidth()"
    :height="ui.phone ? '100%' : 'min(820px, calc(100vh - 40px))'"
    test-id="print-dialog"
    body-class="sl-print-body"
    @close="emit('close')"
  >
    <div :class="`sl-print${showOptions ? '' : ' no-options'}`">
      <div class="sl-print-viewer">
        <div class="sl-print-toolbar">
          <button type="button" class="sl-print-tool" :aria-label="showOptions ? 'Hide options' : 'Show options'" @click="showOptions = !showOptions">
            <Icon name="menu" />
          </button>
          <span class="sl-print-name">{{ title || 'spreadsheet' }}.pdf</span>
          <span class="sl-print-count" data-testid="print-page-count"><span class="current">{{ pages ? current : 0 }}</span> / {{ pages?.length ?? 0 }}</span>
          <span class="sep" />
          <button type="button" class="sl-print-tool" aria-label="Zoom out" @click="zoomBy(1 / 1.2)">
            <Icon name="remove" />
          </button>
          <span class="sl-print-zoom">{{ Math.round(scale * 100) }}%</span>
          <button type="button" class="sl-print-tool" aria-label="Zoom in" @click="zoomBy(1.2)">
            <Icon name="add" />
          </button>
          <button type="button" class="sl-print-tool" aria-label="Fit to width" @click="zoom = null">
            <Icon name="fit_screen" />
          </button>
          <span class="grow" />
          <button type="button" class="sl-print-tool" aria-label="Download PDF" title="Download PDF" data-testid="print-download" :disabled="!pages || busy" @click="download">
            <Icon :name="busy ? 'hourglass_empty' : 'download'" />
          </button>
          <button type="button" class="sl-print-tool" aria-label="Print" title="Print" data-testid="print-print" :disabled="!pages" @click="print">
            <Icon name="print" />
          </button>
        </div>
        <div ref="pane" class="sl-print-preview" @scroll="onScroll">
          <div v-if="!pages" class="sl-print-wait">Preparing pages…</div>
          <template v-else>
            <PagePreview v-for="(p, i) in pages" :key="`${i}-${pages.length}`" :page="p" :index="i" :width="pageWidth(p)" :root="pane" />
          </template>
        </div>
      </div>
      <div v-if="showOptions" class="sl-print-options" data-testid="print-options">
        <label v-for="c in PRINT_CHOICES" :key="c.key" class="opt">
          <span>{{ c.label }}:</span>
          <select class="sl-select" :value="String(opts[c.key])" :data-testid="`print-${c.key}`" @change="choose(c.key, $event)">
            <option v-for="o in c.options" :key="String(o.value)" :value="String(o.value)">{{ o.label }}</option>
          </select>
        </label>
        <div class="sl-print-switches">
          <Switch v-for="s in PRINT_SWITCHES" :key="s.key" :label="s.label" :checked="!!opts[s.key]" :test-id="`print-${s.key}`" @change="set(s.key, $event as never)" />
        </div>
        <div class="sl-print-actions">
          <button type="button" class="sl-button" :disabled="!pages || busy" @click="download">{{ busy ? 'Creating PDF…' : 'Download PDF' }}</button>
          <button type="button" class="sl-button" :disabled="!pages" @click="print">Print</button>
        </div>
      </div>
    </div>
  </Modal>
</template>
