<script setup lang="ts">
/** Search and replace: find text in this or all worksheets, step through matches, replace. */
import { cellName, type SearchHit } from '@sumline/core/views';
import { computed, ref, shallowRef, watch } from 'vue';
import { useUI } from '../context';
import Modal from '../Modal.vue';
import Switch from '../Switch.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const { doc } = ui.value;
const query = ref('');
const replacement = ref('');
const scope = ref<'sheet' | 'all'>('sheet');
const matchCase = ref(false);
const wholeCell = ref(false);
const inFormulas = ref(false);
const hiddenToo = ref(false);
const hits = shallowRef<SearchHit[] | null>(null);
const index = ref(0);
const status = ref('');
const opts = computed(() => ({ matchCase: matchCase.value, wholeCell: wholeCell.value, inFormulas: inFormulas.value, allSheets: scope.value === 'all' }));

// Options or the query changed: the result list is stale.
watch([query, opts, hiddenToo], () => {
  hits.value = null;
  status.value = '';
});

// The match highlight goes when the dialog closes.
watch(
  () => ui.value.grid,
  (grid, _, onCleanup) => {
    onCleanup(() => {
      if (grid) {
        grid.searchHit = undefined;
        grid.requestRender();
      }
    });
  },
  { immediate: true },
);

const run = (): SearchHit[] => {
  let found = doc.find(query.value, opts.value);
  if (!hiddenToo.value) found = found.filter((h) => !h.sheet.isRowHidden(h.r) && !h.sheet.isColHidden(h.c));
  hits.value = found;
  return found;
};

const show = (list: SearchHit[], i: number) => {
  if (!list.length) {
    status.value = 'No matches';
    return;
  }
  const k = ((i % list.length) + list.length) % list.length;
  index.value = k;
  const hit = list[k];
  const si = doc.wb.sheets.indexOf(hit.sheet);
  if (si !== doc.activeIndex) doc.setActiveSheet(si);
  doc.selectCell(hit.r, hit.c);
  const { grid } = ui.value;
  if (grid) {
    grid.searchHit = { r: hit.r, c: hit.c };
    grid.requestRender();
  }
  status.value = `${k + 1} of ${list.length}${scope.value === 'all' ? ` · ${hit.sheet.name}!${cellName(hit.r, hit.c)}` : ''}`;
};

const search = () => show(run(), 0);
const step = (d: number) => {
  const list = hits.value ?? run();
  show(list, hits.value ? index.value + d : d > 0 ? 0 : list.length - 1);
};
const replaceOne = () => {
  const list = hits.value ?? run();
  if (!list.length) {
    status.value = 'No matches';
    return;
  }
  const hit = list[index.value] ?? list[0];
  doc.replace(hit, query.value, replacement.value, opts.value);
  const next = run();
  show(next, Math.min(index.value, next.length - 1));
};
const replaceAll = () => {
  if (!query.value) return;
  let n = 0;
  doc.transact('Replace all', () => {
    for (const hit of run()) if (doc.replace(hit, query.value, replacement.value, opts.value)) n++;
  });
  hits.value = null;
  status.value = n ? `${n} cell${n === 1 ? '' : 's'} replaced` : 'No matches';
};
const findKey = (e: KeyboardEvent) => {
  if (e.key === 'Enter') {
    if (hits.value) step(1);
    else search();
  }
};
</script>

<template>
  <Modal title="Search and replace" :width="560" test-id="search-dialog" @close="emit('close')">
    <div class="sl-form" :style="{ gridTemplateColumns: 'auto 1fr', alignItems: 'center', columnGap: '12px' }">
      <label for="sl-find">Find</label>
      <input id="sl-find" class="sl-input" :value="query" data-testid="search-find" data-autofocus="true" @input="query = ($event.target as HTMLInputElement).value" @keydown="findKey" />
      <label for="sl-repl">Replace by</label>
      <input id="sl-repl" class="sl-input" :value="replacement" data-testid="search-replace" @input="replacement = ($event.target as HTMLInputElement).value" />
      <label for="sl-scope">Search</label>
      <select id="sl-scope" class="sl-select" :value="scope" data-testid="search-scope" @change="scope = ($event.target as HTMLSelectElement).value as 'sheet' | 'all'">
        <option value="sheet">This worksheet</option>
        <option value="all">All worksheets</option>
      </select>
      <span />
      <div :style="{ display: 'grid', gap: '10px' }">
        <Switch label="Case sensitive" :checked="matchCase" test-id="search-case" @change="matchCase = $event" />
        <Switch label="Match all cell contents" :checked="wholeCell" test-id="search-whole" @change="wholeCell = $event" />
        <Switch label="Search inside formulas" :checked="inFormulas" test-id="search-formulas" @change="inFormulas = $event" />
        <Switch label="Replace cells that were hidden by a filter" :checked="hiddenToo" test-id="search-hidden" @change="hiddenToo = $event" />
      </div>
    </div>
    <div :style="{ display: 'flex', flexWrap: 'wrap', gap: '4px', marginTop: '22px' }">
      <button type="button" class="sl-button" :disabled="!query" data-testid="search-go" @click="search">Search</button>
      <button type="button" class="sl-button" :disabled="!query" data-testid="search-prev" @click="step(-1)">Prev</button>
      <button type="button" class="sl-button" :disabled="!query" data-testid="search-next" @click="step(1)">Next</button>
      <button type="button" class="sl-button" :disabled="!query || ui.state.readOnly" data-testid="search-replace-one" @click="replaceOne">Replace</button>
      <button type="button" class="sl-button" :disabled="!query || ui.state.readOnly" data-testid="search-replace-all" @click="replaceAll">Replace All</button>
    </div>
    <div class="sl-hint" :style="{ minHeight: '20px', marginTop: '10px' }" role="status" data-testid="search-status">{{ status }}</div>
  </Modal>
</template>
