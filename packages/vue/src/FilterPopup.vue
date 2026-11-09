<script setup lang="ts">
/** The filter pop-up of a column of the auto filter: sort, a condition, and the values to show. */
import type { ColumnFilter, FilterOperator } from '@sumline/core/views';
import { computed, ref, shallowRef } from 'vue';
import Anchored from './Anchored.vue';
import { useUI } from './context';
import { vChecked } from './directives';

const props = defineProps<{ col: number; at: { x: number; y: number } }>();
const emit = defineEmits<{ close: [] }>();

const OPERATORS: { key: FilterOperator; label: string }[] = [
  { key: 'contains', label: 'Contains' },
  { key: 'notContains', label: 'Does not contain' },
  { key: 'beginsWith', label: 'Begins with' },
  { key: 'endsWith', label: 'Ends with' },
  { key: 'equal', label: 'Equal' },
  { key: 'notEqual', label: 'Not equal' },
  { key: 'greaterThan', label: 'Greater than' },
  { key: 'lessThan', label: 'Lower than' },
];

const { doc } = useUI().value;
const current: ColumnFilter | undefined = doc.sheet.autoFilter?.columns.get(props.col);
// The column's values when the pop-up opened.
const values = computed(() => doc.filterValues(props.col));
const search = ref(current?.conditions?.[0]?.value ?? '');
const op = ref<FilterOperator>(current?.conditions?.[0]?.op ?? 'contains');
const checked = shallowRef<Set<string>>(
  current?.values ? new Set([...current.values.map((v) => v.toLowerCase()), ...(current.blanks ? [''] : [])]) : new Set(values.value.map((v) => v.text.toLowerCase())),
);
const shown = computed(() => values.value.filter((v) => !search.value || op.value !== 'contains' || v.text.toLowerCase().includes(search.value.toLowerCase())));
const all = computed(() => values.value.every((v) => checked.value.has(v.text.toLowerCase())));
const rect = computed(() => new DOMRect(props.at.x, props.at.y, 0, 0));

const apply = () => {
  if (search.value.trim()) {
    doc.setColumnFilter(props.col, { conditions: [{ op: op.value, value: search.value.trim() }] });
  } else if (all.value) {
    doc.setColumnFilter(props.col, null);
  } else {
    const keep = values.value.filter((v) => checked.value.has(v.text.toLowerCase()) && v.text !== '').map((v) => v.text);
    doc.setColumnFilter(props.col, { values: keep, blanks: checked.value.has('') });
  }
  emit('close');
};
const sort = (ascending: boolean) => {
  doc.sortFilterColumn(props.col, ascending);
  emit('close');
};
const checkAll = (e: Event) => {
  checked.value = (e.target as HTMLInputElement).checked ? new Set(values.value.map((v) => v.text.toLowerCase())) : new Set();
};
const check = (e: Event, text: string) => {
  const next = new Set(checked.value);
  if ((e.target as HTMLInputElement).checked) next.add(text.toLowerCase());
  else next.delete(text.toLowerCase());
  checked.value = next;
};
</script>

<template>
  <Anchored :rect="rect" box-class="sl-filter" test-id="filter-popup" label="Filter Options" @close="emit('close')">
    <div class="sl-filter-sort">
      <button type="button" data-testid="filter-sort-asc" @click="sort(true)">Sort smallest to largest</button>
      <button type="button" data-testid="filter-sort-desc" @click="sort(false)">Sort largest to smallest</button>
    </div>
    <div class="sl-filter-search">
      <input
        class="sl-input"
        placeholder="Search"
        :value="search"
        aria-label="Search"
        data-testid="filter-search"
        @input="search = ($event.target as HTMLInputElement).value"
        @keydown="$event.key === 'Enter' && apply()"
      />
      <select class="sl-select" :value="op" aria-label="Condition" data-testid="filter-operator" @change="op = ($event.target as HTMLSelectElement).value as FilterOperator">
        <option v-for="o in OPERATORS" :key="o.key" :value="o.key">{{ o.label }}</option>
      </select>
    </div>
    <div class="sl-filter-values" data-testid="filter-values">
      <label><input v-checked="all" type="checkbox" @change="checkAll" />(Select all)</label>
      <label v-for="v in shown" :key="v.text || '(blank)'"><input v-checked="checked.has(v.text.toLowerCase())" type="checkbox" @change="check($event, v.text)" />{{ v.text === '' ? '(Blanks)' : v.text }}</label>
    </div>
    <div class="sl-filter-actions">
      <button type="button" class="sl-button" data-testid="filter-ok" @click="apply">Ok</button>
      <button type="button" class="sl-button" @click="emit('close')">Cancel</button>
    </div>
  </Anchored>
</template>
