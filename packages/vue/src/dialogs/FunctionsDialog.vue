<script setup lang="ts">
/** The function list: search by name or description, filter by category, click to insert. */
import { exampleFor, functionCatalog, functionCategories } from '@sumline/core/views';
import { computed, ref } from 'vue';
import { useUI } from '../context';
import Modal from '../Modal.vue';

const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const query = ref('');
const category = ref('All');
const categories = ['All', ...functionCategories()];
const items = computed(() => {
  const q = query.value.trim().toUpperCase();
  return [...functionCatalog().values()].filter((f) => (category.value === 'All' || f.category === category.value) && (!q || f.name.includes(q) || f.description.toUpperCase().includes(q)));
});
const insert = (name: string) => {
  emit('close');
  const { state, grid } = ui.value;
  if (state.readOnly || !grid) return;
  if (grid.isEditing() && grid.editState.text.startsWith('=')) {
    const { text, caret } = grid.editState;
    grid.replaceEditText(`${text.slice(0, caret)}${name}(${text.slice(caret)}`, caret + name.length + 1);
  } else {
    grid.startEdit(`=${name}(`, 'enter');
  }
};
</script>

<template>
  <Modal :title="`Function list (${items.length})`" icon="functions" :width="560" height="min(640px, calc(100vh - 16px))" test-id="functions-dialog" @close="emit('close')">
    <div class="sl-row" :style="{ marginBottom: '12px' }">
      <input class="sl-input" placeholder="Search functions" :value="query" data-testid="functions-search" @input="query = ($event.target as HTMLInputElement).value" />
      <select class="sl-select" :value="category" aria-label="Category" @change="category = ($event.target as HTMLSelectElement).value">
        <option v-for="c in categories" :key="c">{{ c }}</option>
      </select>
    </div>
    <div>
      <button
        v-for="f in items.slice(0, 400)"
        :key="f.name"
        type="button"
        class="sl-suggest-item"
        :style="{ display: 'block', width: '100%', border: 0, background: 'none', textAlign: 'left' }"
        @click="insert(f.name)"
      >
        <div class="sl-suggest-sig">{{ f.name }}({{ f.args.join(', ') }})</div>
        <div v-if="f.description" class="sl-suggest-desc">{{ f.description }}</div>
        <div class="sl-suggest-example">{{ exampleFor(f) }} · {{ f.category }}</div>
      </button>
    </div>
  </Modal>
</template>
