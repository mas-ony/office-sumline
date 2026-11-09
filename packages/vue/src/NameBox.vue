<script setup lang="ts">
/** The name box: the selection's reference, a field to go to a cell or name, and the list of defined names. */
import { computed, ref, shallowRef, watch } from 'vue';
import { useUI } from './context';
import Icon from './Icon.vue';
import Popup from './Popup.vue';

const ui = useUI();
const name = ref(ui.value.state.rangeRef);
const focused = ref(false);
const list = shallowRef<DOMRect | null>(null);
const wrap = ref<HTMLDivElement | null>(null);
watch([() => ui.value.state.rangeRef, focused], ([rangeRef, f]) => {
  if (!f) name.value = rangeRef;
});
const names = computed(() => ui.value.doc.wb.names.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.')));

const goTo = () => {
  const { doc, grid } = ui.value;
  const text = name.value.trim();
  if (text && !doc.goTo(text)) ui.value.notify(`“${text}” is not a valid cell reference or name.`);
  focused.value = false;
  grid?.focusIdle();
};
const focus = (e: FocusEvent) => {
  focused.value = true;
  (e.target as HTMLInputElement).select();
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    goTo();
  } else if (e.key === 'Escape') {
    e.preventDefault();
    name.value = ui.value.state.rangeRef;
    ui.value.grid?.focusIdle();
  }
};
const toggle = () => (list.value = list.value ? null : wrap.value!.getBoundingClientRect());
const pick = (n: string) => {
  list.value = null;
  ui.value.doc.goTo(n);
  ui.value.grid?.focusIdle();
};
const edit = (e: MouseEvent, n?: string) => {
  e.stopPropagation();
  list.value = null;
  ui.value.openDialog({ kind: 'name-edit', name: n });
};
const add = () => {
  list.value = null;
  ui.value.openDialog({ kind: 'name-edit' });
};
</script>

<template>
  <div ref="wrap" class="sl-namebox">
    <input
      aria-label="Name box"
      data-testid="name-box"
      :value="name"
      spellcheck="false"
      autocapitalize="characters"
      autocomplete="off"
      enterkeyhint="go"
      @focus="focus"
      @blur="focused = false"
      @input="name = ($event.target as HTMLInputElement).value"
      @keydown="key"
    />
    <button type="button" class="sl-namebox-caret" aria-label="Defined names" data-testid="names-button" :aria-expanded="!!list" @mousedown.prevent @click="toggle">
      <Icon name="arrow_drop_down" />
    </button>
    <Popup v-if="list" :anchor="list" label="Defined names" test-id="names-list" @close="list = null">
      <div :style="{ padding: '6px 0', minWidth: '260px' }">
        <div v-if="names.length === 0" class="sl-hint" :style="{ padding: '6px 14px' }">No defined names</div>
        <button v-for="n in names" :key="`${n.name}-${n.localSheet ?? ''}`" type="button" class="sl-picker-item" @click="pick(n.name)">
          <span :style="{ flex: 1 }">{{ n.name }}</span>
          <span class="sl-hint" :style="{ fontSize: '12px' }">{{ n.formula }}</span>
          <span v-if="!ui.state.readOnly && n.localSheet === undefined" role="button" tabindex="0" :aria-label="`Edit ${n.name}`" :style="{ display: 'flex', color: '#666' }" @click="edit($event, n.name)">
            <Icon name="edit" :size="16" />
          </span>
        </button>
        <div v-if="!ui.state.readOnly" :style="{ padding: '8px 12px 4px' }">
          <button type="button" class="sl-button wide" data-testid="names-add" @click="add">Add new entry</button>
        </div>
      </div>
    </Popup>
  </div>
</template>
