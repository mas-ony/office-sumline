<script setup lang="ts">
/**
 * Worksheet tabs above the toolbar: select, add (+), rename (double click or the tab menu),
 * reorder by dragging, and the tab context menu (right click or long press).
 */
import { cmd, tabMenu } from '@sumline/core/views';
import { onMounted, ref, shallowRef, watch } from 'vue';
import { useUI } from './context';
import { vFocus } from './directives';
import Icon from './Icon.vue';
import Menu from './Menu.vue';

const ui = useUI();
const editing = ref<number | null>(null);
const draft = ref('');
const menu = shallowRef<{ index: number; at: { x: number; y: number } } | null>(null);
const drag = shallowRef<{ from: number; over: number; side: 'before' | 'after' } | null>(null);
let press: { timer: number; index: number; x: number; y: number; moved: boolean; pointerId: number } | null = null;
const list = ref<HTMLDivElement | null>(null);

// Keep the active tab in view.
const showActive = () => list.value?.querySelector<HTMLElement>('.sl-tab.selected')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
onMounted(showActive);
watch(() => ui.value.state.activeIndex, showActive, { flush: 'post' });

const commitRename = () => {
  if (editing.value === null) return;
  const { doc } = ui.value;
  const index = editing.value;
  const name = draft.value.trim();
  editing.value = null;
  if (!name || name === doc.wb.sheets[index]?.name) return;
  const error = doc.renameSheet(index, name);
  if (error) ui.value.notify(error);
};

const openMenu = (index: number, x: number, y: number) => {
  ui.value.doc.setActiveSheet(index);
  menu.value = { index, at: { x, y } };
};

const down = (e: PointerEvent, index: number) => {
  if (e.button !== 0 || editing.value !== null) return;
  const timer = window.setTimeout(() => {
    if (press && !press.moved) {
      openMenu(index, press.x, press.y);
      press = null;
    }
  }, 500);
  press = { timer, index, x: e.clientX, y: e.clientY, moved: false, pointerId: e.pointerId };
};

const move = (e: PointerEvent) => {
  const p = press;
  const d = drag.value;
  if (!p || p.pointerId !== e.pointerId) return;
  if (!p.moved && Math.hypot(e.clientX - p.x, e.clientY - p.y) > 8) {
    p.moved = true;
    window.clearTimeout(p.timer);
    // Mouse drags reorder; touch moves scroll the tab strip.
    if (e.pointerType === 'mouse' && !ui.value.state.readOnly) {
      (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
      drag.value = { from: p.index, over: p.index, side: 'after' };
    }
  }
  if (d) {
    const el = document.elementFromPoint(e.clientX, e.clientY)?.closest<HTMLElement>('[data-tab-index]');
    if (el) {
      const over = Number(el.dataset.tabIndex);
      const rect = el.getBoundingClientRect();
      drag.value = { ...d, over, side: e.clientX < rect.left + rect.width / 2 ? 'before' : 'after' };
    }
  }
};

const up = (e: PointerEvent, index: number) => {
  const { doc, grid } = ui.value;
  const p = press;
  if (p) window.clearTimeout(p.timer);
  press = null;
  if (drag.value) {
    const { from, over, side } = drag.value;
    drag.value = null;
    let to = side === 'before' ? over : over + 1;
    if (from < to) to--;
    if (to !== from) doc.moveSheet(from, Math.max(0, Math.min(doc.wb.sheets.length - 1, to)));
    return;
  }
  if (p && !p.moved && e.button === 0) {
    if (grid?.isEditing() && !grid.editState.text.startsWith('=')) grid.commitEdit();
    doc.setActiveSheet(index);
  }
};

const cancel = () => {
  if (press) window.clearTimeout(press.timer);
  press = null;
  drag.value = null;
};

const rename = (index: number, name: string) => {
  if (ui.value.state.readOnly) return;
  draft.value = name;
  editing.value = index;
};

const contextMenu = (e: MouseEvent, index: number) => {
  e.preventDefault();
  openMenu(index, e.clientX, e.clientY);
};

const renameKey = (e: KeyboardEvent) => {
  if (e.key === 'Enter') commitRename();
  else if (e.key === 'Escape') editing.value = null;
};

const dropClass = (index: number) => {
  const d = drag.value;
  return d && d.over === index && d.from !== index ? ` drop-${d.side}` : '';
};
</script>

<template>
  <div class="sl-tabs" data-testid="sheet-tabs">
    <div ref="list" class="sl-tabs-list" role="tablist" aria-label="Worksheets">
      <template v-for="s in ui.state.sheets" :key="s.index">
        <template v-if="s.state === 'visible'">
          <input
            v-if="editing === s.index"
            v-focus
            class="sl-tab-input"
            aria-label="Worksheet name"
            data-testid="tab-rename-input"
            :value="draft"
            @input="draft = ($event.target as HTMLInputElement).value"
            @blur="commitRename"
            @keydown="renameKey"
          />
          <button
            v-else
            type="button"
            role="tab"
            :aria-selected="s.index === ui.state.activeIndex"
            :data-tab-index="s.index"
            :data-testid="`tab-${s.name}`"
            :class="`sl-tab${s.index === ui.state.activeIndex ? ' selected' : ''}${dropClass(s.index)}`"
            @pointerdown="down($event, s.index)"
            @pointermove="move"
            @pointerup="up($event, s.index)"
            @pointercancel="cancel"
            @dblclick="rename(s.index, s.name)"
            @contextmenu="contextMenu($event, s.index)"
          >
            <span v-if="s.color" class="sl-tab-color" :style="{ background: s.color }" />{{ s.name }}
          </button>
        </template>
      </template>
    </div>
    <button v-if="!ui.state.readOnly" type="button" class="sl-tab-add" aria-label="Add worksheet" title="Add worksheet" data-testid="tab-add" @click="cmd.newWorksheet(ui)">
      <Icon name="add" />
    </button>
    <Menu
      v-if="menu"
      :items="tabMenu(ui, menu.index)"
      :at="menu.at"
      :sheet="ui.phone"
      :title="ui.doc.wb.sheets[menu.index]?.name"
      label="Worksheet"
      test-id="tab-menu"
      @close="menu = null"
    />
  </div>
</template>
