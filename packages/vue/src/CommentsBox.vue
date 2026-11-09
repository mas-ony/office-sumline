<script setup lang="ts">
/** The comment thread box at a cell: the comments, with edit / resolve / delete, and a reply field. */
import { cellName, type MenuItem } from '@sumline/core/views';
import { computed, ref, shallowRef } from 'vue';
import Anchored from './Anchored.vue';
import { useUI } from './context';
import { vFocus } from './directives';
import Icon from './Icon.vue';
import Menu from './Menu.vue';

const props = defineProps<{ r: number; c: number; rect: DOMRect }>();
const emit = defineEmits<{ close: [] }>();

const ui = useUI();
const text = ref('');
const editing = shallowRef<{ id: string; text: string } | null>(null);
const menu = shallowRef<{ id: string; at: DOMRect } | null>(null);
// A fresh object on every change: the thread is changed in place.
const box = computed(() => ({ thread: ui.value.doc.sheet.commentAt(props.r, props.c) }));
const ro = computed(() => ui.value.state.readOnly);

function when(iso?: string): string {
  if (!iso) return '';
  const t = Date.parse(iso.endsWith('Z') || /[+-]\d\d:\d\d$/.test(iso) ? iso : `${iso}Z`);
  if (!Number.isFinite(t)) return '';
  const s = Math.round((Date.now() - t) / 1000);
  if (s < 60) return 'Just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  return new Date(t).toLocaleDateString();
}

const save = () => {
  if (!text.value.trim()) return;
  ui.value.doc.addComment(props.r, props.c, text.value.trim());
  text.value = '';
};
const saveEdit = (id: string) => {
  if (editing.value) ui.value.doc.editComment(id, editing.value.text);
  editing.value = null;
};
const key = (e: KeyboardEvent) => {
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) save();
};
const menuItems = computed<MenuItem[]>(() => {
  const m = menu.value;
  if (!m) return [];
  const { doc } = ui.value;
  const t = box.value.thread;
  return [
    { title: 'Edit', icon: 'edit', onClick: () => (editing.value = { id: m.id, text: t?.comments.find((x) => x.id === m.id)?.text ?? '' }) },
    { title: t?.resolved ? 'Reopen' : 'Resolve', icon: 'check', onClick: () => doc.setResolved(m.id, !t?.resolved) },
    { title: 'Delete', icon: 'delete', onClick: () => doc.deleteComment(m.id) },
  ];
});
</script>

<template>
  <Anchored :rect="rect" box-class="sl-comments" test-id="comments-box" :label="`Comments ${cellName(r, c)}`" @close="emit('close')">
    <template v-if="box.thread">
      <div v-for="(cm, i) in box.thread.comments" :key="cm.id" class="sl-comment" data-testid="comment">
        <div class="who">
          <span :style="{ display: 'flex', alignItems: 'center', gap: '8px' }">
            <span
              :style="{ width: '28px', height: '28px', borderRadius: '50%', border: '1px solid #ccc', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', color: '#555' }"
            >{{ (cm.author || '?').slice(0, 1).toUpperCase() }}</span>
            <span>{{ cm.author || 'Guest' }}<div class="when">{{ when(cm.created) || (box.thread.kind === 'note' ? 'Note' : '') }}{{ i === 0 && box.thread.resolved ? ' · Resolved' : '' }}</div></span>
          </span>
          <button
            v-if="!ro"
            type="button"
            class="sl-icon-btn"
            :style="{ width: '26px', height: '26px' }"
            aria-label="Comment options"
            data-testid="comment-options"
            @click="menu = { id: cm.id, at: ($event.currentTarget as HTMLElement).getBoundingClientRect() }"
          >
            <Icon name="arrow_drop_down" />
          </button>
        </div>
        <div v-if="editing?.id === cm.id" :style="{ display: 'grid', gap: '6px', marginTop: '6px' }">
          <textarea class="sl-textarea" :value="editing.text" @input="editing = { id: cm.id, text: ($event.target as HTMLTextAreaElement).value }" />
          <div :style="{ display: 'flex', gap: '6px' }">
            <button type="button" class="sl-button" @click="saveEdit(cm.id)">Save</button>
            <button type="button" class="sl-button light" @click="editing = null">Cancel</button>
          </div>
        </div>
        <div v-else class="body">{{ cm.text }}</div>
      </div>
    </template>
    <div v-if="!ro" class="sl-comments-input">
      <textarea
        v-focus
        class="sl-textarea"
        :value="text"
        :placeholder="box.thread ? 'Reply' : 'Comment'"
        aria-label="Comment"
        data-testid="comment-input"
        @input="text = ($event.target as HTMLTextAreaElement).value"
        @keydown="key"
      />
      <div :style="{ display: 'flex', gap: '6px' }">
        <button type="button" class="sl-button" :style="{ flex: 1 }" :disabled="!text.trim()" data-testid="comment-save" @click="save">Save</button>
        <button type="button" class="sl-button" :style="{ flex: 1 }" @click="emit('close')">Cancel</button>
      </div>
    </div>
    <Menu v-if="menu" :at="menu.at" :items="menuItems" @close="menu = null" />
  </Anchored>
</template>
