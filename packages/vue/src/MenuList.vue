<script setup lang="ts">
/** A pop-up menu with submenus, moved with the arrow keys. */
import { type MenuAnchor, type MenuItem, placeMenu } from '@sumline/core/views';
import { onMounted, onUnmounted, ref, shallowRef, watch } from 'vue';
import Icon from './Icon.vue';

const props = defineProps<{
  items: MenuItem[];
  at: MenuAnchor;
  side: 'below' | 'right';
  depth: number;
  label?: string;
  testId?: string;
}>();
const emit = defineEmits<{
  close: [];
  /** Back to the parent menu (ArrowLeft and Escape in a submenu). */
  back: [];
}>();

const el = ref<HTMLDivElement | null>(null);
const active = ref(-1);
const open = shallowRef<{ index: number; rect: DOMRect } | null>(null);
let hoverTimer = 0;

const place = () => {
  if (el.value) placeMenu(el.value, props.at, props.side);
};
onMounted(() => {
  place();
  if (props.depth === 0) el.value?.focus({ preventScroll: true });
});
watch(() => [props.at, props.side, props.items], place, { flush: 'post' });
onUnmounted(() => window.clearTimeout(hoverTimer));

const enabled = () => props.items.map((it, i) => (!it.type && !it.disabled ? i : -1)).filter((i) => i >= 0);
const activate = (i: number, btn?: HTMLElement | null) => {
  const it = props.items[i];
  if (!it || it.type || it.disabled) return;
  if (it.submenu) {
    const b = btn ?? (el.value?.querySelector(`[data-index="${i}"]`) as HTMLElement | null);
    if (b) open.value = { index: i, rect: b.getBoundingClientRect() };
    return;
  }
  if (!it.keepOpen) emit('close');
  it.onClick?.();
};

const onKeyDown = (e: KeyboardEvent) => {
  if (open.value) return;
  const list = enabled();
  const pos = list.indexOf(active.value);
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    active.value = list[(pos + 1) % list.length] ?? -1;
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    active.value = list[(pos - 1 + list.length) % list.length] ?? -1;
  } else if (e.key === 'Enter' || e.key === ' ' || (e.key === 'ArrowRight' && props.items[active.value]?.submenu)) {
    e.preventDefault();
    if (active.value >= 0) activate(active.value);
  } else if (e.key === 'Escape' || (e.key === 'ArrowLeft' && props.depth > 0)) {
    e.preventDefault();
    e.stopPropagation();
    if (props.depth > 0) emit('back');
    else emit('close');
  }
};

const enter = (e: MouseEvent, it: MenuItem, i: number) => {
  active.value = i;
  window.clearTimeout(hoverTimer);
  const btn = e.currentTarget as HTMLElement;
  if (it.submenu && !it.disabled) hoverTimer = window.setTimeout(() => (open.value = { index: i, rect: btn.getBoundingClientRect() }), 120);
  else if (open.value && open.value.index !== i) hoverTimer = window.setTimeout(() => (open.value = null), 200);
};
const back = () => {
  open.value = null;
  el.value?.focus({ preventScroll: true });
};
</script>

<template>
  <div
    ref="el"
    class="sl-popup"
    role="menu"
    :aria-label="label"
    tabindex="-1"
    :data-testid="testId"
    :style="{ zIndex: 100 + depth }"
    @keydown="onKeyDown"
    @contextmenu.prevent
  >
    <template v-for="(it, i) in items" :key="i">
      <div v-if="it.type === 'line'" class="sl-menu-sep" role="separator" />
      <div v-else-if="it.type === 'header'" class="sl-menu-header">{{ it.title }}</div>
      <button
        v-else
        type="button"
        :role="it.checked !== undefined ? 'menuitemcheckbox' : 'menuitem'"
        :aria-checked="it.checked"
        :aria-haspopup="it.submenu ? 'menu' : undefined"
        :aria-expanded="it.submenu ? open?.index === i : undefined"
        :data-index="i"
        :data-testid="it.testId"
        :class="['sl-menu-item', { active: active === i || open?.index === i }]"
        :disabled="it.disabled"
        @mouseenter="enter($event, it, i)"
        @click="activate(i, $event.currentTarget as HTMLElement)"
      >
        <span class="icon-slot">
          <Icon v-if="it.icon" :name="it.icon" />
          <Icon v-else-if="it.checked" name="check" class="check" />
        </span>
        <span class="title">{{ it.title }}</span>
        <span v-if="it.shortcut" class="shortcut">{{ it.shortcut }}</span>
        <Icon v-if="it.submenu" name="arrow_right" class="sub" />
      </button>
    </template>
  </div>
  <MenuList
    v-if="open && items[open.index]?.submenu"
    :items="items[open.index].submenu!"
    :at="open.rect"
    side="right"
    :depth="depth + 1"
    :label="items[open.index].title"
    @close="emit('close')"
    @back="back"
  />
</template>
