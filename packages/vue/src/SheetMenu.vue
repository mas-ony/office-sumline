<script setup lang="ts">
/** A menu as a bottom sheet (phones): submenus open in place with a back button. */
import type { MenuItem } from '@sumline/core/views';
import { computed, shallowRef } from 'vue';
import Icon from './Icon.vue';

const props = defineProps<{ items: MenuItem[]; title?: string; testId?: string }>();
const emit = defineEmits<{ close: [] }>();

const stack = shallowRef<{ title?: string; items: MenuItem[] }[]>([{ title: props.title, items: props.items }]);
const top = computed(() => stack.value[stack.value.length - 1]);
const pick = (it: MenuItem) => {
  if (it.submenu) {
    stack.value = [...stack.value, { title: it.title, items: it.submenu }];
    return;
  }
  if (!it.keepOpen) emit('close');
  it.onClick?.();
};
</script>

<template>
  <div
    class="sl-sheet-backdrop"
    @pointerdown="
      $event.preventDefault();
      emit('close');
    "
  />
  <div class="sl-bottom-sheet" role="menu" :aria-label="top.title" :data-testid="testId">
    <div class="sl-sheet-grip" />
    <div v-if="stack.length > 1 || top.title" class="sl-sheet-title">
      <button v-if="stack.length > 1" type="button" aria-label="Back" @click="stack = stack.slice(0, -1)">
        <Icon name="arrow_right" class="sl-flip" />
      </button>
      <span v-else :style="{ width: '8px' }" />
      <span>{{ top.title }}</span>
    </div>
    <template v-for="(it, i) in top.items" :key="i">
      <div v-if="it.type === 'line'" class="sl-menu-sep" role="separator" />
      <div v-else-if="it.type === 'header'" class="sl-menu-header">{{ it.title }}</div>
      <button v-else type="button" role="menuitem" :data-testid="it.testId" class="sl-menu-item" :disabled="it.disabled" @click="pick(it)">
        <span class="icon-slot">
          <Icon v-if="it.icon" :name="it.icon" />
          <Icon v-else-if="it.checked" name="check" class="check" />
        </span>
        <span class="title">{{ it.title }}</span>
        <Icon v-if="it.submenu" name="arrow_right" class="sub" />
      </button>
    </template>
  </div>
</template>
