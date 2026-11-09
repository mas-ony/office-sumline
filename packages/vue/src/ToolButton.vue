<script setup lang="ts">
/** A toolbar button: an icon (or other content) with a tooltip, optionally pressed. */
import type { IconName } from '@sumline/core/views';
import Icon from './Icon.vue';

withDefaults(defineProps<{ icon?: IconName; label: string; active?: boolean; disabled?: boolean; testId?: string; expanded?: boolean }>(), {
  icon: undefined,
  active: undefined,
  disabled: undefined,
  testId: undefined,
  expanded: undefined,
});
const emit = defineEmits<{ click: [e: MouseEvent] }>();
</script>

<template>
  <button
    type="button"
    :class="`sl-tool${active ? ' active' : ''}`"
    :title="label"
    :aria-label="label"
    :aria-pressed="active"
    :aria-expanded="expanded"
    :disabled="disabled"
    :data-testid="testId"
    @mousedown.prevent
    @click="emit('click', $event)"
  >
    <Icon v-if="icon" :name="icon" />
    <slot />
  </button>
</template>
