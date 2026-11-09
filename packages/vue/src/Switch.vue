<script setup lang="ts">
/** Labeled on/off switch. */
import { vChecked } from './directives';

defineProps<{ label: string; checked: boolean; testId?: string }>();
const emit = defineEmits<{ change: [value: boolean] }>();
</script>

<template>
  <label class="sl-switch">
    <input v-checked="checked" type="checkbox" role="switch" :data-testid="testId" @change="emit('change', ($event.target as HTMLInputElement).checked)" />
    <span>{{ label }}</span>
  </label>
</template>
