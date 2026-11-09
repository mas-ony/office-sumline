<script setup lang="ts">
/** The full-page app: the editor with the options and page connections of the shell. */
import { SumlineEditor } from '@sumline/vue';
import { markRaw, shallowRef } from 'vue';
import { connectShell, shellOptions, watchLaunches } from '@sumline/web-shared/shell';

const options = shallowRef(shellOptions());
// A file opened through the installed app replaces the workbook.
watchLaunches((handle) => (options.value = { ...options.value, document: markRaw(handle) }));
</script>

<template>
  <SumlineEditor v-bind="options" @ready="connectShell" />
</template>
