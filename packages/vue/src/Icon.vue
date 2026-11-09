<script setup lang="ts">
/** An icon of the editor's set, as an inline SVG that takes the text colour (bundled package assets, not user content). */
import { type IconName, iconSvg } from '@sumline/core/views';

defineProps<{ name: IconName; size?: number }>();
</script>

<template>
  <span class="mi" aria-hidden="true" :style="size ? { width: `${size}px`, height: `${size}px` } : undefined" v-html="iconSvg(name)" />
</template>
