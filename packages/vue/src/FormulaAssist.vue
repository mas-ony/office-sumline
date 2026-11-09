<script setup lang="ts">
/**
 * Typing a function name shows matching functions with their signature, description and an
 * example; inside a function's parentheses, its arguments with the current one in bold.
 */
import { exampleFor, highlightedArg, type UIContextValue } from '@sumline/core/views';
import { computed, onMounted, onUpdated, ref, shallowRef, watch } from 'vue';
import { useUI } from './context';

const ui = useUI();
const box = ref<HTMLDivElement | null>(null);
const pos = shallowRef<{ left: number; top: number } | null>(null);
const view = computed(() => ui.value.ui.assist.view);

function anchorRect(ctx: UIContextValue): DOMRect | null {
  if (ctx.edit.origin === 'bar') return ctx.ui.rootElement?.querySelector<HTMLElement>('[data-testid="formula-input"]')?.getBoundingClientRect() ?? null;
  return ctx.grid?.editorClientRect() ?? null;
}

const place = () => {
  const ctx = ui.value;
  if (!ctx.ui.assist.view.visible) {
    if (pos.value) pos.value = null;
    return;
  }
  const a = anchorRect(ctx);
  const el = box.value;
  if (!a || !el) return;
  const vh = window.visualViewport?.height ?? window.innerHeight;
  const vw = window.innerWidth;
  const h = el.offsetHeight;
  const w = el.offsetWidth;
  let top = a.bottom + 2;
  if (top + h > vh - 4 && a.top - h - 2 > 4) top = a.top - h - 2;
  const left = Math.max(4, Math.min(a.left, vw - w - 4));
  if (!pos.value || pos.value.left !== left || pos.value.top !== top) pos.value = { left, top };
};
onMounted(place);
onUpdated(place);

const scrollActive = () => box.value?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' });
onMounted(scrollActive);
watch([() => view.value.index, () => view.value.showList], scrollActive, { flush: 'post' });

const boxStyle = computed(() => ({ position: 'fixed' as const, left: `${pos.value ? pos.value.left : -9999}px`, top: `${pos.value ? pos.value.top : 0}px` }));
</script>

<template>
  <Teleport v-if="view.visible" :to="ui.layer">
    <div v-if="view.showList" ref="box" class="sl-suggest" :style="boxStyle" role="listbox" aria-label="Functions" data-testid="formula-suggestions">
      <div
        v-for="(s, i) in view.suggestions"
        :key="s.name"
        role="option"
        :aria-selected="i === view.index"
        :class="`sl-suggest-item${i === view.index ? ' active' : ''}`"
        @mousedown.prevent
        @pointerdown.prevent
        @click="ui.ui.assist.apply(s)"
      >
        <div class="sl-suggest-sig">{{ s.category === 'Name' ? s.name : `${s.name}(${s.args.join(', ')})` }}</div>
        <div v-if="s.description" class="sl-suggest-desc">{{ s.description }}</div>
        <div v-if="s.category !== 'Name'" class="sl-suggest-example">{{ exampleFor(s) }}</div>
      </div>
    </div>
    <div v-else-if="view.call" ref="box" class="sl-arghint" :style="boxStyle" data-testid="formula-hint">
      <div class="sl-suggest-sig"><b>{{ view.call.info.name }}</b>(<span v-for="(a, i) in view.call.info.args" :key="i">{{ i > 0 ? ', ' : '' }}<span v-if="i === highlightedArg(view.call.info, view.call.argIndex)" class="cur">{{ a }}</span><span v-else>{{ a }}</span></span>)</div>
      <div v-if="view.call.info.description" class="sl-suggest-desc">{{ view.call.info.description }}</div>
    </div>
    <span v-else />
  </Teleport>
</template>
