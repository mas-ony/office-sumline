<script setup lang="ts">
/** The toolbar's border picker: where to draw (all, outer, inner…), the colour, thickness and line kind. */
import { BORDER_KINDS, type BorderKind, type BorderLine, borderStyleFor, LINE_KINDS, type LineKind, rgbSpec } from '@sumline/core/views';
import { ref, shallowRef } from 'vue';
import ColorPicker from './ColorPicker.vue';
import { useUI } from './context';
import Icon from './Icon.vue';

const emit = defineEmits<{ close: [] }>();

const WIDTHS = [1, 2, 3, 4, 5];
const ui = useUI();
const color = ref('#000000');
const width = ref(1);
const kind = ref<LineKind>('solid');
const colorRect = shallowRef<DOMRect | null>(null);
const menu = ref<'width' | 'kind' | null>(null);
const colorBtn = ref<HTMLButtonElement | null>(null);

const apply = (k: BorderKind) => {
  const { doc, grid } = ui.value;
  const line: BorderLine = { style: borderStyleFor(width.value, kind.value), color: rgbSpec(color.value) };
  if (grid?.isEditing()) grid.commitEdit();
  doc.setBorder(k, k === 'none' ? null : line);
  emit('close');
};
const sample = (w: number, k: LineKind) => (k === 'solid' ? { height: `${w}px`, background: '#333' } : { borderTop: `${k === 'double' ? 3 : 2}px ${k} #333` });
const toggleColor = () => (colorRect.value = colorRect.value ? null : colorBtn.value!.getBoundingClientRect());
const setWidth = (w: number) => {
  width.value = w;
  menu.value = null;
};
const setKind = (k: LineKind) => {
  kind.value = k;
  menu.value = null;
};
</script>

<template>
  <div class="sl-border-picker" data-testid="border-picker">
    <div class="sl-picker-icons cols-5">
      <button
        v-for="b in BORDER_KINDS"
        :key="b.kind"
        type="button"
        class="sl-icon-btn"
        :title="b.label"
        :aria-label="b.label"
        :data-testid="`border-${b.kind}`"
        :disabled="!ui.state.multiCell && (b.kind === 'inside' || b.kind === 'horizontal' || b.kind === 'vertical')"
        @click="apply(b.kind)"
      >
        <Icon :name="b.icon" />
      </button>
    </div>
    <div class="sl-border-options">
      <button ref="colorBtn" type="button" class="sl-icon-btn" title="Border color" aria-label="Border color" data-testid="border-color" @click="toggleColor">
        <Icon name="color_lens" />
        <span :style="{ position: 'absolute', width: '14px', height: '3px', marginTop: '22px', background: color }" />
      </button>
      <div :style="{ position: 'relative' }">
        <button type="button" class="sl-select-btn" aria-label="Border width" data-testid="border-width" @click="menu = menu === 'width' ? null : 'width'">
          <span class="sl-line-sample" :style="sample(width, 'solid')" />
          <Icon name="arrow_drop_down" />
        </button>
        <div v-if="menu === 'width'" class="sl-picker" :style="{ position: 'absolute', top: '100%', left: 0 }">
          <button v-for="w in WIDTHS" :key="w" type="button" :class="`sl-picker-item${w === width ? ' selected' : ''}`" :aria-label="`${w}px`" @click="setWidth(w)">
            <span class="sl-line-sample" :style="sample(w, 'solid')" />
          </button>
        </div>
      </div>
      <div :style="{ position: 'relative' }">
        <button type="button" class="sl-select-btn" aria-label="Border style" data-testid="border-style" @click="menu = menu === 'kind' ? null : 'kind'">
          <span class="sl-line-sample" :style="sample(2, kind)" />
          <Icon name="arrow_drop_down" />
        </button>
        <div v-if="menu === 'kind'" class="sl-picker" :style="{ position: 'absolute', top: '100%', left: 0 }">
          <button
            v-for="k in LINE_KINDS"
            :key="k"
            type="button"
            :class="`sl-picker-item${k === kind ? ' selected' : ''}`"
            :aria-label="k"
            :data-testid="`border-style-${k}`"
            @click="setKind(k)"
          >
            <span class="sl-line-sample" :style="sample(2, k)" />
          </button>
        </div>
      </div>
    </div>
    <ColorPicker v-if="colorRect" :anchor="colorRect" :value="color" @change="color = $event ?? '#000000'" @close="colorRect = null" />
  </div>
</template>
