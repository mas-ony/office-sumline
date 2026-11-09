<script setup lang="ts">
/** The Shape Settings dialog: fill, border, font colour and text of a shape. */
import { DEFAULT_SHAPE_FILL, DEFAULT_SHAPE_LINE, LINE_PRESETS } from '@sumline/core/views';
import { computed, ref } from 'vue';
import ColorField from '../ColorField.vue';
import { useUI } from '../context';
import Modal from '../Modal.vue';
import Switch from '../Switch.vue';

const props = defineProps<{ id: string }>();
const emit = defineEmits<{ close: [] }>();

const WIDTHS = ['0', '1', '2', '3', '4', '5', '6', '8', '10'];
const ui = useUI();
const { doc } = ui.value;
const shape = computed(() => ui.value.doc.sheet.shapes.find((s) => s.id === props.id));
const first = shape.value;
const line = computed(() => (shape.value ? LINE_PRESETS.has(shape.value.preset) : false));
const fill = ref(first?.fill ?? DEFAULT_SHAPE_FILL);
const transparent = ref(first ? first.fill === null : false);
const border = ref(first?.line?.color ?? DEFAULT_SHAPE_LINE);
const width = ref(first?.line ? String(Math.round(first.line.width)) : '0');
const fontColor = ref(first?.fontColor ?? '#000000');
const text = ref(first?.text ?? '');

const save = () => {
  const s = shape.value;
  if (!s) return;
  const w = Number(width.value);
  doc.updateShape(props.id, {
    fill: line.value || transparent.value ? null : fill.value || DEFAULT_SHAPE_FILL,
    line: w > 0 ? { color: border.value || DEFAULT_SHAPE_LINE, width: w, dash: s.line?.dash } : null,
    fontColor: fontColor.value || '#000000',
    text: line.value ? undefined : text.value || undefined,
  });
  emit('close');
};
</script>

<template>
  <Modal v-if="shape" title="Shape Settings" icon="interests" :width="380" test-id="shape-settings" @close="emit('close')">
    <div class="sl-form">
      <div v-if="!line" class="sl-row sl-row-center">
        <div class="sl-field">
          <label>Background Color</label>
          <ColorField :value="fill" :disabled="transparent" test-id="shape-fill" @change="fill = $event" />
        </div>
        <Switch label="Transparent" :checked="transparent" test-id="shape-transparent" @change="transparent = $event" />
      </div>
      <div class="sl-row">
        <div class="sl-field">
          <label>Border Color</label>
          <ColorField :value="border" test-id="shape-border" @change="border = $event" />
        </div>
        <div class="sl-field">
          <label for="shape-width">Border Width</label>
          <select id="shape-width" class="sl-select" :value="width" data-testid="shape-border-width" @change="width = ($event.target as HTMLSelectElement).value">
            <option v-for="v in WIDTHS" :key="v" :value="v">{{ v }}</option>
          </select>
        </div>
      </div>
      <template v-if="!line">
        <div class="sl-row">
          <div class="sl-field">
            <label>Font Color</label>
            <ColorField :value="fontColor" test-id="shape-font-color" @change="fontColor = $event" />
          </div>
        </div>
        <div class="sl-field">
          <label for="shape-text">Text</label>
          <textarea id="shape-text" class="sl-textarea" rows="2" :value="text" data-testid="shape-text" @input="text = ($event.target as HTMLTextAreaElement).value" />
        </div>
      </template>
    </div>
    <div class="sl-dialog-buttons">
      <button type="button" class="sl-button wide" data-testid="shape-save" @click="save">Save</button>
      <button type="button" class="sl-button wide" @click="emit('close')">Cancel</button>
    </div>
  </Modal>
</template>
