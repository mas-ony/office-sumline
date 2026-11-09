<script setup lang="ts">
/**
 * Data validations: the rules of the selected cells, and a form to add or edit one (cell range,
 * action, message, type, criteria, values, blanks, list source, format style).
 */
import {
  blankValidationForm, type DataValidation, describeValidation, quoteSheetName, rangeName, rangesIntersect, ruleFromForm, TEXT_CRITERIA,
  type TextCriteria, type ValidationAction, VALIDATION_CRITERIA, VALIDATION_KINDS, type ValidationCriteria, type ValidationForm, type ValidationKind,
  validationFormOf,
} from '@sumline/core/views';
import { computed, ref, shallowRef } from 'vue';
import ColorPicker from '../ColorPicker.vue';
import { useUI } from '../context';
import Icon from '../Icon.vue';
import Modal from '../Modal.vue';
import Switch from '../Switch.vue';

const emit = defineEmits<{ close: [] }>();

const STYLE_COLORS = ['color', 'background'] as const;
const ui = useUI();
const { doc } = ui.value;
const mode = shallowRef<{ edit: number | null } | null>(null);
const error = ref('');
const colorPick = shallowRef<{ which: 'color' | 'background'; rect: DOMRect } | null>(null);

// The rules of the active sheet that touch the selection (worked out again on every change).
const list = computed(() => {
  const d = ui.value.doc;
  const selection = d.range;
  return d.sheet.validations.map((v, i) => ({ v, i })).filter(({ v }) => v.sqref.some((r) => rangesIntersect(r, selection)));
});

const blank = () => blankValidationForm(ui.value.doc.sheet.name, ui.value.doc.range);
const form = shallowRef<ValidationForm>(blank());
const set = (patch: Partial<ValidationForm>) => (form.value = { ...form.value, ...patch });

const save = () => {
  const res = ruleFromForm(form.value, ui.value.doc.sheet.name, doc.wb.date1904);
  if ('error' in res) {
    error.value = res.error;
    return;
  }
  const m = mode.value;
  if (m?.edit !== null && m?.edit !== undefined) doc.updateValidation(m.edit, res.rule);
  else doc.addValidation(res.rule);
  mode.value = null;
  error.value = '';
};

const edit = (v: DataValidation, i: number) => {
  form.value = validationFormOf(v, ui.value.doc.sheet.name, doc.wb.date1904);
  mode.value = { edit: i };
};
const add = () => {
  form.value = blank();
  mode.value = { edit: null };
};
const back = () => {
  mode.value = null;
  error.value = '';
};
const actionOf = (v: DataValidation) => (v.formatStyle ? 'Format style' : v.errorStyle === 'stop' ? 'Reject entry' : v.errorStyle === 'information' ? 'Information' : 'Warning message');
const value = (e: Event) => (e.target as HTMLInputElement | HTMLSelectElement).value;
const two = computed(() => form.value.criteria === 'between' || form.value.criteria === 'notBetween');
const valueType = computed(() => (form.value.kind === 'date' ? 'date' : form.value.kind === 'time' ? 'time' : 'text'));
const numeric = computed(() => form.value.kind === 'number' || form.value.kind === 'date' || form.value.kind === 'time' || form.value.kind === 'textLength');
const pickColor = (which: 'color' | 'background', e: MouseEvent) => (colorPick.value = { which, rect: (e.currentTarget as HTMLElement).getBoundingClientRect() });
const picked = (hex: string | null) => {
  const pick = colorPick.value;
  if (pick) set({ style: { ...form.value.style, [pick.which]: hex ?? undefined } });
};
</script>

<template>
  <Modal title="Validations" icon="rule" :width="480" test-id="validations-dialog" @close="emit('close')">
    <template v-if="!mode">
      <p v-if="list.length === 0" data-testid="dv-empty" :style="{ margin: '0 0 22px' }">No validation found for this selection</p>
      <div v-else :style="{ marginBottom: '16px' }">
        <div v-for="{ v, i } in list" :key="i" class="sl-validation-card" data-testid="dv-card">
          <div>
            <div>{{ describeValidation(v) }}</div>
            <div class="meta">{{ v.sqref.map((r) => rangeName(r)).join(' ') }} · {{ actionOf(v) }}</div>
          </div>
          <div :style="{ display: 'flex', gap: '2px' }">
            <button type="button" class="sl-icon-btn" aria-label="Edit validation" data-testid="dv-edit" @click="edit(v, i)">
              <Icon name="edit" />
            </button>
            <button type="button" class="sl-icon-btn" aria-label="Delete validation" data-testid="dv-delete" @click="doc.removeValidation(i)">
              <Icon name="delete" />
            </button>
          </div>
        </div>
      </div>
      <button type="button" class="sl-button wide" data-testid="dv-add" @click="add">Add new validation</button>
    </template>
    <template v-else>
      <div class="sl-form">
        <div class="sl-row">
          <div class="sl-field">
            <label for="dv-range">Cell range</label>
            <div class="sl-range-input">
              <input id="dv-range" class="sl-input" :value="form.range" data-testid="dv-range" @input="set({ range: value($event) })" />
              <button type="button" title="Use the selection" aria-label="Use the selection" @click="set({ range: `${quoteSheetName(ui.doc.sheet.name)}!${rangeName(ui.doc.range)}` })">
                <Icon name="highlight_alt" :size="18" />
              </button>
            </div>
          </div>
          <div class="sl-field">
            <label for="dv-action">Action</label>
            <select id="dv-action" class="sl-select" :value="form.action" data-testid="dv-action" @change="set({ action: value($event) as ValidationAction })">
              <option value="warning">Warning message</option>
              <option value="reject">Reject entry</option>
              <option value="format">Format style</option>
            </select>
          </div>
        </div>
        <div v-if="form.action !== 'format'" class="sl-field">
          <label for="dv-message">{{ form.action === 'reject' ? 'Error message' : 'Warning message' }}</label>
          <input id="dv-message" class="sl-input" :value="form.message" data-testid="dv-message" @input="set({ message: value($event) })" />
        </div>
        <div class="sl-row">
          <div class="sl-field">
            <label for="dv-type">Type</label>
            <select id="dv-type" class="sl-select" :value="form.kind" data-testid="dv-type" @change="set({ kind: value($event) as ValidationKind, value1: '', value2: '' })">
              <option v-for="k in VALIDATION_KINDS" :key="k.key" :value="k.key">{{ k.label }}</option>
            </select>
          </div>
          <div v-if="numeric" class="sl-field">
            <label for="dv-criteria">Criteria</label>
            <select id="dv-criteria" class="sl-select" :value="form.criteria" data-testid="dv-criteria" @change="set({ criteria: value($event) as ValidationCriteria })">
              <option v-for="c in VALIDATION_CRITERIA" :key="c.key" :value="c.key">{{ c.label }}</option>
            </select>
          </div>
          <div v-if="form.kind === 'text'" class="sl-field">
            <label for="dv-tcriteria">Criteria</label>
            <select id="dv-tcriteria" class="sl-select" :value="form.textCriteria" data-testid="dv-text-criteria" @change="set({ textCriteria: value($event) as TextCriteria })">
              <option v-for="c in TEXT_CRITERIA" :key="c.key" :value="c.key">{{ c.label }}</option>
            </select>
          </div>
        </div>
        <div v-if="numeric" class="sl-field">
          <label>Value</label>
          <div class="sl-row" :style="{ gridTemplateColumns: two ? '1fr 1fr' : '1fr' }">
            <input class="sl-input" :type="valueType" :step="form.kind === 'time' ? 1 : undefined" :value="form.value1" data-testid="dv-value1" @input="set({ value1: value($event) })" />
            <input
              v-if="two"
              class="sl-input"
              :type="valueType"
              :step="form.kind === 'time' ? 1 : undefined"
              :value="form.value2"
              data-testid="dv-value2"
              @input="set({ value2: value($event) })"
            />
          </div>
        </div>
        <div v-if="form.kind === 'text' && form.textCriteria !== 'email' && form.textCriteria !== 'url'" class="sl-field">
          <label for="dv-text">Value</label>
          <input id="dv-text" class="sl-input" :value="form.value1" data-testid="dv-value1" @input="set({ value1: value($event) })" />
        </div>
        <template v-if="form.kind === 'list'">
          <div class="sl-field">
            <label for="dv-list">Items (comma separated) or a range like =$D$1:$D$5</label>
            <input id="dv-list" class="sl-input" :value="form.value1" data-testid="dv-value1" @input="set({ value1: value($event) })" />
          </div>
          <Switch label="Create list as a dropdown" :checked="form.dropdown" test-id="dv-dropdown" @change="set({ dropdown: $event })" />
        </template>
        <div v-if="form.kind === 'formula'" class="sl-field">
          <label for="dv-formula">Formula (TRUE when the value is valid)</label>
          <input id="dv-formula" class="sl-input" :value="form.value1" placeholder="=A1>0" spellcheck="false" data-testid="dv-value1" @input="set({ value1: value($event) })" />
        </div>
        <Switch v-if="form.kind !== 'empty' && form.kind !== 'notEmpty'" label="Allow blank" :checked="form.allowBlank" test-id="dv-blank" @change="set({ allowBlank: $event })" />
        <div v-if="form.action === 'format'" class="sl-field">
          <div class="label">Formatting Options</div>
          <div :style="{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }">
            <button v-for="which in STYLE_COLORS" :key="which" type="button" class="sl-select-btn" :data-testid="`dv-style-${which}`" @click="pickColor(which, $event)">
              <span :style="{ width: '16px', height: '16px', borderRadius: '2px', border: '1px solid #ccc', background: form.style[which] ?? 'transparent' }" />{{ which === 'color' ? 'Color' : 'Background color' }}
            </button>
            <Switch label="Bold" :checked="!!form.style.bold" test-id="dv-style-bold" @change="set({ style: { ...form.style, bold: $event } })" />
          </div>
        </div>
        <div v-if="error" class="sl-error" role="alert">{{ error }}</div>
        <button type="button" class="sl-button wide" data-testid="dv-save" @click="save">Save</button>
        <button type="button" class="sl-button light wide" data-testid="dv-back" @click="back">Back</button>
      </div>
      <ColorPicker v-if="colorPick" :anchor="colorPick.rect" :value="form.style[colorPick.which] ?? null" live @change="picked" @close="colorPick = null" />
    </template>
  </Modal>
</template>
