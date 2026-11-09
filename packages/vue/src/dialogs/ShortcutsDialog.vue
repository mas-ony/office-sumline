<script setup lang="ts">
/** The keyboard shortcuts. */
import { MOD } from '@sumline/core/views';
import Modal from '../Modal.vue';

const emit = defineEmits<{ close: [] }>();

const SHORTCUTS: [string, string][] = [
  ['Undo / Redo', `${MOD}+Z / ${MOD}+Y`],
  ['Cut / Copy / Paste', `${MOD}+X / ${MOD}+C / ${MOD}+V`],
  ['Find and replace', `${MOD}+F`],
  ['Save (download XLSX)', `${MOD}+S`],
  ['Print / PDF', `${MOD}+P`],
  ['Bold / Italic / Underline', `${MOD}+B / ${MOD}+I / ${MOD}+U`],
  ['Remove all format', `${MOD}+\\`],
  ['Select all', `${MOD}+A`],
  ['Edit the cell', 'F2 or double click (double tap on phones)'],
  ['New line in a cell', 'Alt+Enter'],
  ['Fill down / right', `${MOD}+D / ${MOD}+R`],
  ['Jump to the data edge', `${MOD}+Arrow keys`],
  ['Extend the selection', 'Shift+Arrow keys'],
  ['Absolute / relative reference', 'F4 while editing a formula'],
  ['AutoSum', 'Alt+='],
  ['Next / previous worksheet', `${MOD}+Page Down / ${MOD}+Page Up`],
];
</script>

<template>
  <Modal title="Keyboard shortcuts" icon="article" :width="520" test-id="shortcuts-dialog" @close="emit('close')">
    <table :style="{ width: '100%', borderCollapse: 'collapse', fontSize: '13.5px' }">
      <tbody>
        <tr v-for="[what, keys] in SHORTCUTS" :key="what" :style="{ borderBottom: '1px solid #eee' }">
          <td :style="{ padding: '7px 8px 7px 0' }">{{ what }}</td>
          <td :style="{ padding: '7px 0', color: '#555', textAlign: 'right' }">{{ keys }}</td>
        </tr>
      </tbody>
    </table>
  </Modal>
</template>
