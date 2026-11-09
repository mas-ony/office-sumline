<script setup lang="ts">
/** About the editor. */
import Modal from '../Modal.vue';

const emit = defineEmits<{ close: [] }>();
</script>

<template>
  <Modal title="About" icon="info" :width="420" test-id="about-dialog" @close="emit('close')">
    <p :style="{ marginTop: '0px' }"><b>Sumline</b></p>
    <p>Open, edit and save Excel workbooks (XLSX, XLSM, XLSB, XLS) and CSV files in the browser, on desktop and on phones. Files never leave your device.</p>
    <p class="sl-hint" :style="{ marginBottom: '0px' }">Formulas, charts, shapes, comments, notes, filters, data validation and conditional formatting are kept in the saved XLSX file.</p>
  </Modal>
</template>
