<script setup lang="ts">
/** The start screen: a drop zone with "Create new spreadsheet" and "Upload file". */
import type { UIState } from '@sumline/core/views';
import { ref } from 'vue';
import Icon from './Icon.vue';

defineProps<{ restore: UIState['restore'] }>();
const emit = defineEmits<{ new: []; open: []; file: [file: File]; restore: []; discard: [] }>();

const dragging = ref(false);
const over = (e: DragEvent) => {
  if ([...(e.dataTransfer?.types ?? [])].includes('Files')) {
    e.preventDefault();
    dragging.value = true;
  }
};
const leave = (e: DragEvent) => {
  if (!(e.currentTarget as HTMLElement).contains(e.relatedTarget as Node)) dragging.value = false;
};
const drop = (e: DragEvent) => {
  e.preventDefault();
  e.stopPropagation();
  dragging.value = false;
  const file = e.dataTransfer?.files[0];
  if (file) emit('file', file);
};
const when = (savedAt: number) => new Date(savedAt).toLocaleString();
</script>

<template>
  <main class="sl-start" data-testid="start-screen">
    <div class="sl-start-head">
      <h1>Open and edit Excel spreadsheets in your browser</h1>
      <p>Open an XLSX, edit it the way you would in Excel and save it back. Formulas, charts, conditional formatting, data validation and multiple worksheets, on desktop and on your phone. Your file stays on your device.</p>
    </div>
    <div :class="['sl-dropzone', { dragging }]" data-testid="dropzone" @dragover="over" @dragleave="leave" @drop="drop">
      <div>
        <Icon name="upload_file" />
        <p><strong>Drop your Excel file here</strong></p>
        <p class="caption">XLSX · XLS · XLSM · XLSB · CSV</p>
        <div class="sl-start-actions">
          <button type="button" class="sl-btn primary" data-testid="start-new" @click="emit('new')">Create new spreadsheet</button>
          <button type="button" class="sl-btn" data-testid="start-upload" @click="emit('open')">Upload file</button>
        </div>
      </div>
    </div>
    <div v-if="restore" class="sl-start-recent" role="alert" data-testid="restore-banner">
      <div class="sl-validation-card">
        <div>
          <div><b>Unsaved changes found</b> in “{{ restore.fileName }}”</div>
          <div class="meta">{{ when(restore.savedAt) }}</div>
        </div>
        <div :style="{ display: 'flex', gap: '6px' }">
          <button type="button" class="sl-button" data-testid="restore-recover" @click="emit('restore')">Recover</button>
          <button type="button" class="sl-button light" data-testid="restore-discard" @click="emit('discard')">Discard</button>
        </div>
      </div>
    </div>
    <div class="sl-start-notes">
      <div>
        <h3>1. Upload or start fresh</h3>
        <p>Drag an XLSX, XLS, XLSM, XLSB or CSV file here, or start with a blank workbook.</p>
      </div>
      <div>
        <h3>2. Edit and format</h3>
        <p>Type formulas, sort and filter, add charts, shapes, comments and data validation.</p>
      </div>
      <div>
        <h3>3. Download or print</h3>
        <p>Save as XLSX, CSV, TSV or PDF, or print. Nothing is uploaded anywhere.</p>
      </div>
    </div>
  </main>
</template>
