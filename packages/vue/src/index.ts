/** @sumline/vue: the Sumline spreadsheet editor as a Vue component. */
export { default as SumlineEditor } from './SumlineEditor.vue';
export { mount, type MountedEditor } from './mount';
export type {
  DocumentInput,
  DocumentSource,
  DownloadFormat,
  FileActions,
  SpreadsheetApi,
  SumlineOptions,
} from '@sumline/core';
