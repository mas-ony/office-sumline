<script setup lang="ts">
/**
 * The editor of one workbook: menu bar, worksheet tabs, toolbar, formula bar and grid, with the
 * context menus, dialogs and cell pop-ups.
 */
import { topMenus, type UIState } from '@sumline/core/views';
import { computed } from 'vue';
import CommentsBox from './CommentsBox.vue';
import { useUI } from './context';
import FilterPopup from './FilterPopup.vue';
import FooterRow from './FooterRow.vue';
import FormulaAssist from './FormulaAssist.vue';
import FormulaBar from './FormulaBar.vue';
import GridHost from './GridHost.vue';
import KeyBar from './KeyBar.vue';
import Menu from './Menu.vue';
import NoteEditor from './NoteEditor.vue';
import SheetTabs from './SheetTabs.vue';
import Toolbar from './Toolbar.vue';
import AboutDialog from './dialogs/AboutDialog.vue';
import ChartInsertDialog from './dialogs/ChartInsertDialog.vue';
import ChartSettings from './dialogs/ChartSettings.vue';
import ConfirmDialog from './dialogs/ConfirmDialog.vue';
import FormatDialog from './dialogs/FormatDialog.vue';
import FunctionsDialog from './dialogs/FunctionsDialog.vue';
import LinkDialog from './dialogs/LinkDialog.vue';
import NameDialog from './dialogs/NameDialog.vue';
import PrintDialog from './dialogs/PrintDialog.vue';
import PromptDialog from './dialogs/PromptDialog.vue';
import SearchDialog from './dialogs/SearchDialog.vue';
import ShapeSettings from './dialogs/ShapeSettings.vue';
import ShapesDialog from './dialogs/ShapesDialog.vue';
import ShortcutsDialog from './dialogs/ShortcutsDialog.vue';
import ValidationsDialog from './dialogs/ValidationsDialog.vue';

const props = defineProps<{ state: UIState }>();

const ctx = useUI();
const { ui } = ctx.value;
const close = () => ui.closeDialog();
const menus = computed(() => topMenus(ctx.value));
const rect = (e: Event) => (e.currentTarget as HTMLElement).getBoundingClientRect();
const renameColumn = (v: string | null) => {
  const d = props.state.dialog;
  if (v !== null && d?.kind === 'column-title') ctx.value.doc.setColumnTitle(d.col, v.trim() || null);
  close();
};
</script>

<template>
  <div class="sl-editor" data-testid="editor">
    <div class="sl-menubar" role="menubar" aria-label="Menu">
      <div class="sl-menubar-items">
        <button
          v-for="(m, i) in menus"
          :key="m.title"
          type="button"
          role="menuitem"
          aria-haspopup="menu"
          :aria-expanded="state.topMenu === i"
          class="sl-menu-title"
          :data-testid="`menu-${m.title.toLowerCase()}`"
          @mousedown.prevent
          @click="ui.clickTopMenu(i, rect($event))"
          @mouseenter="ui.hoverTopMenu(i, rect($event))"
        >
          {{ m.title }}
        </button>
      </div>
      <div class="sl-filename" :title="ctx.state.fileName">
        <span v-if="ctx.state.modified" class="sl-dirty" title="Unsaved changes" />
        <span :style="{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }" data-testid="file-name">{{ ctx.state.fileName }}</span>
        <span v-if="ctx.state.readOnly" class="sl-hint">(view only)</span>
      </div>
    </div>
    <SheetTabs />
    <div class="sl-sheet-frame">
      <Toolbar />
      <FormulaBar />
      <GridHost />
      <FooterRow v-if="ctx.doc.sheet.footer" />
    </div>
  </div>
  <FormulaAssist />
  <KeyBar v-if="ctx.phone && state.edit.active" />
  <Menu v-if="state.menu" :items="state.menu.items" :at="state.menu.at" :sheet="ctx.phone" :title="state.menu.title" :test-id="state.menu.testId" @close="ui.closeMenu" />
  <FormatDialog v-if="state.dialog?.kind === 'format'" :initial-tab="state.dialog.tab" @close="close" />
  <SearchDialog v-if="state.dialog?.kind === 'search'" @close="close" />
  <ValidationsDialog v-if="state.dialog?.kind === 'validations'" @close="close" />
  <NameDialog v-if="state.dialog?.kind === 'name-edit'" :name="state.dialog.name" @close="close" />
  <LinkDialog v-if="state.dialog?.kind === 'link'" @close="close" />
  <AboutDialog v-if="state.dialog?.kind === 'about'" @close="close" />
  <ShortcutsDialog v-if="state.dialog?.kind === 'shortcuts'" @close="close" />
  <FunctionsDialog v-if="state.dialog?.kind === 'functions'" @close="close" />
  <ChartInsertDialog v-if="state.dialog?.kind === 'chart-insert'" @close="close" />
  <ChartSettings v-if="state.dialog?.kind === 'chart-settings'" :id="state.dialog.id" @close="close" />
  <ShapesDialog v-if="state.dialog?.kind === 'shapes'" @close="close" />
  <ShapeSettings v-if="state.dialog?.kind === 'shape-settings'" :id="state.dialog.id" @close="close" />
  <PrintDialog v-if="state.dialog?.kind === 'print'" @close="close" />
  <PromptDialog
    v-if="state.dialog?.kind === 'column-title'"
    title="Rename this column"
    message="Column title (shown in the column header)"
    :value="ctx.doc.sheet.colTitles?.get(state.dialog.col) ?? ''"
    @answer="renameColumn"
  />
  <CommentsBox v-if="state.comments" :r="state.comments.r" :c="state.comments.c" :rect="state.comments.rect" @close="ui.closeComments" />
  <NoteEditor v-if="state.note" :r="state.note.r" :c="state.note.c" :rect="state.note.rect" @close="ui.closeNote" />
  <FilterPopup v-if="state.filter" :col="state.filter.col" :at="state.filter.at" @close="ui.closeFilter" />
  <ConfirmDialog
    v-if="state.pending?.kind === 'confirm'"
    :message="state.pending.message"
    :title="state.pending.opts?.title"
    :ok="state.pending.opts?.ok"
    :cancel="state.pending.opts?.cancel"
    :danger="state.pending.opts?.danger"
    @answer="ui.answer($event)"
  />
  <PromptDialog
    v-if="state.pending?.kind === 'prompt'"
    :message="state.pending.message"
    :value="state.pending.value"
    :title="state.pending.opts?.title"
    :ok="state.pending.opts?.ok"
    @answer="ui.answer($event)"
  />
  <div v-if="state.toast" class="sl-toast" role="status" data-testid="toast">
    <span>{{ state.toast }}</span>
    <button type="button" @click="ui.dismissToast">OK</button>
  </div>
</template>
