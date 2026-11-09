/**
 * What the official interfaces (@sumline/react and @sumline/vue) are built from: the public API
 * plus the interface's commands and menu data, and the engine pieces its dialogs use. It follows
 * those packages and is not a stable API of its own.
 */
export * from './index';

// Engine pieces the dialogs, pop-ups and the footer row use.
export {
  absRef, chartBounds, chartFromRange, type ChartLayout, chartLines, guessLayout, layoutOfChart, localRef,
  readChartData, resolveRef, seriesForLayout, trimRange,
} from './core/chartData';
export { type BorderKind, type SearchHit } from './core/document';
export { editText, formatValue } from './core/format/display';
export { parseInput } from './core/format/input';
export { ArrayValue, RefValue } from './core/formula/runtime';
export { tokenize } from './core/formula/tokenizer';
export { SheetLayout } from './core/layout';
export { rgbSpec } from './core/model/color';
export {
  type ChartSeries, type ChartType, DEFAULT_SHAPE_FILL, DEFAULT_SHAPE_LINE, type LineDash, type SheetChart,
  type Side, type Trendline,
} from './core/model/drawing';
export type { ColumnFilter, DataValidation, FilterOperator, ValidationFormatStyle } from './core/model/sheet';
export type { BorderLine, BorderLineStyle } from './core/model/style';
export { cellName, parseRange, quoteSheetName, type Range, rangeName, rangesIntersect } from './core/refs';
export { describeValidation } from './core/validation';
export { CellError, type Scalar } from './core/values';
export { type EditState, REF_COLORS } from './grid/GridView';
export { type GalleryItem, LINE_PRESETS, SHAPE_GALLERY, shapeIconSvg } from './grid/shapes';
export type { ViewTransform } from './grid/viewport';
export { canvasMeasure, loadPrintFonts, printPages, rasterize, saveFile } from './print/browser';
export {
  buildPages, DEFAULT_PRINT_OPTIONS, MARGINS, type Measure, type Page, PAPER_SIZES, type PaperSize,
  type PrintOptions,
} from './print/pages';
export { pagesToPdf } from './print/pdf';
export { pageSvg } from './print/svg';

// The interface.
export {
  CHART_SIDES, type ChartConfig, chartSettingsOf, chartSettingsPatch, type ChartStyle, insertChart, LINE_DASHES, numText, parseNum,
  relayoutChart, type SeriesStyle, TRENDLINES, withSeriesStyle,
} from './ui/chartDialogs';
export { CHART_TYPE_INFO, chartTypeInfo, type ChartTypeInfo } from './ui/chartTypes';
export * as cmd from './ui/commands';
export {
  computeDocState, CUSTOM_FORMATS, DOC_STATE_EVENTS, type DocState, FORMAT_PRESETS, type FormatState,
  MENU_FONT_SIZES, TOOLBAR_FONTS, TOOLBAR_SIZES, ZOOM_LEVELS,
} from './ui/docState';
export { isTouchDevice, matchesMedia, PHONE_QUERY, watchMedia } from './ui/dom';
export {
  applyCellEditor, EDITOR_TYPES, type EditorType, FORMAT_CATEGORIES, type FormatCategory, maskCode, maskFromCode,
  type MaskState,
} from './ui/formatMask';
export {
  analyzeFormula, applySuggestion, type AssistView, FormulaAssistState, type FormulaContext, type FormulaSegment,
  formulaSegments, highlightedArg,
} from './ui/formulaAssist';
export { exampleFor, functionCatalog, functionCategories, type FunctionInfo } from './ui/functionCatalog';
export { type IconName, iconSvg } from './ui/icons';
export { type MenuAnchor, type MenuItem, placeMenu } from './ui/menuItems';
export { hexToRgb, normalizeHex, PALETTE_ROWS, rgbToHex } from './ui/palette';
export { PRINT_CHOICES, PRINT_SWITCHES, type PrintChoice, PT_TO_PX } from './ui/printDialog';
export {
  blankValidationForm, ruleFromForm, TEXT_CRITERIA, type TextCriteria, type ValidationAction, VALIDATION_CRITERIA, VALIDATION_KINDS,
  type ValidationCriteria, type ValidationForm, type ValidationKind, validationFormOf,
} from './ui/validationForm';
export {
  ALIGN_ICONS, BORDER_KINDS, borderStyleFor, fitToolbar, isDivider, LINE_KINDS, type LineKind, TOOLBAR_ITEMS, type ToolKey,
  VALIGN_ICONS,
} from './ui/toolbar';
export { cellMenu, columnMenu, MOD, objectMenu, rowMenu, tabMenu, topMenus } from './ui/menus';
