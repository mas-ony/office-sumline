# @sumline/core

The engine of Sumline, an Excel (XLSX) editor for desktop and mobile that runs in the browser. It
holds everything the React and Vue interfaces share:

- the workbook model and its commands (`SpreadsheetDocument`), with undo history
- a formula engine with 476 functions, including dynamic arrays and defined names
- XLSX and XLSM reading and writing (the saved file patches the opened package, so what the
  editor does not change is kept), XLS and XLSB readers, CSV and TSV
- the canvas grid with conditional formats, charts drawn from live data, pictures and shapes
- print layout and a PDF writer
- the page API (`SpreadsheetApi`, `installPageApi`) and autosave to IndexedDB
- the interface's state and actions without a framework (`SumlineUI`), which the interfaces
  render

To put the editor in an application, install `@sumline/react` or `@sumline/vue`. Both depend on
this package and re-export its option types. Use this package directly for the page API, or to
write an interface for another framework.

## Install

```bash
npm install @sumline/core
```

## Page API

Each editor has a page API object, which the components pass to `onReady` (`@ready` in Vue). It
follows the workbook the editor shows. `installPageApi(api)` puts it on the page as
`window.spreadsheetEditor` and returns the function that removes it again.

```js
import { installPageApi } from '@sumline/core';

// onReady={(api) => installPageApi(api)}, then from anywhere on the page:
const ed = window.spreadsheetEditor;
ed.setData([['Item', 'Price'], ['Pen', 1.5]], 'A1');
ed.setValue('C2', '=B2*2');
ed.on('change', () => console.log(ed.getValue('C2')));
ed.orderBy(1, true);
const values = ed.getData('Sheet1');
ed.download('xlsx');
```

`ed.doc` gives access to the native document model (`SpreadsheetDocument`) and `ed.grid` to the
grid.

## Entry points

| Import | |
| --- | --- |
| `@sumline/core` | The workbook model, file reading and writing, the page API, autosave and the interface state |
| `@sumline/core/views` | What the React and Vue interfaces are built from (commands, menu data, dialog helpers). It follows those packages and is not a stable API of its own |
| `@sumline/core/style.css` | The fonts, the grid and the interface |

The `style.css` of `@sumline/react` and `@sumline/vue` imports this package's. The stylesheet
imports the fonts from the `@fontsource` packages, so the application's bundler has to handle CSS
`@import` and font files (Vite does by default).

## License

MIT. The bundled icons come from `@material-design-icons/svg` under the Apache License 2.0
(`dist/THIRD-PARTY-LICENSES.md`). The fonts (Carlito, Arimo, Tinos and Cousine, from their
`@fontsource` packages) are under the SIL Open Font License 1.1.
