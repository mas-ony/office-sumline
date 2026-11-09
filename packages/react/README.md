# @sumline/react

Sumline is an Excel (XLSX) editor for desktop and mobile that runs in the browser: a menu bar,
toolbar, formula bar and worksheet tabs, dialogs and context menus, on its own canvas grid, formula
engine and XLSX reader and writer, with charts, shapes, sorting and filters, data validation,
comments and notes. It also opens XLS, XLSB and CSV files, prints and writes PDF in the browser,
and has a touch layout on phones.

This package is the React interface. `@sumline/vue` is the same interface written in Vue, with the
same options.

## Install

```bash
npm install @sumline/react react react-dom
```

It needs React 19.

## Use

```tsx
import { SumlineEditor } from '@sumline/react';
import '@sumline/react/style.css';

export function Budget() {
  return (
    <SumlineEditor
      style={{ height: '80vh' }}
      document="/reports/budget.xlsx"
      author="Jordan Lee"
      onReady={(api) => api.on('change', () => console.log(api.getValue('B10')))}
    />
  );
}
```

Import `style.css` once: it brings the fonts and the editor's styles. The editor fills the height
of its element, so give it (or its container) a height.

`mount()` takes the same options, for pages that are not written in React:

```js
import { mount } from '@sumline/react';
import '@sumline/react/style.css';

const editor = mount(document.querySelector('#editor'), { startScreen: true });
const api = await editor.ready;
editor.update({ readOnly: true });
editor.destroy();
```

## Options

| Option | |
| --- | --- |
| `document` | A URL, a file handle (File System Access API: Save writes back to the file), a File, Blob, ArrayBuffer or bytes, or `{ url or data, name }`; opened again when it changes. Without it the editor starts with a blank workbook |
| `startScreen` | Start on the start screen (a drop zone, "Create new spreadsheet" and "Upload file") instead of a blank workbook; it also shows while `document` loads. Read at start |
| `readOnly` | Workbooks open view-only |
| `author` | The name on comments and notes |
| `storageKey` | Keep the recent colours and print options in localStorage under `<storageKey>-…` |
| `autosave`, `offerRecovery` | Keep unsaved changes in an IndexedDB database of this name; with `startScreen`, the start screen offers to recover them (unless `offerRecovery` is false) |
| `files` | Replace some of the file actions (New, Open, Save, Download, Print, Rename, Share) |
| `className`, `style` | For the editor's element |

The types are exported as `SumlineOptions`, `DocumentInput`, `FileActions` and so on.

## Events

`onReady(api)` runs once a workbook is open (the initial document, or the first one opened from the
start screen), with the editor's page API.

## Page API

The object passed to `onReady` drives the editor from code. It follows the workbook the editor
shows.

```js
api.setData([['Item', 'Price'], ['Pen', 1.5]], 'A1');
api.setValue('C2', '=B2*2');
api.setStyle('A1:B1', 'font-weight', 'bold');
api.setFormat('B2', '"$" #,##0.00');
api.on('change', () => console.log(api.getValue('C2')));
api.setFreezeRows(1);
api.download('xlsx');   // also 'csv', 'tsv'; 'pdf' opens the print preview
await api.open(blob, 'Budget.xlsx');
```

`installPageApi(api)` from `@sumline/core` also puts it on the page as
`window.spreadsheetEditor`.

## Page styles

The editor keeps its classes on its own element and on a layer it adds to the body for menus,
dialogs and pop-ups: `sl-` for the interface, `xgrid-` and `xchart-` for the grid. It changes
nothing on `<html>` or `<body>`, and several editors can share a page. In a window up to 720 pixels
wide it uses its phone layout.

## License

MIT
