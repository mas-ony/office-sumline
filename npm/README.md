# Sumline

Sumline is an Excel (XLSX) editor for desktop and mobile that runs in the browser, with formulas,
charts, shapes, sorting and filters, and printing to PDF, on its own canvas grid, formula engine
and XLSX reader and writer. On phones and tablets it has a touch layout, and it also works over a
plain HTTP address on the local network.

This package is the built app: static files to put on a web server. It has no dependencies and
nothing to build. The React and Vue packages, for putting the editor inside an app instead, are in
the [repository](https://github.com/mas-ony/office-sumline) (`packages/react` and `packages/vue`) and
are not on npm yet.

## Install

```bash
npm install sumline
```

## Host

The app is the package's `dist` folder: `index.html`, `assets/`, the service worker, the manifest
and the icons. Its paths are relative, so it can be served from any path. Copy the folder into a
site, or serve it as it is:

```bash
cp -r node_modules/sumline/dist public/sheets               # then open /sheets/
python3 -m http.server -d node_modules/sumline/dist 8080    # then open http://localhost:8080
```

Serve it over HTTP: browsers do not run its scripts from a `file://` address.

Over HTTPS (or on `localhost`) the app also works offline and can be installed, and an installed
app opens spreadsheet files from the system's "Open with". Over plain HTTP, such as a local
network address, the clipboard API, file pickers, the share sheet and offline use are not
available, and the app works without them.

## Files

It opens XLSX and XLSM (saving patches the opened package, so everything the editor does not
change is kept), XLTX and XLTM templates, XLS (Excel 97-2003), XLSB, CSV and TSV, from its start
screen, by drag and drop or with `?src=`. It saves XLSX, back to the opened file where the browser
allows it (XLS and XLSB files are saved as XLSX), downloads XLSX, CSV, TSV and PDF, and prints
through the browser. Files are opened, edited and saved in the browser; the app sends nothing to a
server.

## Address options

| Option | |
| --- | --- |
| `?src=<url>` | Opens a workbook from a URL the page can fetch (its own server, or one that allows it with CORS) |
| `?new=1` | Starts with a blank workbook, without offering to recover unsaved work |
| `?readonly=1` or `?mode=view` | Opens workbooks view-only |

## Page API

Once a workbook is open, the app puts a programmatic interface on `window.spreadsheetEditor`. A
page that shows the app in a same-origin `<iframe>` reaches it as
`iframe.contentWindow.spreadsheetEditor`. It follows the workbook the app shows, also after
another workbook opens.

```js
const ed = window.spreadsheetEditor;
ed.setData([['Item', 'Price'], ['Pen', 1.5]], 'A1');
ed.setValue('C2', '=B2*2');
ed.setStyle('A1:B1', 'font-weight', 'bold');
ed.setFormat('B2', '"$" #,##0.00');
ed.on('change', () => console.log(ed.getValue('C2')));
ed.setFreezeRows(1);
ed.download('xlsx');   // also 'csv', 'tsv'; 'pdf' opens the print preview
```

It also has `getText`, `getData`, `insertRow`, `insertColumn`, `deleteRow`, `deleteColumn`,
`orderBy`, `setFreezeColumns`, `select`, `getSelected`, `undo`, `redo`, `worksheets`,
`createWorksheet`, `renameWorksheet`, `deleteWorksheet`, `open` and `save`. Its events are
`change`, `selection`, `activeSheet`, `sheets`, `file` (the workbook was renamed or saved) and
`open` (another workbook opened); `on` returns a function that removes the listener.

## Storage

The app keeps unsaved work in IndexedDB (the database `sumline`) and offers to recover it on the
next visit. It keeps its settings in localStorage under `sumline-` keys, and its offline copy in
the cache `sumline-v1`. Copies served from the same origin share them.

## Browsers

Tested in Chrome on desktop and on Android.

## License

MIT. The app bundles third-party packages, icons and fonts: `dist/THIRD-PARTY-LICENSES.md` lists
them with their licences (MIT; Apache-2.0 for the Material Design icons and `bessel`; the SIL Open
Font License 1.1 for the Arimo, Carlito, Cousine and Tinos fonts).
