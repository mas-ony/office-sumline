# Sumline features

What Sumline does, area by area. [The README](../README.md) covers running it, the components
and the tests.

## Platforms and files

| Feature | Details |
| --- | --- |
| Browsers | Tested in Chrome on desktop and on Android |
| Phones | Touch layout: menus and context menus as bottom sheets, long press for the cell menu, selection handles, pinch zoom, a key bar above the keyboard, full-screen dialogs, charts inserted under the data |
| Plain-HTTP LAN address | No secure-context API is required; clipboard, file pickers, sharing and offline caching are used only where available |
| Start screen | A drop zone, Create new spreadsheet and Upload file; files also open from `?src=` and the system's "Open with" for the installed app |
| Open files | XLSX, XLSM, XLS, XLSB, CSV, TSV and TXT |
| XLS (Excel 97-2003) | Own BIFF8 reader: values, formulas, styles, merges, sizes, frozen panes, names, links, notes; Excel 5.0/95 files are refused |
| XLSB | Own BIFF12 reader: values, formulas, styles, merges, sizes, frozen panes, names, links |
| CSV | Also TSV and TXT; delimiter and encoding detection |
| New spreadsheet | 13 x 50 cells, 100 px columns, 26 px rows; the grid grows beyond them |
| Save (toolbar, Ctrl+S) | Downloads `<name>.xlsx` (`.xlsm` kept for macro workbooks), or writes back to the opened file where the File System Access API exists |
| Download | XLSX, CSV, TSV and PDF |
| PDF | Own PDF writer: vector text in the standard PDF fonts (metric-compatible fonts are used for layout), pictures and charts embedded |
| Print | The PDF pages through the browser's print dialog (this also works on Android, where PDFs in frames do not display) |
| Share | Web Share API where available |
| Rename | File › Rename |
| Autosave and recovery | A copy in IndexedDB, offered back with a restore banner |
| Installable app, offline | PWA with file handlers for XLSX, XLSM, XLSB, XLS, CSV (offline caching needs HTTPS or localhost) |
| XLSX output | The opened package is patched: unknown parts and elements, threaded comments, cached values, styles and charts are kept |

## Grid and editing

| Feature | Details |
| --- | --- |
| Selection and editing | Selection, keyboard navigation and in-cell editing, with Excel conventions such as Enter after a run of Tabs returning to the first column |
| Text alignment | Excel's general alignment (text left, numbers right), as stored in XLSX |
| Formula engine | 476 functions, dynamic arrays, names; recalculated as you type |
| Formula suggestions | In the formula bar and while typing in a cell: signature, description, example, argument hints |
| Name box | Defined names: add, edit and delete |
| Context menus | Cell, row, column, tab and object menus; bottom sheets on phones |
| Paste Special | Full, values, values and formulas, style |
| Cells, rows and columns | Insert and delete rows and columns, merge, borders, colours, fonts, sizes, alignment, rotation |
| Number formats | Text, Number, Percentage, Currency, Date, Hour, Date and hour, Custom; a Format dialog with Mask and Editor tabs |
| Sorting | A to Z, Z to A |
| Filters | Auto filter with drop-down buttons: sort, value lists, conditions; saved to XLSX |
| Data validations | Rules list and form; warning, reject or format actions; drop-down lists; saved to XLSX |
| Freeze panes | Rows and columns (none, 1, 2, 3) |
| Footer row | View › Toggle footer row: totals of every used column under the grid (not stored in XLSX) |
| Column titles | Rename this column: a title in place of the column letter (not stored in XLSX) |
| Search and replace | Find values and replace them |
| Comments | Threads with replies, stored as threaded comments |
| Notes | An inline note editor; stored as notes |
| Links | Insert and edit hyperlinks |
| Gridlines and zoom | In the View menu |
| Images | Upload, move, resize, stacking order; saved to XLSX |
| Conditional formatting | Rules from files are shown and kept; the validation "format" action creates rules |
| Undo and redo | Including charts, shapes and structure changes |

## Charts

| Feature | Details |
| --- | --- |
| Insert chart | A gallery of 23 types |
| Placement | Next to the data (under it on phones), so the data stays visible |
| Look | Seven-colour palette, round legend markers, light gridlines |
| Settings | Configuration, style and series: type, orientation, range, labels, series, headers; titles and positions, legend, gridlines, value labels, axis titles, Y min/max; colour, line width and style, trendline, markers, secondary axis |
| Hover information | The browser's tooltip for each point |
| Move, resize, order, edit, delete | Drag handles (large on touch screens), arrow keys, object menu |
| Live data | Charts follow their data |
| Saved to XLSX | Classic chart parts for 17 types and chartex parts for histogram, Pareto, box and whisker, waterfall, funnel and treemap |
| Charts in opened files | Shown and editable; unedited charts are written back unchanged |
| Printing | Charts are printed with the sheet |

## Shapes

| Feature | Details |
| --- | --- |
| Gallery | 8 categories, 125 DrawingML preset shapes |
| Default look | Fill `#4F81BD`, border `#385D8A` 3 px, 400 x 300 px at the active cell |
| Shape Settings | Background, transparent, border colour and width, font colour, and a text field |
| Saved to XLSX | Preset shapes and connectors (arrowheads, dashes); shapes from files become editable |

## Printing and PDF

| Feature | Details |
| --- | --- |
| Options | Scope, page size, orientation, scale, margins, page order, horizontal and vertical alignment; A4, landscape and centred by default |
| Switches | Grid lines, notes, row and column identifiers, page number, spreadsheet title, worksheet title, date, time, repeat frozen rows/columns |
| Pictures, charts and shapes | Printed (can be switched off) |
| Printed area | The used range with its objects, or the file's print area |
| Preview | Own preview with page count, zoom, download and print; on phones too |

## Programming interface

The editor is embedded with the `SumlineEditor` component of `@sumline/react` or `@sumline/vue`,
or with their `mount()`. Its page API (`packages/core/src/api/api.ts`), which the apps install as
`window.spreadsheetEditor`, addresses cells with A1 references (`'B2'`, `'Sheet2!A1:C3'`).
`ed.doc` gives access to the native document model (`packages/core/src/core/document.ts`) and
`ed.grid` to the grid (`packages/core/src/grid/GridView.ts`).

| Feature | Details |
| --- | --- |
| Components | `SumlineEditor` for React and for Vue, each written for its framework, with the same props and events, and `mount(element, options)` for other pages; options for the start screen, a document to open, view-only mode, autosave and replaced file actions; styles scoped to the editor, several editors on a page |
| Cell values | `getValue`, `getText`, `setValue`, `getData`, `setData(data, at)` |
| Selection | `getSelected()` returns the selected range as an A1 reference; `select(ref)` |
| Undo and redo | `undo`, `redo` |
| Styles and number formats | `setStyle(ref, property, value)` with CSS properties (weight, style, decoration, colours, font, size, alignment, wrapping); `setFormat(ref, code)` |
| Rows and columns | `insertRow(count, row, before)`, `insertColumn(count, col, before)`, `deleteRow(row, count)`, `deleteColumn(col, count)` |
| Sorting | `orderBy(col, descending)` sorts the used range |
| Freeze panes | `setFreezeRows(n)`, `setFreezeColumns(n)`; 0 unfreezes |
| Worksheets | `createWorksheet`, `renameWorksheet`, `deleteWorksheet`, `worksheets` |
| Files | `open(file or handle, name)`, `download('xlsx' / 'csv' / 'tsv' / 'pdf')` (PDF opens the print preview), `save()`; the `document` option |
| Events | `on('change' / 'selection' / 'activeSheet' / 'sheets' / 'file' / 'open', fn)` returns a function that unsubscribes; the components' `onReady` (`@ready`) |

## Limitations

- Excel 5.0/95 `.xls` files are not supported; XLS and XLSB files are saved as XLSX.
- Formulas that use tokens the binary readers cannot express (external workbook references, data
  tables) keep their last calculated value.
- Functions beyond the 476 implemented are not available.
- Charts from files that use features the chart model does not cover (3-D, combination charts
  with different types) are drawn approximately; they are written back unchanged unless edited.
- Conditional formatting rules from files are shown and kept, but there is no general rule editor
  (the validation "format" action creates highlighting rules).
- Text that the standard PDF fonts cannot show (for example CJK) is embedded in the PDF as
  pictures.
- Scripts cannot reach the clipboard, the footer row or the dialogs; those are used from the
  menus, toolbar and keyboard.
- Move to trash, History, Checkbox and Alternate colors appear in the menus but are disabled.
