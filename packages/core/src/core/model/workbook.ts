import { defaultTheme, type Theme } from './color';
import { Sheet } from './sheet';
import { defaultStyle, StyleTable, type CellStyle } from './style';
import type { OoxmlPackage } from '../xlsx/package';
import type { XmlElement } from '../xml/xml';

export interface DefinedName {
  name: string;
  /** Formula text without '=' (e.g. Sheet1!$A$1:$B$5). */
  formula: string;
  /** Index of the sheet the name is local to (localSheetId), if any. */
  localSheet?: number;
  hidden?: boolean;
  comment?: string;
  /** Other attributes of the source element. */
  extra?: Record<string, string>;
}

export interface Person {
  id: string;
  displayName: string;
  userId?: string;
  providerId?: string;
}

/** The parsed source package, kept so that unknown parts survive a save. */
export interface SourceInfo {
  pkg: OoxmlPackage;
  workbookPath: string;
  workbookXml: XmlElement;
  stylesXml?: XmlElement;
  /** Index of each imported style in the source cellXfs. */
  dxfsXml?: XmlElement;
  extraNamespaces: Record<string, string>;
}

export class Workbook {
  sheets: Sheet[] = [];
  styles: StyleTable;
  theme: Theme = defaultTheme();
  names: DefinedName[] = [];
  persons: Person[] = [];
  date1904 = false;
  activeSheetIndex = 0;
  /** Title, author and similar document properties. */
  props: Record<string, string> = {};
  source?: SourceInfo;
  /** Set when the file was produced by CSV import (save offers CSV again). */
  fromCsv = false;
  /**
   * Differential formats (`<dxf>`) used by conditional formats, in file order. Replaced (not
   * mutated) when the editor adds one, so caches keyed by the array stay correct.
   */
  dxfs: XmlElement[] = [];
  /**
   * Cell references used by charts, per chart part (`xl/charts/chart1.xml` -> its `<c:f>` texts in
   * document order). Kept here so structural edits update them; written back on save.
   */
  chartRefs = new Map<string, string[]>();

  constructor(base: CellStyle = defaultStyle()) {
    this.styles = new StyleTable(base);
  }

  static blank(sheetName = 'Sheet1'): Workbook {
    const wb = new Workbook();
    wb.addSheet(sheetName);
    return wb;
  }

  get activeSheet(): Sheet {
    return this.sheets[this.activeSheetIndex] ?? this.sheets[0];
  }

  sheetByName(name: string): Sheet | undefined {
    const lower = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === lower);
  }

  sheetByUid(uid: number): Sheet | undefined {
    return this.sheets.find((s) => s.uid === uid);
  }

  indexOf(sheet: Sheet): number {
    return this.sheets.indexOf(sheet);
  }

  nextSheetId(): number {
    return this.sheets.reduce((m, s) => Math.max(m, s.sheetId), 0) + 1;
  }

  addSheet(name: string, index = this.sheets.length): Sheet {
    const sheet = new Sheet(name, this.nextSheetId());
    this.sheets.splice(index, 0, sheet);
    return sheet;
  }

  /** A sheet name that does not exist yet: base + number. */
  uniqueSheetName(base = 'Sheet'): string {
    let n = 1;
    while (this.sheetByName(`${base}${n}`)) n++;
    return `${base}${n}`;
  }

  visibleSheets(): Sheet[] {
    return this.sheets.filter((s) => s.state === 'visible');
  }

  nameLookup(name: string, sheetIndex: number): DefinedName | undefined {
    const upper = name.toUpperCase();
    return (
      this.names.find((n) => n.localSheet === sheetIndex && n.name.toUpperCase() === upper) ??
      this.names.find((n) => n.localSheet === undefined && n.name.toUpperCase() === upper)
    );
  }
}

/** Excel's rules for sheet names. Returns an error message or null when valid. */
export function validateSheetName(name: string, wb: Workbook, current?: Sheet): string | null {
  const n = name.trim();
  if (!n) return 'The sheet name cannot be empty.';
  if (n.length > 31) return 'Sheet names can have at most 31 characters.';
  if (/[\\/?*[\]:]/.test(n)) return 'Sheet names cannot contain \\ / ? * [ ] or :';
  if (n.startsWith("'") || n.endsWith("'")) return 'Sheet names cannot start or end with an apostrophe.';
  if (n.toLowerCase() === 'history') return '"History" is a reserved name.';
  const other = wb.sheetByName(n);
  if (other && other !== current) return 'This name already exists. Please enter another name.';
  return null;
}
