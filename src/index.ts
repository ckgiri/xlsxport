/**
 * xlsxport — export JSON to a real Excel (.xlsx) file in one call.
 *
 *   const bytes = await toExcel(users)                    // Uint8Array
 *   await download(users, 'users.xlsx')                   // browser
 *   await writeExcel('users.xlsx', users)                 // Node: import from 'xlsxport/node'
 */

import { zip, zipSync, webDeflate, type DeflateRaw, type ZipEntry } from './zip';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/** A formula cell: `{ formula: 'SUM(B2:B10)' }` (a leading '=' is optional and removed). */
export interface FormulaCell {
  formula: string;
}

/** A hyperlink cell: `{ link: 'https://example.com', text: 'Open' }`. */
export interface LinkCell {
  link: string;
  text?: string;
}

export type CellValue = string | number | boolean | bigint | Date | null | undefined | FormulaCell | LinkCell | object;

export type Format =
  | 'text'
  | 'integer'
  | 'decimal'
  | 'percent'
  | 'currency'
  | 'date'
  | 'datetime'
  | 'time'
  | 'link'
  | (string & {});

export interface Column<T = any> {
  /** Property to read from each row object (supports 'a.b.c' paths). */
  key?: string;
  /** Header text. Defaults to the key. */
  header?: string;
  /** Compute the cell value from the row. Overrides `key`. */
  value?: (row: T, index: number) => CellValue;
  /** Column width in characters. Default: fit content (8-60). */
  width?: number;
  /**
   * Number/date format: 'integer' (1,234), 'decimal' (1,234.50), 'percent' (0.5 -> 50.00%),
   * 'currency', 'date', 'datetime', 'time', 'text', 'link', or any Excel format code like '0.0"kg"'.
   */
  format?: Format;
  /** Currency symbol for `format: 'currency'`. Default '$'. */
  currency?: string;
}

export interface SheetOptions<T = any> {
  /** Sheet tab name (max 31 chars; invalid characters are replaced). Default 'Sheet1'. */
  name?: string;
  /** Pick, order, rename and format columns. Strings are shorthand for `{ key }`. Default: every key found in the rows. */
  columns?: Array<Column<T> | string>;
  /** Write a header row. Default `true` for objects, and for arrays of arrays when `columns` is given; otherwise `false`. */
  header?: boolean;
  /** Turn 'firstName' into 'First Name' in auto-generated headers. Default `false`. */
  humanizeHeaders?: boolean;
  /** Keep the header visible while scrolling. Default `true` when there is a header. */
  freezeHeader?: boolean;
  /** Add filter/sort dropdowns to the header. Default `true` when there is a header. */
  autoFilter?: boolean;
  /** Bold header row. Default `true`. */
  boldHeader?: boolean;
  /** Format used for Date values at midnight. Default 'yyyy-mm-dd'. */
  dateFormat?: string;
  /** Format used for Date values with a time. Default 'yyyy-mm-dd hh:mm:ss'. */
  dateTimeFormat?: string;
  /** Write dates using UTC or local time. Default 'local'. */
  timezone?: 'local' | 'utc';
}

export interface Sheet<T = any> extends SheetOptions<T> {
  rows: readonly T[];
}

export type Rows = readonly Record<string, any>[] | readonly (readonly CellValue[])[];

/** A single array of rows, or `{ 'Sheet name': rows | { rows, ...options } }` for multiple sheets. */
export type WorkbookInput = Rows | Record<string, Rows | Sheet>;

export interface WorkbookOptions<T = any> extends SheetOptions<T> {
  /** Author stored in document properties. */
  creator?: string;
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
const MAX_CELL_TEXT = 32_767;
// Characters that are not allowed in XML 1.0 (incl. unpaired surrogates)
// eslint-disable-next-line no-control-regex
const INVALID_XML = /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

function esc(value: string): string {
  return value
    .replace(INVALID_XML, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** 0 -> 'A', 25 -> 'Z', 26 -> 'AA' */
export function columnName(index: number): string {
  let name = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function humanize(key: string): string {
  return key
    .replace(/[_\-.]+/g, ' ')
    .replace(/([a-z\d])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/(^|\s)\S/g, (c) => c.toUpperCase());
}

function getPath(obj: any, path: string): unknown {
  if (obj == null) return undefined;
  if (path in Object(obj)) return obj[path];
  let cur = obj;
  for (const part of path.split('.')) {
    if (cur == null) return undefined;
    cur = cur[part];
  }
  return cur;
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const BUILTIN_FORMATS: Record<string, number> = {
  General: 0,
  '0': 1,
  '0.00': 2,
  '#,##0': 3,
  '#,##0.00': 4,
  '0%': 9,
  '0.00%': 10,
  '@': 49,
};

class Styles {
  private numFmts = new Map<string, number>();
  private xfs = new Map<string, number>([['0|0', 0]]);
  private xfList: Array<[font: number, numFmt: number]> = [[0, 0]];

  /** font: 0 normal, 1 bold, 2 hyperlink */
  get(font: number, format: string | undefined): number {
    const numFmt = format === undefined ? 0 : this.numFmtId(format);
    const key = `${font}|${numFmt}`;
    let id = this.xfs.get(key);
    if (id === undefined) {
      id = this.xfList.length;
      this.xfList.push([font, numFmt]);
      this.xfs.set(key, id);
    }
    return id;
  }

  private numFmtId(code: string): number {
    if (code in BUILTIN_FORMATS) return BUILTIN_FORMATS[code]!;
    let id = this.numFmts.get(code);
    if (id === undefined) {
      id = 164 + this.numFmts.size;
      this.numFmts.set(code, id);
    }
    return id;
  }

  xml(): string {
    const numFmts = [...this.numFmts]
      .map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`)
      .join('');
    const xfs = this.xfList
      .map(([font, numFmt]) =>
        `<xf numFmtId="${numFmt}" fontId="${font}" fillId="0" borderId="0" xfId="0"` +
        `${numFmt ? ' applyNumberFormat="1"' : ''}${font ? ' applyFont="1"' : ''}/>`)
      .join('');
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (this.numFmts.size ? `<numFmts count="${this.numFmts.size}">${numFmts}</numFmts>` : '') +
      '<fonts count="3">' +
      '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
      '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
      '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
      '</fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfList.length}">${xfs}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>'
    );
  }
}

class SharedStrings {
  private map = new Map<string, number>();
  private list: string[] = [];
  count = 0;

  id(text: string): number {
    this.count++;
    let id = this.map.get(text);
    if (id === undefined) {
      id = this.list.length;
      this.list.push(text);
      this.map.set(text, id);
    }
    return id;
  }

  xml(): string {
    const items = this.list.map((s) => {
      const space = /^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : '';
      return `<si><t${space}>${esc(s)}</t></si>`;
    });
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.count}" uniqueCount="${this.list.length}">` +
      items.join('') +
      '</sst>'
    );
  }
}

// ---------------------------------------------------------------------------
// Worksheet
// ---------------------------------------------------------------------------

function formatCode(format: Format | undefined, column: Column | undefined): string | undefined {
  switch (format) {
    case undefined:
    case 'link':
      return undefined;
    case 'text':
      return '@';
    case 'integer':
      return '#,##0';
    case 'decimal':
      return '#,##0.00';
    case 'percent':
      return '0.00%';
    case 'currency': {
      const symbol = (column?.currency ?? '$').replace(/"/g, '');
      return `"${symbol}"#,##0.00;-"${symbol}"#,##0.00`;
    }
    case 'date':
      return 'yyyy-mm-dd';
    case 'datetime':
      return 'yyyy-mm-dd hh:mm:ss';
    case 'time':
      return 'hh:mm:ss';
    default:
      return format;
  }
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

function dateToSerial(date: Date, utc: boolean): number {
  const ms = utc
    ? date.getTime()
    : Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds());
  return (ms - EXCEL_EPOCH) / 86_400_000;
}

function isMidnight(date: Date, utc: boolean): boolean {
  return utc
    ? date.getUTCHours() === 0 && date.getUTCMinutes() === 0 && date.getUTCSeconds() === 0 && date.getUTCMilliseconds() === 0
    : date.getHours() === 0 && date.getMinutes() === 0 && date.getSeconds() === 0 && date.getMilliseconds() === 0;
}

interface BuiltSheet {
  name: string;
  xml: string;
  rels?: string;
  filterRef?: string;
  hasFormula: boolean;
}

interface ResolvedColumn {
  header: string;
  get: (row: any, index: number) => unknown;
  width?: number;
  format?: Format;
  source?: Column;
}

function resolveColumns(rows: readonly any[], opts: SheetOptions): { columns: ResolvedColumn[]; isArrays: boolean } {
  const isArrays = rows.length > 0 && Array.isArray(rows[0]);
  if (opts.columns) {
    return {
      isArrays,
      columns: opts.columns.map((c, i) => {
        const col: Column = typeof c === 'string' ? { key: c } : c;
        const key = col.key;
        const header = col.header ?? (key !== undefined ? (opts.humanizeHeaders ? humanize(key) : key) : `Column ${i + 1}`);
        const get = col.value
          ? col.value
          : key !== undefined
            ? (row: any) => getPath(row, key)
            : isArrays
              ? (row: any) => row[i]
              : () => undefined;
        return { header, get, width: col.width, format: col.format, source: col };
      }),
    };
  }
  if (isArrays) {
    let width = 0;
    for (const row of rows) width = Math.max(width, (row as unknown[]).length);
    return { isArrays, columns: Array.from({ length: width }, (_, i) => ({ header: `Column ${i + 1}`, get: (row: any) => row[i] })) };
  }
  const keys = new Set<string>();
  for (const row of rows) {
    if (row && typeof row === 'object') for (const key of Object.keys(row)) keys.add(key);
  }
  return {
    isArrays,
    columns: [...keys].map((key) => ({ header: opts.humanizeHeaders ? humanize(key) : key, get: (row: any) => row?.[key] })),
  };
}

function isFormula(v: unknown): v is FormulaCell {
  return typeof v === 'object' && v !== null && typeof (v as FormulaCell).formula === 'string';
}

function isLink(v: unknown): v is LinkCell {
  return typeof v === 'object' && v !== null && typeof (v as LinkCell).link === 'string';
}

function displayLength(v: unknown): number {
  if (v == null) return 0;
  if (v instanceof Date) return isNaN(v.getTime()) ? 0 : 19;
  if (isLink(v)) return (v.text ?? v.link).length;
  if (isFormula(v)) return 10;
  if (typeof v === 'number') return Math.min(15, String(v).length + 2);
  const s = typeof v === 'object' ? safeJson(v) : String(v);
  let max = 0;
  for (const line of s.split('\n')) max = Math.max(max, line.length);
  return max;
}

function safeJson(v: unknown): string {
  try {
    return JSON.stringify(v, (_k, val) => (typeof val === 'bigint' ? val.toString() : val)) ?? '';
  } catch {
    return String(v);
  }
}

function buildSheet(rows: readonly any[], opts: SheetOptions, name: string, styles: Styles, strings: SharedStrings): BuiltSheet {
  const { columns, isArrays } = resolveColumns(rows, opts);
  if (columns.length > MAX_COLS) throw new RangeError(`xlsxport: sheet '${name}' has ${columns.length} columns; Excel allows ${MAX_COLS}.`);
  const header = opts.header ?? (!isArrays || !!opts.columns);
  const totalRows = rows.length + (header ? 1 : 0);
  if (totalRows > MAX_ROWS) throw new RangeError(`xlsxport: sheet '${name}' has ${totalRows} rows; Excel allows ${MAX_ROWS}.`);

  const utc = opts.timezone === 'utc';
  const dateStyle = styles.get(0, opts.dateFormat ?? 'yyyy-mm-dd');
  const dateTimeStyle = styles.get(0, opts.dateTimeFormat ?? 'yyyy-mm-dd hh:mm:ss');
  const headerStyle = opts.boldHeader === false ? 0 : styles.get(1, undefined);
  const linkStyle = styles.get(2, undefined);
  const colStyles = columns.map((c) => {
    const code = formatCode(c.format, c.source);
    return code === undefined ? 0 : styles.get(0, code);
  });
  const textColumn = columns.map((c) => c.format === 'text');
  const linkColumn = columns.map((c) => c.format === 'link');
  const explicitDateFormat = columns.map((c) => c.format === 'date' || c.format === 'datetime' || c.format === 'time' || (c.format !== undefined && /[dmyhs]/i.test(formatCode(c.format, c.source) ?? '') && !/[#0]/.test(formatCode(c.format, c.source) ?? '')));

  const letters = columns.map((_, i) => columnName(i));
  const widths = columns.map((c) => c.width ?? 0);
  const measure = columns.map((c) => c.width === undefined);
  const MEASURE_ROWS = 2000;

  const links: Array<[ref: string, url: string]> = [];
  let hasFormula = false;
  const out: string[] = [];
  let r = 0;

  if (header) {
    r = 1;
    let cells = '';
    columns.forEach((c, i) => {
      if (measure[i]) widths[i] = Math.max(widths[i]!, c.header.length + (opts.autoFilter === false ? 0 : 2));
      cells += `<c r="${letters[i]}1" t="s"${headerStyle ? ` s="${headerStyle}"` : ''}><v>${strings.id(c.header)}</v></c>`;
    });
    out.push(`<row r="1">${cells}</row>`);
  }

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    r++;
    let cells = '';
    for (let c = 0; c < columns.length; c++) {
      let value: unknown = columns[c]!.get(row, i);
      if (value === undefined || value === null) continue;
      if (i < MEASURE_ROWS && measure[c]) widths[c] = Math.max(widths[c]!, displayLength(value));
      const ref = letters[c]! + r;
      const style = colStyles[c]!;
      const s = (id: number) => (id ? ` s="${id}"` : '');

      if (linkColumn[c] && typeof value === 'string' && value !== '') value = { link: value };

      if (isLink(value)) {
        links.push([ref, value.link]);
        cells += `<c r="${ref}" t="s"${s(linkStyle)}><v>${strings.id(truncate(value.text ?? value.link))}</v></c>`;
        continue;
      }
      if (isFormula(value)) {
        hasFormula = true;
        cells += `<c r="${ref}"${s(style)}><f>${esc(value.formula.replace(/^=/, ''))}</f></c>`;
        continue;
      }
      if (textColumn[c] && !(value instanceof Date)) {
        cells += `<c r="${ref}" t="s"${s(style)}><v>${strings.id(truncate(typeof value === 'object' ? safeJson(value) : String(value)))}</v></c>`;
        continue;
      }
      switch (typeof value) {
        case 'number':
          if (Number.isFinite(value)) cells += `<c r="${ref}"${s(style)}><v>${value}</v></c>`;
          break;
        case 'bigint':
          cells += Number.isSafeInteger(Number(value))
            ? `<c r="${ref}"${s(style)}><v>${value}</v></c>`
            : `<c r="${ref}" t="s"${s(style)}><v>${strings.id(value.toString())}</v></c>`;
          break;
        case 'boolean':
          cells += `<c r="${ref}" t="b"${s(style)}><v>${value ? 1 : 0}</v></c>`;
          break;
        case 'string':
          if (value === '') break;
          cells += `<c r="${ref}" t="s"${s(style)}><v>${strings.id(truncate(value))}</v></c>`;
          break;
        default:
          if (value instanceof Date) {
            if (isNaN(value.getTime())) break;
            const id = explicitDateFormat[c] ? style : isMidnight(value, utc) ? dateStyle : dateTimeStyle;
            cells += `<c r="${ref}"${s(id)}><v>${dateToSerial(value, utc)}</v></c>`;
          } else if (typeof value === 'object') {
            cells += `<c r="${ref}" t="s"${s(style)}><v>${strings.id(truncate(safeJson(value)))}</v></c>`;
          } else {
            cells += `<c r="${ref}" t="s"${s(style)}><v>${strings.id(truncate(String(value)))}</v></c>`;
          }
      }
    }
    out.push(cells ? `<row r="${r}">${cells}</row>` : `<row r="${r}"/>`);
  }

  const lastCol = columnName(Math.max(0, columns.length - 1));
  const dimension = columns.length && r ? `A1:${lastCol}${r}` : 'A1';
  const freeze = header && (opts.freezeHeader ?? true);
  const filter = header && columns.length > 0 && (opts.autoFilter ?? true);
  const filterRef = filter ? `A1:${lastCol}${Math.max(r, 1)}` : undefined;

  const cols = columns.length
    ? '<cols>' +
      widths.map((w, i) => {
        const width = measure[i] ? Math.min(60, Math.max(8, w + 2)) : w;
        return `<col min="${i + 1}" max="${i + 1}" width="${width}" customWidth="1"/>`;
      }).join('') +
      '</cols>'
    : '';

  const view = freeze
    ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    : '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';

  const hyperlinks = links.length
    ? '<hyperlinks>' + links.map(([ref], i) => `<hyperlink ref="${ref}" r:id="rId${i + 1}"/>`).join('') + '</hyperlinks>'
    : '';

  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<dimension ref="${dimension}"/>` +
    view +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    cols +
    `<sheetData>${out.join('')}</sheetData>` +
    (filterRef ? `<autoFilter ref="${filterRef}"/>` : '') +
    hyperlinks +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>';

  const rels = links.length
    ? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      links.map(([, url], i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(url)}" TargetMode="External"/>`).join('') +
      '</Relationships>'
    : undefined;

  return { name, xml, rels, filterRef, hasFormula };
}

function truncate(text: string): string {
  return text.length > MAX_CELL_TEXT ? text.slice(0, MAX_CELL_TEXT) : text;
}

// ---------------------------------------------------------------------------
// Workbook
// ---------------------------------------------------------------------------

function cleanSheetName(name: string, used: Set<string>): string {
  let base = String(name).replace(/[\\/?*[\]:]/g, '_').replace(/^'+|'+$/g, '').trim() || 'Sheet';
  if (base.toLowerCase() === 'history') base = 'History_';
  base = base.slice(0, 31);
  let candidate = base;
  for (let n = 2; used.has(candidate.toLowerCase()); n++) {
    const suffix = ` (${n})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function normalize(input: WorkbookInput, options: WorkbookOptions): Array<{ name: string; rows: readonly any[]; opts: SheetOptions }> {
  if (Array.isArray(input)) return [{ name: options.name ?? 'Sheet1', rows: input, opts: options }];
  if (input === null || typeof input !== 'object') {
    throw new TypeError('xlsxport: expected an array of rows or an object of { sheetName: rows }.');
  }
  const entries = Object.entries(input as Record<string, Rows | Sheet>);
  if (entries.length === 0) throw new TypeError('xlsxport: the workbook needs at least one sheet.');
  return entries.map(([name, value]) => {
    if (Array.isArray(value)) return { name, rows: value, opts: { ...options, name: undefined } };
    const sheet = value as Sheet;
    if (!sheet || !Array.isArray(sheet.rows)) throw new TypeError(`xlsxport: sheet '${name}' must be an array or { rows: [...] }.`);
    const { rows, ...rest } = sheet;
    return { name: rest.name ?? name, rows, opts: { ...options, ...rest } };
  });
}

const encoder = new TextEncoder();

function buildEntries(input: WorkbookInput, options: WorkbookOptions = {}): ZipEntry[] {
  const styles = new Styles();
  const strings = new SharedStrings();
  const used = new Set<string>();
  const sheets = normalize(input, options).map(({ name, rows, opts }) =>
    buildSheet(rows, opts, cleanSheetName(name, used), styles, strings));
  const hasFormula = sheets.some((s) => s.hasFormula);

  const definedNames = sheets
    .map((s, i) => (s.filterRef
      ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${esc(s.name.replace(/'/g, "''"))}'!${s.filterRef.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}</definedName>`
      : ''))
    .join('');

  const workbook =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<bookViews><workbookView/></bookViews>' +
    '<sheets>' + sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets>' +
    (definedNames ? `<definedNames>${definedNames}</definedNames>` : '') +
    (hasFormula ? '<calcPr calcId="191029" fullCalcOnLoad="1"/>' : '') +
    '</workbook>';

  const n = sheets.length;
  const workbookRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
    `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
    '</Relationships>';

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const core =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    (options.creator ? `<dc:creator>${esc(options.creator)}</dc:creator>` : '') +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    '</cp:coreProperties>';

  const app =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>xlsxport</Application></Properties>';

  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>';

  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '</Relationships>';

  const files: Array<[string, string]> = [
    ['[Content_Types].xml', contentTypes],
    ['_rels/.rels', rootRels],
    ['docProps/core.xml', core],
    ['docProps/app.xml', app],
    ['xl/workbook.xml', workbook],
    ['xl/_rels/workbook.xml.rels', workbookRels],
    ['xl/styles.xml', styles.xml()],
    ['xl/sharedStrings.xml', strings.xml()],
  ];
  sheets.forEach((s, i) => {
    files.push([`xl/worksheets/sheet${i + 1}.xml`, s.xml]);
    if (s.rels) files.push([`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, s.rels]);
  });
  return files.map(([name, text]) => ({ name, data: encoder.encode(text) }));
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/** MIME type for .xlsx files. */
export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/** @internal Used by 'xlsxport/node' to plug in zlib. */
export async function toExcelWith(input: WorkbookInput, options: WorkbookOptions | undefined, deflate: DeflateRaw | undefined): Promise<Uint8Array> {
  return zip(buildEntries(input, options), deflate);
}

/**
 * Build an .xlsx file from JSON rows. Returns the file bytes (compressed when `CompressionStream('deflate-raw')` is available, otherwise stored uncompressed).
 *
 *   const bytes = await toExcel([{ name: 'Ada', joined: new Date() }])
 *   const bytes = await toExcel({ Users: users, Orders: { rows: orders, columns: ['id', 'total'] } })
 */
export async function toExcel<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>): Promise<Uint8Array> {
  return zip(buildEntries(input, options), webDeflate());
}

/**
 * Synchronous version of `toExcel`. The file is not compressed (larger, but valid),
 * unless you pass a sync deflate function such as `zlib.deflateRawSync`.
 */
export function toExcelSync<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>, deflateRawSync?: (data: Uint8Array) => Uint8Array): Uint8Array {
  return zipSync(buildEntries(input, options), deflateRawSync);
}

/** Build the .xlsx file as a Blob (browser, Node 18+, Deno, Bun, edge runtimes). */
export async function toBlob<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>): Promise<Blob> {
  return new Blob([(await toExcel(input, options)) as BlobPart], { type: XLSX_MIME });
}

/** Build a `Response` that downloads the file — for Next.js route handlers, Hono, Cloudflare Workers, etc. */
export async function toResponse<T = any>(input: WorkbookInput, filename = 'export.xlsx', options?: WorkbookOptions<T>): Promise<Response> {
  const name = ensureExt(filename);
  return new Response((await toExcel(input, options)) as BodyInit, {
    headers: {
      'Content-Type': XLSX_MIME,
      'Content-Disposition': `attachment; filename="${name.replace(/[^\x20-\x7e]|"/g, '_')}"; filename*=UTF-8''${encodeURIComponent(name)}`,
    },
  });
}

/** Browser only: build the file and trigger a download. */
export async function download<T = any>(input: WorkbookInput, filename = 'export.xlsx', options?: WorkbookOptions<T>): Promise<void> {
  const doc = (globalThis as any).document;
  if (!doc) throw new Error("xlsxport: download() needs a browser. In Node use writeExcel() from 'xlsxport/node'.");
  const url = URL.createObjectURL(await toBlob(input, options));
  const a = doc.createElement('a');
  a.href = url;
  a.download = ensureExt(filename);
  a.style.display = 'none';
  doc.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function ensureExt(name: string): string {
  return /\.xlsx$/i.test(name) ? name : `${name}.xlsx`;
}
