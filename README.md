# xlsxport

Export JSON to a real Excel `.xlsx` file in one call.

[![npm version](https://img.shields.io/npm/v/xlsxport.svg)](https://www.npmjs.com/package/xlsxport) [![bundle size](https://img.shields.io/bundlephobia/minzip/xlsxport)](https://bundlephobia.com/package/xlsxport) [![license](https://img.shields.io/npm/l/xlsxport.svg)](LICENSE)

`xlsxport` turns an array of objects into an Excel file. You get a header row, sized columns, real dates and numbers, without any setup. Use it for "Export to Excel" buttons, reports and API downloads, in place of `exceljs` or `xlsx` when you only need to write files.

## Features

- **One call** from JSON to an `.xlsx` file: `writeExcel('users.xlsx', users)`.
- **Ready header row**: bold, frozen (stays visible when you scroll) and with filter buttons.
- **Column widths fit the content**, so nothing is cut off.
- **Real Excel values**: dates, numbers and booleans, not text.
- **Formats by name**: `currency`, `percent`, `date`, `integer` and more, or any Excel format code.
- **Links, formulas and multiple sheets** in the same file.
- **Runs everywhere**: Node.js, browsers, Bun, Deno and edge runtimes (Cloudflare Workers, Vercel).
- **Safe text**: a string that starts with `=` is always text, never a formula.
- **Zero dependencies**, about 6 KB gzipped, ESM and CommonJS, TypeScript types included.

## Installation

```sh
npm install xlsxport
```

```sh
pnpm add xlsxport
yarn add xlsxport
bun add xlsxport
```

## Quick start

```js
import { writeExcel } from 'xlsxport/node'

const users = [
  { name: 'Ada', email: 'ada@example.com', joined: new Date(2024, 0, 15), active: true },
  { name: 'Linus', email: 'linus@example.com', joined: new Date(2024, 5, 1), active: false },
]

await writeExcel('users.xlsx', users)
// users.xlsx, sheet "Sheet1":
// | name  | email             | joined     | active |  <- bold, frozen, with filters
// | Ada   | ada@example.com   | 2024-01-15 | TRUE   |
// | Linus | linus@example.com | 2024-06-01 | FALSE  |
```

In the browser or on an edge runtime, import from `'xlsxport'` instead (see [Usage](#usage)).

With CommonJS, use `const { writeExcel } = require('xlsxport/node')` or `const { toExcel } = require('xlsxport')`.

## Table of contents

- [Usage](#usage)
- [API](#api)
- [Comparison](#comparison)
- [Notes and limitations](#notes-and-limitations)
- [Requirements](#requirements)
- [License](#license)

## Usage

### Export an array of objects

Without options, every key found in the rows becomes a column, in the order the keys first appear. A missing value gives an empty cell.

```js
import { toExcel } from 'xlsxport'

const rows = [
  { id: 1, name: 'Ada' },
  { id: 2, name: 'Grace', team: 'Navy' },   // a new key adds a new column
]

const bytes = await toExcel(rows)   // Uint8Array with the .xlsx file
// | id | name  | team |
// | 1  | Ada   |      |
// | 2  | Grace | Navy |
```

### Choose, order and rename columns

Pass `columns` to pick which columns to write and in what order. A string is the same as `{ key: 'name' }`.

```js
import { toExcel } from 'xlsxport'

const orders = [
  { id: 1, customer: { name: 'Ada' }, total: 100, discount: 0.1, secret: 'x' },
  { id: 2, customer: { name: 'Grace' }, total: 50, discount: 0, secret: 'y' },
]

const bytes = await toExcel(orders, {
  columns: [
    'id',                                          // same as { key: 'id' }
    { key: 'customer.name', header: 'Customer' },  // nested path, new header text
    { header: 'Net', value: (order) => order.total * (1 - order.discount) },  // computed
  ],
})
// | id | Customer | Net |   ('total', 'discount' and 'secret' are not written)
// | 1  | Ada      | 90  |
// | 2  | Grace    | 50  |
```

To turn keys like `firstName` into readable headers, set `humanizeHeaders`. It only changes headers you did not write yourself.

```js
await toExcel([{ firstName: 'Ada', created_at: new Date() }], { humanizeHeaders: true })
// | First Name | Created At |
```

### Format numbers, currency and dates

Set `format` on a column. The cell keeps the real number, so Excel can still sum and sort it. Only the display changes.

```js
import { writeExcel } from 'xlsxport/node'

await writeExcel('orders.xlsx', orders, {
  columns: [
    { key: 'total', format: 'currency', currency: '₹' },  // 1234.5 -> ₹1,234.50
    { key: 'qty', format: 'integer' },                     // 1234 -> 1,234
    { key: 'discount', format: 'percent' },                // 0.15 -> 15.00%
    { key: 'createdAt', header: 'Date', format: 'date' },  // 2024-01-15
    { key: 'pincode', format: 'text' },                    // stored as text: '0042' keeps its zeros
    { key: 'weight', format: '0.0" kg"', width: 12 },      // any Excel format code: 72.5 kg
  ],
})
```

| Format | Excel format code | Example |
|---|---|---|
| `integer` | `#,##0` | `1234` shows as `1,234` |
| `decimal` | `#,##0.00` | `1234.5` shows as `1,234.50` |
| `percent` | `0.00%` | `0.15` shows as `15.00%` |
| `currency` | `"$"#,##0.00;-"$"#,##0.00` | `1234.5` shows as `$1,234.50` (symbol from the `currency` option) |
| `date` | `yyyy-mm-dd` | `2024-01-15` |
| `datetime` | `yyyy-mm-dd hh:mm:ss` | `2024-01-15 14:30:00` |
| `time` | `hh:mm:ss` | `14:30:00` |
| `text` | `@` | Every value is stored as text, numbers too. |
| `link` | none | A string value becomes a clickable link. |
| Any other string | the string itself | `'0.0" kg"'` shows `72.5` as `72.5 kg` |

`Date` values need no format. A date at midnight uses `dateFormat` (default `yyyy-mm-dd`). A date with a time uses `dateTimeFormat` (default `yyyy-mm-dd hh:mm:ss`). You can change both for the whole file:

```js
await toExcel(rows, { dateFormat: 'dd/mm/yyyy', timezone: 'utc' })
```

`timezone` decides which clock is written. `'local'` (the default) writes the date as it looks on the machine that runs the code. `'utc'` writes the UTC time.

### Add links and formulas

Use `{ link, text }` for a clickable link and `{ formula }` for a formula. Excel calculates formulas when it opens the file.

```js
import { toExcel } from 'xlsxport'

const bytes = await toExcel([
  { item: 'Pen', price: 2, page: { link: 'https://example.com/pen', text: 'View' } },
  { item: 'Book', price: 12, page: { link: 'https://example.com/book' } },  // shows the URL
  { item: 'Total', price: { formula: 'SUM(B2:B3)' } },                       // Excel shows 14
  { item: '=1+1' },                                                           // stays text: =1+1
])
```

A plain string is never run as a formula, even if it starts with `=`. So user input in your data cannot turn into a formula (this is called formula injection).

If a column holds URL strings, `format: 'link'` makes each one clickable: `{ key: 'url', format: 'link' }`.

### Export several sheets

Pass an object instead of an array. Each key is a sheet name. A value can be an array of rows, or `{ rows, ...options }` to set options for that sheet only.

```js
import { writeExcel } from 'xlsxport/node'

await writeExcel('report.xlsx', {
  Users: users,
  Orders: { rows: orders, columns: ['id', 'total'] },
  Summary: [
    { metric: 'Revenue', value: { formula: 'SUM(Orders!B2:B1000)' } },
  ],
})
// 3 sheets: "Users", "Orders" (2 columns: id, total) and "Summary"
```

Options you pass as the last argument apply to every sheet. Options inside a sheet object win over them.

Sheet names are cleaned to follow Excel's rules: at most 31 characters, the characters `[ ] : * ? / \` are replaced with `_`, and repeated names (ignoring upper and lower case) get a number, as in `Users (2)`.

### Download a file in the browser

`download()` builds the file and starts the download. It adds `.xlsx` to the name if it is missing.

```js
import { download } from 'xlsxport'

button.onclick = () => download(rows, 'orders.xlsx')
```

To upload the file instead, use `toBlob()`:

```js
import { toBlob } from 'xlsxport'

const form = new FormData()
form.append('file', await toBlob(rows), 'orders.xlsx')
await fetch('/upload', { method: 'POST', body: form })
```

### Return a file from an API route (Next.js, Hono, Cloudflare Workers, Remix)

`toResponse()` returns a standard `Response` with the `Content-Type` and `Content-Disposition` headers already set, so the browser downloads the file.

```js
import { toResponse } from 'xlsxport'

export async function GET() {
  const orders = await db.orders.findMany()
  return toResponse(orders, 'orders.xlsx')
}
```

### Send a file from an Express route

```js
import { toExcel } from 'xlsxport/node'

app.get('/export', async (req, res) => {
  const bytes = await toExcel(users)
  res.attachment('users.xlsx')   // sets Content-Disposition and Content-Type
  res.send(Buffer.from(bytes))
})
```

### Write a file in Node.js

Import from `'xlsxport/node'`. It uses Node's built-in `zlib` for compression and adds `writeExcel()`.

```js
import { writeExcel, toExcel, toExcelSync } from 'xlsxport/node'

const size = await writeExcel('report.xlsx', rows)   // => number of bytes written
const bytes = await toExcel(rows)                    // => Uint8Array, for example to upload to S3
const syncBytes = toExcelSync(rows)                  // => Uint8Array, built synchronously
```

### Export arrays of arrays

Rows can also be arrays. Then no header row is added, unless you pass `columns`. Use `{ header }` objects, because a string in `columns` is read as an object key.

```js
import { toExcel } from 'xlsxport'

await toExcel([['Name', 'Age'], ['Ada', 36]])
// | Name | Age |   <- a normal row: not bold, not frozen
// | Ada  | 36  |

await toExcel([['Ada', 36], ['Grace', 45]], {
  columns: [{ header: 'Name' }, { header: 'Age' }],
})
// | Name  | Age |   <- a real header row
// | Ada   | 36  |
// | Grace | 45  |
```

## API

There are two entry points:

- `xlsxport`: `toExcel`, `toExcelSync`, `toBlob`, `toResponse`, `download`, `columnName`, `XLSX_MIME`. Works in every runtime.
- `xlsxport/node`: all of the above, plus `writeExcel`. Here `toExcel` and `toExcelSync` compress with Node's `zlib`.

### toExcel(input, options?)

```ts
function toExcel<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>): Promise<Uint8Array>
```

Builds the `.xlsx` file and returns its bytes. `input` is an array of rows, or an object of sheets (see [Export several sheets](#export-several-sheets)). For `options`, see [Options](#options).

The `xlsxport` version compresses with the web-standard `CompressionStream`. If the runtime does not support it, the file is written without compression (larger, but valid). The `xlsxport/node` version always compresses.

### toExcelSync(input, options?, deflateRawSync?)

```ts
// from 'xlsxport'
function toExcelSync<T = any>(
  input: WorkbookInput,
  options?: WorkbookOptions<T>,
  deflateRawSync?: (data: Uint8Array) => Uint8Array,
): Uint8Array

// from 'xlsxport/node'
function toExcelSync<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>): Uint8Array
```

Synchronous version of `toExcel()`. From `'xlsxport'`, the file is not compressed (larger, but valid), unless you pass a sync compress function such as `zlib.deflateRawSync`. From `'xlsxport/node'`, it is always compressed.

### toBlob(input, options?)

```ts
function toBlob<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>): Promise<Blob>
```

Builds the file as a `Blob` with the `.xlsx` MIME type. Works in browsers, Node.js, Bun, Deno and edge runtimes.

### toResponse(input, filename?, options?)

```ts
function toResponse<T = any>(
  input: WorkbookInput,
  filename?: string,          // default 'export.xlsx'
  options?: WorkbookOptions<T>,
): Promise<Response>
```

Builds a `Response` that downloads the file. It sets `Content-Type` to `XLSX_MIME` and `Content-Disposition` to `attachment` with the file name. Names with non-ASCII characters work. `.xlsx` is added to the name if it is missing.

### download(input, filename?, options?)

```ts
function download<T = any>(
  input: WorkbookInput,
  filename?: string,          // default 'export.xlsx'
  options?: WorkbookOptions<T>,
): Promise<void>
```

Browser only. Builds the file and starts a download. `.xlsx` is added to the name if it is missing. Throws an `Error` when there is no `document`, for example in Node.js.

### writeExcel(path, input, options?)

```ts
// from 'xlsxport/node'
function writeExcel<T = any>(path: string, input: WorkbookInput, options?: WorkbookOptions<T>): Promise<number>
```

Builds the file and writes it to `path`. Returns the number of bytes written. The path is used as given (`.xlsx` is not added).

### columnName(index)

```ts
function columnName(index: number): string
```

Turns a zero-based column number into an Excel column letter. Useful when you build formulas.

```js
import { columnName } from 'xlsxport'

columnName(0)    // => 'A'
columnName(25)   // => 'Z'
columnName(26)   // => 'AA'
```

### XLSX_MIME

```ts
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
```

The MIME type of `.xlsx` files, for your own `Content-Type` headers.

### Options

All options are optional. The last argument of `toExcel()`, `toExcelSync()`, `toBlob()`, `toResponse()`, `download()` and `writeExcel()` takes them. Inside a sheet object (`{ rows, ...options }`), every option except `creator` can be set for that sheet only.

| Option | Type | Default | Description |
|---|---|---|---|
| `name` | `string` | `'Sheet1'` | Sheet name when `input` is one array. In a sheet object, it replaces the key as the name. Cleaned to Excel's rules. |
| `columns` | `Array<Column \| string>` | every key in the rows | Pick, order, rename and format columns. See [Column](#column). |
| `header` | `boolean` | `true` | Write a header row. For arrays of arrays the default is `false`, unless you pass `columns`. |
| `humanizeHeaders` | `boolean` | `false` | Turn keys into readable headers: `firstName` becomes `First Name`. Headers you set with `header` are not changed. |
| `freezeHeader` | `boolean` | `true` | Keep the header row visible when you scroll. Only when there is a header. |
| `autoFilter` | `boolean` | `true` | Add filter and sort buttons to the header. Only when there is a header. |
| `boldHeader` | `boolean` | `true` | Make the header row bold. |
| `dateFormat` | `string` | `'yyyy-mm-dd'` | Excel format for `Date` values at midnight. |
| `dateTimeFormat` | `string` | `'yyyy-mm-dd hh:mm:ss'` | Excel format for `Date` values with a time. |
| `timezone` | `'local' \| 'utc'` | `'local'` | Write dates using local time or UTC. |
| `creator` | `string` | none | Author saved in the file's properties. For the whole file only. |

### Column

An item in `columns`. A string `'name'` is short for `{ key: 'name' }`.

| Option | Type | Default | Description |
|---|---|---|---|
| `key` | `string` | none | Property to read from each row. Dot paths like `'customer.name'` read nested values. |
| `header` | `string` | the key | Header text. Without `key` or `header`, it is `Column 1`, `Column 2`, and so on. |
| `value` | `(row, index) => CellValue` | none | Compute the cell value from the row. Used instead of `key`. |
| `width` | `number` | fit content (8 to 60) | Column width in characters. |
| `format` | `Format` | none | A format name or an Excel format code. See [the formats](#format-numbers-currency-and-dates). |
| `currency` | `string` | `'$'` | Currency symbol for `format: 'currency'`. |

### Cell values

How each JavaScript value is written:

| Value | Written as |
|---|---|
| `number`, `bigint` | Number. A `bigint` too large to be exact as a number is written as text. |
| `boolean` | `TRUE` or `FALSE` |
| `Date` | Excel date (`yyyy-mm-dd`, or with the time if it has one) |
| `string` | Text. A string that starts with `=` is never run as a formula. |
| `{ formula: 'SUM(A1:A9)' }` | Formula, calculated when Excel opens the file. A leading `=` is allowed. |
| `{ link: 'https://...', text: 'Open' }` | Clickable link. Without `text`, the URL is shown. |
| `null`, `undefined`, `''`, `NaN`, `Infinity`, invalid `Date` | Empty cell |
| Other objects and arrays | JSON text |

### Types

All types are exported:

`WorkbookInput`, `WorkbookOptions`, `SheetOptions`, `Sheet`, `Rows`, `Column`, `Format`, `CellValue`, `FormulaCell`, `LinkCell`.

```ts
import { toExcel, type Column } from 'xlsxport'

interface Order { id: number; total: number }

const columns: Column<Order>[] = [
  { key: 'id' },
  { header: 'Total', value: (order) => order.total, format: 'currency' },
]
```

`toExcelWith` is also exported, but it is internal (used by `xlsxport/node`). Do not use it.

## Comparison

The two popular options have problems. `xlsx` (SheetJS) on npm was last published in 2022 (0.18.5) and has known security advisories, because newer versions are only on SheetJS's own CDN. `exceljs` has not had a release since 2023, pulls in 9 dependencies and takes 15+ lines for a basic export.

| | `exceljs` | `xlsx` (npm) | **xlsxport** |
|---|---|---|---|
| JSON to file in one call | No | Partial | Yes |
| Auto column widths | No | No | Yes |
| Frozen bold header with filters | Manual | No | Yes (default) |
| Dates as real Excel dates | Yes | Partial | Yes |
| Runs on edge (Workers, Vercel) | Partial (needs Node APIs) | Yes | Yes |
| Dependencies | 9 | 7 | **0** |
| 100k rows x 6 columns (Node 20) | 16.5 s | Not measured | **1.9 s** |
| Last release | 2023 | 2022 | 2026 |

`xlsxport` only writes spreadsheets. To read `.xlsx` files, use another package.

## Notes and limitations

- **Writing only.** `xlsxport` cannot read or edit existing `.xlsx` files.
- **Little styling.** The header can be bold and links are blue and underlined. There are no colors, borders or merged cells.
- **Formulas have no saved result.** Excel calculates them when it opens the file. Programs that do not calculate formulas show them empty.
- **Column widths are measured from the first 2,000 rows** only.
- **Text is cleaned so the file always opens.** Characters that are not allowed in XML are removed, and text longer than 32,767 characters (Excel's limit per cell) is cut.
- **Excel's size limits apply.** A sheet with more than 1,048,576 rows or 16,384 columns throws a `RangeError`.
- **Bad input throws a `TypeError`**, for example `null` or an empty object of sheets.
- **The whole file is built in memory.** There is no streaming.

## Requirements

- Node.js 18 or newer. Also works in browsers, Bun, Deno and edge runtimes (Cloudflare Workers, Vercel).
- ESM and CommonJS.
- TypeScript types included.
- Zero dependencies.

## License

MIT
