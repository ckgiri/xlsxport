# Changelog

All notable changes to this package are documented here. This project follows [Semantic Versioning](https://semver.org/).

## 1.1.0 - Unreleased

Documentation only. No code or behavior changes.

- Rewrote the README: plain wording, a full API reference with options tables, and examples that are checked against the build.
- Fixed type doc comments shown in editors: the `header` default, formula cells and the uncompressed fallback are described correctly.

## 1.0.0 - 2026-11-24

First release.

- `toExcel`, `toExcelSync`, `toBlob`, `toResponse`, `download`; `xlsxport/node` adds zlib compression and `writeExcel`.
- Columns from object keys, or configured with `key` (nested paths), `header`, `value`, `width`, `format`, `currency`.
- Formats: integer, decimal, percent, currency, date, datetime, time, text, link, or any Excel format code.
- Bold frozen header, autofilter and auto-sized columns by default.
- Real Excel dates (local or UTC), booleans, bigints, hyperlinks and formulas.
- Multiple sheets, arrays of arrays, sheet-name cleaning, XML-safe text, Excel's cell-length limit.
- Zero dependencies (built-in ZIP writer).
