import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { inflateRawSync } from 'node:zlib';
import { toExcel, toExcelSync, toBlob, toResponse, columnName, XLSX_MIME } from '../src/index';
import { toExcel as nodeToExcel, toExcelSync as nodeToExcelSync, writeExcel } from '../src/node';
import { crc32 } from '../src/zip';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

async function read(bytes: Uint8Array) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes) as any);
  return wb;
}

/** Tiny unzip for inspecting raw XML. */
function unzip(bytes: Uint8Array): Record<string, string> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const files: Record<string, string> = {};
  let pos = 0;
  while (view.getUint32(pos, true) === 0x04034b50) {
    const method = view.getUint16(pos + 8, true);
    const size = view.getUint32(pos + 18, true);
    const nameLen = view.getUint16(pos + 26, true);
    const name = new TextDecoder().decode(bytes.subarray(pos + 30, pos + 30 + nameLen));
    const data = bytes.subarray(pos + 30 + nameLen, pos + 30 + nameLen + size);
    files[name] = new TextDecoder().decode(method === 8 ? inflateRawSync(data) : data);
    pos += 30 + nameLen + size;
  }
  return files;
}

const users = [
  { id: 1, name: 'Ada', email: 'ada@example.com', active: true, joined: new Date(2024, 0, 15), score: 0.875 },
  { id: 2, name: 'Linus <&> "Torvalds"', email: 'linus@example.com', active: false, joined: new Date(2024, 5, 1, 14, 30), score: 0.5 },
  { id: 3, name: 'Grace', extra: 'only here' },
];

describe('toExcel', () => {
  it('writes a valid workbook that Excel readers open', async () => {
    const wb = await read(await toExcel(users));
    const ws = wb.getWorksheet('Sheet1')!;
    expect(ws.getRow(1).values).toEqual([undefined, 'id', 'name', 'email', 'active', 'joined', 'score', 'extra']);
    expect(ws.getCell('B2').value).toBe('Ada');
    expect(ws.getCell('B3').value).toBe('Linus <&> "Torvalds"');
    expect(ws.getCell('A2').value).toBe(1);
    expect(ws.getCell('D2').value).toBe(true);
    expect(ws.getCell('D3').value).toBe(false);
    expect(ws.getCell('G4').value).toBe('only here');
    expect(ws.getCell('E4').value).toBeNull();
    expect(ws.rowCount).toBe(4);
  });

  it('writes dates as real Excel dates with sensible formats', async () => {
    const rows = [{ d: new Date(Date.UTC(2024, 0, 15)) }, { d: new Date(Date.UTC(2024, 5, 1, 14, 30)) }];
    const wb = await read(await toExcel(rows, { timezone: 'utc' }));
    const ws = wb.worksheets[0]!;
    expect(ws.getCell('A2').numFmt).toBe('yyyy-mm-dd');
    expect(ws.getCell('A3').numFmt).toBe('yyyy-mm-dd hh:mm:ss');
    expect((ws.getCell('A2').value as Date).getTime()).toBe(rows[0]!.d.getTime());
    expect((ws.getCell('A3').value as Date).getTime()).toBe(rows[1]!.d.getTime());
  });

  it('local timezone keeps the wall-clock time', async () => {
    const files = unzip(await toExcel([{ d: new Date(2024, 5, 1, 14, 30) }]));
    // 2024-06-01 14:30 => serial 45444.604166...
    expect(files['xl/worksheets/sheet1.xml']).toMatch(/<v>45444\.6041666/);
  });

  it('bolds, freezes and filters the header by default', async () => {
    const files = unzip(await toExcel(users));
    const sheet = files['xl/worksheets/sheet1.xml']!;
    expect(sheet).toContain('state="frozen"');
    expect(sheet).toContain('<autoFilter ref="A1:G4"/>');
    expect(files['xl/workbook.xml']).toContain(`'Sheet1'!$A$1:$G$4`);
    const wb = await read(await toExcel(users));
    expect(wb.worksheets[0]!.getCell('A1').font?.bold).toBe(true);
  });

  it('auto-sizes columns', async () => {
    const wb = await read(await toExcel([{ short: 'a', long: 'x'.repeat(40), huge: 'y'.repeat(500) }]));
    const ws = wb.worksheets[0]!;
    expect(ws.getColumn(1).width).toBe(9); // header 'short' + filter button
    expect(ws.getColumn(2).width).toBe(42);
    expect(ws.getColumn(3).width).toBe(60);
  });

  it('supports columns: pick, rename, compute, format, nested paths', async () => {
    const rows = [{ user: { name: 'Ada' }, price: 1234.5, qty: 3, rate: 0.25, zip: '01234' }];
    const bytes = await toExcel(rows, {
      columns: [
        { key: 'user.name', header: 'Customer' },
        { key: 'price', format: 'currency', currency: '₹' },
        { key: 'qty', format: 'integer' },
        { header: 'Total', value: (r) => r.price * r.qty, format: 'decimal' },
        { key: 'rate', format: 'percent' },
        { key: 'zip', format: 'text' },
        'qty',
      ],
    });
    const ws = (await read(bytes)).worksheets[0]!;
    expect(ws.getRow(1).values).toEqual([undefined, 'Customer', 'price', 'qty', 'Total', 'rate', 'zip', 'qty']);
    expect(ws.getCell('A2').value).toBe('Ada');
    expect(ws.getCell('B2').numFmt).toBe('"₹"#,##0.00;-"₹"#,##0.00');
    expect(ws.getCell('C2').numFmt).toBe('#,##0');
    expect(ws.getCell('D2').value).toBe(3703.5);
    expect(ws.getCell('D2').numFmt).toBe('#,##0.00');
    expect(ws.getCell('E2').numFmt).toBe('0.00%');
    expect(ws.getCell('F2').value).toBe('01234');
  });

  it('humanizes headers on request', async () => {
    const ws = (await read(await toExcel([{ firstName: 'a', created_at: 'b', userID: 'c' }], { humanizeHeaders: true }))).worksheets[0]!;
    expect(ws.getRow(1).values).toEqual([undefined, 'First Name', 'Created At', 'User ID']);
  });

  it('writes multiple sheets with per-sheet options', async () => {
    const wb = await read(await toExcel({
      Users: users,
      'Orders/2024': { rows: [{ id: 'A1', total: 10 }], columns: ['id', 'total'] },
    }));
    expect(wb.worksheets.map((w) => w.name)).toEqual(['Users', 'Orders_2024']);
    expect(wb.getWorksheet('Orders_2024')!.getCell('B2').value).toBe(10);
  });

  it('dedupes and truncates sheet names', async () => {
    const wb = await read(await toExcel({ ['x'.repeat(40)]: [{ a: 1 }], ['X'.repeat(40)]: [{ a: 2 }] }));
    expect(wb.worksheets.map((w) => w.name)).toEqual(['x'.repeat(31), 'X'.repeat(27) + ' (2)']);
  });

  it('supports links and formulas', async () => {
    const bytes = await toExcel(
      [
        { site: 'https://example.com', docs: { link: 'https://example.com/docs', text: 'Docs' }, n: 2 },
        { site: 'https://b.example', docs: null, n: 3 },
        { site: null, docs: null, n: { formula: 'SUM(C2:C3)' } },
      ],
      { columns: [{ key: 'site', format: 'link' }, 'docs', 'n'] },
    );
    const ws = (await read(bytes)).worksheets[0]!;
    expect(ws.getCell('A2').value).toMatchObject({ text: 'https://example.com', hyperlink: 'https://example.com' });
    expect(ws.getCell('B2').value).toMatchObject({ text: 'Docs', hyperlink: 'https://example.com/docs' });
    expect(ws.getCell('C4').value).toMatchObject({ formula: 'SUM(C2:C3)' });
    expect(unzip(bytes)['xl/workbook.xml']).toContain('fullCalcOnLoad="1"');
  });

  it('handles awkward values safely', async () => {
    const bytes = await toExcel([
      {
        ctrl: 'bad\u0000chars\u0007here',
        formulaLike: '=HYPERLINK("http://evil")',
        nan: NaN,
        inf: Infinity,
        big: 2n ** 70n,
        small: 5n,
        obj: { a: [1, 2] },
        spaced: '  padded  ',
        multi: 'line1\nline2',
        emoji: 'ok 👍',
        invalidDate: new Date('nope'),
        long: 'z'.repeat(40_000),
      },
    ]);
    const ws = (await read(bytes)).worksheets[0]!;
    expect(ws.getCell('A2').value).toBe('badcharshere');
    expect(ws.getCell('B2').value).toBe('=HYPERLINK("http://evil")'); // stored as text, never executed
    expect(ws.getCell('C2').value).toBeNull();
    expect(ws.getCell('D2').value).toBeNull();
    expect(ws.getCell('E2').value).toBe('1180591620717411303424');
    expect(ws.getCell('F2').value).toBe(5);
    expect(ws.getCell('G2').value).toBe('{"a":[1,2]}');
    expect(ws.getCell('H2').value).toBe('  padded  ');
    expect(ws.getCell('I2').value).toBe('line1\nline2');
    expect(ws.getCell('J2').value).toBe('ok 👍');
    expect(ws.getCell('K2').value).toBeNull();
    expect((ws.getCell('L2').value as string).length).toBe(32_767);
  });

  it('writes arrays of arrays', async () => {
    const ws = (await read(await toExcel([['Name', 'Age'], ['Ada', 36]]))).worksheets[0]!;
    expect(ws.getCell('A1').value).toBe('Name');
    expect(ws.getCell('B2').value).toBe(36);
    expect(unzip(await toExcel([['a']]))['xl/worksheets/sheet1.xml']).not.toContain('frozen');
  });

  it('writes an empty sheet', async () => {
    const wb = await read(await toExcel([]));
    expect(wb.worksheets).toHaveLength(1);
  });

  it('rejects bad input clearly', async () => {
    await expect(toExcel('nope' as any)).rejects.toThrow(/expected an array/);
    await expect(toExcel({} as any)).rejects.toThrow(/at least one sheet/);
    await expect(toExcel({ S: { nope: 1 } } as any)).rejects.toThrow(/must be an array/);
  });

  it('compresses', async () => {
    const rows = Array.from({ length: 2000 }, (_, i) => ({ id: i, name: `User ${i}`, email: `user${i}@example.com` }));
    const compressed = await toExcel(rows);
    const stored = toExcelSync(rows);
    expect(compressed.length).toBeLessThan(stored.length / 3);
    expect((await read(stored)).worksheets[0]!.rowCount).toBe(2001);
  });
});

describe('outputs', () => {
  it('toBlob and toResponse', async () => {
    const blob = await toBlob(users);
    expect(blob.type).toBe(XLSX_MIME);
    const res = await toResponse(users, 'Users Report');
    expect(res.headers.get('content-type')).toBe(XLSX_MIME);
    expect(res.headers.get('content-disposition')).toContain('filename="Users Report.xlsx"');
    expect((await read(new Uint8Array(await res.arrayBuffer()))).worksheets[0]!.rowCount).toBe(4);
  });

  it('node helpers', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsxport-'));
    const path = join(dir, 'users.xlsx');
    const size = await writeExcel(path, users);
    expect(readFileSync(path).length).toBe(size);
    expect((await read(readFileSync(path))).worksheets[0]!.getCell('B2').value).toBe('Ada');
    expect((await read(await nodeToExcel(users))).worksheets[0]!.rowCount).toBe(4);
    expect((await read(nodeToExcelSync(users))).worksheets[0]!.rowCount).toBe(4);
  });
});

describe('helpers', () => {
  it('columnName', () => {
    expect([0, 25, 26, 51, 52, 701, 702, 16383].map(columnName)).toEqual(['A', 'Z', 'AA', 'AZ', 'BA', 'ZZ', 'AAA', 'XFD']);
  });

  it('crc32', () => {
    expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926);
  });
});
