/**
 * Node.js helpers: zlib compression and writing straight to disk.
 *
 *   import { writeExcel } from 'xlsxport/node'
 *   await writeExcel('users.xlsx', users)
 */

import { deflateRaw, deflateRawSync } from 'node:zlib';
import { writeFile } from 'node:fs/promises';
import { promisify } from 'node:util';
import { toExcelWith, toExcelSync as toExcelSyncCore, type WorkbookInput, type WorkbookOptions } from './index';

export * from './index';

const deflateAsync = promisify(deflateRaw);
const deflate = async (data: Uint8Array) => new Uint8Array(await deflateAsync(data));

/** Build an .xlsx file using Node's zlib (works on every Node version). */
export function toExcel<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>): Promise<Uint8Array> {
  return toExcelWith(input, options, deflate);
}

/** Build a compressed .xlsx file synchronously. */
export function toExcelSync<T = any>(input: WorkbookInput, options?: WorkbookOptions<T>): Uint8Array {
  return toExcelSyncCore(input, options, (data) => new Uint8Array(deflateRawSync(data)));
}

/** Build an .xlsx file and write it to `path`. Returns the number of bytes written. */
export async function writeExcel<T = any>(path: string, input: WorkbookInput, options?: WorkbookOptions<T>): Promise<number> {
  const bytes = await toExcel(input, options);
  await writeFile(path, bytes);
  return bytes.length;
}
