/** Minimal ZIP writer (store or deflate), enough for Office Open XML files. */

export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

export type DeflateRaw = (data: Uint8Array) => Uint8Array | Promise<Uint8Array>;

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) crc = CRC_TABLE[(crc ^ data[i]!) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(date: Date): [time: number, day: number] {
  const year = Math.max(1980, date.getFullYear());
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const day = ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return [time, day];
}

interface Prepared {
  name: Uint8Array;
  data: Uint8Array;
  crc: number;
  size: number;
  method: 0 | 8;
}

function assemble(files: Prepared[], date: Date): Uint8Array {
  const [time, day] = dosDateTime(date);
  let total = 22;
  for (const f of files) total += 30 + f.name.length + f.data.length + 46 + f.name.length;
  if (total > 0xffffffff || files.length > 0xffff) throw new RangeError('xlsxport: workbook is too large for a ZIP32 file.');

  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  let pos = 0;
  const offsets: number[] = [];

  for (const f of files) {
    offsets.push(pos);
    view.setUint32(pos, 0x04034b50, true);
    view.setUint16(pos + 4, 20, true); // version needed
    view.setUint16(pos + 6, 0x0800, true); // UTF-8 names
    view.setUint16(pos + 8, f.method, true);
    view.setUint16(pos + 10, time, true);
    view.setUint16(pos + 12, day, true);
    view.setUint32(pos + 14, f.crc, true);
    view.setUint32(pos + 18, f.data.length, true);
    view.setUint32(pos + 22, f.size, true);
    view.setUint16(pos + 26, f.name.length, true);
    view.setUint16(pos + 28, 0, true);
    out.set(f.name, pos + 30);
    out.set(f.data, pos + 30 + f.name.length);
    pos += 30 + f.name.length + f.data.length;
  }

  const cdStart = pos;
  files.forEach((f, i) => {
    view.setUint32(pos, 0x02014b50, true);
    view.setUint16(pos + 4, 20, true); // version made by
    view.setUint16(pos + 6, 20, true); // version needed
    view.setUint16(pos + 8, 0x0800, true);
    view.setUint16(pos + 10, f.method, true);
    view.setUint16(pos + 12, time, true);
    view.setUint16(pos + 14, day, true);
    view.setUint32(pos + 16, f.crc, true);
    view.setUint32(pos + 20, f.data.length, true);
    view.setUint32(pos + 24, f.size, true);
    view.setUint16(pos + 28, f.name.length, true);
    // extra len, comment len, disk start, internal attrs = 0
    view.setUint32(pos + 38, 0, true); // external attrs
    view.setUint32(pos + 42, offsets[i]!, true);
    out.set(f.name, pos + 46);
    pos += 46 + f.name.length;
  });

  view.setUint32(pos, 0x06054b50, true);
  view.setUint16(pos + 8, files.length, true);
  view.setUint16(pos + 10, files.length, true);
  view.setUint32(pos + 12, pos - cdStart, true);
  view.setUint32(pos + 16, cdStart, true);
  return out;
}

const encoder = new TextEncoder();

function prepare(entry: ZipEntry, compressed: Uint8Array | null): Prepared {
  const useCompressed = compressed !== null && compressed.length < entry.data.length;
  return {
    name: encoder.encode(entry.name),
    data: useCompressed ? compressed! : entry.data,
    crc: crc32(entry.data),
    size: entry.data.length,
    method: useCompressed ? 8 : 0,
  };
}

export function zipSync(entries: ZipEntry[], deflate?: (d: Uint8Array) => Uint8Array, date = new Date()): Uint8Array {
  return assemble(entries.map((e) => prepare(e, deflate ? deflate(e.data) : null)), date);
}

export async function zip(entries: ZipEntry[], deflate?: DeflateRaw, date = new Date()): Promise<Uint8Array> {
  const prepared = await Promise.all(entries.map(async (e) => prepare(e, deflate ? await deflate(e.data) : null)));
  return assemble(prepared, date);
}

/** Deflate using the web-standard CompressionStream, if this runtime supports 'deflate-raw'. */
export function webDeflate(): DeflateRaw | undefined {
  const CS = (globalThis as any).CompressionStream;
  if (typeof CS !== 'function') return undefined;
  try {
    new CS('deflate-raw');
  } catch {
    return undefined;
  }
  return async (data) => {
    const stream = new Blob([data as BlobPart]).stream().pipeThrough(new CS('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  };
}
