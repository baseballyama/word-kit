import { unzipSync, zipSync } from "fflate";

export interface ZipEntry {
  readonly name: string;
  readonly data: Uint8Array;
}

/**
 * Read a ZIP archive into ordered entries.
 *
 * fflate's {@link unzipSync} returns a plain object whose key order matches
 * the central directory order, which we surface as an array. Per-entry
 * compression method is not exposed by the sync API, so we record only the
 * decompressed payload.
 */
export function readZip(bytes: Uint8Array): ZipEntry[] {
  const decoded = unzipSync(bytes);
  return Object.entries(decoded).map(([name, data]) => ({ name, data }));
}

export interface WriteZipEntry {
  readonly name: string;
  readonly data: Uint8Array;
  readonly compression: "store" | "deflate";
}

/**
 * Serialize entries back to a ZIP archive. Entries are written in the given
 * order. `compression: "store"` uses no compression (DEFLATE level 0 in
 * fflate's terms); `"deflate"` uses level 6.
 */
export function writeZip(entries: readonly WriteZipEntry[]): Uint8Array {
  const input: Record<string, [Uint8Array, { level: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 }]> = {};
  for (const entry of entries) {
    input[entry.name] = [entry.data, { level: entry.compression === "store" ? 0 : 6 }];
  }
  return zipSync(input);
}
