import type { PartName } from "./types.js";

/**
 * Normalize a string into a canonical OPC part name.
 *
 * Canonical form: leading `/`, no trailing `/`, no `.` or `..` segments,
 * no double slashes. Per ECMA-376 Part 2 §9.1.1.
 */
export function normalizePartName(input: string): PartName {
  if (input.length === 0) {
    throw new Error("Empty part name");
  }
  const withSlash = input.startsWith("/") ? input : `/${input}`;
  const segments = withSlash.split("/").slice(1);
  const cleaned: string[] = [];
  for (const seg of segments) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") {
      if (cleaned.length === 0) {
        throw new Error(`Part name escapes package root: ${input}`);
      }
      cleaned.pop();
      continue;
    }
    cleaned.push(seg);
  }
  if (cleaned.length === 0) {
    throw new Error(`Part name resolves to package root: ${input}`);
  }
  return `/${cleaned.join("/")}`;
}

/** Convert a canonical part name into the ZIP entry name (no leading slash). */
export function partNameToZipEntry(name: PartName): string {
  return name.startsWith("/") ? name.slice(1) : name;
}

/** Convert a ZIP entry name into a canonical part name. */
export function zipEntryToPartName(entry: string): PartName {
  return normalizePartName(entry);
}

/** Lowercase a part name for case-insensitive comparison (ASCII only). */
export function lowerPartName(name: PartName): PartName {
  return name.toLowerCase();
}

/**
 * Return the parent "folder" of a part name (with trailing `/`).
 *
 * `/word/document.xml` → `/word/`
 * `/word/_rels/document.xml.rels` → `/word/_rels/`
 * `/foo.xml` → `/`
 */
export function partFolder(name: PartName): string {
  const idx = name.lastIndexOf("/");
  return idx <= 0 ? "/" : name.slice(0, idx + 1);
}

/** Return the last path segment (file portion) of a part name. */
export function partBaseName(name: PartName): string {
  const idx = name.lastIndexOf("/");
  return idx < 0 ? name : name.slice(idx + 1);
}

/**
 * Compute the relationships part name for a given source part.
 *
 * Sources at the package root use `/_rels/.rels`. Sources at `/foo/bar.xml`
 * use `/foo/_rels/bar.xml.rels`.
 */
export function relsPartNameFor(source: PartName | "package"): PartName {
  if (source === "package") return "/_rels/.rels";
  const folder = partFolder(source);
  const base = partBaseName(source);
  return `${folder}_rels/${base}.rels`;
}

/**
 * Resolve a relationship target URI (Internal) against the source part's
 * folder, producing a canonical absolute part name.
 *
 * `source = /word/document.xml`, `target = "media/image1.png"`
 *   → `/word/media/image1.png`
 * `source = /word/_rels/document.xml.rels`, `target = "../media/image1.png"`
 *   → wrong direction; rels file targets resolve against the **source part's**
 *     folder, not the .rels file's folder. Callers must pass the source part.
 */
export function resolveInternalTarget(source: PartName, target: string): PartName {
  if (target.startsWith("/")) {
    return normalizePartName(target);
  }
  return normalizePartName(`${partFolder(source)}${target}`);
}

/**
 * Inverse of {@link resolveInternalTarget}: compute the relative target URI
 * to be written in a `.rels` file given the absolute source and target parts.
 *
 * Word/Office use relative paths when source and target are siblings or
 * descendants; otherwise an absolute path (starting with `/`).
 */
export function relativizeTarget(source: PartName, target: PartName): string {
  const sourceSegs = source.split("/").slice(1, -1); // drop leading "" and filename
  const targetSegs = target.split("/").slice(1);
  let common = 0;
  while (
    common < sourceSegs.length &&
    common < targetSegs.length - 1 &&
    sourceSegs[common] === targetSegs[common]
  ) {
    common++;
  }
  const upHops = sourceSegs.length - common;
  const downSegs = targetSegs.slice(common);
  if (upHops === 0) {
    return downSegs.join("/");
  }
  return `${"../".repeat(upHops)}${downSegs.join("/")}`;
}
