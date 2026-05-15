import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  sourcemap: true,
  clean: true,
  target: "es2022",
  treeshake: true,
  // docx-preview ships only as ESM and resolves JSZip in its own bundle;
  // do not pull it into our dist — let it remain a runtime peer.
  deps: { neverBundle: ["docx-preview"] },
});
