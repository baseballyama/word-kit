import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  sourcemap: true,
  clean: true,
  target: "es2022",
  treeshake: true,
  // Keep the core library a runtime dependency rather than inlining it, so the
  // editor and the app share one copy of `@office-kit/docx`.
  deps: { neverBundle: ["@office-kit/docx"] },
});
