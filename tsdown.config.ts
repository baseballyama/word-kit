import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { defineConfig } from "tsdown";

const pkg = JSON.parse(
  readFileSync(fileURLToPath(new URL("./package.json", import.meta.url)), "utf8"),
) as { version: string };

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  sourcemap: true,
  clean: true,
  target: "es2022",
  treeshake: true,
  // Replace `__DOCX_KIT_VERSION__` in the source with the literal version
  // string from package.json so VERSION never drifts from the published tag.
  define: {
    __DOCX_KIT_VERSION__: JSON.stringify(pkg.version),
  },
});
