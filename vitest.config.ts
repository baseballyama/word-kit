import { resolve } from "node:path";
import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      // Let packages/preview/src/* import from '@office-kit/docx' without a
      // built dist/. The alias wires the package name to the TypeScript source
      // so vitest exercises the same code the tests cover.
      "@office-kit/docx": resolve("./src/index.ts"),
    },
  },
  test: {
    include: ["src/**/*.test.ts", "packages/*/src/**/*.test.ts"],
    coverage: {
      provider: "v8",
      reporter: ["text", "html"],
      include: ["src/**/*.ts", "packages/*/src/**/*.ts"],
      exclude: ["**/*.test.ts", "**/dist/**"],
    },
  },
});
