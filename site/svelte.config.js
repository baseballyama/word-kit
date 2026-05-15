import adapter from "@sveltejs/adapter-static";
import { vitePreprocess } from "@sveltejs/vite-plugin-svelte";
import { mdsvex, escapeSvelte } from "mdsvex";
import { createHighlighter } from "shiki";

const theme = "github-dark";
const langs = ["ts", "tsx", "js", "json", "sh", "bash", "xml", "svelte", "html"];

const highlighter = await createHighlighter({ themes: [theme], langs });

/** @type {import('mdsvex').MdsvexOptions} */
const mdsvexOptions = {
  extensions: [".svx", ".md"],
  highlight: {
    highlighter: async (code, lang = "text") => {
      const safeLang = langs.includes(lang) ? lang : "text";
      const html = escapeSvelte(highlighter.codeToHtml(code, { lang: safeLang, theme }));
      return `{@html \`${html}\`}`;
    },
  },
};

const basePath = process.env.BASE_PATH ?? "";

/** @type {import('@sveltejs/kit').Config} */
const config = {
  extensions: [".svelte", ".svx", ".md"],
  preprocess: [vitePreprocess(), mdsvex(mdsvexOptions)],
  kit: {
    adapter: adapter({ fallback: "404.html" }),
    prerender: { entries: ["*"] },
    paths: { base: basePath, relative: true },
  },
};

export default config;
