// @vitest-environment happy-dom
//
// happy-dom (rather than jsdom) is the test DOM here because jsdom's
// global `TextEncoder` returns a `Uint8Array` from a different realm
// than the test's, so `instanceof Uint8Array` fails inside fflate's
// internal type guards and the produced zip is malformed (path/0/,
// path/1/ … directory-marker entries instead of the requested file).
// happy-dom shares the global Uint8Array with the test realm, so
// fflate produces correct output and docx-preview can parse it.

import { appendParagraph, createDocx, toUint8Array, type Docx } from "@word-kit/core";
import { describe, expect, it } from "vitest";
import { previewToDOM } from "./index.js";

function freshDoc(text: string): Docx {
  const doc = createDocx({ paragraphs: [] });
  appendParagraph(doc, text);
  return doc;
}

describe("previewToDOM (v0 bridge over docx-preview)", () => {
  it("renders a fresh Docx into the supplied container", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const handle = await previewToDOM(freshDoc("Hello, preview!"), container);
    expect(container.textContent).toContain("Hello, preview!");
    handle.dispose();
    container.remove();
  });

  it("accepts raw bytes (Uint8Array) as well as a Docx value", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const bytes = toUint8Array(freshDoc("From bytes"));
    const handle = await previewToDOM(bytes, container);
    expect(container.textContent).toContain("From bytes");
    handle.dispose();
    container.remove();
  });

  it("dispose() removes the rendered DOM and is idempotent", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const handle = await previewToDOM(freshDoc("disposable"), container);
    expect(container.querySelector("[data-word-kit-preview]")).not.toBeNull();
    handle.dispose();
    expect(container.querySelector("[data-word-kit-preview]")).toBeNull();
    expect(() => handle.dispose()).not.toThrow();
    container.remove();
  });

  it("two simultaneous previews on the same container coexist", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const a = await previewToDOM(freshDoc("aaaaa"), container);
    const b = await previewToDOM(freshDoc("bbbbb"), container);
    expect(container.textContent).toContain("aaaaa");
    expect(container.textContent).toContain("bbbbb");
    expect(container.querySelectorAll("[data-word-kit-preview]")).toHaveLength(2);
    a.dispose();
    b.dispose();
    expect(container.querySelectorAll("[data-word-kit-preview]")).toHaveLength(0);
    container.remove();
  });

  it("custom classPrefix is honoured by the underlying renderer", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);
    const handle = await previewToDOM(freshDoc("prefixed"), container, {
      classPrefix: "custom-",
    });
    const classed = container.querySelectorAll('[class*="custom-"]');
    expect(classed.length).toBeGreaterThan(0);
    handle.dispose();
    container.remove();
  });
});
