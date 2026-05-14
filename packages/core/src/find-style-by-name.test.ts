import { describe, expect, it } from "vitest";
import { addStyle, createDocx, findStyleIdByName } from "./docx.js";

describe("findStyleIdByName", () => {
  it("resolves a style by its display name", () => {
    const doc = createDocx({ paragraphs: [] });
    addStyle(doc, { type: "paragraph", styleId: "MyHeading", name: "My Heading" });
    expect(findStyleIdByName(doc, "My Heading")).toBe("MyHeading");
  });

  it("returns undefined for an unknown display name", () => {
    const doc = createDocx({ paragraphs: [] });
    expect(findStyleIdByName(doc, "No Such Style")).toBeUndefined();
  });

  it("is case-sensitive on the display name", () => {
    const doc = createDocx({ paragraphs: [] });
    addStyle(doc, { type: "paragraph", styleId: "Sub", name: "Subtitle" });
    expect(findStyleIdByName(doc, "Subtitle")).toBe("Sub");
    expect(findStyleIdByName(doc, "subtitle")).toBeUndefined();
  });
});
