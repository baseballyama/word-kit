import { getPart, hasPart } from "@word-kit/opc";
import { describe, expect, it } from "vitest";
import { Docx, MARGINS_NORMAL, PAGE_SIZE_A4 } from "./index.js";

const TINY_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0xfa, 0xcf, 0x00, 0x00,
  0x00, 0x02, 0x00, 0x01, 0xe5, 0x27, 0xde, 0xfc, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
  0xae, 0x42, 0x60, 0x82,
]);

describe("kitchen sink — every Docx feature in one document", () => {
  it("constructs, edits, saves, reopens, and verifies a complex docx", () => {
    const doc = Docx.create({ paragraphs: [] });

    // Page setup
    doc.setPageSize(PAGE_SIZE_A4);
    doc.setPageMargins(MARGINS_NORMAL);
    doc.setPageOrientation("portrait");

    // Headers and footers
    doc.addHeader("Confidential — {{company}}");
    doc.addFooter("Page");

    // Custom style
    doc.addStyle({
      type: "paragraph",
      styleId: "MyHeading",
      name: "My Heading",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      bold: true,
      fontSizeHalfPoints: 32,
      color: "1F497D",
    });

    // Title paragraph with style
    doc.appendParagraph("Quarterly Report — {{quarter}} {{year}}", {
      style: "MyHeading",
      bold: true,
    });

    // Plain paragraphs with template placeholders
    doc.appendParagraph("Prepared by {{author}}.");
    doc.appendParagraph("Submitted to {{recipient}}.");

    // Comment on a paragraph
    const para = doc.paragraphs.at(-1);
    if (para) {
      doc.addComment(para, {
        author: "Reviewer",
        initials: "R",
        text: "Confirm the recipient name before sending.",
      });
    }

    // Bullet list
    doc.addBulletList(["Revenue ↑ 12%", "Costs ↓ 3%", "Headcount: 47"]);
    // Numbered list
    doc.addNumberedList(["Plan Q3", "Hire 5 engineers", "Ship v2.0"]);

    // Table
    doc.addTable([
      ["Metric", "Q1", "Q2", "Q3"],
      ["Users", "1000", "1500", "2200"],
      ["MRR", "$10k", "$15k", "$24k"],
    ]);

    // Image
    doc.addImage(TINY_PNG, {
      widthEmu: 914400,
      heightEmu: 914400,
      altText: "logo",
    });

    // Hyperlink
    doc.addHyperlink("https://example.com/q3-report", "Q3 full report", {
      tooltip: "Open in browser",
    });

    // Template replacement (run-spanning safe)
    expect(
      doc.replaceText(/\{\{(\w+)\}\}/g, (m) => {
        const dict: Record<string, string> = {
          company: "Acme",
          quarter: "Q3",
          year: "2026",
          author: "Yamada",
          recipient: "Board",
        };
        return dict[m.captures[0] ?? ""] ?? m.text;
      }),
    ).toBeGreaterThan(0);

    // Save and reopen
    const bytes = doc.toUint8Array();
    expect(bytes.byteLength).toBeGreaterThan(0);
    const reopened = Docx.open(bytes);

    // Verify everything came back
    expect(reopened.paragraphs.length).toBeGreaterThan(5);
    expect(reopened.tables).toHaveLength(1);
    expect(reopened.tables[0]?.rows).toHaveLength(3);
    expect(hasPart(reopened.opc, "/word/styles.xml")).toBe(true);
    expect(hasPart(reopened.opc, "/word/numbering.xml")).toBe(true);
    expect(hasPart(reopened.opc, "/word/comments.xml")).toBe(true);
    expect(hasPart(reopened.opc, "/word/header1.xml")).toBe(true);
    expect(hasPart(reopened.opc, "/word/footer1.xml")).toBe(true);
    expect(hasPart(reopened.opc, "/word/media/image1.png")).toBe(true);
    expect(
      reopened.stylesPart?.styles.some((s) =>
        s.attrs.some((a) => a.name.local === "styleId" && a.value === "MyHeading"),
      ),
    ).toBe(true);
    expect(reopened.commentsPart?.comments).toHaveLength(1);
    // Template placeholders inside document.xml (title and plain paragraphs)
    // are resolved by replaceText.
    expect(reopened.text).toContain("Q3 2026");
    expect(reopened.text).toContain("Yamada");
    expect(reopened.text).toContain("Board");
    // Template placeholders inside header/footer parts are not part of
    // document.xml's flat text — verify them by sniffing the header bytes.
    const headerPart = getPart(reopened.opc, "/word/header1.xml");
    expect(new TextDecoder().decode(headerPart?.data ?? new Uint8Array())).toContain("{{company}}");
  });

  it("docx.toBlob() yields a properly-typed Blob", () => {
    const doc = Docx.create();
    const blob = doc.toBlob();
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    expect(blob.size).toBeGreaterThan(0);
  });
});
