import type { XmlAttr, XmlElement } from "@word-kit/ooxml-xml";
import { WML_NS } from "./namespaces.js";

export type NumberFormat =
  | "decimal"
  | "upperRoman"
  | "lowerRoman"
  | "upperLetter"
  | "lowerLetter"
  | "ordinal"
  | "cardinalText"
  | "ordinalText"
  | "bullet"
  | "none";

export interface BuildAbstractNumLevelOptions {
  /** Indentation level (0-8). */
  readonly ilvl: number;
  /** Starting number. */
  readonly start?: number;
  /** Number format. */
  readonly numFmt: NumberFormat;
  /** Display text. `%1` is the current level number; `%2` the next, etc. */
  readonly lvlText: string;
  /** Justification of the level marker. */
  readonly lvlJc?: "left" | "center" | "right";
  /** Left indent in twips. */
  readonly indentLeft?: number;
  /** Hanging indent in twips. */
  readonly indentHanging?: number;
  /** Bullet font (only meaningful for bullet levels). */
  readonly bulletFont?: string;
}

export interface BuildAbstractNumOptions {
  readonly abstractNumId: number;
  readonly levels: readonly BuildAbstractNumLevelOptions[];
}

/** Construct a `<w:abstractNum>` element with the given level definitions. */
export function buildAbstractNum(options: BuildAbstractNumOptions): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "abstractNum", prefix: "w" },
    attrs: [wmlAttr("abstractNumId", String(options.abstractNumId))],
    children: options.levels.map((lvl) => buildLvl(lvl)),
    xmlSpace: "default",
    selfClosing: false,
  };
}

function buildLvl(level: BuildAbstractNumLevelOptions): XmlElement {
  const children: XmlElement[] = [];
  if (level.start !== undefined) {
    children.push(wmlSelfClose("start", [wmlAttr("val", String(level.start))]));
  }
  children.push(wmlSelfClose("numFmt", [wmlAttr("val", level.numFmt)]));
  children.push(wmlSelfClose("lvlText", [wmlAttr("val", level.lvlText)]));
  if (level.lvlJc) {
    children.push(wmlSelfClose("lvlJc", [wmlAttr("val", level.lvlJc)]));
  }
  if (level.indentLeft !== undefined || level.indentHanging !== undefined) {
    const indAttrs: XmlAttr[] = [];
    if (level.indentLeft !== undefined) {
      indAttrs.push(wmlAttr("left", String(level.indentLeft)));
    }
    if (level.indentHanging !== undefined) {
      indAttrs.push(wmlAttr("hanging", String(level.indentHanging)));
    }
    children.push(wmlElement("pPr", [], [wmlSelfClose("ind", indAttrs)]));
  }
  if (level.bulletFont) {
    children.push(
      wmlElement(
        "rPr",
        [],
        [
          wmlSelfClose("rFonts", [
            wmlAttr("ascii", level.bulletFont),
            wmlAttr("hAnsi", level.bulletFont),
            wmlAttr("hint", "default"),
          ]),
        ],
      ),
    );
  }
  return {
    kind: "element",
    name: { uri: WML_NS, local: "lvl", prefix: "w" },
    attrs: [wmlAttr("ilvl", String(level.ilvl))],
    children,
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Construct a `<w:num>` element that references the given abstract numbering. */
export function buildNum(numId: number, abstractNumId: number): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "num", prefix: "w" },
    attrs: [wmlAttr("numId", String(numId))],
    children: [wmlSelfClose("abstractNumId", [wmlAttr("val", String(abstractNumId))])],
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Bullet-list abstract definition with sensible defaults across 9 levels. */
export function bulletAbstractNumLevels(): BuildAbstractNumLevelOptions[] {
  const markers = ["", "o", "", "", "o", "", "", "o", ""];
  return Array.from({ length: 9 }, (_, i) => ({
    ilvl: i,
    numFmt: "bullet" as const,
    lvlText: markers[i] ?? "",
    lvlJc: "left" as const,
    indentLeft: 720 * (i + 1),
    indentHanging: 360,
    bulletFont: "Symbol",
  }));
}

/** Decimal-numbered abstract definition (1., 2., 3., …) with 9 levels. */
export function decimalAbstractNumLevels(): BuildAbstractNumLevelOptions[] {
  return Array.from({ length: 9 }, (_, i) => ({
    ilvl: i,
    start: 1,
    numFmt: "decimal" as const,
    lvlText: `%${i + 1}.`,
    lvlJc: "left" as const,
    indentLeft: 720 * (i + 1),
    indentHanging: 360,
  }));
}

/** Minimal stand-alone `numbering.xml` skeleton (no abstractNums or nums). */
export const EMPTY_NUMBERING_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
].join("");

/** Build `<w:numPr>` to attach to a paragraph's `<w:pPr>`. */
export function buildNumPr(numIdValue: number, ilvl: number): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local: "numPr", prefix: "w" },
    attrs: [],
    children: [
      wmlSelfClose("ilvl", [wmlAttr("val", String(ilvl))]),
      wmlSelfClose("numId", [wmlAttr("val", String(numIdValue))]),
    ],
    xmlSpace: "default",
    selfClosing: false,
  };
}

/** Build `<w:pPr>` with an embedded `<w:numPr>` and optional pStyle. */
export function buildPPrWithNumPr(numIdValue: number, ilvl: number, pStyle?: string): XmlElement {
  const children: XmlElement[] = [];
  if (pStyle) {
    children.push(wmlSelfClose("pStyle", [wmlAttr("val", pStyle)]));
  }
  children.push(buildNumPr(numIdValue, ilvl));
  return {
    kind: "element",
    name: { uri: WML_NS, local: "pPr", prefix: "w" },
    attrs: [],
    children,
    xmlSpace: "default",
    selfClosing: false,
  };
}

function wmlAttr(local: string, value: string): XmlAttr {
  return { name: { uri: WML_NS, local, prefix: "w" }, value, isNamespaceDecl: false };
}

function wmlSelfClose(local: string, attrs: XmlAttr[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children: [],
    xmlSpace: "default",
    selfClosing: true,
  };
}

function wmlElement(local: string, attrs: XmlAttr[], children: XmlElement[]): XmlElement {
  return {
    kind: "element",
    name: { uri: WML_NS, local, prefix: "w" },
    attrs,
    children,
    xmlSpace: "default",
    selfClosing: false,
  };
}
