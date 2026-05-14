import { OpcPackage } from "./package.js";

const WML_DOC_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const OFFICE_DOCUMENT_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const MINIMAL_DOCUMENT_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  "\r\n",
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  "<w:body><w:p/></w:body>",
  "</w:document>",
].join("");

/**
 * Build a minimal but valid `.docx` package containing a single empty
 * paragraph. The result opens cleanly in Microsoft Word and LibreOffice
 * Writer.
 *
 * Used as a test fixture and as the seed package for `Docx.create()` once
 * the higher-level API is wired up.
 */
export function buildMinimalDocx(): OpcPackage {
  const pkg = OpcPackage.empty();
  pkg.contentTypes.setDefault("rels", "application/vnd.openxmlformats-package.relationships+xml");
  pkg.contentTypes.setDefault("xml", "application/xml");
  pkg.contentTypes.setOverride("/word/document.xml", WML_DOC_CT);

  pkg.packageRelationships.add({
    id: "rId1",
    type: OFFICE_DOCUMENT_REL,
    target: "word/document.xml",
  });

  pkg.addPart({
    name: "/word/document.xml",
    contentType: WML_DOC_CT,
    data: new TextEncoder().encode(MINIMAL_DOCUMENT_XML),
  });

  return pkg;
}
