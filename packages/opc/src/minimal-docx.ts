import { setContentTypeDefault, setContentTypeOverride } from "./content-types.js";
import { addPart, emptyOpcPackage, type OpcPackage, packageRelationships } from "./package.js";
import { addRelationship } from "./relationships.js";

const WML_DOC_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const OFFICE_DOCUMENT_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const MINIMAL_DOCUMENT_XML = [
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
  "\r\n",
  "<w:document ",
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">',
  "<w:body><w:p/></w:body>",
  "</w:document>",
].join("");

/**
 * Build a minimal but valid `.docx` package containing a single empty
 * paragraph. The result opens cleanly in Microsoft Word and LibreOffice
 * Writer.
 *
 * Used as a test fixture and as the seed package for `createDocx()` from
 * `@word-kit/core`.
 */
export function buildMinimalDocx(): OpcPackage {
  const pkg = emptyOpcPackage();
  setContentTypeDefault(
    pkg.contentTypes,
    "rels",
    "application/vnd.openxmlformats-package.relationships+xml",
  );
  setContentTypeDefault(pkg.contentTypes, "xml", "application/xml");
  setContentTypeOverride(pkg.contentTypes, "/word/document.xml", WML_DOC_CT);

  addRelationship(packageRelationships(pkg), {
    id: "rId1",
    type: OFFICE_DOCUMENT_REL,
    target: "word/document.xml",
  });

  addPart(pkg, {
    name: "/word/document.xml",
    contentType: WML_DOC_CT,
    data: new TextEncoder().encode(MINIMAL_DOCUMENT_XML),
  });

  return pkg;
}
