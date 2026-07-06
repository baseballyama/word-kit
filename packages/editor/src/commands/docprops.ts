/**
 * Document-property commands: core properties (title, author, subject, …) and
 * extended app properties (company, manager, counts, …). These patch the
 * existing values via `setCoreProperties` / `setAppProperties`.
 */

import {
  appProperties,
  coreProperties,
  type DocumentAppProperties,
  type DocumentCoreProperties,
  setAppProperties,
  setCoreProperties,
} from "@office-kit/docx";
import type { Command } from "./types.js";

export const setCorePropertiesCommand: Command<Partial<DocumentCoreProperties>> = {
  id: "docprops.setCore",
  group: "advanced",
  label: "Document properties",
  run(model, patch) {
    setCoreProperties(model.doc, { ...coreProperties(model.doc), ...patch });
  },
};

export const setAppPropertiesCommand: Command<Partial<DocumentAppProperties>> = {
  id: "docprops.setApp",
  group: "advanced",
  label: "Extended properties",
  run(model, patch) {
    setAppProperties(model.doc, { ...appProperties(model.doc), ...patch });
  },
};

export const docpropsCommands = [setCorePropertiesCommand, setAppPropertiesCommand];
