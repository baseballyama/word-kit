/**
 * Review commands: add a comment anchored to the caret paragraph, and accept or
 * reject all tracked revisions in the document.
 */

import {
  acceptAllRevisions,
  addComment,
  type AddCommentOptions,
  rejectAllRevisions,
} from "@office-kit/docx";
import { paragraphAt } from "../doc-access.js";
import type { EditorModel } from "../model.js";
import type { Command } from "./types.js";

function caretParagraph(model: EditorModel) {
  const pos = model.selection?.focus;
  return pos ? paragraphAt(model.doc, pos) : undefined;
}

export const addCommentCommand: Command<AddCommentOptions> = {
  id: "review.addComment",
  group: "review",
  label: "New comment",
  run(model, options) {
    const para = caretParagraph(model);
    if (para) addComment(model.doc, para, options);
  },
  isEnabled: (model) => !!caretParagraph(model),
};

export const acceptAllRevisionsCommand: Command<void> = {
  id: "review.acceptAll",
  group: "review",
  label: "Accept all changes",
  run(model) {
    acceptAllRevisions(model.doc);
  },
};

export const rejectAllRevisionsCommand: Command<void> = {
  id: "review.rejectAll",
  group: "review",
  label: "Reject all changes",
  run(model) {
    rejectAllRevisions(model.doc);
  },
};

export const reviewCommands = [
  addCommentCommand,
  acceptAllRevisionsCommand,
  rejectAllRevisionsCommand,
];
