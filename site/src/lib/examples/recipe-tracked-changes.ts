// Resolve tracked changes (Word's "Accept / Reject all changes") in
// bulk. word-kit walks every w:ins / w:del / w:moveFrom / w:moveTo
// marker in the package and either bakes it in or removes it.
//
// `acceptAllRevisions` inlines the inserted ranges and deletes the
// deleted ones. `rejectAllRevisions` does the reverse: keeps deletions
// (= drops the suggested insertions) and restores deleted ranges.

import { acceptAllRevisions, openDocx, rejectAllRevisions, toUint8Array } from "@word-kit/core";

declare const reviewedBytes: Uint8Array;

// Final draft: accept everything the editor suggested.
const accepted = openDocx(reviewedBytes);
const acceptedCount: number = acceptAllRevisions(accepted);
console.log(`accepted ${acceptedCount} revisions`);
const finalBytes: Uint8Array = toUint8Array(accepted);

// Sometimes you want the opposite — revert all proposed changes.
const reverted = openDocx(reviewedBytes);
const rejectedCount: number = rejectAllRevisions(reverted);
console.log(`rejected ${rejectedCount} revisions`);
const originalBytes: Uint8Array = toUint8Array(reverted);

void finalBytes;
void originalBytes;
