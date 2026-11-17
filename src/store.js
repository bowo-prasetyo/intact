// Holds the currently open document. It is a shallowRef because a DocxDocument wraps XML DOM
// trees that must not be made deeply reactive.
import { shallowRef } from 'vue';
import { DocxDocument } from './docx/DocxDocument.js';

export const currentDoc = shallowRef(null);

export function setDocument(doc) {
  currentDoc.value?.dispose();
  currentDoc.value = doc;
}

export async function openFile(file) {
  setDocument(await DocxDocument.open(await file.arrayBuffer(), file.name));
}

export async function newDocument() {
  setDocument(await DocxDocument.createBlank('Untitled.docx'));
}

/**
 * Opens a draft kept in this browser (see drafts.js) where its editing left off: with its unsaved
 * changes, the tracked changes made then still counted as this editing's own, and "Edit anyway"
 * still on if it was. Its edits keep updating the same draft.
 */
export async function openDraft(draft) {
  const doc = await DocxDocument.open(draft.blob, draft.name);
  doc.restored = true;
  doc.draftId = draft.id;
  doc.restoredFrom = { savedAt: draft.savedAt, unlocked: !!draft.unlocked };
  // Revision ids from the first opening's own on are this editing's (see firstNewId).
  if (Number.isInteger(draft.firstNewId) && draft.firstNewId < doc.firstNewId) doc.firstNewId = draft.firstNewId;
  setDocument(doc);
}
