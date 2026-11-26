// Holds the currently open document. It is a shallowRef because a DocxDocument wraps XML DOM
// trees that must not be made deeply reactive.
import { shallowRef } from 'vue';
import { DocxDocument } from './docx/DocxDocument.js';
import { blankDocxFiles } from './docx/blank.js';
import { healthcareStarterFiles } from './docx/healthcareStarters.js';
import { lang, t } from './i18n.js';

export const currentDoc = shallowRef(null);

export function setDocument(doc) {
  currentDoc.value?.dispose();
  currentDoc.value = doc;
}

/**
 * Opens `file`. `handle`, where the browser gave one (see editor/fileHandles.js), is the file on
 * the device that Save writes back to.
 */
export async function openFile(file, handle = null, password = null) {
  const doc = await DocxDocument.open(await file.arrayBuffer(), file.name, { password });
  doc.fileHandle = handle;
  setDocument(doc);
}

/** A blank document: its text in the interface's language (Indonesian, or English), so spelling is checked in it. */
export async function newDocument() {
  const language = lang.value === 'id' ? 'id-ID' : 'en-US';
  setDocument(await DocxDocument.createBlank(`${t('Untitled')}.docx`, blankDocxFiles({ language })));
}

/** A new document from a starter: the healthcare administration set is built locally too. */
export async function newFromStarter(id) {
  const healthcareFiles = healthcareStarterFiles(id);
  if (healthcareFiles) {
    const title = id.replace(/^healthcare-/, '').replace(/-/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase());
    setDocument(await DocxDocument.createBlank(`${title}.docx`, healthcareFiles));
    return;
  }
  setDocument(await DocxDocument.createStarter(id));
}

/**
 * Opens a draft kept in this browser (see drafts.js) where its editing left off: with its unsaved
 * changes, the tracked changes made then still counted as this editing's own, and "Edit anyway"
 * still on if it was. Its edits keep updating the same draft. A password-protected document's draft
 * is encrypted with its password, which `password` gives.
 */
export async function openDraft(draft, password = null) {
  const doc = await DocxDocument.open(draft.blob, draft.name, { password });
  doc.restored = true;
  doc.draftId = draft.id;
  doc.restoredFrom = { savedAt: draft.savedAt, unlocked: !!draft.unlocked, length: draft.length || null };
  // Revision ids from the first opening's own on are this editing's (see firstNewId).
  if (Number.isInteger(draft.firstNewId) && draft.firstNewId < doc.firstNewId) doc.firstNewId = draft.firstNewId;
  setDocument(doc);
}
