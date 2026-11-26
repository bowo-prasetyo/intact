// Holds the currently open document. It is a shallowRef because a DocxDocument wraps XML DOM
// trees that must not be made deeply reactive.
import { shallowRef } from 'vue';
import { lang, t } from './i18n.js';

// The document's engine (DocxDocument, with JSZip, the XML, encryption and the rest) is loaded the
// first time a document is opened or started, not with the start page: the start page shows at once
// ("Instant"), and the engine comes in the same chunk the editor uses. A test keeps the start page
// within its budget (tests/start-page-size.spec.js).
let engine = null;
const docx = () => (engine ||= import('./docx/DocxDocument.js').then((m) => m.DocxDocument));
/** Starts loading the engine without waiting for it (a hint: the start page is up and idle). */
export const preloadEngine = () => { docx().catch(() => { engine = null; }); };

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
  const DocxDocument = await docx();
  const doc = await DocxDocument.open(await file.arrayBuffer(), file.name, { password });
  doc.fileHandle = handle;
  setDocument(doc);
}

/** A blank document: its text in the interface's language (Indonesian, or English), so spelling is checked in it. */
export async function newDocument() {
  const language = lang.value === 'id' ? 'id-ID' : 'en-US';
  const [DocxDocument, { blankDocxFiles }] = await Promise.all([docx(), import('./docx/blank.js')]);
  setDocument(await DocxDocument.createBlank(`${t('Untitled')}.docx`, blankDocxFiles({ language })));
}

/** A new document from a starter (see docx/starters.js): a letter, a memo, minutes, a quote… */
export async function newFromStarter(id) {
  const DocxDocument = await docx();
  setDocument(await DocxDocument.createStarter(id));
}

/**
 * Opens a draft kept in this browser (see drafts.js) where its editing left off: with its unsaved
 * changes, the tracked changes made then still counted as this editing's own, and "Edit anyway"
 * still on if it was. Its edits keep updating the same draft. A password-protected document's draft
 * is encrypted with its password, which `password` gives.
 */
export async function openDraft(draft, password = null) {
  const DocxDocument = await docx();
  const doc = await DocxDocument.open(draft.blob, draft.name, { password });
  doc.restored = true;
  doc.draftId = draft.id;
  doc.restoredFrom = { savedAt: draft.savedAt, unlocked: !!draft.unlocked, length: draft.length || null };
  // Revision ids from the first opening's own on are this editing's (see firstNewId).
  if (Number.isInteger(draft.firstNewId) && draft.firstNewId < doc.firstNewId) doc.firstNewId = draft.firstNewId;
  setDocument(doc);
}
