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
