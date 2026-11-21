// What kind of file was opened, before it is unzipped, and the package-level restrictions of a
// .docx: Word's Restrict Editing (w:documentProtection), "read-only recommended"
// (w:writeProtection) and digital signatures (the package's _xmlsignatures parts).
import { wChild, wAttr, parseXml, serializeXml } from './xml.js';

/** Thrown for a file Intact can't open, with a message that says why in plain language. */
export class UnreadableFileError extends Error {
  constructor(kind, message) {
    super(message);
    this.name = 'UnreadableFileError';
    this.kind = kind;
  }
}

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ZIP_MAGIC = [0x50, 0x4b];

/** True if `text` (bytes read as Latin-1) holds `name` as UTF-16LE, as compound-file names are stored. */
const hasUtf16 = (text, name) => text.includes([...name].map((c) => `${c}\0`).join(''));

/**
 * Why a file can't be opened as a .docx, from its first bytes, or null if it looks like one (a ZIP
 * package). An encrypted .docx is not a ZIP but an OLE compound file holding the encrypted package
 * (MS-OFFCRYPTO): with a password (EncryptionInfo, StrongEncryption) or by rights management
 * (Azure / Purview sensitivity labels that encrypt, and IRM: the DRM transform). A Word 97–2003
 * .doc is a compound file too, with a WordDocument stream.
 * @param {Uint8Array} bytes
 */
export function sniffFile(bytes) {
  if (ZIP_MAGIC.every((b, i) => bytes[i] === b)) return null;
  if (!CFB_MAGIC.every((b, i) => bytes[i] === b)) {
    return new UnreadableFileError('not-docx', 'This file isn’t a .docx document. Every .docx is a ZIP package, and this one isn’t: it may be damaged, or another kind of file with a .docx name.');
  }
  // Stream names are in the compound file's directory, anywhere in the file: search all of it.
  const text = new TextDecoder('latin1').decode(bytes);
  if (hasUtf16(text, 'DRMEncrypted') || hasUtf16(text, 'DRMTransform') || text.includes('{C73DFACD-061F-43B0-8B64-0C620D2A8B50}')
    || hasUtf16(text, '{C73DFACD-061F-43B0-8B64-0C620D2A8B50}')) {
    return new UnreadableFileError('rights-managed', 'This document is protected by rights management: a sensitivity label that encrypts it, or Information Rights Management. Only Microsoft Office, signed in with an account that has permission, can open it, so Intact can’t and won’t try. Open it in Word, or ask its owner for a copy without encryption if policy allows.');
  }
  if (hasUtf16(text, 'EncryptionInfo') || hasUtf16(text, 'EncryptedPackage')) {
    // Opened with its password (see DocxDocument.open and encryption.js), which is asked for.
    return new UnreadableFileError('password', 'This document is encrypted with a password: it opens with its password, which is checked on this device.');
  }
  if (hasUtf16(text, 'WordDocument')) {
    return new UnreadableFileError('doc', 'This is a Word 97–2003 document (.doc) with a .docx name. Intact opens .docx only: in Word, save it as a Word Document (.docx) with File › Save As, then open that.');
  }
  return new UnreadableFileError('not-docx', 'This file isn’t a .docx document: it is an older Office file of another kind. Intact opens .docx files only.');
}

// ------------------------------------------------------------------ Restrict Editing

/** Word's names for the editing restrictions (Review › Restrict Editing). */
export const PROTECTION_LABELS = {
  readOnly: 'No changes (read only)',
  comments: 'Comments only',
  trackedChanges: 'Tracked changes only',
  forms: 'Filling in forms only',
};

/**
 * The editing restriction Word enforces on the document (w:documentProtection in settings.xml), or
 * null: { mode, password }. `mode` is readOnly, comments, trackedChanges or forms; `password` says
 * whether Word asks for a password to stop the protection. Only enforced protection counts, as in
 * Word: a restriction set up but not started (w:enforcement off) restricts nothing.
 */
export function readProtection(settingsRoot) {
  const el = wChild(settingsRoot, 'documentProtection');
  if (!el || attrOn(el, 'enforcement') !== true) return null;
  const mode = wAttr(el, 'edit');
  if (!PROTECTION_LABELS[mode]) return null;
  return { mode, password: hasPassword(el) };
}

/**
 * "Read-only recommended" or a password to modify (w:writeProtection), or null: { recommended,
 * password }. In Word both still let anyone open the document and save their edits as a copy.
 */
export function readWriteProtection(settingsRoot) {
  const el = wChild(settingsRoot, 'writeProtection');
  if (!el) return null;
  const recommended = attrOn(el, 'recommended') === true;
  const password = hasPassword(el);
  return recommended || password ? { recommended, password } : null;
}

/** An on/off attribute (ST_OnOff): true, false, or undefined when absent. */
function attrOn(el, name) {
  const v = wAttr(el, name);
  if (v == null || v === '') return undefined;
  return !(v === '0' || v === 'false' || v === 'off');
}

function hasPassword(el) {
  return ['hash', 'hashValue', 'cryptProviderType'].some((a) => wAttr(el, a));
}

// ------------------------------------------------------------------ digital signatures

const SIG_ORIGIN = 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin';
const SIG_PART = 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature';
const PKG_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/** A path inside the package from a relationship target, relative to the part in `dir`. */
function resolve(dir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = dir ? dir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

const relsPathOf = (path) => {
  const slash = path.lastIndexOf('/');
  return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`;
};

/**
 * The package's digital signatures: { originPath, parts: [paths], count }, or null if it has none.
 * `parts` are every part and relationships part that belongs to them, to take out together.
 */
export async function readSignatures(zip) {
  const rels = zip.file('_rels/.rels');
  if (!rels) return null;
  const relsDoc = parseXml(await rels.async('string'));
  const origin = [...relsDoc.getElementsByTagNameNS(PKG_RELS, 'Relationship')].find((r) => r.getAttribute('Type') === SIG_ORIGIN);
  if (!origin) return null;
  const originPath = resolve('', origin.getAttribute('Target') || '');
  const parts = [originPath];
  let count = 0;
  const originRels = relsPathOf(originPath);
  const sigRels = zip.file(originRels);
  if (sigRels) {
    parts.push(originRels);
    const dir = originPath.includes('/') ? originPath.slice(0, originPath.lastIndexOf('/')) : '';
    for (const r of parseXml(await sigRels.async('string')).getElementsByTagNameNS(PKG_RELS, 'Relationship')) {
      if (r.getAttribute('TargetMode') === 'External') continue;
      const path = resolve(dir, r.getAttribute('Target') || '');
      if (r.getAttribute('Type') === SIG_PART) count++;
      parts.push(path);
      if (zip.file(relsPathOf(path))) parts.push(relsPathOf(path));
    }
  }
  // An origin with no signatures left is what Word leaves after removing them all: nothing is signed.
  return count ? { originPath, parts: parts.filter((p) => zip.file(p)), count } : null;
}

/**
 * Takes the digital signatures out of the package, as Word does when a signed document is edited:
 * their parts, the package relationship to them and their content types. Signature lines in the
 * text stay, unsigned.
 */
export async function removeSignatures(zip, signatures) {
  for (const path of signatures.parts) zip.remove(path);
  // Folders left empty (_xmlsignatures/, its _rels/) go too.
  const files = Object.values(zip.files).filter((f) => !f.dir).map((f) => f.name);
  for (const dir of Object.values(zip.files).filter((f) => f.dir).map((f) => f.name)) {
    if (!files.some((name) => name.startsWith(dir))) zip.remove(dir);
  }
  const relsDoc = parseXml(await zip.file('_rels/.rels').async('string'));
  for (const r of [...relsDoc.getElementsByTagNameNS(PKG_RELS, 'Relationship')]) {
    if (r.getAttribute('Type') === SIG_ORIGIN) r.remove();
  }
  zip.file('_rels/.rels', serializeXml(relsDoc));
  const ct = zip.file('[Content_Types].xml');
  if (!ct) return;
  const ctDoc = parseXml(await ct.async('string'));
  const removed = new Set(signatures.parts.map((p) => `/${p}`.toLowerCase()));
  for (const o of [...ctDoc.getElementsByTagNameNS(CT_NS, 'Override')]) {
    if (removed.has((o.getAttribute('PartName') || '').toLowerCase())) o.remove();
  }
  // The .sigs extension is only for signature origins: no part needs its default any more.
  const sigsLeft = Object.keys(zip.files).some((p) => /\.sigs$/i.test(p));
  for (const d of [...ctDoc.getElementsByTagNameNS(CT_NS, 'Default')]) {
    if (!sigsLeft && (d.getAttribute('Extension') || '').toLowerCase() === 'sigs') d.remove();
  }
  zip.file('[Content_Types].xml', serializeXml(ctDoc));
}
