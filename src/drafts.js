// Drafts: a copy of the document being edited, kept in this browser (IndexedDB) while it has changes
// that haven't been downloaded, so a reload, a crash or a closed tab doesn't lose them. Nothing
// leaves the device. A draft goes when its document is downloaded or closed, when it is discarded
// from the start page, and in any case after MAX_AGE. Storage can be missing or refused (private
// windows, blocked site data): every call then quietly does nothing, and editing works as before.

const DB = 'intact';
const STORE = 'drafts';
export const MAX_AGE = 7 * 24 * 60 * 60 * 1000; // a week
const MAX_DRAFTS = 5;

let opening = null;
function db() {
  opening ||= new Promise((resolve, reject) => {
    const req = indexedDB.open(DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
    req.onblocked = () => reject(new Error('Draft storage is blocked'));
  }).catch((e) => { opening = null; throw e; });
  return opening;
}

async function run(mode, work) {
  const conn = await db();
  return new Promise((resolve, reject) => {
    const tx = conn.transaction(STORE, mode);
    const result = work(tx.objectStore(STORE));
    tx.oncomplete = () => resolve(result?.result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

const quietly = (promise, fallback) => promise.catch((e) => { console.warn('Drafts:', e); return fallback; });

// The document is kept as its bytes (an ArrayBuffer), not as a Blob: Safari's engine can't keep a
// Blob in IndexedDB in a private window (the transaction aborts), and bytes are kept everywhere. A
// draft read back has its Blob again; one kept as a Blob by an earlier version is read as it is.
async function stored({ blob, ...draft }) {
  return blob ? { ...draft, bytes: await blob.arrayBuffer(), type: blob.type } : draft;
}
const withBlob = ({ bytes, type, ...draft }) => (bytes ? { ...draft, blob: new Blob([bytes], { type: type || '' }) } : draft);

/**
 * Keeps a draft: { id, name, blob (kept as its bytes: see stored), savedAt, firstNewId, unlocked, length (a length target), protected
 * (encrypted with the document's password: see DocxDocument.draft) }. Older drafts beyond the most
 * recent few go. Resolves to whether it was kept.
 */
export function saveDraft(draft) {
  return quietly(stored(draft).then((record) => run('readwrite', (store) => store.put({ ...record, savedAt: Date.now() }))).then(async () => {
    const all = await listDrafts();
    await Promise.all(all.slice(MAX_DRAFTS).map((d) => deleteDraft(d.id)));
    return true;
  }), false);
}

/** The drafts kept, newest first; drafts older than MAX_AGE are deleted instead. */
export function listDrafts() {
  return quietly(run('readonly', (store) => store.getAll()).then(async (all = []) => {
    const now = Date.now();
    const old = all.filter((d) => !(now - d.savedAt < MAX_AGE));
    await Promise.all(old.map((d) => deleteDraft(d.id)));
    return all.filter((d) => !old.includes(d)).sort((a, b) => b.savedAt - a.savedAt).map(withBlob);
  }), []);
}

export function deleteDraft(id) {
  return quietly(run('readwrite', (store) => store.delete(id)).then(() => true), false);
}

export function newDraftId() {
  return globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}
