// Saving back to the file that was opened, where the browser can (the File System Access API:
// Chrome and Edge on a computer): the file is opened through the browser's own file picker (or
// dropped), which hands the page a handle to it, and Save writes the new document into that file,
// as Word's Save does, instead of downloading a copy. Save as… picks a file to save to the same way.
// Nothing leaves the device: the browser writes the file where it is. Elsewhere (Firefox, Safari,
// phones) the page has no such handle, and Download keeps a copy in Downloads, as before.

const MIMES = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  docm: 'application/vnd.ms-word.document.macroEnabled.12',
  dotm: 'application/vnd.ms-word.template.macroEnabled.12',
};

/** Whether files can be opened with a handle to save back to. */
export const canOpenWithHandle = () => typeof window !== 'undefined' && typeof window.showOpenFilePicker === 'function';

/** Whether Save as… can pick a file on this device to save to. */
export const canPickSaveFile = () => typeof window !== 'undefined' && typeof window.showSaveFilePicker === 'function';

/**
 * Asks for a Word file with the browser's file picker: { file, handle }, or null if none was chosen.
 * Must be called from a click or a key press.
 */
export async function pickFileToOpen() {
  try {
    const [handle] = await window.showOpenFilePicker({
      id: 'intact-open',
      types: [{ description: 'Word documents and templates', accept: Object.fromEntries(Object.entries(MIMES).map(([ext, mime]) => [mime, [`.${ext}`]])) }],
      excludeAcceptAllOption: false,
      multiple: false,
    });
    if (!handle) return null;
    return { file: await handle.getFile(), handle };
  } catch (e) {
    if (e?.name === 'AbortError') return null;
    throw e;
  }
}

/**
 * The handle of a file dropped on the page, if the browser gives one: `item` is the drop's
 * DataTransferItem. getAsFileSystemHandle must be called while the drop event is handled, so this
 * is called there and its promise awaited after.
 */
export function handleOfDropped(item) {
  try {
    const promise = item?.kind === 'file' && typeof item.getAsFileSystemHandle === 'function' ? item.getAsFileSystemHandle() : null;
    return promise ? promise.then((h) => (h?.kind === 'file' ? h : null), () => null) : Promise.resolve(null);
  } catch {
    return Promise.resolve(null);
  }
}

/**
 * Asks where to save a new file named `name` (of extension `ext`): its handle, or null if cancelled.
 * Must be called from a click or a key press, before anything slow.
 */
export async function pickFileToSave(name, ext) {
  try {
    return await window.showSaveFilePicker({
      id: 'intact-save',
      suggestedName: name,
      types: [{ description: 'Word document', accept: { [MIMES[ext] || MIMES.docx]: [`.${ext}`] } }],
    });
  } catch (e) {
    if (e?.name === 'AbortError') return null;
    throw e;
  }
}

/**
 * Makes sure the page may write to `handle`, asking the person if need be (the browser asks "Save
 * changes to …?"). True if it may. Called from a click or key press, before anything slow.
 */
export async function mayWrite(handle) {
  const opts = { mode: 'readwrite' };
  try {
    if (typeof handle.queryPermission === 'function' && (await handle.queryPermission(opts)) === 'granted') return true;
    if (typeof handle.requestPermission === 'function') return (await handle.requestPermission(opts)) === 'granted';
    return true; // a handle from the save picker may be written to
  } catch {
    return false;
  }
}

/** Writes `blob` into the file of `handle`, replacing what it held. */
export async function writeFile(handle, blob) {
  const writable = await handle.createWritable();
  try {
    await writable.write(blob);
    await writable.close();
  } catch (e) {
    await writable.abort?.().catch(() => {});
    throw e;
  }
}
