// Several documents at once, each in a tab of its own (a CV in one, its cover letter in the next):
// every tab is the whole app with its one document, so tabs only need to open each other and to
// know which drafts the others are editing. Nothing leaves the device and nothing is stored for it:
//  - a new tab opens on the start page, or with a file handed to it from this tab, in memory, over a
//    BroadcastChannel (the new tab asks for it by a one-time token in its address);
//  - a tab editing a document holds a Web Lock named after its draft (see drafts.js), so the start
//    page in another tab shows that draft as open there rather than offering to restore a second
//    copy, whose edits would overwrite each other's.
// A browser without BroadcastChannel or Web Locks still opens tabs; it just can't hand a file over,
// or tell which drafts are open.

const CHANNEL = 'intact-tabs';
const LOCK = 'intact-draft:';
const HANDOFF_WAIT = 60_000; // how long a file waits here for the new tab to ask for it

/** This app's address with `hash` (the route): the same page, wherever the app is served from. */
export function appUrl(hash = '#/') {
  return `${location.href.split('#')[0]}${hash}`;
}

/**
 * Opens the start page in a new tab, to open another document beside this one. Returns whether a
 * tab was opened (a browser may block it, though a click or a key press normally allows it).
 */
export function openStartInNewTab() {
  return !!window.open(appUrl('#/'), '_blank');
}

const pending = new Map(); // token -> File, waiting for the new tab to ask for it
let channel = null;
function sender() {
  if (channel || typeof BroadcastChannel === 'undefined') return channel;
  channel = new BroadcastChannel(CHANNEL);
  channel.onmessage = ({ data }) => {
    const file = data?.want && pending.get(data.want);
    if (!file) return;
    pending.delete(data.want);
    channel.postMessage({ token: data.want, file });
  };
  return channel;
}

const token = () => globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`;

/**
 * Opens `file` in a new tab: the tab opens on the start page, asks this one for the file and opens
 * it. Returns 'opened', 'blocked' (the browser didn't open the tab) or 'unsupported' (no
 * BroadcastChannel: the start page opens in the new tab, for the file to be chosen there).
 */
export function openFileInNewTab(file) {
  if (!sender()) return openStartInNewTab() ? 'unsupported' : 'blocked';
  const id = token();
  pending.set(id, file);
  setTimeout(() => pending.delete(id), HANDOFF_WAIT);
  if (window.open(appUrl(`#/?receive=${encodeURIComponent(id)}`), '_blank')) return 'opened';
  pending.delete(id);
  return 'blocked';
}

/**
 * The file handed to this tab under `id` (from its address) by the tab that opened it, or null if
 * none comes in time (that tab was closed, say).
 */
export function receiveFile(id, wait = 5000) {
  if (typeof BroadcastChannel === 'undefined') return Promise.resolve(null);
  const ch = new BroadcastChannel(CHANNEL);
  return new Promise((resolve) => {
    const done = (file) => { clearTimeout(timer); ch.close(); resolve(file); };
    const timer = setTimeout(() => done(null), wait);
    ch.onmessage = ({ data }) => { if (data?.token === id && data.file) done(data.file); };
    ch.postMessage({ want: id });
  });
}

/**
 * Marks the draft `id` as being edited in this tab, until the returned function is called (or the
 * tab is closed): other tabs see it as open here (see draftsOpenElsewhere).
 */
export function holdDraft(id) {
  if (!navigator.locks?.request) return () => {};
  let release;
  const held = new Promise((resolve) => { release = resolve; });
  navigator.locks.request(`${LOCK}${id}`, () => held).catch(() => {});
  return () => release();
}

/** The ids of the drafts being edited in tabs (this one's included), as a Set. */
export async function draftsOpenElsewhere() {
  if (!navigator.locks?.query) return new Set();
  try {
    const { held = [] } = await navigator.locks.query();
    return new Set(held.map((l) => l.name).filter((n) => n?.startsWith(LOCK)).map((n) => n.slice(LOCK.length)));
  } catch {
    return new Set();
  }
}
