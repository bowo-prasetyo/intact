// Read Aloud, as Word's Review > Read Aloud: the text from the caret (or the selection) read paragraph
// by paragraph in the browser's speech, the paragraph and the word being read highlighted.
//
// Private: only voices on this device are used (SpeechSynthesisVoice.localService). Some browsers
// also offer online voices, which send the text to a server to be spoken; those are never chosen,
// so reading aloud keeps the document on the device as the rest of Intact does. Where there is no
// voice on the device for the text's language, it says so instead of reading.
//
// Nothing is changed in the page: the highlights are the CSS Custom Highlight API's (as Find's),
// so Undo, saving and the document never see them. Hidden text (a test's answers) isn't read.

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6, li';
// Not read: hidden text, list markers and what is drawn over the page or kept as it is.
const SKIP = '[data-hidden], .raw-hidden, .list-marker, [aria-hidden="true"], .docx-hf-band, .docx-notes, .comment-anchor-label';
const canHighlight = () => typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined';

/** The speech synthesis this browser has, or null. */
export function speechAvailable() {
  return typeof window !== 'undefined' && window.speechSynthesis && typeof window.SpeechSynthesisUtterance === 'function' ? window.speechSynthesis : null;
}

/** The voices on this device (none online), once the browser has listed them (it may take a moment). */
export async function localVoices(synth = speechAvailable()) {
  if (!synth) return [];
  let voices = synth.getVoices();
  if (!voices.length) {
    voices = await new Promise((resolve) => {
      const done = () => resolve(synth.getVoices());
      synth.addEventListener?.('voiceschanged', done, { once: true });
      setTimeout(done, 1500);
    });
  }
  return voices.filter((v) => v.localService);
}

/** The on-device voice for language `lang` (en-GB, id-ID…): the same language and region, else the same language; the browser's default first. */
export function voiceFor(voices, lang) {
  const want = (lang || 'en').toLowerCase().replace('_', '-');
  const primary = want.split('-')[0];
  const norm = (v) => (v.lang || '').toLowerCase().replace('_', '-');
  const rank = (v) => (norm(v) === want ? 0 : norm(v).split('-')[0] === primary ? 1 : 9) - (v.default ? 0.5 : 0);
  const best = voices.filter((v) => rank(v) < 9).sort((a, b) => rank(a) - rank(b))[0];
  return best || null;
}

/** The text nodes of `block` that are read, with where each starts in the text read: { text, parts: [{ node, at }] }. */
export function readableText(block) {
  const parts = [];
  let text = '';
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) => (n.parentElement?.closest(SKIP) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    parts.push({ node: n, at: text.length });
    text += n.data.replace(/\t/g, ' ');
  }
  return { text, parts };
}

/** A range over characters `start` to `end` of a block's text read (see readableText). */
function rangeOf(parts, start, end) {
  const find = (offset, isEnd) => {
    for (let i = parts.length - 1; i >= 0; i--) {
      const p = parts[i];
      if (offset > p.at || (!isEnd && offset === p.at) || (isEnd && offset === p.at && i === 0)) return [p.node, Math.min(offset - p.at, p.node.length)];
    }
    return parts[0] ? [parts[0].node, 0] : null;
  };
  const a = find(start, false);
  const b = find(end, true);
  if (!a || !b) return null;
  const r = document.createRange();
  r.setStart(...a);
  r.setEnd(...b);
  return r;
}

/**
 * A reader for the paragraphs of `page`: start(from, { lang }) reads from top-level paragraph `from`
 * (or the paragraphs of a selection, `only`), pause(), resume(), stop(). `onState` hears
 * { reading, paused, block, error }; `langOf(block)` gives a paragraph's language.
 */
export function createReader(page, { onState = () => {}, langOf = () => null } = {}) {
  const synth = speechAvailable();
  let queue = [];
  let current = null;
  let voices = null;
  let paused = false;
  let token = 0;

  const clear = () => {
    if (!canHighlight()) return;
    CSS.highlights.delete('read-aloud');
    CSS.highlights.delete('read-aloud-word');
  };
  const say = (s) => onState({ reading: !!current, paused, block: current?.block || null, ...s });

  function finish(error = null) {
    token++;
    queue = [];
    current = null;
    paused = false;
    synth?.cancel();
    clear();
    say({ reading: false, error });
  }

  async function next(run) {
    if (run !== token) return;
    const block = queue.shift();
    if (!block) return finish();
    if (!block.isConnected) return next(run);
    const { text, parts } = readableText(block);
    if (!text.trim()) return next(run);
    const lang = langOf(block);
    const voice = voiceFor(voices, lang);
    if (!voice) return finish(`no-voice:${lang || ''}`);
    current = { block, parts };
    if (canHighlight()) {
      const all = rangeOf(parts, 0, text.length);
      CSS.highlights.set('read-aloud', all ? new Highlight(all) : new Highlight());
      CSS.highlights.delete('read-aloud-word');
    }
    block.scrollIntoView?.({ block: 'nearest', behavior: 'smooth' });
    const u = new window.SpeechSynthesisUtterance(text);
    u.voice = voice;
    u.lang = voice.lang || lang || '';
    u.onboundary = (e) => {
      if (run !== token || e.name === 'sentence' || !canHighlight()) return;
      const len = e.charLength || (/^\S+/.exec(text.slice(e.charIndex))?.[0].length ?? 0);
      const r = rangeOf(parts, e.charIndex, e.charIndex + len);
      if (r) CSS.highlights.set('read-aloud-word', new Highlight(r));
    };
    u.onend = () => next(run);
    u.onerror = (e) => {
      if (run !== token || e.error === 'interrupted' || e.error === 'canceled') return;
      finish(`speech:${e.error || 'error'}`);
    };
    say({ reading: true });
    synth.speak(u);
  }

  return {
    get reading() { return !!current; },
    get paused() { return paused; },
    /** Reads from top-level paragraph `from` to the end, or only the paragraphs `only`. */
    async start(from, only = null) {
      if (!synth) return finish('unsupported');
      finish();
      const run = ++token;
      const all = [...page.querySelectorAll(BLOCKS)].filter((b) => !b.closest(SKIP) && !b.querySelector(BLOCKS));
      queue = only?.length ? all.filter((b) => only.some((o) => o === b || o.contains(b))) : all.slice(Math.max(0, all.findIndex((b) => b === from || b.contains(from) || from?.contains?.(b))));
      voices = await localVoices(synth);
      if (run !== token) return undefined;
      if (!voices.length) return finish('no-local-voices');
      return next(run);
    },
    pause() {
      if (!current || paused) return;
      paused = true;
      synth.pause();
      say({});
    },
    resume() {
      if (!current || !paused) return;
      paused = false;
      synth.resume();
      say({});
    },
    stop: () => finish(),
  };
}
