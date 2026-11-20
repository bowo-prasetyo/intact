// Character formatting set on runs: a font, size, colour or character style from the document, or
// direct formatting cleared. The selected text's run properties (as they would be saved) are changed,
// and each piece of text is rebuilt as one run with them (DocxDocument.flattenRun), so the page shows
// exactly what is saved.
//
// Tracking formatting: before a formatting command, the selected text is noted with the run
// properties it would be saved with; after it, text whose properties changed gets a run of its own
// whose properties record the old ones (<w:rPrChange>), as Word records a tracked formatting change.
// Pieces are found again by their offsets in the text of their editing host, which formatting
// doesn't change.
import { portions, isolateText } from './inline.js';
import { NS, wAttr, ensureOrdered, removeW, RPR_ORDER, elementChildren } from '../docx/xml.js';

const W = NS.w;
const set = (el, name, value) => el.setAttributeNS(W, `w:${name}`, value);
const unset = (el, ...names) => names.forEach((n) => el.removeAttributeNS(W, n));

/** A font: a theme's ('theme:major' / 'theme:minor', as Word writes them) or one by name. */
export const setFont = (value) => (rPr) => {
  const f = ensureOrdered(rPr, 'rFonts', RPR_ORDER);
  if (value.startsWith('theme:')) {
    const theme = `${value.slice(6)}HAnsi`;
    unset(f, 'ascii', 'hAnsi');
    set(f, 'asciiTheme', theme);
    set(f, 'hAnsiTheme', theme);
  } else {
    unset(f, 'asciiTheme', 'hAnsiTheme'); // a theme font would win over the name
    set(f, 'ascii', value);
    set(f, 'hAnsi', value);
  }
};

/** A size in points (w:sz and w:szCs, in half-points). */
export const setSize = (pt) => (rPr) => {
  for (const local of ['sz', 'szCs']) set(ensureOrdered(rPr, local, RPR_ORDER), 'val', String(Math.round(pt * 2)));
};

// Word's Grow Font and Shrink Font (Ctrl+Shift+> and <) step through the sizes its menu offers;
// below 8 pt and above 72 pt they go by 1 pt and by 10 pt. Word's sizes run from 1 to 1638 pt.
const FONT_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export const MIN_SIZE = 1;
export const MAX_SIZE = 1638;

/** The size after `pt` (dir 1) or before it (dir -1) in Word's Grow Font / Shrink Font steps. */
export function nextFontSize(pt, dir) {
  if (dir > 0) {
    if (pt < 8) return Math.floor(pt) + 1;
    return FONT_STEPS.find((s) => s > pt) ?? Math.min(MAX_SIZE, (Math.floor(pt / 10) + 1) * 10);
  }
  if (pt <= 8) return Math.max(MIN_SIZE, Math.ceil(pt) - 1);
  if (pt > 72) return Math.max(72, (Math.ceil(pt / 10) - 1) * 10);
  return [...FONT_STEPS].reverse().find((s) => s < pt);
}

/**
 * Each run's size changed by `change(pt)` from the size it has: its own w:sz, else `sizeOf(node)`
 * (the size it shows, from its style). As Word's Grow Font: text in several sizes grows each.
 */
export const changeSize = (change, sizeOf) => (rPr, node) => {
  const own = elementChildren(rPr).find((c) => c.localName === 'sz' && c.namespaceURI === W);
  const pt = own ? Number(wAttr(own, 'val')) / 2 : sizeOf(node);
  if (!Number.isFinite(pt)) return;
  setSize(Math.min(MAX_SIZE, Math.max(MIN_SIZE, change(pt))))(rPr);
};

/** A colour: { hex, theme? } (theme: w:themeColor's name), or null for Automatic. */
export const setColor = (color) => (rPr) => {
  const c = ensureOrdered(rPr, 'color', RPR_ORDER);
  unset(c, 'themeColor', 'themeTint', 'themeShade');
  set(c, 'val', color ? color.hex : 'auto');
  if (color?.theme) set(c, 'themeColor', color.theme);
};

/**
 * Hidden text (w:vanish) on or off, as Word's Font > Hidden (Ctrl+Shift+H): what a test's answer key
 * holds, shown here (faintly, with a dotted line under it) and printed only when asked.
 */
export const setHidden = (on) => (rPr) => {
  removeW(rPr, 'vanish');
  if (on) ensureOrdered(rPr, 'vanish', RPR_ORDER);
};

/** A character style by id, or none (null). */
export const setCharacterStyle = (id) => (rPr) => {
  if (id) set(ensureOrdered(rPr, 'rStyle', RPR_ORDER), 'val', id);
  else removeW(rPr, 'rStyle');
};

// What clearing formatting keeps, as Word's Ctrl+Space: highlight, language and proofing, the
// script settings, and a tracked change's record.
const KEPT = new Set(['highlight', 'lang', 'noProof', 'rtl', 'cs', 'specVanish', 'oMath', 'rPrChange']);

/** Takes direct character formatting (and a character style) off; a link keeps its Hyperlink style. */
export const clearDirect = (isHyperlinkStyle) => (rPr) => {
  for (const c of elementChildren(rPr)) {
    if (KEPT.has(c.localName)) continue;
    if (c.localName === 'rStyle' && isHyperlinkStyle(wAttr(c, 'val'))) continue;
    rPr.removeChild(c);
  }
};


const hostOf = (node) => (node.nodeType === 1 ? node : node.parentElement)?.closest('[contenteditable="true"]');

/** The text nodes of `host`, in order, with their offsets in its text: [{ node, start }]. */
function textNodes(host) {
  const out = [];
  let at = 0;
  const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    out.push({ node: n, start: at });
    at += n.length;
  }
  return out;
}

/** The offset of DOM point (node, offset) in the text of `host`. */
function offsetIn(host, node, offset) {
  if (node.nodeType === 3) {
    for (const t of textNodes(host)) if (t.node === node) return t.start + offset;
    return 0;
  }
  const r = document.createRange();
  r.setStart(host, 0);
  r.setEnd(node, offset);
  return r.toString().length;
}

/** The text nodes holding exactly [start, end) of `host`'s text, split where they need to be. */
function nodesBetween(host, start, end) {
  const out = [];
  for (const { node, start: s } of textNodes(host)) {
    const e = s + node.length;
    if (e <= start || s >= end || !node.length) continue;
    out.push(isolateText({ node, s: Math.max(0, start - s), e: Math.min(node.length, end - s) }));
  }
  return out;
}

/** A range over [start, end) of `host`'s text. */
function rangeBetween(host, start, end) {
  const nodes = textNodes(host);
  const point = (at) => {
    for (const { node, start: s } of nodes) if (at <= s + node.length) return [node, at - s];
    return [host, host.childNodes.length];
  };
  const r = document.createRange();
  r.setStart(...point(start));
  r.setEnd(...point(end));
  return r;
}

/**
 * Changes the run properties of the selected text: `mutate(rPr)` changes a copy of each piece's
 * <w:rPr>. Returns a range over the text, or null if nothing editable was selected.
 */
export function restyle(doc, range, mutate) {
  const nodes = portions(range).map(isolateText);
  if (!nodes.length) return null;
  for (const node of nodes) doc.flattenRun(node, doc.mutatedRPr(doc.rPrOf(node), (rPr) => mutate(rPr, node)));
  const r = document.createRange();
  r.setStart(nodes[0], 0);
  const last = nodes[nodes.length - 1];
  r.setEnd(last, last.length);
  return r;
}

/** Notes the selected text and the run properties it would be saved with, before a formatting command. */
export function noteRuns(doc, range) {
  const host = hostOf(range.startContainer);
  if (!host) return null;
  const pieces = portions(range).filter((p) => hostOf(p.node) === host).map((p) => {
    const start = offsetIn(host, p.node, p.s);
    return { start, end: start + (p.e - p.s), old: doc.rPrOf(p.node) };
  });
  return { host, pieces, start: offsetIn(host, range.startContainer, range.startOffset), end: offsetIn(host, range.endContainer, range.endOffset) };
}

/**
 * After a formatting command: the text noted by noteRuns whose run properties changed records the
 * old ones as a tracked change by `opts` ({ author, date }). Returns a range over the noted text.
 */
export function trackRuns(doc, noted, opts) {
  if (!noted?.host.isConnected) return null;
  const { host } = noted;
  for (const { start, end, old } of noted.pieces) {
    for (const node of nodesBetween(host, start, end)) {
      const now = doc.rPrOf(node);
      if (doc.sameProps(doc.withoutChange(now), doc.withoutChange(old))) continue;
      doc.flattenRun(node, doc.trackedRPr(old, now, opts));
    }
  }
  return rangeBetween(host, noted.start, noted.end);
}
