// Minimal numbering.xml support: computes a static marker ("•", "1.", "a)") per list paragraph, and
// adds the list definitions behind the Bullets and Numbering buttons.
import { NS, wChild, wChildren, wAttr, parseXml } from './xml.js';

const BULLETS = ['•', '◦', '▪'];
const DASHES = new Set(['-', '–', '—']);
export const MAX_LEVEL = 8;

function roman(n) {
  const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of map) while (n >= v) { s += r; n -= v; }
  return s;
}
function letter(n) {
  let s = '';
  while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}
/** Number `n` as list format `fmt` writes it ('decimal' 3, 'lowerLetter' c, 'upperRoman' III…). */
export function formatNumber(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return letter(n);
    case 'upperLetter': return letter(n).toUpperCase();
    case 'lowerRoman': return roman(n);
    case 'upperRoman': return roman(n).toUpperCase();
    case 'decimalZero': return String(n).padStart(2, '0');
    case 'none': return '';
    default: return String(n);
  }
}

export class Numbering {
  constructor(doc) {
    this.abstract = new Map(); // abstractNumId -> Map(ilvl -> {fmt, text, start, pPr})
    this.nums = new Map(); // numId -> { abstractId, starts: Map(ilvl -> startOverride) }
    this.counters = new Map(); // numId -> number[]
    if (!doc) return;
    const root = doc.documentElement;
    for (const an of wChildren(root, 'abstractNum')) {
      const levels = new Map();
      for (const lvl of wChildren(an, 'lvl')) {
        levels.set(Number(wAttr(lvl, 'ilvl')), {
          fmt: wAttr(wChild(lvl, 'numFmt'), 'val') || 'decimal',
          text: wAttr(wChild(lvl, 'lvlText'), 'val') ?? '',
          start: Number(wAttr(wChild(lvl, 'start'), 'val') ?? 1),
          // w:lvlRestart: the level whose use restarts this one (1-based; 0 never), else any higher one.
          restart: wChild(lvl, 'lvlRestart') ? Number(wAttr(wChild(lvl, 'lvlRestart'), 'val')) : null,
          pPr: wChild(lvl, 'pPr'),
        });
      }
      this.abstract.set(wAttr(an, 'abstractNumId'), levels);
    }
    for (const num of wChildren(root, 'num')) {
      const starts = new Map();
      for (const o of wChildren(num, 'lvlOverride')) {
        const start = wAttr(wChild(o, 'startOverride'), 'val');
        if (start != null) starts.set(Number(wAttr(o, 'ilvl') ?? 0), Number(start));
      }
      this.nums.set(wAttr(num, 'numId'), { abstractId: wAttr(wChild(num, 'abstractNumId'), 'val'), starts });
    }
  }

  levels(numId) {
    return this.abstract.get(this.nums.get(numId)?.abstractId);
  }

  start(numId, ilvl) {
    return this.nums.get(numId)?.starts.get(ilvl) ?? this.levels(numId)?.get(ilvl)?.start ?? 1;
  }

  hasLevel(numId, ilvl) {
    return !!this.levels(numId)?.has(ilvl);
  }

  /** 'bullet' or 'number': what kind of list a paragraph at this level shows. */
  kind(numId, ilvl) {
    const levels = this.levels(numId);
    const lvl = levels?.get(ilvl) || levels?.get(0);
    return !lvl || lvl.fmt === 'bullet' ? 'bullet' : 'number';
  }

  /** Returns { marker, pPr } for the paragraph at `ilvl` of list `numId`, advancing counters. */
  next(numId, ilvl = 0) {
    if (!numId || numId === '0') return null;
    const levels = this.levels(numId);
    const lvl = levels?.get(ilvl);
    if (!lvl) return { marker: BULLETS[ilvl % 3], pPr: null };
    let counts = this.counters.get(numId);
    if (!counts) this.counters.set(numId, (counts = []));
    for (let i = 0; i <= ilvl; i++) {
      if (counts[i] === undefined) counts[i] = this.start(numId, i) - (i === ilvl ? 1 : 0);
    }
    counts[ilvl]++;
    // The levels below restart, but for those whose w:lvlRestart says otherwise (Pasal numbers that
    // run on through every BAB of a regulation: w:lvlRestart w:val="0").
    for (let i = ilvl + 1; i < counts.length; i++) {
      const restart = levels.get(i)?.restart ?? i;
      if (ilvl < restart) counts[i] = undefined;
    }
    // A dash bullet (as AutoFormat makes from "- ") shows as typed; symbol-font bullets as •, ◦, ▪.
    if (lvl.fmt === 'bullet') return { marker: DASHES.has(lvl.text) ? lvl.text : BULLETS[ilvl % 3], pPr: lvl.pPr };
    const marker = lvl.text.replace(/%(\d)/g, (_, d) => {
      const i = Number(d) - 1;
      return formatNumber(counts[i] ?? this.start(numId, i), levels.get(i)?.fmt);
    });
    return { marker, pPr: lvl.pPr };
  }
}

// ------------------------------------------------------------------ new list definitions

const BULLET_LEVELS = [
  ['', 'Symbol'], ['o', 'Courier New'], ['', 'Wingdings'],
];
const NUMBER_LEVELS = [['decimal', 'left'], ['lowerLetter', 'left'], ['lowerRoman', 'right']];

/**
 * The nine levels Word uses for its default bullet or numbered list (0.5" steps, 0.25" hanging).
 * `first` changes the first level, as Word's AutoFormat does for a list started by typing:
 * { fmt ('decimal', 'lowerLetter'…), text ('%1.', '(%1)'…) } for numbers, { char } for bullets.
 */
function abstractNumXml(id, kind, first = null) {
  let lvls = '';
  for (let i = 0; i <= MAX_LEVEL; i++) {
    const ind = `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="${kind === 'number' && i % 3 === 2 ? 180 : 360}"/></w:pPr>`;
    if (kind === 'bullet') {
      const [char, font] = BULLET_LEVELS[i % 3];
      // A typed bullet (a dash) is in the text's own font, as Word writes it.
      const rPr = i === 0 && first?.char ? '' : `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>`;
      lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${i === 0 && first?.char ? first.char : char}"/><w:lvlJc w:val="left"/>${ind}`
        + `${rPr}</w:lvl>`;
    } else {
      let [fmt, jc] = NUMBER_LEVELS[i % 3];
      let text = `%${i + 1}.`;
      if (i === 0 && first?.fmt) {
        fmt = first.fmt;
        text = first.text || text;
        jc = /Roman$/.test(fmt) ? 'right' : 'left';
      }
      lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="${jc}"/>${ind}</w:lvl>`;
    }
  }
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`;
}

function importW(doc, xml) {
  const root = parseXml(`<root xmlns:w="${NS.w}">${xml}</root>`).documentElement;
  return doc.importNode(root.firstElementChild, true);
}

const nextId = (els, attr) => Math.max(-1, ...els.map((e) => Number(wAttr(e, attr)) || 0)) + 1;

/**
 * Adds a list to numbering.xml and returns its numId. Each call starts a new list (numbered lists
 * restart at 1, or at `first.start`), sharing one abstract definition per kind (and first level) per
 * editing session. `first`: see abstractNumXml.
 */
export function addList(doc, kind, cache, first = null) {
  const root = doc.documentElement;
  const key = first ? `${kind}:${first.fmt || ''}:${first.text || ''}:${first.char || ''}` : kind;
  let abstractId = cache[key];
  if (abstractId == null) {
    abstractId = String(nextId(wChildren(root, 'abstractNum'), 'abstractNumId'));
    // Schema order: numPicBullet*, abstractNum*, num*, numIdMacAtCleanup?
    const before = wChild(root, 'num') || wChild(root, 'numIdMacAtCleanup');
    root.insertBefore(importW(doc, abstractNumXml(abstractId, kind, first)), before);
    cache[key] = abstractId;
  }
  const numId = String(Math.max(1, nextId(wChildren(root, 'num'), 'numId')));
  const start = Math.max(0, Math.floor(Number(first?.start ?? 1)));
  const restart = kind === 'number' ? `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${start}"/></w:lvlOverride>` : '';
  root.insertBefore(
    importW(doc, `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/>${restart}</w:num>`),
    wChild(root, 'numIdMacAtCleanup'),
  );
  return numId;
}

// ------------------------------------------------------------------ multilevel lists (Word's List Library)

// Ready-made multilevel lists, as Word's Multilevel List gallery offers them, with Indonesia's legal
// numbering as its regulations write it (UU No. 12/2011, its Lampiran II): a Pasal on its own line,
// centred, its ayat (1), (2), then huruf a., b., angka 1., 2., and a), 1) below them. Each level is
// { fmt, text, jc (the paragraph's own alignment), left, hanging (twips), restart (w:lvlRestart),
// keepNext }. A level without its own text and number is Word's next one down, as Word fills them.
const ID_LEGAL = [
  { fmt: 'decimal', text: '(%L)', left: 567, hanging: 567 },
  { fmt: 'lowerLetter', text: '%L.', left: 1134, hanging: 567 },
  { fmt: 'decimal', text: '%L.', left: 1701, hanging: 567 },
  { fmt: 'lowerLetter', text: '%L)', left: 2268, hanging: 567 },
  { fmt: 'decimal', text: '%L)', left: 2835, hanging: 567 },
];
const PASAL = { fmt: 'decimal', text: 'Pasal %L', jc: 'center', left: 0, hanging: 0, keepNext: true, before: 240 };
export const LIST_PRESETS = {
  // Pasal 1 → (1) → a. → 1. → a) → 1)
  pasal: [PASAL, ...ID_LEGAL],
  // BAB I → Pasal 1 (numbered on through the BABs) → (1) → a. → 1. → a)
  bab: [{ fmt: 'upperRoman', text: 'BAB %L', jc: 'center', left: 0, hanging: 0, keepNext: true, before: 240 }, { ...PASAL, restart: 0 }, ...ID_LEGAL.slice(0, 4)],
  // 1. → 1.1. → 1.1.1. (a report's numbered headings and paragraphs)
  outline: [1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => ({
    fmt: 'decimal', text: Array.from({ length: n }, (_, i) => `%${i + 1}`).join('.') + '.', left: 432 + 144 * (n - 1) * 2, hanging: 432 + 144 * (n - 1),
  })),
};

function presetXml(id, preset) {
  const levels = LIST_PRESETS[preset];
  let lvls = '';
  for (let i = 0; i <= MAX_LEVEL; i++) {
    const l = levels[Math.min(i, levels.length - 1)];
    // Past the preset's levels: one more step in, as decimals.
    const extra = i - (levels.length - 1);
    const left = extra > 0 ? l.left + 567 * extra : l.left;
    const text = (extra > 0 ? '%L.' : l.text).replace('%L', `%${i + 1}`);
    const fmt = extra > 0 ? 'decimal' : l.fmt;
    const restart = l.restart != null && extra <= 0 ? `<w:lvlRestart w:val="${l.restart}"/>` : '';
    const ind = l.hanging ? `<w:ind w:left="${left}" w:hanging="${l.hanging}"/>` : `<w:ind w:left="${left}" w:firstLine="0"/>`;
    const pPr = `<w:pPr>${l.keepNext && extra <= 0 ? '<w:keepNext/>' : ''}${l.before && extra <= 0 ? `<w:spacing w:before="${l.before}"/>` : ''}${ind}${l.jc && extra <= 0 ? `<w:jc w:val="${l.jc}"/>` : ''}</w:pPr>`;
    lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>${restart}<w:suff w:val="${l.hanging ? 'tab' : 'space'}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>${pPr}</w:lvl>`;
  }
  return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="multilevel"/>${lvls}</w:abstractNum>`;
}

/**
 * Adds a list of a ready-made multilevel kind (`preset`: a key of LIST_PRESETS) to numbering.xml and
 * returns its numId, sharing one definition per preset per editing session.
 */
export function addPresetList(doc, preset, cache) {
  const root = doc.documentElement;
  const key = `preset:${preset}`;
  let abstractId = cache[key];
  if (abstractId == null) {
    abstractId = String(nextId(wChildren(root, 'abstractNum'), 'abstractNumId'));
    root.insertBefore(importW(doc, presetXml(abstractId, preset)), wChild(root, 'num') || wChild(root, 'numIdMacAtCleanup'));
    cache[key] = abstractId;
  }
  const numId = String(Math.max(1, nextId(wChildren(root, 'num'), 'numId')));
  root.insertBefore(importW(doc, `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`), wChild(root, 'numIdMacAtCleanup'));
  return numId;
}
