// Minimal numbering.xml support: computes a static marker ("•", "1.", "a)") per list paragraph, and
// adds the list definitions behind the Bullets and Numbering buttons.
import { NS, wChild, wChildren, wAttr, parseXml } from './xml.js';

const BULLETS = ['•', '◦', '▪'];
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
function format(n, fmt) {
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
    counts.length = ilvl + 1;
    if (lvl.fmt === 'bullet') return { marker: BULLETS[ilvl % 3], pPr: lvl.pPr };
    const marker = lvl.text.replace(/%(\d)/g, (_, d) => {
      const i = Number(d) - 1;
      return format(counts[i] ?? this.start(numId, i), levels.get(i)?.fmt);
    });
    return { marker, pPr: lvl.pPr };
  }
}

// ------------------------------------------------------------------ new list definitions

const BULLET_LEVELS = [
  ['', 'Symbol'], ['o', 'Courier New'], ['', 'Wingdings'],
];
const NUMBER_LEVELS = [['decimal', 'left'], ['lowerLetter', 'left'], ['lowerRoman', 'right']];

/** The nine levels Word uses for its default bullet or numbered list (0.5" steps, 0.25" hanging). */
function abstractNumXml(id, kind) {
  let lvls = '';
  for (let i = 0; i <= MAX_LEVEL; i++) {
    const ind = `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="${kind === 'number' && i % 3 === 2 ? 180 : 360}"/></w:pPr>`;
    if (kind === 'bullet') {
      const [char, font] = BULLET_LEVELS[i % 3];
      lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${char}"/><w:lvlJc w:val="left"/>${ind}`
        + `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr></w:lvl>`;
    } else {
      const [fmt, jc] = NUMBER_LEVELS[i % 3];
      lvls += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="${jc}"/>${ind}</w:lvl>`;
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
 * restart at 1), sharing one abstract definition per kind per editing session.
 */
export function addList(doc, kind, cache) {
  const root = doc.documentElement;
  let abstractId = cache[kind];
  if (abstractId == null) {
    abstractId = String(nextId(wChildren(root, 'abstractNum'), 'abstractNumId'));
    // Schema order: numPicBullet*, abstractNum*, num*, numIdMacAtCleanup?
    const before = wChild(root, 'num') || wChild(root, 'numIdMacAtCleanup');
    root.insertBefore(importW(doc, abstractNumXml(abstractId, kind)), before);
    cache[kind] = abstractId;
  }
  const numId = String(Math.max(1, nextId(wChildren(root, 'num'), 'numId')));
  const restart = kind === 'number' ? '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' : '';
  root.insertBefore(
    importW(doc, `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/>${restart}</w:num>`),
    wChild(root, 'numIdMacAtCleanup'),
  );
  return numId;
}
