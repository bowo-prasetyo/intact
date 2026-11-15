// Minimal numbering.xml support: computes a static marker ("•", "1.", "a)") per list paragraph.
import { wChild, wChildren, wAttr } from './xml.js';

const BULLETS = ['•', '◦', '▪'];

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
    this.nums = new Map(); // numId -> abstractNumId
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
      this.nums.set(wAttr(num, 'numId'), wAttr(wChild(num, 'abstractNumId'), 'val'));
    }
  }

  /** Returns { marker, pPr } for the paragraph's numPr, advancing counters. */
  next(numPr) {
    const numId = wAttr(wChild(numPr, 'numId'), 'val');
    const ilvl = Number(wAttr(wChild(numPr, 'ilvl'), 'val') ?? 0);
    if (!numId || numId === '0') return null;
    const levels = this.abstract.get(this.nums.get(numId));
    const lvl = levels?.get(ilvl);
    if (!lvl) return { marker: BULLETS[ilvl % 3], pPr: null };
    let counts = this.counters.get(numId);
    if (!counts) this.counters.set(numId, (counts = []));
    for (let i = 0; i <= ilvl; i++) {
      if (counts[i] === undefined) counts[i] = (levels.get(i)?.start ?? 1) - (i === ilvl ? 1 : 0);
    }
    counts[ilvl]++;
    counts.length = ilvl + 1;
    if (lvl.fmt === 'bullet') return { marker: BULLETS[ilvl % 3], pPr: lvl.pPr };
    const marker = lvl.text.replace(/%(\d)/g, (_, d) => {
      const i = Number(d) - 1;
      return format(counts[i] ?? levels.get(i)?.start ?? 1, levels.get(i)?.fmt);
    });
    return { marker, pPr: lvl.pPr };
  }
}
