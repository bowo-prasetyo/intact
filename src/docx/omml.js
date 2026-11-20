// Equations: Office Math (OMML, the m: elements Word writes) shown as MathML, which browsers draw.
// Read-only: the equation's XML is kept and written back as it was; this is only how it looks.
// Covers what Word's equation editor writes: runs (with their style and script), fractions, scripts,
// radicals, n-ary operators (sums, integrals), delimiters, functions, limits, accents, bars, group
// characters, boxes, equation arrays, matrices and phantoms. Anything else shows its content.
// Letters in a style other than italic or upright (bold, script, double-struck…) are written as
// Unicode's mathematical letters, which every browser draws; MathML Core has no other way.

import { NS } from './xml.js';

export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const MML = 'http://www.w3.org/1998/Math/MathML';

const kids = (el) => [...(el?.children || [])];
const mChild = (el, local) => kids(el).find((c) => c.namespaceURI === MATH_NS && c.localName === local) || null;
const mChildren = (el, local) => kids(el).filter((c) => c.namespaceURI === MATH_NS && c.localName === local);
/** An m:val property: absent → undefined; present without a value → ''. */
function prop(pr, local) {
  const el = mChild(pr, local);
  if (!el) return undefined;
  return el.getAttributeNS(MATH_NS, 'val') ?? el.getAttribute('m:val') ?? '';
}
const on = (v) => v !== undefined && !['0', 'false', 'off'].includes(v);

// ------------------------------------------------------------------ mathematical letters

// First code point of A, a and 0 in each alphabet, and the letters Unicode had already elsewhere.
const ALPHABETS = {
  bold: [0x1d400, 0x1d41a, 0x1d7ce, {}],
  'bold-italic': [0x1d468, 0x1d482, 0x1d7ce, {}],
  italic: [0x1d434, 0x1d44e, null, { h: 0x210e }],
  script: [0x1d49c, 0x1d4b6, null, { B: 0x212c, E: 0x2130, F: 0x2131, H: 0x210b, I: 0x2110, L: 0x2112, M: 0x2133, R: 0x211b, e: 0x212f, g: 0x210a, o: 0x2134 }],
  'bold-script': [0x1d4d0, 0x1d4ea, 0x1d7ce, {}],
  fraktur: [0x1d504, 0x1d51e, null, { C: 0x212d, H: 0x210c, I: 0x2111, R: 0x211c, Z: 0x2128 }],
  'bold-fraktur': [0x1d56c, 0x1d586, 0x1d7ce, {}],
  'double-struck': [0x1d538, 0x1d552, 0x1d7d8, { C: 0x2102, H: 0x210d, N: 0x2115, P: 0x2119, Q: 0x211a, R: 0x211d, Z: 0x2124 }],
  'sans-serif': [0x1d5a0, 0x1d5ba, 0x1d7e2, {}],
  'bold-sans-serif': [0x1d5d4, 0x1d5ee, 0x1d7ec, {}],
  'sans-serif-italic': [0x1d608, 0x1d622, 0x1d7e2, {}],
  'sans-serif-bold-italic': [0x1d63c, 0x1d656, 0x1d7ec, {}],
  monospace: [0x1d670, 0x1d68a, 0x1d7f6, {}],
};

/** `ch` in mathematical alphabet `variant`, or `ch` itself if the alphabet has no such letter. */
export function mathLetter(ch, variant) {
  const a = ALPHABETS[variant];
  if (!a) return ch;
  if (a[3][ch]) return String.fromCodePoint(a[3][ch]);
  const c = ch.codePointAt(0);
  if (c >= 65 && c <= 90) return String.fromCodePoint(a[0] + c - 65);
  if (c >= 97 && c <= 122) return String.fromCodePoint(a[1] + c - 97);
  if (c >= 48 && c <= 57 && a[2]) return String.fromCodePoint(a[2] + c - 48);
  return ch;
}

/** The alphabet for a run's m:scr (script) and m:sty (p, b, i, bi), or null for the default. */
function variantOf(scr, sty) {
  const bold = sty === 'b' || sty === 'bi';
  const italic = sty === undefined || sty === 'i' || sty === 'bi';
  switch (scr) {
    case 'script': return bold ? 'bold-script' : 'script';
    case 'fraktur': return bold ? 'bold-fraktur' : 'fraktur';
    case 'double-struck': return 'double-struck';
    case 'sans-serif':
      if (bold) return italic ? 'sans-serif-bold-italic' : 'bold-sans-serif';
      return italic ? 'sans-serif-italic' : 'sans-serif';
    case 'monospace': return 'monospace';
    default:
      if (sty === 'b') return 'bold';
      if (sty === 'bi') return 'bold-italic';
      return null; // italic letters (the default) or upright ('p')
  }
}

// Accents are stored as combining marks; drawn above a base they need their spacing forms.
const SPACING = {
  '̀': '`', '́': '´', '̂': 'ˆ', '̃': '˜', '̄': '¯', '̅': '‾',
  '̆': '˘', '̇': '˙', '̈': '¨', '̌': 'ˇ', '⃖': '←', '⃗': '→', '⃡': '↔',
  '⃐': '↼', '⃑': '⇀', '̿': '‗',
};

const INTEGRALS = /[∫-∳⨋-⨜]/;

// ------------------------------------------------------------------ conversion

class Converter {
  constructor(doc) {
    this.doc = doc;
    this.inArray = 0;
  }

  el(name, children = [], attrs = {}) {
    const e = this.doc.createElementNS(MML, name);
    for (const [k, v] of Object.entries(attrs)) if (v != null) e.setAttribute(k, v);
    for (const c of children) e.append(c);
    return e;
  }

  /** One node for a list of nodes: itself, or an <mrow> around them. */
  row(nodes) {
    return nodes.length === 1 ? nodes[0] : this.el('mrow', nodes);
  }

  /** An argument (m:e, m:num, m:sub…) as one node; a missing or empty one is an empty <mrow>. */
  arg(el) {
    return this.row(el ? this.all(el) : []);
  }

  all(el) {
    return kids(el).flatMap((c) => this.node(c));
  }

  op(text, attrs = {}) {
    return this.el('mo', [text], attrs);
  }

  node(el) {
    if (el.namespaceURI === NS.w) {
      // Revisions and markers around math: show inserted content, not deleted or hidden markers.
      if (['ins', 'moveTo', 'smartTag', 'customXml', 'sdt', 'sdtContent'].includes(el.localName)) return this.all(el);
      if (el.localName === 'r') {
        const text = kids(el).filter((t) => t.namespaceURI === NS.w && t.localName === 't').map((t) => t.textContent).join('');
        return text ? [this.el('mtext', [text])] : [];
      }
      return [];
    }
    if (el.namespaceURI !== MATH_NS) return [];
    const L = el.localName;
    if (L.endsWith('Pr')) return []; // properties (fPr, ctrlPr, rPr…) are read where they apply
    switch (L) {
      case 'oMath': return this.all(el);
      case 'r': return this.run(el);
      case 'f': return [this.fraction(el)];
      case 'sSup': return [this.el('msup', [this.arg(mChild(el, 'e')), this.arg(mChild(el, 'sup'))])];
      case 'sSub': return [this.el('msub', [this.arg(mChild(el, 'e')), this.arg(mChild(el, 'sub'))])];
      case 'sSubSup': return [this.el('msubsup', [this.arg(mChild(el, 'e')), this.arg(mChild(el, 'sub')), this.arg(mChild(el, 'sup'))])];
      case 'sPre': return [this.el('mmultiscripts', [this.arg(mChild(el, 'e')), this.el('mprescripts'), this.arg(mChild(el, 'sub')), this.arg(mChild(el, 'sup'))])];
      case 'rad': return [this.radical(el)];
      case 'nary': return this.nary(el);
      case 'd': return [this.delimiter(el)];
      case 'func': return [this.el('mrow', [this.arg(mChild(el, 'fName')), this.op('⁡'), this.arg(mChild(el, 'e'))])];
      case 'limLow': return [this.el('munder', [this.arg(mChild(el, 'e')), this.arg(mChild(el, 'lim'))])];
      case 'limUpp': return [this.el('mover', [this.arg(mChild(el, 'e')), this.arg(mChild(el, 'lim'))])];
      case 'acc': return [this.accent(el)];
      case 'bar': return [this.bar(el)];
      case 'groupChr': return [this.groupChr(el)];
      case 'box': return [this.arg(mChild(el, 'e'))];
      case 'borderBox': return [this.el('mrow', [this.arg(mChild(el, 'e'))], { style: 'border: 1px solid; padding: 0.1em 0.2em' })];
      case 'eqArr': return [this.equationArray(el)];
      case 'm': return [this.matrix(el)];
      case 'phant': return [this.phantom(el)];
      default: return this.all(el); // e, num, den, sub, sup, deg, lim, fName and anything unknown
    }
  }

  /** A run's text as identifiers (letters), numbers and operators, in the run's style. */
  run(r) {
    const rPr = mChild(r, 'rPr');
    const text = mChildren(r, 't').map((t) => t.textContent).join('');
    if (!this.inArray || !text.includes('&')) return this.runText(text, rPr);
    // In an equation array, an & is an alignment point: a mark the array lines its rows up at.
    return text.split('&').flatMap((part, i) => [...(i ? [this.el('mspace', [], { 'data-align': '' })] : []), ...this.runText(part, rPr)]);
  }

  /** A run's text (`rPr` its properties) as MathML: numbers, identifiers, operators and spaces. */
  runText(text, rPr) {
    if (!text) return [];
    if (on(prop(rPr, 'nor')) || prop(rPr, 'lit') === '1') return [this.el('mtext', [text])];
    const sty = prop(rPr, 'sty');
    const variant = variantOf(prop(rPr, 'scr'), sty);
    const upright = sty === 'p';
    const out = [];
    for (const [, num, word, space, other] of text.matchAll(/(\d+(?:[.,]\d+)*)|(\p{L}+)|(\s+)|([^\s\d\p{L}])/gu)) {
      if (num) out.push(this.el('mn', [[...num].map((c) => mathLetter(c, variant)).join('')]));
      else if (word) {
        // Upright words (sin, lim, max) are one identifier; italic letters each their own, as Word sets them.
        if (variant) for (const c of word) out.push(this.el('mi', [mathLetter(c, variant)], { mathvariant: 'normal' }));
        else if (upright) out.push(this.el('mi', [word], word.length === 1 ? { mathvariant: 'normal' } : {}));
        else for (const c of word) out.push(this.el('mi', [c]));
      } else if (space) out.push(this.el('mspace', [], { width: `${0.25 * space.length}em` }));
      else out.push(this.op(other === '-' ? '\u2212' : other)); // Word shows a hyphen as a minus sign
    }
    return out;
  }

  fraction(f) {
    const type = prop(mChild(f, 'fPr'), 'type');
    const num = this.arg(mChild(f, 'num'));
    const den = this.arg(mChild(f, 'den'));
    if (type === 'lin' || type === 'skw') return this.el('mrow', [num, this.op('/'), den]);
    return this.el('mfrac', [num, den], type === 'noBar' ? { linethickness: '0' } : {});
  }

  radical(rad) {
    const deg = mChild(rad, 'deg');
    const e = this.arg(mChild(rad, 'e'));
    if (on(prop(mChild(rad, 'radPr'), 'degHide')) || !deg || !deg.textContent.trim()) return this.el('msqrt', [e]);
    return this.el('mroot', [e, this.arg(deg)]);
  }

  nary(n) {
    const pr = mChild(n, 'naryPr');
    const chr = prop(pr, 'chr') || '∫';
    const integral = INTEGRALS.test(chr);
    const limLoc = prop(pr, 'limLoc') || (integral ? 'subSup' : 'undOvr');
    const sub = on(prop(pr, 'subHide')) ? null : this.arg(mChild(n, 'sub'));
    const sup = on(prop(pr, 'supHide')) ? null : this.arg(mChild(n, 'sup'));
    const base = this.op(chr, { largeop: 'true', movablelimits: 'false' });
    const [both, under, over] = limLoc === 'undOvr' ? ['munderover', 'munder', 'mover'] : ['msubsup', 'msub', 'msup'];
    let lims = base;
    if (sub && sup) lims = this.el(both, [base, sub, sup]);
    else if (sub) lims = this.el(under, [base, sub]);
    else if (sup) lims = this.el(over, [base, sup]);
    return [this.el('mrow', [lims, this.arg(mChild(n, 'e'))])];
  }

  delimiter(d) {
    const pr = mChild(d, 'dPr');
    const beg = prop(pr, 'begChr') ?? '(';
    const end = prop(pr, 'endChr') ?? ')';
    const sep = prop(pr, 'sepChr') ?? '|';
    const fence = (c) => this.op(c, { fence: 'true', stretchy: 'true' });
    const items = [];
    if (beg) items.push(fence(beg));
    mChildren(d, 'e').forEach((e, i) => {
      if (i && sep) items.push(this.op(sep, { separator: 'true' }));
      items.push(this.arg(e));
    });
    if (end) items.push(fence(end));
    return this.el('mrow', items);
  }

  accent(acc) {
    const chr = prop(mChild(acc, 'accPr'), 'chr') || '̂';
    return this.el('mover', [this.arg(mChild(acc, 'e')), this.op(SPACING[chr] || chr, { stretchy: 'true' })], { accent: 'true' });
  }

  bar(bar) {
    const top = prop(mChild(bar, 'barPr'), 'pos') === 'top';
    const line = this.op(top ? '‾' : '_', { stretchy: 'true' });
    return this.el(top ? 'mover' : 'munder', [this.arg(mChild(bar, 'e')), line], top ? { accent: 'true' } : { accentunder: 'true' });
  }

  groupChr(g) {
    const pr = mChild(g, 'groupChrPr');
    const chr = prop(pr, 'chr') || '⏟';
    const top = prop(pr, 'pos') === 'top';
    return this.el(top ? 'mover' : 'munder', [this.arg(mChild(g, 'e')), this.op(chr, { stretchy: 'true' })]);
  }

  /**
   * An equation array: a row for each line, lined up at its alignment points (the &s Word puts in
   * the text) as Word lines them up: the parts between them in columns, right-aligned then
   * left-aligned in turn (x | = 2 + 3 / | = 5), as LaTeX's aligned does. Without any, centred.
   */
  equationArray(arr) {
    this.inArray++;
    try {
      const isMark = (n) => n.localName === 'mspace' && n.hasAttribute('data-align');
      const cells = mChildren(arr, 'e').map((e) => {
        const parts = [[]];
        for (const n of this.all(e)) {
          if (isMark(n)) parts.push([]);
          else parts[parts.length - 1].push(n);
        }
        return parts;
      });
      const cols = Math.max(1, ...cells.map((c) => c.length));
      const rows = cells.map((parts) => this.el('mtr', Array.from({ length: cols }, (_, i) => this.el('mtd', [this.row(parts[i] || [])]))));
      const align = cols > 1 ? Array.from({ length: cols }, (_, i) => (i % 2 ? 'left' : 'right')).join(' ') : null;
      const table = this.el('mtable', rows, { displaystyle: 'true', columnalign: align, columnspacing: cols > 1 ? '0' : null });
      // MathML Core lays out columns but not columnalign: each cell aligned itself, too.
      if (align) for (const tr of rows) [...tr.children].forEach((td, i) => { td.style.textAlign = i % 2 ? 'left' : 'right'; td.style.paddingInline = '0'; });
      return table;
    } finally {
      this.inArray--;
    }
  }

  matrix(m) {
    const rows = mChildren(m, 'mr').map((mr) => this.el('mtr', mChildren(mr, 'e').map((e) => this.el('mtd', [this.arg(e)]))));
    return this.el('mtable', rows);
  }

  phantom(p) {
    const e = this.arg(mChild(p, 'e'));
    return on(prop(mChild(p, 'phantPr'), 'show') ?? '1') ? e : this.el('mphantom', [e]);
  }
}

/**
 * MathML for an equation: an m:oMath (in line with the text) or an m:oMathPara (on its own line,
 * one <math display="block"> per equation in it). Returns the <math> elements.
 */
export function ommlToMathml(node, doc = document) {
  const conv = new Converter(doc);
  const block = node.localName === 'oMathPara';
  const maths = block ? mChildren(node, 'oMath') : node.localName === 'oMath' ? [node] : [...node.getElementsByTagNameNS(MATH_NS, 'oMath')];
  return maths.map((m) => {
    const math = conv.el('math', [conv.row(conv.all(m))], block ? { display: 'block' } : {});
    return math;
  });
}

/** True if `node` is (or holds) an equation. */
export function isEquation(node) {
  return (node.namespaceURI === MATH_NS && /^oMath(Para)?$/.test(node.localName))
    || node.getElementsByTagNameNS?.(MATH_NS, 'oMath').length > 0;
}
