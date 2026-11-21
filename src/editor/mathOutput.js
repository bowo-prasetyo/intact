// An equation already in a document (Word's Office Math) as LaTeX, the part of it the Equation
// dialog takes (see mathInput.js), to edit it there: runs, fractions, scripts, roots, sums and
// integrals with their limits, brackets, functions, limits, accents, lines over and under, matrices,
// cases and lines of working. An equation with anything else in it (letter styles such as bold or
// script, boxes, phantoms, prescripts, tracked changes) gives null: it stays as Word wrote it,
// rather than come back without what LaTeX here can't say.
import { SYMBOLS, NARY, FUNCTIONS, ACCENTS, MATRICES } from './mathInput.js';
import { MATH_NS } from '../docx/omml.js';

const kids = (el) => [...(el?.children || [])];
const mChild = (el, local) => kids(el).find((c) => c.namespaceURI === MATH_NS && c.localName === local) || null;
const mChildren = (el, local) => kids(el).filter((c) => c.namespaceURI === MATH_NS && c.localName === local);
function prop(pr, local) {
  const el = mChild(pr, local);
  if (!el) return undefined;
  return el.getAttributeNS(MATH_NS, 'val') ?? el.getAttribute('m:val') ?? '';
}
const on = (v) => v !== undefined && !['0', 'false', 'off'].includes(v);

class Unsupported extends Error {}
const no = () => { throw new Unsupported(); };

// Characters as the commands that make them (the first name for each), and those LaTeX escapes.
const NAMES = new Map();
for (const [name, ch] of Object.entries(SYMBOLS)) if (/^[A-Za-z]+$/.test(name) && [...ch].length === 1 && ch.trim() && !NAMES.has(ch)) NAMES.set(ch, `\\${name}`);
const ESCAPED = { '{': '\\{', '}': '\\}', '%': '\\%', $: '\\$', '#': '\\#', _: '\\_', '&': '\\&', '−': '-', '∗': '*' };
const NARY_NAMES = new Map(Object.entries(NARY).map(([name, [chr]]) => [chr, name]));
const ACCENT_NAMES = new Map();
for (const [name, chr] of Object.entries(ACCENTS)) if (!ACCENT_NAMES.has(chr)) ACCENT_NAMES.set(chr, name);
// Spacing forms Word also writes for accents.
for (const [chr, name] of [['^', 'hat'], ['ˆ', 'hat'], ['~', 'tilde'], ['˜', 'tilde'], ['¯', 'bar'], ['‾', 'bar'], ['→', 'vec'], ['˙', 'dot'], ['¨', 'ddot'], ['ˇ', 'check']]) ACCENT_NAMES.set(chr, name);
const BRACKETS = { '(': '(', ')': ')', '[': '[', ']': ']', '|': '|', '': '.', '{': '\\{', '}': '\\}', '⟨': '\\langle', '⟩': '\\rangle', '〈': '\\langle', '〉': '\\rangle' };
const MATRIX_ENVS = new Map(Object.entries(MATRICES).filter(([, b]) => b).map(([name, [open, close]]) => [`${open}${close}`, name]));

/** Pieces of LaTeX joined, with a space after a command a letter would otherwise run on from. */
function join(parts) {
  let out = '';
  for (const p of parts) {
    if (!p) continue;
    if (/\\[A-Za-z]+$/.test(out) && /^[A-Za-z]/.test(p)) out += ' ';
    out += p;
  }
  return out;
}

/** True if `tex` is one thing a script or a command takes without braces: a character or a command. */
const single = (tex) => [...tex].length === 1 || /^\\[A-Za-z]+$/.test(tex);
const braced = (tex) => `{${tex}}`;
const base = (tex) => (single(tex) ? tex : braced(tex));

/** A run's text as LaTeX: its letters and symbols, \mathrm{…} upright, a function by its name. */
function run(r, inArray) {
  const rPr = mChild(r, 'rPr');
  const sty = prop(rPr, 'sty');
  if (prop(rPr, 'scr') !== undefined && prop(rPr, 'scr') !== 'roman') no();
  if (sty === 'b' || sty === 'bi') no();
  for (const c of kids(r)) {
    if (c.namespaceURI === MATH_NS && !['rPr', 't'].includes(c.localName)) no();
    if (c.namespaceURI !== MATH_NS && c.localName !== 'rPr') no(); // w:br, w:sym… inside the maths
  }
  const text = mChildren(r, 't').map((t) => t.textContent).join('');
  if (!text) return '';
  if (sty === 'p' || on(prop(rPr, 'nor'))) {
    if (FUNCTIONS.has(text)) return `\\${text}`;
    if (/[{}]/.test(text)) no();
    return `\\mathrm{${text}}`;
  }
  return join([...text].map((ch) => {
    if (ch === '&' && inArray) return '&';
    if (ESCAPED[ch]) return ESCAPED[ch];
    if (ch === ' ' || ch === ' ' || ch === ' ') return '\\,';
    if (ch === '\\' || ch === '^') no();
    return NAMES.get(ch) || ch;
  }));
}

/** The LaTeX of the m: elements in `el` (an argument: m:e, m:num, m:sub…), in order. */
function content(el, inArray = false) {
  if (!el) return '';
  return join(kids(el).map((c) => node(c, inArray)));
}

function node(el, inArray) {
  if (el.namespaceURI !== MATH_NS) {
    // Word's run properties beside the maths (w:rPr in m:ctrlPr) aren't content; anything else is.
    if (el.localName === 'rPr' || el.localName === 'bookmarkStart' || el.localName === 'bookmarkEnd') return '';
    return no();
  }
  switch (el.localName) {
    case 'r': return run(el, inArray);
    case 'ctrlPr': case 'argPr': case 'fPr': case 'radPr': case 'naryPr': case 'dPr': case 'accPr': case 'barPr':
    case 'sSupPr': case 'sSubPr': case 'sSubSupPr': case 'funcPr': case 'limLowPr': case 'limUppPr': case 'mPr': case 'eqArrPr': case 'boxPr':
      return '';
    case 'f':
      if (prop(mChild(el, 'fPr'), 'type') === 'noBar') no();
      return `\\frac${braced(content(mChild(el, 'num')))}${braced(content(mChild(el, 'den')))}`;
    case 'sSup': return `${base(content(mChild(el, 'e')))}^${braced(content(mChild(el, 'sup')))}`;
    case 'sSub': return `${base(content(mChild(el, 'e')))}_${braced(content(mChild(el, 'sub')))}`;
    case 'sSubSup': return `${base(content(mChild(el, 'e')))}_${braced(content(mChild(el, 'sub')))}^${braced(content(mChild(el, 'sup')))}`;
    case 'rad': {
      const deg = content(mChild(el, 'deg'));
      return `\\sqrt${on(prop(mChild(el, 'radPr'), 'degHide')) || !deg ? '' : `[${deg}]`}${braced(content(mChild(el, 'e')))}`;
    }
    case 'nary': {
      const pr = mChild(el, 'naryPr');
      const name = NARY_NAMES.get(prop(pr, 'chr') || '∫');
      if (!name) no();
      const sub = on(prop(pr, 'subHide')) ? '' : content(mChild(el, 'sub'));
      const sup = on(prop(pr, 'supHide')) ? '' : content(mChild(el, 'sup'));
      return `\\${name}${sub ? `_${braced(sub)}` : ''}${sup ? `^${braced(sup)}` : ''}${braced(content(mChild(el, 'e')))}`;
    }
    case 'd': return delimiter(el);
    case 'func': {
      const name = content(mChild(el, 'fName'));
      return join([name, braced(content(mChild(el, 'e')))]);
    }
    case 'limLow': return `${base(content(mChild(el, 'e')))}_${braced(content(mChild(el, 'lim')))}`;
    case 'limUpp': return `${base(content(mChild(el, 'e')))}^${braced(content(mChild(el, 'lim')))}`;
    case 'acc': {
      const name = ACCENT_NAMES.get(prop(mChild(el, 'accPr'), 'chr') || '̂');
      if (!name) no();
      return `\\${name}${braced(content(mChild(el, 'e')))}`;
    }
    case 'bar': return `\\${prop(mChild(el, 'barPr'), 'pos') === 'top' ? 'overline' : 'underline'}${braced(content(mChild(el, 'e')))}`;
    case 'box': return content(mChild(el, 'e'));
    case 'm': return environment('matrix', matrixRows(el));
    case 'eqArr': return environment('aligned', mChildren(el, 'e').map((e) => content(e, true)));
    default: return no();
  }
}

/** A matrix's rows, each its cells joined by &. */
const matrixRows = (m) => mChildren(m, 'mr').map((mr) => mChildren(mr, 'e').map((e) => content(e)).join(' & '));
const environment = (name, rows) => `\\begin{${name}} ${rows.join(' \\\\ ')} \\end{${name}}`;

/** Brackets round their content: a matrix in them as pmatrix, bmatrix…, a brace before lines as cases. */
function delimiter(el) {
  const pr = mChild(el, 'dPr');
  const open = prop(pr, 'begChr') ?? '(';
  const close = prop(pr, 'endChr') ?? ')';
  const parts = mChildren(el, 'e');
  if (parts.length !== 1) no();
  const inner = kids(parts[0]).filter((c) => !(c.namespaceURI === MATH_NS && c.localName === 'ctrlPr'));
  if (inner.length === 1 && inner[0].namespaceURI === MATH_NS) {
    const env = MATRIX_ENVS.get(`${open}${close}`);
    if (inner[0].localName === 'm' && env) return environment(env, matrixRows(inner[0]));
    if (inner[0].localName === 'eqArr' && open === '{' && close === '') {
      return environment('cases', mChildren(inner[0], 'e').map((e) => content(e, true)));
    }
  }
  if (BRACKETS[open] === undefined || BRACKETS[close] === undefined) no();
  return `\\left${BRACKETS[open]} ${content(parts[0])} \\right${BRACKETS[close]}`;
}

/**
 * An equation (m:oMath, or m:oMathPara with one m:oMath in it) as { latex, display } (display: on
 * its own line), or null if it has something the Equation dialog can't write back.
 */
export function ommlToLatex(el) {
  try {
    let math = el;
    let display = false;
    if (el.localName === 'oMathPara') {
      const maths = mChildren(el, 'oMath');
      if (maths.length !== 1 || kids(el).some((c) => !['oMath', 'oMathParaPr'].includes(c.localName))) return null;
      [math] = maths;
      display = true;
    }
    if (math?.localName !== 'oMath') return null;
    const latex = content(math).trim();
    return latex ? { latex, display } : null;
  } catch (e) {
    if (e instanceof Unsupported) return null;
    throw e;
  }
}
