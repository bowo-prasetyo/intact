// Typing an equation, as Word's equation editor takes LaTeX: a small part of LaTeX's maths, written
// as Word's Office Math (OMML), which Word then edits as its own. Fractions (\frac{a}{b}),
// superscripts and subscripts (x^2, x_{i}), roots (\sqrt{x}, \sqrt[3]{x}), sums, products and
// integrals with limits (\sum_{i=1}^{n}), brackets that grow (\left( … \right)), Greek letters and
// the usual symbols (\alpha, \pi, \times, \le, \infty …), and functions (\sin, \log …) upright.
// Matrices (\begin{pmatrix} a & b \\ c & d \end{pmatrix}, and matrix, bmatrix, Bmatrix, vmatrix,
// Vmatrix), cases, and working lined up at its = signs (\begin{aligned} x &= 1 \\ &= 2 \end{aligned},
// as Word's equation arrays, its alignment points the &s in the text); accents (\hat, \bar, \vec,
// \dot, \ddot, \tilde) and lines over and under (\overline, \underline).

export const SYMBOLS = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ',
  vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ',
  tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  times: '×', cdot: '⋅', div: '÷', pm: '±', mp: '∓', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠',
  approx: '≈', equiv: '≡', sim: '∼', propto: '∝', infty: '∞', partial: '∂', nabla: '∇', to: '→', rightarrow: '→',
  leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐', leftrightarrow: '↔', iff: '⇔', in: '∈', notin: '∉', subset: '⊂',
  subseteq: '⊆', cup: '∪', cap: '∩', emptyset: '∅', forall: '∀', exists: '∃', neg: '¬', land: '∧', lor: '∨',
  degree: '°', circ: '∘', ldots: '…', cdots: '⋯', angle: '∠', perp: '⊥', parallel: '∥', prime: '′',
  ',': ' ', ';': ' ', quad: ' ', qquad: '  ', '{': '{', '}': '}', '%': '%', $: '$', '#': '#', '&': '&', _: '_',
};
// Large operators, with where their limits go (under and over, or beside), as Word places them.
export const NARY = { sum: ['∑', 'undOvr'], prod: ['∏', 'undOvr'], coprod: ['∐', 'undOvr'], bigcup: ['⋃', 'undOvr'], bigcap: ['⋂', 'undOvr'], int: ['∫', 'subSup'], iint: ['∬', 'subSup'], iiint: ['∭', 'subSup'], oint: ['∮', 'subSup'] };
export const FUNCTIONS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh', 'log', 'ln', 'exp', 'lim', 'max', 'min', 'det', 'gcd']);
// Accents, as Word's (m:acc with its character; a hat is Word's default).
export const ACCENTS = { hat: '\u0302', widehat: '\u0302', bar: '\u0305', vec: '\u20d7', dot: '\u0307', ddot: '\u0308', tilde: '\u0303', widetilde: '\u0303', check: '\u030c', acute: '\u0301', grave: '\u0300', breve: '\u0306' };
// Matrices and the brackets around them (none for matrix), as Word puts a matrix in a delimiter.
export const MATRICES = { matrix: null, pmatrix: ['(', ')'], bmatrix: ['[', ']'], Bmatrix: ['{', '}'], vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'] };
// Lines of working, each its own row of an equation array, lined up at their &s.
const ARRAYS = new Set(['aligned', 'align', 'align*', 'gathered', 'gather', 'gather*', 'split', 'eqnarray', 'eqnarray*']);
const DELIMS = { '(': '(', ')': ')', '[': '[', ']': ']', '|': '|', '.': '', '\\{': '{', '\\}': '}', '\\langle': '⟨', '\\rangle': '⟩', '\\lvert': '|', '\\rvert': '|' };

class Parser {
  constructor(src) {
    this.src = src;
    this.i = 0;
  }

  error(msg) {
    throw new Error(msg);
  }

  peek() {
    while (this.i < this.src.length && /\s/.test(this.src[this.i])) this.i++;
    return this.src[this.i];
  }

  command() {
    // After a backslash: a name of letters, or one other character.
    const m = /^[A-Za-z]+|^./.exec(this.src.slice(this.i));
    if (!m) this.error('A \\ needs a name after it, such as \\frac or \\alpha');
    this.i += m[0].length;
    return m[0];
  }

  /** Nodes up to the end, or to `stop` (a closing brace, \right or ]). */
  list(stop = null) {
    const out = [];
    for (;;) {
      const c = this.peek();
      if (c === undefined) {
        if (stop === 'cell') this.error(`A \\begin{${this.env}} has no \\end{${this.env}} to close it`);
        if (stop === '}') this.error('A { has no } to close it');
        if (stop === 'right') this.error('A \\left has no \\right to close it');
        if (stop === ']') this.error('A [ has no ] to close it');
        return out;
      }
      if (stop === '}' && c === '}') { this.i++; return out; }
      if (stop === ']' && c === ']') { this.i++; return out; }
      if (stop === 'right' && this.src.startsWith('\\right', this.i)) { this.i += 6; return out; }
      if (stop === 'cell') {
        // A cell of a matrix or a line of working ends at &, at \\ (a new row) or at \end.
        if (c === '&') { this.i++; this.ended = '&'; return out; }
        if (this.src.startsWith('\\\\', this.i)) { this.i += 2; this.ended = 'row'; return out; }
        if (this.src.startsWith('\\end', this.i)) { this.ended = 'end'; return out; }
      }
      if (c === '}') this.error('A } has no { before it');
      if (c === '^' || c === '_') {
        this.i++;
        const base = out.pop() || { t: 'run', text: '' };
        // "ab^2": the script goes on the last letter only.
        let b = base;
        if (base.t === 'run' && [...base.text].length > 1 && !base.fn) {
          const chars = [...base.text];
          out.push({ t: 'run', text: chars.slice(0, -1).join('') });
          b = { t: 'run', text: chars[chars.length - 1] };
        }
        out.push(this.script(b, c));
        continue;
      }
      const node = this.atom();
      const last = out[out.length - 1];
      if (node.t === 'run' && !node.fn && !node.op && last?.t === 'run' && !last.fn && !last.op) last.text += node.text;
      else out.push(node);
    }
  }

  /** `base` with the script starting at `kind` (^ or _), and the other one if it follows. */
  script(base, kind) {
    let sub = null;
    let sup = null;
    if (kind === '^') sup = this.argument(); else sub = this.argument();
    const next = this.peek();
    if ((next === '^' && !sup) || (next === '_' && !sub)) {
      this.i++;
      if (next === '^') sup = this.argument(); else sub = this.argument();
    }
    if (base.t === 'nary') return { ...base, sub: sub ?? base.sub, sup: sup ?? base.sup };
    if (sub && sup) return { t: 'subsup', e: [base], sub, sup };
    return sup ? { t: 'sup', e: [base], sup } : { t: 'sub', e: [base], sub };
  }

  /** A braced group, or one character or command. */
  argument() {
    const c = this.peek();
    if (c === undefined) this.error('Something should follow ^, _, \\frac or \\sqrt');
    if (c === '{') {
      this.i++;
      return this.list('}');
    }
    return [this.atom()];
  }

  atom() {
    const c = this.peek();
    if (c === '{') {
      this.i++;
      return { t: 'group', e: this.list('}') };
    }
    if (c !== '\\') {
      this.i++;
      // Digits and letters run together; an operator on its own.
      return /[+\-=<>±×÷*/,]/.test(c) ? { t: 'run', text: c === '-' ? '−' : c === '*' ? '∗' : c, op: true } : { t: 'run', text: c };
    }
    this.i++;
    const name = this.command();
    if (name === 'frac' || name === 'dfrac' || name === 'tfrac') return { t: 'frac', num: this.argument(), den: this.argument() };
    if (name === 'sqrt') {
      let deg = null;
      if (this.peek() === '[') {
        this.i++;
        deg = this.list(']');
      }
      return { t: 'rad', deg, e: this.argument() };
    }
    if (name === 'left') {
      const open = this.delimiter();
      const e = this.list('right');
      const close = this.delimiter();
      return { t: 'delim', open, close, e };
    }
    if (name === 'right') this.error('A \\right has no \\left before it');
    if (name === 'begin') return this.environment();
    if (name === 'end') this.error('An \\end has no \\begin before it');
    if (ACCENTS[name]) return { t: 'acc', chr: ACCENTS[name], e: this.argument() };
    if (name === 'overline' || name === 'underline') return { t: 'bar', pos: name === 'overline' ? 'top' : 'bot', e: this.argument() };
    if (NARY[name]) {
      const [chr, limLoc] = NARY[name];
      const node = { t: 'nary', chr, limLoc, sub: null, sup: null, e: [] };
      // Its limits, then what it applies to: the next term.
      let n = node;
      while (this.peek() === '_' || this.peek() === '^') {
        const k = this.src[this.i++];
        n = { ...n, [k === '_' ? 'sub' : 'sup']: this.argument() };
      }
      if (this.peek() !== undefined && this.peek() !== '}' && !this.src.startsWith('\\right', this.i)) {
        const term = this.atom();
        let e = [term];
        if (this.peek() === '^' || this.peek() === '_') {
          const k = this.src[this.i++];
          e = [this.script(term, k)];
        }
        n.e = e;
      }
      return n;
    }
    if (FUNCTIONS.has(name)) return { t: 'run', text: name, fn: true };
    if (name === 'text' || name === 'mathrm' || name === 'operatorname') {
      // Its text as typed, spaces too (elsewhere spaces are LaTeX's, and left out).
      if (this.peek() === '{') {
        const end = this.src.indexOf('}', this.i);
        if (end < 0) this.error('A { has no } to close it');
        const text = this.src.slice(this.i + 1, end);
        this.i = end + 1;
        return { t: 'run', text, fn: true };
      }
      return { t: 'run', text: this.argument().map(textOf).join(''), fn: true };
    }
    if (SYMBOLS[name] !== undefined) return { t: 'run', text: SYMBOLS[name], op: /[^\p{L}\p{N}]/u.test(SYMBOLS[name]) };
    this.error(`\\${name} isn’t one this editor knows: try \\frac, \\sqrt, \\sum, \\alpha or ^ and _`);
    return null;
  }

  /** The name in braces after \begin or \end. */
  envName() {
    if (this.peek() !== '{') this.error('\\begin and \\end need a name in braces, such as {pmatrix}');
    const end = this.src.indexOf('}', this.i);
    if (end < 0) this.error('A { has no } to close it');
    const name = this.src.slice(this.i + 1, end).trim();
    this.i = end + 1;
    return name;
  }

  /**
   * \begin{name} … \end{name}: its rows (split at \\) of cells (split at &), as a matrix in its
   * brackets, cases (a brace on the left, the conditions lined up), or lines of working.
   */
  environment() {
    const name = this.envName();
    const known = name in MATRICES || name === 'cases' || ARRAYS.has(name);
    if (!known) this.error(`{${name}} isn’t one this editor knows: try pmatrix, bmatrix, cases or aligned`);
    const outer = this.env;
    this.env = name;
    const rows = [[]];
    for (;;) {
      const cell = this.list('cell');
      rows[rows.length - 1].push(cell);
      if (this.ended === 'row') rows.push([]);
      if (this.ended === 'end') break;
    }
    this.i += 4; // \end
    const closing = this.envName();
    if (closing !== name) this.error(`\\begin{${name}} is closed by \\end{${closing}}`);
    this.env = outer;
    // A last \\ before \end makes no row.
    const last = rows[rows.length - 1];
    if (rows.length > 1 && last.length === 1 && !last[0].length) rows.pop();
    if (name in MATRICES) {
      const m = { t: 'matrix', rows };
      const brackets = MATRICES[name];
      return brackets ? { t: 'delim', open: brackets[0], close: brackets[1], e: [m] } : m;
    }
    const array = { t: 'eqArr', rows };
    return name === 'cases' ? { t: 'delim', open: '{', close: '', e: [array] } : array;
  }

  delimiter() {
    const c = this.peek();
    if (c === undefined) this.error('\\left and \\right need a bracket after them, such as ( or [');
    if (c === '\\') {
      const start = this.i;
      this.i++;
      const name = `\\${this.command()}`;
      if (DELIMS[name] === undefined) {
        this.i = start;
        this.error(`${name} isn’t a bracket \\left and \\right take`);
      }
      return DELIMS[name];
    }
    this.i++;
    if (DELIMS[c] === undefined) this.error(`${c} isn’t a bracket \\left and \\right take`);
    return DELIMS[c];
  }
}

const textOf = (n) => (n.t === 'run' ? n.text : n.e ? n.e.map(textOf).join('') : '');
const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const FONT = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';

function omml(nodes) {
  return nodes.map((n) => {
    switch (n.t) {
      case 'run':
        if (!n.text) return '';
        return `<m:r>${n.fn ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}${FONT}<m:t xml:space="preserve">${esc(n.text)}</m:t></m:r>`;
      case 'group': return omml(n.e);
      case 'frac': return `<m:f><m:num>${omml(n.num)}</m:num><m:den>${omml(n.den)}</m:den></m:f>`;
      case 'sup': return `<m:sSup><m:e>${omml(n.e)}</m:e><m:sup>${omml(n.sup)}</m:sup></m:sSup>`;
      case 'sub': return `<m:sSub><m:e>${omml(n.e)}</m:e><m:sub>${omml(n.sub)}</m:sub></m:sSub>`;
      case 'subsup': return `<m:sSubSup><m:e>${omml(n.e)}</m:e><m:sub>${omml(n.sub)}</m:sub><m:sup>${omml(n.sup)}</m:sup></m:sSubSup>`;
      case 'rad': return n.deg
        ? `<m:rad><m:deg>${omml(n.deg)}</m:deg><m:e>${omml(n.e)}</m:e></m:rad>`
        : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${omml(n.e)}</m:e></m:rad>`;
      case 'nary': return `<m:nary><m:naryPr><m:chr m:val="${n.chr}"/><m:limLoc m:val="${n.limLoc}"/>${n.sub ? '' : '<m:subHide m:val="1"/>'}${n.sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr>`
        + `<m:sub>${omml(n.sub || [])}</m:sub><m:sup>${omml(n.sup || [])}</m:sup><m:e>${omml(n.e)}</m:e></m:nary>`;
      case 'acc': return `<m:acc>${n.chr === '\u0302' ? '' : `<m:accPr><m:chr m:val="${n.chr}"/></m:accPr>`}<m:e>${omml(n.e)}</m:e></m:acc>`;
      case 'bar': return `<m:bar><m:barPr><m:pos m:val="${n.pos}"/></m:barPr><m:e>${omml(n.e)}</m:e></m:bar>`;
      case 'matrix': {
        // Every column centred, as Word's matrices are; as many columns as the longest row.
        const cols = Math.max(...n.rows.map((r) => r.length));
        const rows = n.rows.map((r) => `<m:mr>${Array.from({ length: cols }, (_, i) => `<m:e>${omml(r[i] || [])}</m:e>`).join('')}</m:mr>`).join('');
        return `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${cols}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>${rows}</m:m>`;
      }
      case 'eqArr':
        // Each line a row; where a & was typed, Word's alignment point: an & in the text, in a run of
        // its own, which Word lines up from row to row (and doesn't show).
        return `<m:eqArr>${n.rows.map((r) => `<m:e>${r.map((cell, i) => (i ? `<m:r>${FONT}<m:t>&amp;</m:t></m:r>` : '') + omml(cell)).join('')}</m:e>`).join('')}</m:eqArr>`;
      case 'delim': {
        const pr = n.open === '(' && n.close === ')' ? '' : `<m:dPr><m:begChr m:val="${esc(n.open)}"/><m:endChr m:val="${esc(n.close)}"/></m:dPr>`;
        return `<m:d>${pr}<m:e>${omml(n.e)}</m:e></m:d>`;
      }
      default: return '';
    }
  }).join('');
}

/**
 * The equation typed as `src` (LaTeX), as OMML: { xml } (an m:oMath, or with `display` an
 * m:oMathPara holding one, for an equation on its own line), or { error } saying what is wrong.
 */
export function latexToOmml(src, { display = false } = {}) {
  const text = (src || '').trim().replace(/^\$+|\$+$/g, '').trim();
  if (!text) return { error: 'Type an equation, such as x^2 + \\frac{1}{2}' };
  try {
    const body = omml(new Parser(text).list());
    const math = `<m:oMath>${body}</m:oMath>`;
    return { xml: display ? `<m:oMathPara>${math}</m:oMathPara>` : math };
  } catch (e) {
    return { error: e.message };
  }
}
